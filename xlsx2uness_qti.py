#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx2uness_qti.py
=================
Convertit une banque de questions au format Excel (export/template Wooclap :
colonnes `Type, Title, Correct, Choice, Choice, ...`) en archive ZIP au
format QTI 2.1 (import UNESS / Theia) : un fichier XML par question +
un fichier `pool.xml` qui centralise les références.

Format d'entrée attendu (feuille "Questions" du fichier .xlsx) :
    Type   | Title                 | Correct           | Choice 1 | Choice 2 | ...
    Poll   | Énoncé de la question | <réponse attendue> | prop. A | prop. B | ...

Colonne "Correct", trois écrits acceptés :
  - le TEXTE EXACT d'une des propositions (recommandé, format banque Wooclap) ;
  - le NUMÉRO de la bonne proposition (ex. "3" = 3e proposition) ;
  - la LETTRE de la bonne proposition (ex. "B").
  (vide ou introuvable -> la 1re proposition est retenue + avertissement)

Utilisation :
    python xlsx2uness_qti.py mesquestions.xlsx
    python xlsx2uness_qti.py mesquestions.xlsx -o mon_export.zip
    python xlsx2uness_qti.py mesquestions.xlsx --titre "OCE - Dossier 2026"
    python xlsx2uness_qti.py mesquestions.xlsx --feuille Questions

Aucune dépendance : bibliothèque standard Python 3 uniquement
(zipfile, xml, argparse). Fonctionne sur Windows, macOS et Linux.

Le ZIP produit s'importe tel quel dans UNESS via l'import QTI.
Testé avec les banques : HAS2026 (16 QCM, OCE endodontie/antibiothérapie).
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

# --------------------------------------------------------------------------
# Lecture d'un fichier .xlsx (bibliothèque standard uniquement)
# --------------------------------------------------------------------------

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_DOCREL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_letters_to_index(ref):
    """'B12' -> 1 (index de colonne, 0-based)."""
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return 0
    idx = 0
    for ch in m.group(1):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _text_of_si(si):
    """Texte complet d'un <si> de sharedStrings (gère le rich-text)."""
    parts = []
    for t in si.iter(NS_MAIN + "t"):
        parts.append(t.text or "")
    return "".join(parts)


def read_xlsx_rows(path):
    """Retourne une liste de feuilles : [(nom, lignes), ...]
    où lignes = [[valeur_deinit colonne, ...], ...] par ligne."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        # chaînes partagées (fichiers enregistrés par Excel, avec <t="sharedStrings">)
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.iter(NS_MAIN + "si"):
                shared.append(_text_of_si(si))

        # relations workbook -> feuilles
        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            for rel in root.iter(NS_PKGREL + "Relationship"):
                rels[rel.get("Id")] = (rel.get("Target") or "").lstrip("/")

        # ordre et noms des feuilles
        sheet_files = []
        if "xl/workbook.xml" in names:
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            for sh in wb.iter(NS_MAIN + "sheet"):
                rid = sh.get(NS_DOCREL + "id")
                target = rels.get(rid, "")
                if target and not target.startswith("xl/"):
                    target = "xl/" + target
                if target and target in names:
                    sheet_files.append((sh.get("name") or "Sheet", target))
        if not sheet_files and "xl/worksheets/sheet1.xml" in names:
            sheet_files = [("Questions", "xl/worksheets/sheet1.xml")]

        sheets = []
        for sname, sfile in sheet_files:
            root = ET.fromstring(z.read(sfile))
            rows = []
            for row in root.iter(NS_MAIN + "row"):
                cells = {}
                for c in row.iter(NS_MAIN + "c"):
                    col = _col_letters_to_index(c.get("r"))
                    ctype = c.get("t")
                    value = ""
                    if ctype == "inlineStr":
                        is_el = c.find(NS_MAIN + "is")
                        if is_el is not None:
                            value = "".join(
                                (t.text or "") for t in is_el.iter(NS_MAIN + "t")
                            )
                    elif ctype == "s":
                        v = c.find(NS_MAIN + "v")
                        if v is not None and v.text is not None:
                            try:
                                value = shared[int(v.text)]
                            except (ValueError, IndexError):
                                value = v.text
                    else:
                        v = c.find(NS_MAIN + "v")
                        if v is not None:
                            value = v.text or ""
                    cells[col] = value
                if cells:
                    width = max(cells) + 1
                    rows.append([cells.get(i, "") for i in range(width)])
            sheets.append((sname, rows))
        return sheets


# --------------------------------------------------------------------------
# Interprétation des lignes en questions
# --------------------------------------------------------------------------

def normalize(s):
    return re.sub(r"\s+", " ", (s or "").replace("\r", " ").replace("\n", " ")).strip()


def rows_to_questions(rows):
    """Transforme les lignes brutes de la feuille en questions :
    [{title, choices: [...], correct: int, warn: bool}, ...]"""
    if not rows:
        return []
    header = [normalize(h).lower() for h in rows[0]]

    title_col = None
    correct_col = None
    choice_cols = []
    for i, h in enumerate(header):
        if h in ("title", "question", "titre", "énoncé", "enonce", "sujet"):
            title_col = i
        elif h in ("correct", "correct answer", "réponse", "reponse",
                   "bonne réponse", "bonne reponse"):
            correct_col = i
        elif h.startswith("choice") or h.startswith("choix") or h.startswith("option"):
            choice_cols.append(i)

    if title_col is None:
        # pas d'en-tête reconnu : on tente la disposition Wooclap standard
        title_col = 1
        correct_col = 2
        choice_cols = [3, 4, 5, 6, 7]
    if correct_col is None:
        correct_col = 2 if title_col != 2 else 1
    if not choice_cols:
        choice_cols = [i for i in range(len(header))
                       if i not in (title_col, correct_col) and i > min(title_col, correct_col)]
        choice_cols = [i for i in choice_cols if i > max(title_col, correct_col) - 1]

    questions = []
    for line in rows[1:]:
        line = [normalize(c) for c in line]
        need = max([title_col, correct_col] + choice_cols) + 1
        if len(line) < need:
            line = line + [""] * (need - len(line))
        title = line[title_col]
        choices = [line[i] for i in choice_cols if i < len(line) and line[i]]
        if not title or len(choices) < 2:
            continue  # ligne vide / incomplète : ignorée
        correct_raw = line[correct_col] if correct_col < len(line) else ""
        correct_index = _find_correct(correct_raw, choices)
        questions.append({
            "title": title,
            "choices": choices,
            "correct": correct_index,
            "warn": correct_raw == "",
        })
    return questions


def _find_correct(raw, choices):
    """Déduit l'index de la bonne réponse depuis la colonne Correct."""
    raw = normalize(raw)
    if not raw:
        return 0
    for i, ch in enumerate(choices):            # 1) texte exact d'une proposition
        if normalize(ch) == raw:
            return i
    if re.fullmatch(r"\d+", raw):               # 2) numéro (1-based)
        idx = int(raw) - 1
        if 0 <= idx < len(choices):
            return idx
    if re.fullmatch(r"[A-Za-z]", raw):         # 3) lettre (A..Z)
        idx = ord(raw.upper()) - 65
        if 0 <= idx < len(choices):
            return idx
    for i, ch in enumerate(choices):           # 4) texte contenu dans une proposition
        if raw.lower() in ch.lower() and len(raw) >= 5:
            return i
    return 0


# --------------------------------------------------------------------------
# Génération du format QTI 2.1 (modèles à l'identique d'un export Theia/UNESS)
# --------------------------------------------------------------------------

NS_ATTR = ('xmlns="http://www.imsglobal.org/xsd/imsqti_v2p1" '
           'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"')
SCHEMA = ('xsi:schemaLocation="http://www.imsglobal.org/xsd/imsqti_v2p1 '
          'http://www.imsglobal.org/xsd/qti/qtiv2p1/imsqti_v2p1.xsd"')


def question_xml(number, title, choices, correct_index):
    qid = "SINGLECHOICE_QUESTION_9000%05d" % number
    cids = ["CHOICE_159%06d" % (number * 100 + j) for j in range(len(choices))]
    x = []
    x.append('<?xml version="1.0" encoding="UTF-8"?>')
    x.append('<assessmentItem %s identifier="%s" title="Q%d" timeDependent="false" '
             'adaptive="false" toolName="Theia" %s>'
             % (NS_ATTR, qid, number, SCHEMA))
    x.append('  <responseDeclaration identifier="RESPONSE" cardinality="single" '
             'baseType="identifier">')
    x.append('    <correctResponse>')
    x.append('      <value>%s</value>' % cids[correct_index])
    x.append('    </correctResponse>')
    x.append('  </responseDeclaration>')
    x.append('  <outcomeDeclaration identifier="SCORE" cardinality="single" '
             'baseType="integer">')
    x.append('    <defaultValue>')
    x.append('      <value>0</value>')
    x.append('    </defaultValue>')
    x.append('  </outcomeDeclaration>')
    x.append('  <itemBody>')
    x.append('    <choiceInteraction responseIdentifier="%s" maxChoices="1">' % qid)
    x.append('      <prompt>')
    x.append('        <div>%s</div>' % escape("<p>%s</p>" % title))
    x.append('      </prompt>')
    for j, ch in enumerate(choices):
        x.append('      <simpleChoice identifier="%s">' % cids[j])
        x.append('        <div>%s</div>' % escape("<p>%s</p>" % ch))
        x.append('      </simpleChoice>')
    x.append('    </choiceInteraction>')
    x.append('  </itemBody>')
    x.append('</assessmentItem>')
    return qid, "\n".join(x) + "\n"


def pool_xml(qids, title):
    x = []
    x.append('<?xml version="1.0" encoding="UTF-8"?>')
    x.append('<testPart %s identifier="DLPOOL_IMPORT" navigationMode="nonlinear" '
             'submissionMode="individual" %s>' % (NS_ATTR, SCHEMA))
    x.append('  <assessmentSection identifier="DLPOOL_SECTION_IMPORT" required="false" '
             'fixed="false" title="%s" visible="true" keepTogether="true">'
             % escape(title))
    for qid in qids:
        x.append('    <assessmentItemRef identifier="%s" required="true" '
                 'fixed="false" href="%s.xml">' % (qid, qid))
        x.append('      <weight identifier="WEIGHT" value="1"/>')
        x.append('    </assessmentItemRef>')
    x.append('  </assessmentSection>')
    x.append('</testPart>')
    return "\n".join(x) + "\n"


def build_qti_zip(questions, out_path, pool_title="Banque de questions"):
    qids = []
    files = []
    for i, q in enumerate(questions, 1):
        qid, content = question_xml(i, q["title"], q["choices"], q["correct"])
        qids.append(qid)
        files.append((qid + ".xml", content))
    files.append(("pool.xml", pool_xml(qids, pool_title)))
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files:  # pool.xml en dernier
            z.writestr(name, content)
    return len(qids)


# --------------------------------------------------------------------------
# Point d'entrée
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Convertit un fichier Excel de QCM (format banque Wooclap : "
                    "Type/Title/Correct/Choice...) en archive ZIP QTI "
                    "importable dans UNESS.")
    ap.add_argument("xlsx", help="fichier .xlsx d'entrée")
    ap.add_argument("-o", "--output",
                    help="fichier .zip de sortie (défaut : <entrée>-UNESS-QTI.zip)")
    ap.add_argument("--titre", default="Banque de questions",
                    help="titre de la section (visible dans UNESS)")
    ap.add_argument("--feuille", default=None,
                    help="nom de la feuille à lire (défaut : feuille 'Questions', "
                         "sinon la 1re)")
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        sys.exit("Fichier introuvable : %s" % args.xlsx)

    sheets = read_xlsx_rows(args.xlsx)
    if not sheets:
        sys.exit("Aucune feuille exploitable dans le fichier.")

    sheet_name, sheet_rows = None, None
    if args.feuille:
        for name, rows in sheets:
            if name.lower() == args.feuille.lower():
                sheet_name, sheet_rows = name, rows
                break
        if sheet_rows is None:
            sys.exit("Feuille '%s' non trouvée (feuilles : %s)"
                     % (args.feuille, ", ".join(n for n, _ in sheets)))
    else:
        for name, rows in sheets:
            if name.lower().startswith("question"):
                sheet_name, sheet_rows = name, rows
                break
        if sheet_rows is None:
            sheet_name, sheet_rows = sheets[0]

    questions = rows_to_questions(sheet_rows)
    if not questions:
        sys.exit("Aucune question exploitable dans la feuille '%s'. "
                 "Vérifier : en-têtes en ligne 1 (Title, Correct, Choice...), "
                 "une question par ligne." % sheet_name)

    output = args.output
    if not output:
        base = os.path.splitext(os.path.basename(args.xlsx))[0]
        output = "%s-UNESS-QTI.zip" % base

    n = build_qti_zip(questions, output, args.titre)

    print("Feuille lue           : %s" % sheet_name)
    print("Questions converties   : %d" % n)
    warns = [i + 1 for i, q in enumerate(questions) if q.get("warn")]
    if warns:
        print("AVERTISSEMENT : colonne 'Correct' vide pour les questions n° %s "
              "(1re proposition retenue par défaut)" % warns)
    print("Archive QTI générée    : %s" % output)
    print("-> Importez ce ZIP dans UNESS via l'import QTI.")


if __name__ == "__main__":
    main()
