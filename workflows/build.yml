"""Lit liste.xlsx + notes/*.xlsx et produit data.json pour le site."""
import json, glob, os
from openpyxl import load_workbook

# Matières (codes de l'emploi du temps). Le nom du fichier Excel = le code : notes/M.xlsx, notes/PC.xlsx ...
SUBJECTS = {
    "M":    {"fr": "Mathématiques",       "ar": "الرياضيات"},
    "PC":   {"fr": "Physique-Chimie",     "ar": "الفيزياء والكيمياء"},
    "SN":   {"fr": "Sciences naturelles", "ar": "علوم الطبيعة"},
    "HG":   {"fr": "Histoire-Géographie", "ar": "التاريخ والجغرافيا"},
    "FR":   {"fr": "Français",            "ar": "الفرنسية"},
    "ANG":  {"fr": "Anglais",             "ar": "الإنجليزية"},
    "AR":   {"fr": "Arabe",               "ar": "العربية"},
    "IR":   {"fr": "Instruction religieuse", "ar": "التربية الإسلامية"},
    "PH":   {"fr": "Philosophie",         "ar": "الفلسفة"},
    "IC":   {"fr": "Instruction civique", "ar": "التربية المدنية"},
    "INFO": {"fr": "Informatique",        "ar": "المعلوماتية"},
    "EPS":  {"fr": "Éducation physique",  "ar": "التربية البدنية"},
    "LEG":  {"fr": "Législation",         "ar": "التشريع"},
}

# Coefficients par classe (tableau DREN du Trarza). Matière absente = non comptée dans la moyenne générale.
COEF = {
    "1AS": {"AR": 5, "FR": 4, "ANG": 1, "IR": 3, "IC": 1, "M": 5, "SN": 2, "HG": 2, "EPS": 1},
    "2AS": {"AR": 5, "FR": 4, "ANG": 1, "IR": 3, "IC": 1, "M": 5, "SN": 2, "HG": 2, "EPS": 1},
    "3AS": {"AR": 5, "FR": 3, "ANG": 1, "IR": 3, "IC": 1, "M": 5, "PC": 2, "SN": 2, "HG": 2, "EPS": 1},
    "4AS": {"AR": 5, "FR": 3, "ANG": 1, "IR": 3, "IC": 1, "M": 5, "PC": 2, "SN": 2, "HG": 2, "EPS": 1},
    "5D":  {"AR": 3, "FR": 3, "ANG": 2, "PH": 2, "IR": 2, "IC": 1, "M": 4, "PC": 4, "SN": 6, "HG": 2, "EPS": 1},
    "6D":  {"AR": 2, "FR": 2, "ANG": 2, "PH": 2, "IR": 2, "IC": 1, "M": 4, "PC": 6, "SN": 6, "HG": 2, "EPS": 1},
    "7A":  {"AR": 6, "FR": 6, "ANG": 4, "PH": 5, "IR": 2, "M": 2, "HG": 4, "EPS": 1},
}

def clean(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()

def num(v):
    try:
        return float(str(v).replace(",", ".")) if v not in (None, "") else None
    except ValueError:
        return None

def rows(path):
    ws = load_workbook(path, data_only=True).active
    for r in list(ws.iter_rows(values_only=True))[1:]:      # ligne 1 = en-têtes
        if r and r[0] is not None and r[1] is not None:
            yield r

students = {}
# liste.xlsx : colonnes  classe | numero | nom
for r in rows("liste.xlsx"):
    c, n = clean(r[0]).upper(), clean(r[1])
    students[(c, n)] = {"classe": c, "numero": n, "nom": clean(r[2]), "notes": {}}

# notes/CODE.xlsx : colonnes  classe | numero | test | examen
for path in sorted(glob.glob("notes/*.xlsx")):
    code = os.path.splitext(os.path.basename(path))[0].upper()
    for r in rows(path):
        k = (clean(r[0]).upper(), clean(r[1]))
        if k in students:
            students[k]["notes"][code] = [num(r[2]), num(r[3])]

with open("data.json", "w", encoding="utf-8") as f:
    json.dump({"subjects": SUBJECTS, "coef": COEF, "students": list(students.values())}, f, ensure_ascii=False)
print(len(students), "élèves")
