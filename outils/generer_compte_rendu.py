import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_UNDERLINE
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
BLACK = RGBColor(0, 0, 0)
GREY = RGBColor(0x6B, 0x6B, 0x6B)
BODY = "Open Sans"
MONO = "Courier New"
W = Cm(17)

doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.PORTRAIT
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = Cm(1.6); sec.bottom_margin = Cm(1.6)
sec.header_distance = Cm(0.9)

st = doc.styles["Normal"]
st.font.name = BODY; st.font.size = Pt(10.5); st.font.color.rgb = BLACK
rf = st.element.get_or_add_rPr().get_or_add_rFonts()
for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"): rf.set(qn(k), BODY)
st.paragraph_format.space_after = Pt(3); st.paragraph_format.line_spacing = 1.1


def run(p, text, bold=False, italic=False, underline=False, color=None, size=None, font=None):
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if underline: r.underline = WD_UNDERLINE.SINGLE
    if color is not None: r.font.color.rgb = color
    if size: r.font.size = Pt(size)
    if font:
        r.font.name = font
        f = r._element.get_or_add_rPr().get_or_add_rFonts()
        for k in ("w:eastAsia", "w:cs"): f.set(qn(k), font)
    return r


def para(space_before=0, space_after=3, keep=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before); p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.keep_with_next = keep
    return p


def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd"); s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), hex_)
    tcPr.append(s)


def borders(table, color="7F7F7F", size="6"):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), size); e.set(qn("w:color"), color)
        b.append(e)
    tblPr.append(b)


def fix(table, widths):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblW")): tblPr.remove(old)
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(int(sum(widths) / 635))); w.set(qn("w:type"), "dxa"); tblPr.append(w)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for gc, wd in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths): gc.set(qn("w:w"), str(int(wd / 635)))
    for row in table.rows:
        for c, wd in zip(row.cells, widths): c.width = wd


def min_height(row, cm):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight"); h.set(qn("w:val"), str(int(cm * 567))); h.set(qn("w:hRule"), "atLeast")
    trPr.append(h)


def spacer(pt=4):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pt)


def answer(height=3.0, placeholder="Écris ta réponse ici."):
    t = doc.add_table(rows=1, cols=1); borders(t, "A6A6A6"); fix(t, [W])
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    min_height(t.rows[0], height)
    run(t.rows[0].cells[0].paragraphs[0], placeholder, italic=True, color=GREY, size=9.5)
    spacer()


def question(num, verb, rest):
    """Question numérotée, verbe d'action en gras (style des fiches du professeur)."""
    p = para(space_before=6, keep=True)
    p.paragraph_format.left_indent = Cm(1.0); p.paragraph_format.first_line_indent = Cm(-1.0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.0))
    run(p, f"{num}\t", bold=True)
    run(p, verb, bold=True)
    run(p, rest)


def section_title(text):
    p = para(space_before=12, space_after=4, keep=True)
    run(p, text, bold=True, underline=True, size=11.5)


def grid(rows, widths, height=0.75, header_rows=1):
    t = doc.add_table(rows=len(rows), cols=len(rows[0])); borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(rows):
        min_height(t.rows[i], height)
        for j, v in enumerate(r):
            c = t.cell(i, j); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            head = i < header_rows or j == 0
            if head: shade(c, "F2F2F2")
            run(p, v, bold=head, size=10)
    fix(t, widths)
    spacer()
    return t


# ---------- En-tête (style Courier New des fiches) ----------
hp = sec.header.paragraphs[0]
hp.paragraph_format.tab_stops.add_tab_stop(Cm(8.5), WD_TAB_ALIGNMENT.CENTER)
hp.paragraph_format.tab_stops.add_tab_stop(W, WD_TAB_ALIGNMENT.RIGHT)
run(hp, "Physique-Chimie\tLycée Nadar\tMr Gheraibia  2026/2027", font=MONO, size=9)

p = para(space_after=8)
run(p, "Thèmes : Constitution de la matière à l'échelle microscopique et macroscopique", bold=True, font=MONO, size=10)

p = para(space_after=8)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, "TP : Dosage par étalonnage à une échelle de teinte", bold=True, underline=True, size=15)
p = para(space_after=8); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, "Compte rendu à compléter sur ordinateur", bold=True, size=11)

p = para(space_after=8)
run(p, "Ce compte rendu accompagne le TP à distance réalisé avec le simulateur en ligne. Un viticulteur a préparé une solution de sulfate de cuivre qu'il pense être à ")
run(p, "2,5 g·L⁻¹", bold=True)
run(p, ". Tu dois retrouver la concentration de son échantillon à l'aide d'une ")
run(p, "échelle de teinte", bold=True)
run(p, ".")

ident = grid([
    ["Nom", "", "Prénom", ""],
    ["Classe", "", "Date", ""],
    ["Code échantillon", "X-", "Note", "/ 20"],
], widths=[Cm(3.6), Cm(5.0), Cm(3.0), Cm(5.4)], header_rows=0, height=0.8)
for r in ident.rows:
    shade(r.cells[2], "F2F2F2"); r.cells[2].paragraphs[0].runs[0].bold = True
    r.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

p = para(space_before=4, keep=True); run(p, "Consignes :", bold=True)
for verb, rest in [
    ("Ouvrir", " ce document à côté du site du TP et répondre au fur et à mesure. Les numéros sont ceux des questions du site (Q1 à Q11)."),
    ("Écrire", " chaque réponse dans le cadre prévu (effacer le texte gris)."),
    ("Détailler", " tous les calculs : formule littérale, application numérique, résultat avec unité."),
    ("Enregistrer", " le fichier sous le nom NOM_Prenom_TP-echelle-teinte.docx."),
    ("Déposer", " le fichier dans le dossier partagé de la classe."),
]:
    q = doc.add_paragraph(style="List Bullet"); q.paragraph_format.space_after = Pt(1)
    run(q, verb, bold=True); run(q, rest)

# ---------- Contexte ----------
section_title("Situation et hypothèses")
p = para(); run(p, "Problématique : ", bold=True)
run(p, "comment déterminer la concentration en masse d'une solution colorée inconnue, et la solution du viticulteur est-elle bien à 2,5 g·L⁻¹ ?", italic=True)
question("1.", "Formuler", " une hypothèse : d'après les béchers A, B et C, comment varie la couleur d'une solution de sulfate de cuivre avec sa concentration ?")
answer(2.3)
question("2.", "Proposer", " en 3 ou 4 lignes une méthode pour estimer la concentration de l'échantillon à partir de solutions de concentrations connues.")
answer(2.8)

# ---------- Partie 1 ----------
section_title("Partie 1 : Préparation de la solution mère par dissolution")
question("3.", "Calculer", " la concentration en masse Cm0 de la solution mère (m = 2,35 g de sulfate de cuivre pentahydraté ; V = 100,0 mL).")
answer(2.8, "Formule littérale :\nConversion :\nApplication numérique :\nRésultat : Cm0 = ………… g·L⁻¹")

# ---------- Partie 2 ----------
section_title("Partie 2 : Réalisation de l'échelle de teinte")
question("4.", "Montrer", " que Cfille = Cm0 × V0 / Vfiole. Indice : la masse de soluté prélevée ne change pas pendant la dilution.")
answer(3.2)
question("5.", "Compléter", " le tableau (fiole jaugée de 50,0 mL), puis détailler un calcul de F et un calcul de Cm.")
grid([
    ["Tube n°", "1", "2", "3", "4", "5", "6"],
    ["Solution", "S1", "S2", "S3", "S4", "S5", "S6"],
    ["V0 à prélever (mL)", "2,0", "3,0", "4,0", "6,0", "8,0", "10,0"],
    ["Facteur de dilution F", "", "", "", "", "", ""],
    ["Cm (g·L⁻¹)", "", "", "", "", "", ""],
], widths=[Cm(4.4)] + [Cm(2.1)] * 6)
answer(1.8, "Tube 1 : F = …   ;   Cm = …")
question("5 b.", "Relever", " les volumes réellement prélevés sur le simulateur (cahier de manipulation de la paillasse 2).")
grid([
    ["Tube n°", "1", "2", "3", "4", "5", "6", "7", "8"],
    ["V0 prélevé (mL)", "", "", "", "", "", "", "", ""],
], widths=[Cm(4.2)] + [Cm(1.6)] * 8)
question("5 c.", "Décrire", " en 4 étapes la préparation d'une solution fille, en nommant la verrerie utilisée.")
answer(3.0, "1.\n2.\n3.\n4.")

# ---------- Partie 3 ----------
section_title("Partie 3 : Comparaison de la solution inconnue")
question("6.", "Indiquer", " entre quels tubes se situe l'échantillon, puis en déduire un encadrement de sa concentration.")
answer(1.6, "L'échantillon se situe entre le tube … et le tube …\n………… g·L⁻¹ < Cm < ………… g·L⁻¹")
question("6 b.", "Insérer", " une capture d'écran du portoir avec le tube X à sa place (Windows : touches Windows + Maj + S ; Mac : Cmd + Maj + 4 ; puis coller dans le cadre).")
answer(5.0, "Colle ta capture d'écran ici.")

# ---------- Exploitation ----------
section_title("Exploitation et conclusion")
question("7.", "Calculer", " la concentration estimée (milieu de l'encadrement) et l'incertitude U (demi-largeur). Écrire le résultat sous la forme Cm = … ± … g·L⁻¹.")
answer(2.3)
question("8.", "Calculer", " le volume V0 à prélever pour fabriquer un tube complémentaire (tube 7) situé entre les deux tubes qui encadrent X. Donner le nouvel encadrement.")
answer(3.0, "Concentration visée :\nCalcul de V0 :\nVolume prélevé :\nNouvel encadrement : ………… < Cm < ………… g·L⁻¹")
question("9.", "Expliquer", " pourquoi les tubes peu concentrés se distinguent mieux par le dessus, et pourquoi une même solution paraît plus foncée dans un bécher que dans un tube.")
answer(3.0)
question("10.", "Citer", " deux sources d'erreur lors de la préparation de l'échelle en vrai laboratoire et une limite de la méthode.")
answer(3.0, "Erreur 1 :\nErreur 2 :\nLimite :")
question("11.", "Conclure", " : répondre au viticulteur (5 lignes maximum) en donnant le résultat avec son incertitude. Sa solution peut-elle être à 2,5 g·L⁻¹ ?")
answer(3.5)

# ---------- Grille ----------
section_title("Grille d'évaluation (réservée au professeur)")
grid([
    ["Compétence", "Questions", "Barème", "Points"],
    ["S'approprier", "1", "1", ""],
    ["Analyser · Raisonner", "2, 4, 9", "4", ""],
    ["Réaliser", "3, 5, 5 b, 5 c", "6", ""],
    ["Valider", "6, 7, 8, 10", "6", ""],
    ["Communiquer", "6 b, 11, soin de la rédaction", "3", ""],
    ["Total", "", "20", ""],
], widths=[Cm(4.6), Cm(7.0), Cm(2.4), Cm(3.0)], height=0.65)
p = para(keep=True); run(p, "Commentaire :", bold=True)
answer(2.2, " ")

TBL = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
TC = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]
TR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"]


def reorder(el, order):
    kids = list(el)
    key = lambda k: order.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in order else 99
    for k in kids: el.remove(k)
    for k in sorted(kids, key=key): el.append(k)


body = doc.element.body
for t in body.iter(qn("w:tblPr")): reorder(t, TBL)
for t in body.iter(qn("w:tcPr")): reorder(t, TC)
for t in body.iter(qn("w:trPr")): reorder(t, TR)
doc.save(OUT)
print("saved", OUT)
