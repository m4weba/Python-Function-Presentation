from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# Präsentation erstellen
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


# Farben
DUNKELBLAU = RGBColor(31, 55, 86)
BLAU = RGBColor(52, 152, 219)
GRUEN = RGBColor(39, 174, 96)
ORANGE = RGBColor(230, 126, 34)
ROT = RGBColor(192, 57, 43)
GRAU = RGBColor(245, 247, 250)
DUNKELGRAU = RGBColor(50, 50, 50)
WEISS = RGBColor(255, 255, 255)


def hintergrund(slide, farbe=GRAU):
    """Färbt den Hintergrund einer Folie."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        prs.slide_width,
        prs.slide_height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = farbe
    shape.line.fill.background()
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)


def titel(slide, text, farbe=DUNKELBLAU):
    """Fügt einen Titel hinzu."""
    textbox = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.3), Inches(12), Inches(0.8)
    )
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = farbe


def textfeld(slide, text, x, y, breite, hoehe,
             schriftgroesse=22, farbe=DUNKELGRAU, fett=False):
    """Fügt ein Textfeld hinzu."""
    textbox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(breite), Inches(hoehe)
    )
    tf = textbox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(schriftgroesse)
    p.font.color.rgb = farbe
    p.font.bold = fett

    return textbox


def codefeld(slide, code, x=0.8, y=2.0, breite=11.7, hoehe=3.6):
    """Fügt ein Codefeld mit Monospace-Schrift ein."""
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(breite), Inches(hoehe)
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(35, 39, 42)
    box.line.color.rgb = RGBColor(80, 80, 80)

    textbox = slide.shapes.add_textbox(
        Inches(x + 0.25), Inches(y + 0.2),
        Inches(breite - 0.5), Inches(hoehe - 0.4)
    )
    tf = textbox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = code
    p.font.name = "Consolas"
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(235, 235, 235)


def aufzählung(slide, punkte, x=1.0, y=1.7, breite=11, hoehe=4.5):
    """Fügt eine Aufzählung hinzu."""
    textbox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(breite), Inches(hoehe)
    )
    tf = textbox.text_frame
    tf.word_wrap = True

    for nummer, punkt in enumerate(punkte):
        p = tf.paragraphs[0] if nummer == 0 else tf.add_paragraph()
        p.text = punkt
        p.font.size = Pt(23)
        p.font.color.rgb = DUNKELGRAU
        p.space_after = Pt(14)
        p.level = 0


def neue_folie(titel_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    hintergrund(slide)
    titel(slide, titel_text)
    return slide


# Folie 1: Titelfolie
slide = prs.slides.add_slide(prs.slide_layouts[6])
hintergrund(slide, DUNKELBLAU)

textfeld(
    slide,
    "Funktionen in Python",
    1.0, 1.8, 11.3, 1.0,
    schriftgroesse=42,
    farbe=WEISS,
    fett=True
)

textfeld(
    slide,
    "Parameter, Argumente, return sowie globale und lokale Variablen",
    1.0, 3.0, 11.3, 1.0,
    schriftgroesse=24,
    farbe=RGBColor(220, 230, 240)
)

textfeld(
    slide,
    "Eine Einführung mit Beispielen",
    1.0, 5.7, 11.3, 0.5,
    schriftgroesse=18,
    farbe=RGBColor(220, 230, 240)
)


# Folie 2: Übersicht
slide = neue_folie("1. Was sind Funktionen?")
aufzählung(slide, [
    "Eine Funktion ist ein wiederverwendbarer Programmblock.",
    "Sie führt eine bestimmte Aufgabe aus.",
    "Funktionen machen Programme übersichtlicher und vermeiden Wiederholungen.",
    "In Python beginnt eine Funktionsdefinition mit dem Schlüsselwort def."
])

codefeld(slide, """def begruessung():
    print("Hallo!")

begruessung()""")


# Folie 3: Aufbau
slide = neue_folie("2. Aufbau einer Funktion")

textfeld(
    slide,
    "Eine Funktion besteht meistens aus:",
    0.9, 1.3, 11, 0.5,
    schriftgroesse=24, fett=True
)

aufzählung(slide, [
    "def: Schlüsselwort zum Erstellen einer Funktion",
    "Funktionsname: beschreibt, was die Funktion macht",
    "Parameter: Eingabewerte der Funktion",
    "Funktionsblock: eingerückter Code",
    "return: gibt ein Ergebnis zurück"
], x=1.0, y=1.9, hoehe=3.5)

codefeld(slide, """def addiere(zahl1, zahl2):
    ergebnis = zahl1 + zahl2
    return ergebnis""")


# Folie 4: Parameter
slide = neue_folie("3. Parameter")

textfeld(
    slide,
    "Parameter sind Platzhalter für Werte, die eine Funktion benötigt.",
    0.9, 1.2, 11.5, 0.7,
    schriftgroesse=24
)

codefeld(slide, """def begruesse(name):
    print("Hallo", name)

begruesse("Lena")
begruesse("Max")""")

textfeld(
    slide,
    "Der Parameter name wird beim Aufruf durch einen konkreten Wert ersetzt.",
    1.0, 6.0, 11.2, 0.5,
    schriftgroesse=20, farbe=BLAU, fett=True
)


# Folie 5: Argumente
slide = neue_folie("4. Argumente")

textfeld(
    slide,
    "Argumente sind die konkreten Werte, die beim Funktionsaufruf übergeben werden.",
    0.9, 1.2, 11.7, 0.8,
    schriftgroesse=23
)

codefeld(slide, """def multipliziere(a, b):   # a und b sind Parameter
    return a * b

ergebnis = multipliziere(4, 5)  # 4 und 5 sind Argumente
print(ergebnis)                 # Ausgabe: 20""")

aufzählung(slide, [
    "Parameter stehen in der Funktionsdefinition.",
    "Argumente stehen beim Aufruf der Funktion.",
    "Eine Funktion kann mehrere Parameter besitzen."
], x=1.0, y=5.9, hoehe=1.2)


# Folie 6: return
slide = neue_folie("5. Das Schlüsselwort return")

textfeld(
    slide,
    "Mit return gibt eine Funktion ein Ergebnis an die aufrufende Stelle zurück.",
    0.9, 1.2, 11.6, 0.8,
    schriftgroesse=23
)

codefeld(slide, """def quadrat(zahl):
    return zahl * zahl

ergebnis = quadrat(6)
print(ergebnis)  # Ausgabe: 36""")

aufzählung(slide, [
    "return beendet die Funktion.",
    "Der zurückgegebene Wert kann gespeichert oder weiterverwendet werden.",
    "Ohne return gibt eine Funktion standardmäßig None zurück."
], x=1.0, y=5.9, hoehe=1.2)


# Folie 7: print und return
slide = neue_folie("6. Unterschied zwischen print und return")

textfeld(
    slide,
    "print zeigt einen Wert nur auf dem Bildschirm an. return gibt ihn zurück.",
    0.9, 1.2, 11.6, 0.8,
    schriftgroesse=23
)

codefeld(slide, """def mit_print():
    print("Das ist nur eine Ausgabe")

def mit_return():
    return "Das ist ein Rückgabewert"

wert = mit_return()
print(wert)""")

textfeld(
    slide,
    "Merksatz: print zeigt etwas – return liefert etwas zurück.",
    1.0, 6.1, 11.2, 0.5,
    schriftgroesse=22, farbe=ORANGE, fett=True
)


# Folie 8: lokale Variablen
slide = neue_folie("7. Lokale Variablen")

textfeld(
    slide,
    "Eine lokale Variable wird innerhalb einer Funktion erstellt und ist normalerweise nur dort verfügbar.",
    0.9, 1.2, 11.7, 0.9,
    schriftgroesse=22
)

codefeld(slide, """def beispiel():
    nachricht = "Hallo"
    print(nachricht)

beispiel()
# print(nachricht)  # Fehler: Variable nicht außerhalb verfügbar""")

textfeld(
    slide,
    "Der Gültigkeitsbereich einer lokalen Variable ist auf die Funktion begrenzt.",
    1.0, 6.1, 11.2, 0.5,
    schriftgroesse=21, farbe=ROT, fett=True
)


# Folie 9: globale Variablen
slide = neue_folie("8. Globale Variablen")

textfeld(
    slide,
    "Eine globale Variable wird außerhalb von Funktionen definiert und kann im gesamten Programm gelesen werden.",
    0.9, 1.2, 11.7, 0.9,
    schriftgroesse=22
)

codefeld(slide, """sprache = "Python"  # globale Variable

def zeige_sprache():
    print(sprache)

zeige_sprache()
print(sprache)""")

textfeld(
    slide,
    "Globale Variablen sollten sparsam verwendet werden, weil sie Programme unübersichtlicher machen können.",
    1.0, 6.1, 11.2, 0.5,
    schriftgroesse=19, farbe=ORANGE, fett=True
)


# Folie 10: global
slide = neue_folie("9. Das Schlüsselwort global")

textfeld(
    slide,
    "Mit global kann eine Funktion eine globale Variable verändern.",
    0.9, 1.2, 11.6, 0.7,
    schriftgroesse=23
)

codefeld(slide, """punktzahl = 0

def punkt_hinzufuegen():
    global punktzahl
    punktzahl += 10

punkt_hinzufuegen()
print(punktzahl)  # Ausgabe: 10""")

textfeld(
    slide,
    "Besser ist häufig, einen Wert mit return zurückzugeben.",
    1.0, 6.1, 11.2, 0.5,
    schriftgroesse=21, farbe=GRUEN, fett=True
)


# Folie 11: Kombination
slide = neue_folie("10. Alles zusammen")

codefeld(slide, """def berechne_preis(preis, rabatt):
    neuer_preis = preis - rabatt
    return neuer_preis

grundpreis = 50
rabatt = 10

endpreis = berechne_preis(grundpreis, rabatt)
print("Endpreis:", endpreis)""")

aufzählung(slide, [
    "preis und rabatt sind Parameter.",
    "grundpreis und rabatt beim Aufruf sind Argumente.",
    "neuer_preis ist eine lokale Variable.",
    "return gibt den berechneten Endpreis zurück."
], x=1.0, y=5.4, hoehe=1.5)


# Folie 12: Übung
slide = neue_folie("11. Kleine Übung")

textfeld(
    slide,
    "Aufgabe: Schreibe eine Funktion, die den Durchschnitt von drei Zahlen berechnet.",
    0.9, 1.3, 11.7, 0.8,
    schriftgroesse=24
)

codefeld(slide, """def durchschnitt(a, b, c):
    # Dein Code hier
    pass

print(durchschnitt(4, 6, 8))""")

textfeld(
    slide,
    "Tipp: Addiere die drei Zahlen und teile das Ergebnis durch 3.",
    1.0, 6.1, 11.2, 0.5,
    schriftgroesse=21, farbe=BLAU, fett=True
)


# Folie 13: Zusammenfassung
slide = neue_folie("12. Zusammenfassung")

aufzählung(slide, [
    "Funktionen bündeln wiederverwendbaren Code.",
    "Parameter sind Platzhalter in einer Funktion.",
    "Argumente sind konkrete Werte beim Funktionsaufruf.",
    "return gibt ein Ergebnis zurück.",
    "Lokale Variablen gelten nur innerhalb einer Funktion.",
    "Globale Variablen gelten im gesamten Programm.",
    "Funktionen machen Code übersichtlich und leichter testbar."
], x=1.0, y=1.5, hoehe=4.8)

textfeld(
    slide,
    "Vielen Dank für eure Aufmerksamkeit!",
    1.0, 6.4, 11.2, 0.5,
    schriftgroesse=26, farbe=GRUEN, fett=True
)


# Präsentation speichern
prs.save("Python_Funktionen.pptx")
print("Die Präsentation wurde als Python_Funktionen.pptx gespeichert.")
