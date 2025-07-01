## @package testing_scenarios.py
# @brief Enthält Methoden und einen Ablauf um definierte Testszenarien auszuführen und das Ergebnis in einer PPP darzustellen

from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt


## @brief Erstellt einen Screenshot des Visumnetzes und speichert ihn als .jpg Datei.
## @param name_screenshot Name der zu speichernden Screenshot-Datei.
## @param gpa_path Optionaler Pfad zu gespeicherten Grafikparametern.
## @return Der Pfad zur gespeicherten .jpg-Datei.
def take_screenshot(name_screenshot, gpa_path: str = ""):
    """
    Diese Funktion erstellt einen Screenshot des aktuellen Visumnetzes.
    Falls ein Pfad zu gespeicherten Grafikparametern übergeben wird, werden diese geladen.
    Die Datei wird unter dem Namen `name_screenshot` als .jpg gespeichert.
    """
    if gpa_path != "":
        # Lade die gespeicherten Grafikparameter, wenn ein Pfad angegeben ist
        Visum.Net.GraphicParameters.Open(Path.cwd() / gpa_path)

    # Das gesamte Netz wird in den Viewport geladen und angezeigt
    Visum.Graphic.DisplayEntireNetwork()

    # Der Screenshot wird erstellt und als .jpg Datei gespeichert
    jpg_path = Path(Visum.GetPath(41)) / f"{name_screenshot}.jpg"
    Visum.Graphic.Screenshot(jpg_path, 3)  # Parameter 3 gibt die Qualitätsstufe an
    return jpg_path


## @brief Fügt eine neue Folie mit einem Screenshot hinzu.
## @param prs PowerPoint-Präsentationsobjekt.
## @param jpg_path Pfad zum Screenshot-Bild.
## @param title Titel, der auf der Folie angezeigt werden soll.
## @return Die erstellte Folie.
def add_slide_screenshot(prs, jpg_path, title):
    """
    Diese Funktion fügt eine neue Folie zur Präsentation hinzu.
    Die Folie enthält einen Screenshot, der unter dem Pfad `jpg_path` gespeichert ist,
    und einen Titel `title`, der oben auf der Folie angezeigt wird.
    """
    # Einfügen einer leeren Folie (Layout 9 ist eine leere Folie)
    slide = prs.slides.add_slide(prs.slide_layouts[9])
    # Setze den Titel der Folie
    slide.shapes.title.text = title
    # Füge das Bild auf die Folie, definiere die Größe und Position
    slide.shapes.add_picture(str(jpg_path), top=Cm(2.23), left=Cm(2.05), width=Cm(29.77))
    return slide


## @brief Filtert die Zonen basierend auf einem gegebenen Attribut.
## @param attr_source Das Attribut, auf dessen Wert die Zonen gefiltert werden.
def filter_zones_source(attr_source):
    """
    Diese Funktion filtert Zonen, basierend auf einem gegebenen Attribut.
    Nur Zonen, bei denen das Attribut größer als 0 ist, werden berücksichtigt.
    """
    filter = Visum.Filters.ZoneFilter()
    filter.Init()
    if attr_source is not None:
        # Setze die Bedingung für den Filter: Das Attribut muss größer als 0 sein
        filter.AddCondition("OP_NONE", False, attr_source, "GreaterVal", 0)
        filter.UseFilter = True  # Aktiviere den Filter


## @brief Filtert die Strecken nach den gegebenen Streckentypen.
## @param set_linktypes Eine Liste oder Menge der zu filternden Streckentypen.
def filter_links_vfs(set_linktypes):
    """
    Diese Funktion filtert die Strecken, basierend auf den gegebenen Streckentypen.
    Nur Strecken, deren Typ in `set_linktypes` enthalten ist, werden berücksichtigt.
    """
    filter = Visum.Filters.LinkFilter()
    filter.Init()
    # Filtere Strecken nach ihrem Typ (TypeNo), basierend auf der gegebenen Menge von Streckentypen
    filter.AddCondition("OP_NONE", False, "TypeNo", "ContainedIn", ",".join(str(x) for x in set_linktypes))
    filter.UseFilter = True  # Aktiviere den Filter


if __name__ == '__main__':
    # Standardbibliotheken laden
    from pathlib import Path
    import luftlinientool as llt  # Annahme: luftlinientool ist eine benutzerdefinierte Bibliothek

    # Definition der Testszenarien
    dict_scenarios = {
        "n=1, k=0,":
            {"k": 0, "n": 1, "attr_cfl": "TypeNo", "attr_origin": None, "attr_ziel": None},
        "n=1, k=0, VFS = AddVal2":
            {"k": 0, "n": 1, "attr_cfl": "AddVal2", "attr_origin": None, "attr_ziel": None},
        "n=1, k=0, istQuelle = istZiel = IstUntersuchungsgebiet":
            {"k": 0, "n": 1, "attr_cfl": "TypeNo",
             "attr_origin": "IstUntersuchungsgebiet", "attr_ziel": "IstUntersuchungsgebiet"},
        "n=1, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = None": {
            "k": 0, "n": 1, "attr_cfl": "TypeNo", "attr_origin": "IstUntersuchungsgebiet", "attr_ziel": None},
        "n=1, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = AddVal1 (1)":
            {"k": 0, "n": 1, "attr_cfl": "TypeNo", "attr_origin": "IstUntersuchungsgebiet", "attr_ziel": "AddVal1"},
        "n=1, k=1, istQuelle = IstUntersuchungsgebiet, istZiel = None":
            {"k": 1, "n": 1, "attr_cfl": "TypeNo", "attr_origin": "IstUntersuchungsgebiet", "attr_ziel": None},
        "n=2, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = None":
            {"k": 0, "n": 2, "attr_cfl": "TypeNo", "attr_origin": "IstUntersuchungsgebiet", "attr_ziel": None},
    }

    # Definition der Verbindungsfunktionsstufen (VFS)
    dict_vfs = {"VFS 0": 0, "VFS 1": 1, "VFS 2": 2}

    # Pfad zur Netzdatei und Laden der Visum-Version
    path_source = Path().cwd() / "Version"
    file_source = 'Beispielnetz.ver'
    Visum = llt.open_visum(path_source / file_source)

    # Initialisierung der PowerPoint-Präsentation mit Layout
    prs = Presentation("Layout.pptx")
    path_ppp = "Testszenarien.pptx"

    # Schleife über die Testszenarien
    for scenario, param in dict_scenarios.items():
        try:
            # Luftlinien-Tool initialisieren und Berechnungen durchführen
            llt1 = llt.LuftlinienCalculator(Visum, attr_cfl=param["attr_cfl"], dict_cfl=dict_vfs,
                                            max_distance=param["n"], no_suppliers=param["k"],
                                            attr_orig=param["attr_origin"], attr_dest=param["attr_ziel"])
            llt1.calculate_main()

            # Schleife über die definierten VFS und Ergebnisse exportieren
            for vfs in dict_vfs.keys():
                # Exportieren des Netzes und Erstellung von Screenshots
                llt1.export_net(list_vfs=[vfs], links_additive=True)
                filter_zones_source(param["attr_origin"])  # Filtere Zonen basierend auf Quelle
                filter_links_vfs(list(llt1.dict_export_linktypes.values()))  # Filtere Strecken basierend auf Typ

                # Screenshot erstellen und in PowerPoint einfügen
                jpg_path = take_screenshot(f"{vfs}, {scenario}", gpa_path="Links_VFS.gpa")
                add_slide_screenshot(prs, jpg_path, title=f"{vfs}, {scenario}")

                # Lade die Visum-Datei erneut, um alle Änderungen zurückzusetzen
                Visum.LoadVersion(path_source / file_source)
        except:
            # Fehler auf einer separaten Folie anzeigen
            slide = prs.slides.add_slide(prs.slide_layouts[9])
            slide.shapes.title.text = f"Fehler {scenario}"

        # Speichern der Präsentation
        prs.save(path_ppp)

    # Abschlussmeldung
    print("The END")
