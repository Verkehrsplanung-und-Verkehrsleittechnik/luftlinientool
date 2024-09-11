from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt

def take_screenshot(name_screenshot, gpa_path: str = ""):
    if gpa_path != "":
        Visum.Net.GraphicParameters.Open(Path.cwd() / gpa_path)

    ## Optional aber empfohlen: Zoom auf ein Netzobjekt. z. B. Gebiet
    #Visum.Graphic.Autozoom(Visum.Net.Territories.ItemByKey(1))
    Visum.Graphic.DisplayEntireNetwork()
    # Visum.Graphic.SetScaleOfView(2)
    jpg_path = Path(Visum.GetPath(41)) / f"{name_screenshot}.jpg"
    Visum.Graphic.Screenshot(jpg_path, 3)
    return jpg_path

def add_slide_screenshot(prs, jpg_path, title):
    # Einfügen einer Folie mit bestimmtem Layout
    slide = prs.slides.add_slide(prs.slide_layouts[9])
    slide.shapes.title.text = title
    slide.shapes.add_picture(str(jpg_path), top=Cm(2.23), left=Cm(2.05), width=Cm(29.77))
    return slide

def filter_zones_source(attr_source):
    filter = Visum.Filters.ZoneFilter()
    filter.Init()
    if attr_source is not None:
        filter.AddCondition("OP_NONE", False, attr_source, "GreaterVal", 0)
        filter.UseFilter = True

def  filter_links_vfs(set_linktypes):
    filter = Visum.Filters.LinkFilter()
    filter.Init()
    filter.AddCondition("OP_NONE", False, "TypeNo", "ContainedIn", ",".join(str(x) for x in set_linktypes))
    filter.UseFilter = True

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    dict_scenarios = {
        "n=1, k=0,":
            {"k": 0, "n": 1, "attr_vfs": "TypeNo", "attr_quelle": None, "attr_ziel": None},
        "n=1, k=0, VFS = AddVal2":
            {"k": 0, "n": 1, "attr_vfs": "AddVal2", "attr_quelle": None, "attr_ziel": None},
        "n=1, k=0, istQuelle = istZiel = IstUntersuchungsgebiet":
            {"k": 0, "n": 1, "attr_vfs": "TypeNo",
             "attr_quelle": "IstUntersuchungsgebiet", "attr_ziel": "IstUntersuchungsgebiet"},
        "n=1, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = None": {
            "k": 0, "n": 1, "attr_vfs": "TypeNo", "attr_quelle": "IstUntersuchungsgebiet", "attr_ziel": None},
        "n=1, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = AddVal1 (1)":
            {"k": 0, "n": 1, "attr_vfs": "TypeNo", "attr_quelle": "IstUntersuchungsgebiet", "attr_ziel": "AddVal1"},
        "n=1, k=1, istQuelle = IstUntersuchungsgebiet, istZiel = None":
            {"k": 1, "n": 1, "attr_vfs": "TypeNo", "attr_quelle": "IstUntersuchungsgebiet", "attr_ziel": None},
        "n=2, k=0, istQuelle = IstUntersuchungsgebiet, istZiel = None":
            {"k": 0, "n": 2, "attr_vfs": "TypeNo", "attr_quelle": "IstUntersuchungsgebiet", "attr_ziel": None},
    }

    dict_vfs = {"VFS 0": 0, "VFS 1": 1, "VFS 2": 2}

    path_source = Path().cwd() / "Version"
    file_source = 'Beispielnetz.ver'

    Visum = llt.open_visum(path_source / file_source)

    # PPP initialisieren
    prs = Presentation("Layout.pptx")
    path_ppp = "Testszenarien.pptx"

    for scenario, param in dict_scenarios.items():
        try:
            # Berechnung
            llt1 = llt.LuftlinienCalculator( Visum, dict_vfs=dict_vfs,
                                             attr_vfs=param["attr_vfs"],
                                             attr_quelle=param["attr_quelle"],
                                             attr_ziel=param["attr_ziel"],
                                             anz_versorger=param["k"],
                                             max_entfernung=param["n"],
            )
            llt1.calculate_main()

            for vfs in dict_vfs.keys():
                # Export & Screenshot
                llt1.export_net(list_vfs =[vfs], links_additive=True)
                filter_zones_source(param["attr_quelle"])
                filter_links_vfs(list(llt1.dict_export_linktypes.values()))

                # Einfügen in PPP
                jpg_path = take_screenshot(f"{vfs}, {scenario}", gpa_path="Links_VFS.gpa")
                add_slide_screenshot(prs,jpg_path, title=f"{vfs}, {scenario}")

                # Clean up
                # Visum.Net.Links.RemoveAll(OnlyActive=True)
                Visum.LoadVersion(path_source / file_source)
        except:
            slide = prs.slides.add_slide(prs.slide_layouts[9])
            slide.shapes.title.text = f"Fehler {scenario}"

        # PPP Speichern
        prs.save(path_ppp)

    print("The END")
