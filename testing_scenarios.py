## @package testing_scenarios.py
# @brief Contains methods and a workflow to execute defined test scenarios and present the results in a PPT presentation

from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt


## @brief Creates a screenshot of the Visum network and saves it as a .jpg file.
## @param name_screenshot Name of the screenshot file to be saved.
## @param gpa_path Optional path to saved graphic parameters.
## @return The path to the saved .jpg file.
def take_screenshot(name_screenshot, gpa_path: str = ""):
    """
    This function creates a screenshot of the current Visum network.
    If a path to saved graphic parameters is provided, they will be loaded.
    The file is saved under the name `name_screenshot` as a .jpg file.
    """
    if gpa_path != "":
        # Load the saved graphic parameters if a path is specified
        Visum.Net.GraphicParameters.Open(Path.cwd() / gpa_path)

    # The entire network is loaded into the viewport and displayed
    Visum.Graphic.DisplayEntireNetwork()

    # The screenshot is created and saved as a .jpg file
    jpg_path = Path(Visum.GetPath(41)) / f"{name_screenshot}.jpg"
    Visum.Graphic.Screenshot(jpg_path, 3)  # Parameter 3 specifies the quality level
    return jpg_path


## @brief Adds a new slide with a screenshot.
## @param prs PowerPoint presentation object.
## @param jpg_path Path to the screenshot image.
## @param title Title to be displayed on the slide.
## @return The created slide.
def add_slide_screenshot(prs, jpg_path, title):
    """
    This function adds a new slide to the presentation.
    The slide contains a screenshot that is saved under the path `jpg_path`,
    and a title `title` that is displayed at the top of the slide.
    """
    # Insert an empty slide (Layout 9 is an empty slide)
    slide = prs.slides.add_slide(prs.slide_layouts[9])
    # Set the title of the slide
    slide.shapes.title.text = title
    # Add the image to the slide, define the size and position
    slide.shapes.add_picture(str(jpg_path), top=Cm(2.23), left=Cm(2.05), width=Cm(29.77))
    return slide


## @brief Filters the zones based on a given attribute.
## @param attr_source The attribute whose value is used to filter the zones.
def filter_zones_source(attr_source):
    """
    This function filters zones based on a given attribute.
    Only zones where the attribute is greater than 0 are considered.
    """
    filter = Visum.Filters.ZoneFilter()
    filter.Init()
    if attr_source is not None:
        # Set the condition for the filter: The attribute must be greater than 0
        filter.AddCondition("OP_NONE", False, attr_source, "GreaterVal", 0)
        filter.UseFilter = True  # Activate the filter


## @brief Filters the links according to the given link types.
## @param set_linktypes A list or set of link types to be filtered.
def filter_links_vfs(set_linktypes):
    """
    This function filters the links based on the given link types.
    Only links whose type is contained in `set_linktypes` are considered.
    """
    filter = Visum.Filters.LinkFilter()
    filter.Init()
    # Filter links by their type (TypeNo), based on the given set of link types
    filter.AddCondition("OP_NONE", False, "TypeNo", "ContainedIn", ",".join(str(x) for x in set_linktypes))
    filter.UseFilter = True  # Activate the filter


if __name__ == '__main__':
    # Load standard libraries
    from pathlib import Path
    import luftlinientool as llt  # Assumption: luftlinientool is a custom library

    # Definition of test scenarios
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

    # Definition of connection function levels (VFS)
    dict_vfs = {"VFS 0": 0, "VFS 1": 1, "VFS 2": 2}

    # Path to the network file and loading the Visum version
    path_source = Path().cwd() / "Version"
    file_source = 'Beispielnetz.ver'
    Visum = llt.open_visum(path_source / file_source)

    # Initialization of the PowerPoint presentation with layout
    prs = Presentation("Layout.pptx")
    path_ppp = "Testszenarien.pptx"

    # Loop over the test scenarios
    for scenario, param in dict_scenarios.items():
        try:
            # Initialize the air-line tool and perform calculations
            llt1 = llt.LuftlinienCalculator(Visum, attr_cfl=param["attr_cfl"], dict_cfl=dict_vfs,
                                            max_distance=param["n"], no_suppliers=param["k"],
                                            attr_orig=param["attr_origin"], attr_dest=param["attr_ziel"])
            llt1.calculate_main()

            # Loop over the defined VFS and export results
            for vfs in dict_vfs.keys():
                # Export the network and create screenshots
                llt1.export_net(links_additive=True, list_cfl=[vfs])
                filter_zones_source(param["attr_origin"])  # Filter zones based on source
                filter_links_vfs(list(llt1.dict_export_linktypes.values()))  # Filter links based on type

                # Create screenshot and insert into PowerPoint
                jpg_path = take_screenshot(f"{vfs}, {scenario}", gpa_path="Links_VFS.gpa")
                add_slide_screenshot(prs, jpg_path, title=f"{vfs}, {scenario}")

                # Reload the Visum file to reset all changes
                Visum.LoadVersion(path_source / file_source)
        except:
            # Display error on a separate slide
            slide = prs.slides.add_slide(prs.slide_layouts[9])
            slide.shapes.title.text = f"Error {scenario}"

        # Save the presentation
        prs.save(path_ppp)

    # Final message
    print("The END")
