
Direct-Line Network Tool (DLNT)
===============================

Calculation of direct-line networks based on traffic cells and their centrality according to the German guidelines for integrated network design (RIN).

Application
-----------

### Application using graphical user interface (GUI)

Two possibilities

*   No running Visum model:  
    Start GUI from extern, running (execute _dlnt\_GUI.py_)
*   Running Visum model:  
    *   Integrate execution of the GUI (_dlnt\_GUI.py_) into the script menu
    *   Start script via script menu  
        

### Command-based application

The direct-line network tool can also be used without the GUI. For this, an instance of the class **DirectLineNetworkCalculator** must be initiated. After that, the methods of the instance (Import, Calculation, Export) can be accessed. An example can be found in _example\_application\_wo\_gui.py_.

Required procedural steps
-------------------------

1.  Set parameters (which CFL, attribute values, etc.)
2.  Initialize Direct Distance Calculator object
    *   Code: Calling the constructor with a parameter assignment
    *   GUI: Execute "Read data" in toolbar
3.  Calculate direct-line connections
    *   Code: Call `calculate_main` method
    *   GUI: Execute 'Perform Calculation' in the toolbar
4.  Export results to modell in desired format
    *   Code: Call of `export_matrix` or `export_net` methods
    *   GUI: Use the corresponding buttons _Mtx_ or _Net_ of the column 'Export to Visum as'.  
        Alternatively, the button _"Export to Visum all CFL: all links and matrices"_ exports the combined results of all connectivity function levels.

### Notes:

*   In the GUI, current calculations are reset if one of the zone attribute selections is changed. A new instance of the DirectLineNetworkCalculator is created.
*   If zone values are changed in Visum, they are not automatically updated in the DirectLineNetworkCalculator instance. Therefore, the procedure must be repeated from step 2.
*   Initial parameter values can be set in the GUI using the button "Default Values"
*   "Reset Calculation" deletes all already existing results
*   Step 3 uses the parameters currently entered in the GUI/set in the DirectLineNetworkCalculator instance. Before calculating with new parameters, it is recommended to delete existing results ("Reset Calculation").

Requirement
-----------

transport modell in PTV Visum including categorized zones:

*   Attribute for centrality of traffic cells: the **smaller** the number, the greater the centrality of the zone
    
    Example: Indication of centrality via the type number
    
    *   0 ... Metropolitan region
    *   1 ... Higher-order centres
    *   2 ... Middle-order centres
    *   3 ... Lower-order centres
    *   4 ... Place without central function
    *   5 ... Suburb
    
*   (optional) Active zone filter
*   (optional) Specification of an attribute defining which zone should be used as an origin {0=No, 1=Yes}
*   (optional) Specification of an attribute defining which zone should be used as a destination {0=No, 1=Yes}