# Application

## Application via Dialog

Two possibilities:

- **Visum version is still closed:**  
  GUI start externally (`ddt_GUI.py` execute)

- **Visum version is already open:**
  - Integrate execution of the GUI (`ddt_GUI.py`) into the script menu
  - Start script via script menu  
    **Note:** If the GUI is started & closed multiple times, an error message will appear. This can be ignored; the functionality is still intact.

## Call via Code

The Direct Distance Tool can also be used without the GUI. For this, an instance of the class `DirectDistanceCalculator` must be created as a code execution. After that, the methods of the instance (Import, Calculation, Export) can be accessed. An example can be seen under `Bsp_Aufruf_ohne_GUI.py`.

### Steps to execute

1. **Set parameters** (which VFS, attribute values, etc.)

2. **Initialize Direct Distance Calculator object**
   - **Code:** Call constructor with parameter passing
   - **GUI:** Execute "Read data" in toolbar

3. **Calculate/create direct distance connections**
   - **Code:** Call `calculate_main` method
   - **GUI:** Execute "Direct Distance Network Calculation" in toolbar

4. **Export results to Visum in desired format**
   - **Code:** Call `export_matrix` / `export_net` method
   - **GUI:** Use the corresponding buttons (Mtx/Net) in the column "create in Visum as". Alternatively, the button "Import to Visum all VFS routes + Mtx" transfers the combined results of all VFS.

### Notes

- In the GUI, current calculations are reset if one of the zone attribute selections is changed. A new instance of the Direct Distance Calculator is created.
- If zone values are changed in Visum, they are not automatically updated in the Direct Distance Calculator. Therefore, the procedure must be repeated from step 2.
- Initial parameter values can be recalled in the GUI via the tool "Default values"
- "Initialize results" enables deletion of already existing results
- Step 3 uses the parameters currently entered in the GUI. Before calculating with new parameters, it is recommended to delete existing results ("Initialize results").

## Requirements

Network with categorized zones:

- **Attribute for centrality (zones):** The smaller the number, the greater the centrality of the zone

  **Example:** Indication of centrality via the type number
  - `0` ... Metropolitan region
  - `1` ... Major centrality
  - `2` ... Medium centrality
  - `3` ... Basic centrality
  - `4` ... Place without central function
  - `5` ... Subdivision

- **(optional)** Active zone filter
- **(optional)** Indication of an attribute defining which zone should be used as a source `{0=No, 1=Yes}`
- **(optional)** Indication of an attribute defining which zone should be used as a target `{0=No, 1=Yes}`