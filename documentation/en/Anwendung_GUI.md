# Direct Distance Tool 2022
Creating direct distance connections for given districts based on the RIN.

## Application
### Usage via Dialog
Two options:
* Visum version is not yet open:  
  Start the GUI externally (run *llt_GUI.py*)  
* Visum version is already open:  
  * Integrate the execution of the GUI (*llt_GUI.py*) into the script menu  
  * Start the script via the script menu  
    
    Note: If the GUI is started & closed several times, an error message will appear. This can be ignored, as the functionality is still available.  

### Execution via Code
The Direct Distance Tool can also be used without the GUI. For this, a code execution must create an instance of the class *DirectDistanceCalculator*.  
After that, the methods of the instance (Import, Calculation, Export) can be accessed.  
An example can be seen in *Bsp_Aufruf_ohne_GUI.py*.  

### Steps to Execute
1. Set parameters (which VFS, attribute values, etc.)  

2. Initialize direct distance calculator object  
   
   * Code: Call constructor with parameter transfer  
   * GUI: Run "Read data" in toolbar  
   
3. Calculate/generate direct distance connections  

   * Code: Call `calculate_main` method  
   * GUI: Run "Calculate Direct Distance Network" in toolbar  
   
4. Export results in the desired form to Visum  

   * Code: Call `export_matrix`/`export_net` method  
   * GUI: Use the corresponding buttons (Mtx/Net) in the column *create in Visum as*. Alternatively, the button *Import to Visum all VFS routes + Mtx* transfers the combined results of all VFS.  

Notes:  
* In the GUI, current calculations are reset if one of the district attributes is changed. A new instance of the Direct Distance calculator is then created.  
* If district values in Visum are changed, they are not automatically updated in the Direct Distance calculator. Therefore, the procedure must be repeated from step 2.  
* The initial parameter values can be recalled in the GUI using the "Default values" tool.  
* "Initialize results" allows the deletion of already existing results.  
* Step 3 uses the parameters currently entered in the GUI. Before calculation with new parameters, it is recommended to delete the existing results ("Initialize results").  

## Requirements
Network with categorized districts:  
* Attribute for centrality (districts): the smaller the number, the greater the centrality of the district  

  Example: Indicating centrality via type number  
  * 0 ... Metropolitan region  
  * 1 ... Major center  
  * 2 ... Medium center  
  * 3 ... Basic center  
  * 4 ... Location without central function  
  * 5 ... Sub-location  
    
* (optional) Active district filter  
* (optional) Specification of an attribute defining which district should be used as the source {0=No, 1=Yes}  
* (optional) Specification of an attribute defining which district should be used as the target {0=No, 1=Yes}  
