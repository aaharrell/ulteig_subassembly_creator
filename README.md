# Created for use by Ulteig Engineering, Inc.  
___

##### By Austen Harrell  

"master_subassemblies_creator.py" should be compiled into an executable using the following:  
  
```pyinstaller -y -F -c master_subassemblies_creator.py```  

#### Source Code External Dependencies (`pip install <lib_name>` in cmd or powershell):  
  * openpyxl
  * pandas
  

## Instructions:  
### Retrieving the subassemblies:  
  1. Go to Xcel's system.
  2. Open Assembly Manager, then open the subassemblies list. Click on an active subassembly ProjectWise link.
     *  Alternatively, enter the following URN: pw:\\PWNSP.Corp.Xcelenergy.com:ProjectWise_NSP_Draw\Documents\Xcel Masters and Templates\Transmission\Masters\XEL\Assemblies\Record\
  3. Sort all subassembly in ascending (A-Z) order.
  4. `Ctrl+a` to select all drawings.
  5. `Right Click` -> "Export..." -> "Next" -> Choose/make a designated folder (e.g. "Subassemblies) in your OneDrive folder -> Export.


### Export the subassembly data from Xcel
  1. Make sure the ProjectWise folder is still sorted as described above.
  2. `Ctrl+a` to select all drawings.
  3. `Right Click` -> "Copy List To" -> "Clipboard Tab Separated"
  4. Open an instance of Microsoft Excel.
  5. Paste into cell A1.
  6. Save the notebook in the same location as all of the exported subassembly PDFs.
     *  IMPORTANT: Must be saved as a .xlsx file.
     *  IMPORTANT: The name of this file must start with "_" (e.g. "_Subassembly Details.xlsx), otherwise the specific naming doesn't matter.
    


  ### Creating the Ulteig reference spreadsheet  
  1. Move back to your Ulteig computer.
  2. Ensure that the contents of the OneDrive folder (subassembly PDFs & Excel subassembly data) have **both** fully downloaded to your computer and contain only the subassembly drawings and single "_ ... .xlsx" file.
  3. Move this entire OneDrive folder to the desired location on an Ulteig network drive.
  4. Open the folder (you should see all of the drawings inside with the Excel spreadsheet at the top).
  5. Copy the address to this folder using the file explorer address bar.
  6. Run "master_subassemblies_creator.exe" from this repo; you can run this application from anywhere on your machine.
     *  This application will open a console window.
     *  You will paste the folder address into this application (may have to use `Right Click` to successfully paste into the console; 'Ctrl + p' should not work).
  7. The final Excel file will be created in the same folder that the application was executed from (see the SUCCESS message for exact location of the Excel file).
