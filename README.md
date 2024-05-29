# Excel_File_Comparison
That sounds like a fantastic initiative! Automation can greatly streamline processes and save valuable time, especially for tasks like comparing Excel sheets which can be tedious and prone to human error when done manually. Offering a script to automate this process not only helps people be more efficient but also empowers them to focus on more important aspects of their work. And presenting it in a professional manner adds to its appeal and usability


**To fulfill the requirement, you can follow these steps:**

**Install Required Libraries:** 

Ensure you have pandas and openpyxl installed. You can install them via pip if you haven't already:



Copy code
**pip install pandas openpyxl**
Save the Script: Copy the provided script into a Python file, for example, excel_comparison_gui.py.



**Run the Script**: 
Execute the script using Python:

Copy code
**python excel_comparison_gui.py**



**GUI Interface:**

Once you run the script, a GUI window titled "Excel Column Comparison" will appear.
You'll see input fields for File 1 Path, File 2 Path, Sheet 1 Name, Sheet 2 Name, and Columns to Compare (comma-separated).
Use the "Browse" buttons to select the Excel files.
Enter the sheet names and columns to compare.
Click the "Compare" button to start the comparison process.




**Error Handling:**

If any errors occur (such as missing files, incorrect sheet names, or missing columns), the script will display relevant error messages in pop-up dialogs.
Address any errors and try running the comparison again.



**
**Comparison Results:****

After successful comparison, the script will save the modified Excel files with highlighted cells.
The results will be saved as file1_comparison.xlsx and file2_comparison.xlsx in the same directory where the script is located.



**Review Results:**

Open the generated Excel files to review the comparison results.
Cells with values present in both files will be highlighted in yellow, while cells with values not present in both files will be highlighted in red.




**Enjoy........**
