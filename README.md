User Guide for Product Information Entry Application
Overview
This application allows users to enter product information, validate the data, and save it to an Excel file. The application is built using Python's tkinter library for the GUI and openpyxl for handling Excel files.
How to Use
1.	Launch the Application:
o	Run the script to open the Product Information Entry window.
2.	Enter Product Information:
o	Fill in all the fields in the form:
	Technician ID
	Serial Number
	Product ESO
	Date of Build (YYYY-MM-DD)
	Gap at 3 o'clock (mm)
	Gap at 6 o'clock (mm)
	Gap at 9 o'clock (mm)
	Gap at 12 o'clock (mm)
	Shim Thickness RF (mm)
	Shim Thickness RR (mm)
	Shim Thickness LF (mm)
	Shim Thickness LR (mm)
3.	Submit Data:
o	Click the Submit button to validate and save the data.
o	If any field is empty or contains invalid data, an error message will be displayed.
4.	Clear Fields:
o	Click the Clear button to reset all fields.
Data Validation
•	Date of Build: Must be in YYYY-MM-DD format.
•	Technician ID: Must be a numeric value.
•	Shim Thickness: Must be between 0.00mm and 12.7mm.
•	Gap Thickness: Must be between 0.0000mm and 0.0762mm.
Saving Data
•	The data is saved to an Excel file located at C:\Users\srat\OneDrive\Documents\product info entry.xlsx.
•	If the file does not exist, it will be created.
•	If the file exists, new data will be appended to it.
Error Handling
•	If there is a permission error while saving the file, an error message will be displayed.
Example Usage
1.	Enter the following data:
o	Technician ID: 56734
o	Serial Number: ZNL00678
o	Product ESO: YHNBG
o	Date of Build: 2025-03-08
o	Gap at 3 o'clock: 0.006
o	Gap at 6 o'clock: 0.007
o	Gap at 9 o'clock: 0.005
o	Gap at 12 o'clock: 0.004
o	Shim Thickness RF: 4.5
o	Shim Thickness RR: 4.6
o	Shim Thickness LF: 5
o	Shim Thickness LR: 4.9
2.	Click Submit to save the data.
3.	If successful, the message "Data saved to product_info.xlsx" will be displayed.
4.	If you are happy with your data, you can click Clear.
