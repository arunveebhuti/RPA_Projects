"# RPA_Projects on Excel made in AUTOMATION ANYWHERE " 

-Folder Bacic Automations :

Some use cases of excel comands like deleting empty rows etc..

- Excel project 1 folder contains source code and the outputfile for the following usecase

Process Use Case
1.	Open the Data.xlsx file.
•	The file contains data for payment mode, bank, and RRN.
2.	Open the Macro Sheet.xlsx file
•	This file contains macro names for the Data file.
3.	Developing a Bot to capture the right macro names and update the Data file.

Business Rules/Conditions to Automate the Process
The Bot you develop is expected to do the following
•	The Bot needs to open the Data file
•	In the Data file, the Bot needs to read payment mode, bank name, and RRN number row by row.
•	The Bot needs to check for all the given data in the Macro Sheet file.
	If payment mode, bank name, and RRN number is a (Match)(Valid) then pick up the macro name and dump it into the Data file in Column E.
	If the bank name is SBI, Corporation Bank, HDFC, or ICICI then consider bank name in Macro Sheet as – (Hyphen)

