'STATS INFO

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Project Krabappel\KRABAPPEL FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'--------- Project Krabappel --------------

EMConnect ""

call excel_open("True")
call excel_open_file("C:\DHS-MAXIS-Scripts\Project Krabappel\Krabappel template.xlsx","false")

'call excel_read(row,col)
'call excel_write(row,col,value)

'Appl Case
'For next for MEMB/MEMI
'ADDR is separate
'Do all STAT panels
'STORE ALL CASE NUMBERS AS AN ARRAY!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Do approval


call Script_End_procedure( "" )