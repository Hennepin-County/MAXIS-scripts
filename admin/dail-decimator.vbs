'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - DAIL DECIMATOR.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
'see below dialog selection
STATS_denomination = "C"       			'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/02/2018", "Added Casey Love as autorized user of the script, and blanked out MAXIS case number for PRIV cases.", "Ilse Ferris, Hennepin County")
call changelog_update("12/30/2017", "Complete updates for INFO, SVES, COLA and ELIG messages.", "Ilse Ferris, Hennepin County")
call changelog_update("12/11/2017", "Added Quality Improvement Team as authorized users of DAIL Decimator script.", "Ilse Ferris, Hennepin County")
call changelog_update("12/05/2017", "Added ELIG DAIL messages as DAILs to decimate!", "Ilse Ferris, Hennepin County")
call changelog_update("10/28/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

Function dail_selection
	'selecting the type of DAIl message
	EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	transmit
	EMWriteScreen "_", 7, 39		'clears the all selection
	
	IF dail_to_decimate = "COLA" then selection_row = 8
	IF dail_to_decimate = "ELIG" then selection_row = 11
	IF dail_to_decimate = "INFO" then selection_row = 13
	IF dail_to_decimate = "SVES" then selection_row = 13
	
	Call write_value_and_transmit("x", selection_row, 39)	
End Function


'END CHANGELOG BLOCK =======================================================================================================

BeginDialog dail_dialog, 0, 0, 266, 110, "Dail Decimator dialog"
  DropListBox 80, 50, 60, 15, "Select one..."+chr(9)+"COLA"+chr(9)+"ELIG"+chr(9)+"INFO"+chr(9)+"SVES", dail_to_decimate
  EditBox 80, 70, 180, 15, worker_number
  CheckBox 15, 95, 135, 10, "Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 155, 90, 50, 15
    CancelButton 210, 90, 50, 15
  Text 15, 75, 60, 10, "Worker number(s):"
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
  Text 20, 20, 235, 20, "This script should be used to remove DAIL messages that have been determined by Quality Improvement staff do not require action."
  Text 40, 55, 35, 10, "Dail type:"
EndDialog

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
worker_signature = "I. Ferris/BZS and QI Teams"

'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Validating user ID
If user_ID_for_validation = "ILFE001" OR _		
	user_ID_for_validation = "WF7638" OR _		
	user_ID_for_validation = "WF1875" OR _ 		
	user_ID_for_validation = "WFQ898" OR _ 		
	user_ID_for_validation = "WFP803" OR _		
	user_ID_for_validation = "WFP106" OR _		
	user_ID_for_validation = "WFK093" OR _ 		
	user_ID_for_validation = "WF1373" OR _ 		
	user_ID_for_validation = "WFU161" OR _ 		
	user_ID_for_validation = "WFS395" OR _ 		
	user_ID_for_validation = "WFU851" OR _ 		
	user_ID_for_validation = "WFX901" OR _ 
	user_ID_for_validation = "CALO001" OR _ 		
	user_ID_for_validation = "WFI021" then 		
    'the dialog
    Do
    	Do
      		err_msg = ""
      		dialog dail_dialog
      		If ButtonPressed = 0 then StopScript
      		If dail_to_decimate = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the type of DAIL message to decimate!"	
      		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."	
      		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."							
      	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine										
      	LOOP until err_msg = ""		
      	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
      Loop until are_we_passworded_out = false					'loops until user passwords back in		
Else 
	script_end_procedure("This script is for Quality Improvement staff only. You do not have access to use this script.")
End if 

If dail_to_decimate = "COLA" then STATS_manualtime = "40"                'manual run time in seconds
If dail_to_decimate = "INFO" then STATS_manualtime = "30"                'manual run time in seconds
If dail_to_decimate = "ELIG" then STATS_manualtime = "25"                'manual run time in seconds

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Deleted DAILS - " & dail_to_decimate 

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array	
	'msgbox worker
	DO 
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then 
			MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if 
	Loop until dail_check = "DAIL"
	
	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'
	
	Call dail_selection
	
	EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed
	
	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
		
		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""
			
		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message... 
				Call write_value_and_transmit("T", dail_row, 3)
				dail_row = 6
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if 
			
			EMReadScreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)
			stats_counter = stats_counter + 1
		
			'If instr(dail_msg, "TPQY RESPONSE") then 
			' 	add_to_excel = True				'added this in for clearing the SVES messages
			If instr(dail_msg, "APPLCT ID CHNGD") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "CASE AUTOMATICALLY DENIED") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") then 
				 add_to_excel = True
			ElseIf instr(dail_msg, "CASE NOTE ENTERED BY") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "CASE NOTE TRANSFER FROM") then 
			 	add_to_excel = True 		
			ElseIf instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") then 
			 	add_to_excel = True
			ElseIf instr(dail_msg, "CASE XFER") then 
			 	add_to_excel = True 		 		 	
			ElseIf instr(dail_msg, "DIRECT DEPOSIT STATUS") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") then 
			 	add_to_excel = True
			ElseIf instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") then 
				 add_to_excel = True
			ElseIf instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") then 
			 	add_to_excel = True	 
			ElseIf instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") then 
			 	add_to_excel = True 		
			ElseIf instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") then 
			 	add_to_excel = True
			ElseIf instr(dail_msg, "PMI MERGED AND AUTO CASE NOTE ENTERED") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") then 
			 	add_to_excel = True 			 		 	
			ElseIf instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "THIS CASE IS ERROR PRONE") then 
			 	add_to_excel = True 
			ElseIf instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") then 
				add_to_excel = True			 		 	
			ElseIf instr(dail_msg, "NEW ELIG RESULTS EXIST") then 
				add_to_excel = True 
			ElseIf instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "WARNING MESSAGES EXIST") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") then 
				add_to_excel = True	
			ElseIf instr(dail_msg, "MSA RECERT DUE - NOT AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "MSA HRF DUE - NOT AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE - REVIEW FOR MEDI UPDATES") then 
				add_to_excel = True			
			ElseIf instr(dail_msg, "NEW MSA ELIG AUTO") then
				add_to_excel = True
			ElseIf instr(dail_msg, "GA: NEW PERSONAL NEEDS STANDARD AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "SNAP: NEW VERSION AUTO") then 
				add_to_excel = True
			ElseIf instr(dail_msg, "GRH: NEW VERSION AUTO") then       
				add_to_excel = True
			Else	
			    add_to_excel = False 
			End if 
						
			IF add_to_excel = True then 
				EMReadScreen maxis_case_number, 8, dail_row - 1, 73
				EMReadScreen dail_month, 8, dail_row, 11
				'--------------------------------------------------------------------...and put that in Excel.
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = trim(maxis_case_number)
				objExcel.Cells(excel_row, 3).Value = trim(dail_type)
				objExcel.Cells(excel_row, 4).Value = trim(dail_month)
				objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
				excel_row = excel_row + 1
				
				'If (DAIL_type = "COLA" and case_note = True) then 
				'	Call write_value_and_transmit("N", dail_row, 3)	
				'	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
				'	If PRIV_check = "PRIV" then
				'		msgbox "PRIV case, clear case and navigate back to the message."
				'	Else 
				'	    PF9
				'	    CALL write_variable_in_case_note(dail_msg)
				'	    CALL write_variable_in_case_note("Case was auto approved due to COLA changes")
				'	    CALL write_variable_in_case_note("---")
				'	    CALL write_variable_in_case_note(worker_signature)
				'	    'Navigating back to DAIL/DAIL by pf3 2 times
				'	    PF3
				'	    PF3
				'	End If 
				'END IF
				
				Call write_value_and_transmit("D", dail_row, 3)	
				EMReadScreen other_worker_error, 13, 24, 2
				If other_worker_error = "** WARNING **" then transmit
				deleted_dails = deleted_dails + 1
			else
				add_to_excel = False
				dail_row = dail_row + 1
			End if
			
			EMReadScreen message_error, 17, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			If message_error = "NO MESSAGES TYPES" then 
				CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				Call write_value_and_transmit(worker, 21, 6)
				transmit   'transmit past 'not your dail message'
				Call dail_selection	
				exit do
			End if 
	    	
			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then 
				PF8
				EMReadScreen last_page_check, 21, 24, 2
				If last_page_check = "THIS IS THE LAST PAGE" then 
					all_done = true
					exit do 
				Else 
					dail_row = 6
				End if 
			End if
		LOOP
		IF all_done = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of " & dail_to_decimate & " messages reviewed"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! Please review the list created for accuracy.")