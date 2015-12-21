'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - PDED LIST GEN.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_cases_into_excel_dialog, 0, 0, 176, 105, "Pull cases into Excel dialog"
  EditBox 75, 10, 90, 15, x_number
  CheckBox 10, 65, 150, 10, "Check HERE to run for entire agency.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 65, 85, 50, 15
    CancelButton 115, 85, 50, 15
  Text 10, 15, 60, 10, "Worker to check:"
  Text 10, 45, 145, 10, "* For multiple, separate with comma."
  Text 10, 30, 145, 10, "* Enter 7-digit worker number ONLY."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

DIALOG pull_cases_into_excel_dialog
	IF ButtonPressed = 0 then stopscript

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True


'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
ObjExcel.Cells(1, 5).Value = "Pickle Disregard?"
ObjExcel.Cells(1, 6).Value = "DAC Disregard?"
ObjExcel.Cells(1, 7).Value = "Unearned Inc Disregard?"
ObjExcel.Cells(1, 8).Value = "Earned Inc Disregard?"
objExcel.Cells(1, 9).Value = "Guardianship Fee"
objExcel.Cells(1, 10).Value = "Rep Payee Fee"
objExcel.Cells(1, 11).Value = "Shel/Spec Need"

FOR i = 1 to 11
	objExcel.Cells(1, i).Font.Bold = True
	objExcel.Columns(i).AutoFit()
NEXT

'Setting the variable for what's to come
excel_row = 2

'If all workers are selected, the script will open the worker list stored on the shared drive, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = 1 then
	CALL create_array_of_all_active_x_numbers_in_county(x_array, right(worker_county_code, 2))
Else
	IF len(x_number) > 3 THEN 
		x_array = split(x_number, ", ")
	ELSE		
		x_array = split(x_number)
	END IF
End if

For each worker in x_array
	IF worker <> "" THEN
		Call navigate_to_screen("rept", "actv")
		IF worker <> "" THEN
			EMWriteScreen worker, 21, 13
			transmit
		END IF
		EMReadScreen user_id, 7, 21, 71
		EMReadScreen check_worker, 7, 21, 13
		IF user_id = check_worker THEN PF7
	
		'Grabbing each case number on screen
		Do
			MAXIS_row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen case_number, 8, MAXIS_row, 12
				If case_number = "        " then exit do
				EMReadScreen client_name, 21, MAXIS_row, 21
				EMReadScreen next_REVW_date, 8, MAXIS_row, 42
				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = case_number
				ObjExcel.Cells(excel_row, 3).Value = client_name
				ObjExcel.Cells(excel_row, 4).Value = replace(next_REVW_date, " ", "/")
				MAXIS_row = MAXIS_row + 1
				excel_row = excel_row + 1
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
next

'Resetting excel_row variable, now we need to start looking people up
excel_row = 2 

Do 
	case_number = ObjExcel.Cells(excel_row, 2).Value
	If case_number = "" then exit do
	call navigate_to_MAXIS_screen("STAT", "PDED")
	
	' >>>>> DETERMINING THE PEOPLE IN THE HOUSEHOLD FOR WHICH TO BE CHECKING THE PDED <<<<<
	pded_hh_array = ""
	pded_row = 5
	DO
		EMReadScreen pded_hh_memb, 2, pded_row, 3
		IF pded_hh_memb = "  " THEN
			EXIT DO
		ELSE
			pded_hh_array = pded_hh_array & pded_hh_memb & " "
			pded_row = pded_row + 1
		END IF
	LOOP UNTIL pded_hh_memb = "  "

	pded_hh_array = trim(pded_hh_array)
	pded_hh_array = split(pded_hh_array)

	FOR EACH hh_memb IN pded_hh_array
		IF hh_memb <> "" THEN 
			CALL write_value_and_transmit(hh_memb, 20, 76)
			EMReadScreen num_of_PDED, 1, 2, 78
			IF num_of_PDED <> "0" THEN 
			  ' >>>>> Reading the values <<<<<
				EMReadScreen pickle_disregard, 1, 6, 60
				EMReadScreen dac_disregard, 1, 8, 60
				EMReadScreen unea_inc_disregard, 8, 10, 62
				EMReadScreen earn_inc_disregard, 8, 11, 62
				EMReadScreen guard_fee, 8, 15, 44
				EMReadScreen payee_fee, 8, 15, 70
				EMReadScreen shel_spec_need, 1, 18, 78
				
				' >>>>> Populating Excel <<<<<
				IF pickle_disregard = "1" THEN 
					objExcel.Cells(excel_row, 5).Value = objExcel.Cells(excel_row, 5).Value & hh_memb & " - Pickle Elig; "
				ELSEIF pickle_disregard = "2" THEN 
					objExcel.Cells(excel_row, 5).Value = objExcel.Cells(excel_row, 5).Value & hh_memb & " - Potential Pickle Elig; "
				END IF
				
				IF dac_disregard = "Y" THEN objExcel.Cells(excel_row, 6).Value = objExcel.Cells(excel_row, 6).Value & hh_memb & " - YES; "
				
				unea_inc_disregard = replace(unea_inc_disregard, "_", "")
				unea_inc_disregard = trim(unea_inc_disregard)
				IF unea_inc_disregard <> "" THEN objExcel.Cells(excel_row, 7).Value = objExcel.Cells(excel_row, 7).Value & hh_memb & " (" & unea_inc_disregard & ")" & "; "
				
				earn_inc_disregard = replace(earn_inc_disregard, "_", "")
				earn_inc_disregard = trim(earn_inc_disregard)
				IF earn_inc_disregard <> "" THEN objExcel.Cells(excel_row, 8).Value = objExcel.Cells(excel_row, 8).Value & hh_memb & " (" & earn_inc_disregard & ")" & "; "
				
				guard_fee = replace(guard_fee, "_", "")
				guard_fee = trim(guard_fee)
				IF guard_fee <> "" THEN objExcel.Cells(excel_row, 9).Value = objExcel.Cells(excel_row, 9).Value & hh_memb & " (" & guard_fee & ")" & "; "
				
				payee_fee = replace(payee_fee, "_", "")
				payee_fee = trim(payee_fee)
				IF payee_fee <> "" THEN objExcel.Cells(excel_row, 10).Value = objExcel.Cells(excel_row, 10).Value & hh_memb & " (" & payee_fee & ")" & "; "
				
				IF shel_spec_need = "Y" THEN objExcel.Cells(excel_row, 11).Value = objExcel.Cells(excel_row, 11).Value & hh_memb & " - YES; "
			END IF
		END IF
	NEXT
	
	' >>>>> Deleting blank rows <<<<<
	IF objExcel.Cells(excel_row, 5).Value = "" AND _
		objExcel.Cells(excel_row, 6).Value = "" AND _
		objExcel.Cells(excel_row, 7).Value = "" AND _
		objExcel.Cells(excel_row, 8).Value = "" AND _
		objExcel.Cells(excel_row, 9).Value = "" AND _
		objExcel.Cells(excel_row, 10).Value = "" AND _
		objExcel.Cells(excel_row, 11).Value = "" THEN 
			SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			objRange.Delete
			excel_row = excel_row - 1
	END IF	
	
	excel_row = excel_row + 1
Loop until case_number = ""

FOR i = 1 to 11
	objExcel.Columns(i).AutoFit()
NEXT

'Logging usage stats
script_end_procedure("Success!!")
