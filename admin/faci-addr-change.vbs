'Required for statistical purposes===============================================================================
name_of_script = "BULK - FACI ADDR CHANGE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'FUNCTIONS that are currently not in the FuncLib that are used in this script----------------------------------------------------------------------------------------------------
Function File_Selection_System_Dialog(file_selected)
    'Creates a Windows Script Host object
    Set wShell=CreateObject("WScript.Shell")

    'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

    'Creates the file_selected variable from the exit
    file_selected = oExec.StdOut.ReadLine
End function

'-------THIS FUNCTION ALLOWS THE USER TO PICK AN EXCEL FILE---------
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("Shell.Application")
    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "Computer")
	IF file is Nothing THEN
		script_end_procedure("The script will end.")
	ELSE
		BrowseForFile = file.self.Path
	END IF
End Function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'function to call up a local/network file
DO
	DO
		'file_location = InputBox("Please enter the file location.")
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(BrowseForFile)
		objExcel.Visible = True
		objExcel.DisplayAlerts = True

		confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
		IF confirm_file = vbCancel THEN
			objWorkbook.Close
			objExcel.Quit
			stopscript
		ELSEIF confirm_file = vbNo THEN
			objWorkbook.Close
			objExcel.Quit
		END IF
	LOOP UNTIL confirm_file = vbYes
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'creating an array of cases to update
excel_row = 2
entry_record = 0
Do
	MAXIS_case_number = objExcel.cells(excel_row, 1).value
	MAXIS_case_number = trim(MAXIS_case_number)
	IF MAXIS_case_number = "" then
		exit do
	Else
		add_to_array = add_to_array & MAXIS_case_number & ","
		entry_record = entry_record + 1			'This increments to the next entry in the array'
	END IF
	excel_row = excel_row + 1
LOOP

If left(add_to_array, 1) = "," then add_to_array = right(add_to_array, len(add_to_array) - 1)
FACI_addr_case_array = Split(add_to_array, ",")

For each MAXIS_case_number in FACI_addr_case_array
	Do
		IF MAXIS_case_number = "" then exit do
		IF MAXIS_case_number <> "" THEN
			back_to_self
			EMWriteScreen "________", 18, 43					'clears case number
			EMWriteScreen MAXIS_case_number, 18, 43
			CALL navigate_to_MAXIS_screen("STAT", "ADDR")
			'Checking for privileged
			EMReadScreen privileged_case, 40, 24, 2
			IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
				privileged_array = privileged_array & MAXIS_case_number & "~~~"
			ELSE
				PF9
				EMReadScreen edit_check, 2, 24, 2
				If edit_check <> "  " then
					msgbox MAXIS_case_number & " cannot be updated. Log the case number and update manually."
					exit do
				ELSE
					EMReadScreen ADDR_panel_check, 4, 2, 44
					IF ADDR_panel_check = "ADDR" then
						Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

						Call access_ADDR_panel("WRITE", notes_on_address, "3231 1st Avenue", "", "", "Minneapolis", "MN", "55408", "27", "SF", "N", "N", "10", reservation_name, "3231 1st Avenue", "", "", "Minneapolis", "MN", "55408", "06/01/2016", addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

						DO
							Transmit
							EMReadScreen continue_check, 5, 24, 2
							If continue_check = "ENTER" then exit do
						LOOP until continue_check = "ENTER"

						PF3
						Transmit

						excel_col = excel_row + 1
						STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
					END IF
				END IF
			END IF
		END IF
		EMSendKey "<PF3>"
    	EMWaitReady 0, 0
    	EMReadScreen SELF_check, 4, 2, 50
	Loop until SELF_check = "SELF"
Next

IF privileged_array <> "" THEN
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a CASE NOTE for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
msgbox STATS_counter
script_end_procedure("Success! The addresses have been updated!")
