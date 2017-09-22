'This script is not complete. It was created to update a single FACI for the time being. 
'This will need to be enhanced in the future to hold the address as a variables, and not hard coded. 

'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("T:\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Required for statistical purposes===============================================================================
name_of_script = "BULK - SWKR CHANGE USING PMI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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
excel_row = 208
Do 
	'establishing variables for case to be updated
	client_PMI = objExcel.cells(excel_row, 2).value
	client_PMI = trim(client_PMI)
	first_name = objExcel.cells(excel_row, 3).value
	first_name = trim(first_name)
	last_name = objExcel.cells(excel_row, 4).value
	last_name = trim(last_name)
	phone_ext = objExcel.cells(excel_row, 5).value
	phone_ext = trim(phone_ext)
	
	IF client_PMI = "" then exit do
	'trims all the 0's off of the PMI number 
	Do 
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)
	Loop until left(client_PMI, 1) <> "0"
	
	back_to_self
	EMWriteScreen "________", 18, 43					'clears case number
	call navigate_to_MAXIS_screen("pers", "____")
	EMWriteScreen client_PMI, 15, 36
	Transmit
	EMReadscreen PMI_confirmation, 10, 8, 71
	PMI_confirmation = trim(PMI_confirmation)
	'msgbox PMI_confirmation
	If PMI_confirmation <> client_PMI then
		msgbox client_PMI & " does not match client. Process manually."
	Else 	
		EMWriteScreen "x", 8, 5		
		Transmit
		
		'chekcing for an active case
		MAXIS_row = 10
		Do 
			EMReadscreen open_case, 5, MAXIS_row, 53
			open_case = trim(open_case)
			If open_case = "" then
				EMReadscreen MAXIS_case_number, 8, MAXIS_row, 6
				MAXIS_case_number = trim(MAXIS_case_number) 
		 		EMWriteScreen "x", MAXIS_row, 4
				Transmit
				Exit do
			Else 
				MAXIS_row = MAXIS_row + 1
			END IF 
		LOOP until MAXIS_row = 19
		If MAXIS_row = 19 then msgbox "Unable to find an open case for " & client_PMI & vbnewline & "excel row: " & excel_row
		
		'navigating to the SWKR panel, and ensuring that we are in that panel
		back_to_self
		EMWriteScreen MAXIS_case_number, 18, 43
		Call navigate_to_MAXIS_screen("STAT", "SWKR")
		'msgbox "what's happening?"
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN 
			msgbox "PRIV case " & client_PMI & vbnewline & "Excel row: " & excel_row
			privileged_array = privileged_array & client_PMI & "~~~"
			EMWriteScreen "________", 18, 43
			Transmit
		ELSE
			Do 
				EMReadscreen SWKR_panel_check, 4, 2, 47
				If SWKR_panel_check <> "SWKR" then call write_value_and_transmit("SWKR", 20, 71)
			Loop until SWKR_panel_check = "SWKR"
				
			EMReadScreen needs_SWKR_panel, 1, 2, 73
			If needs_SWKR_panel = "0" then 
				call write_value_and_transmit("NN", 20, 79)
			ELSE 
				PF9
			END IF 
				
			EMReadScreen edit_check, 2, 24, 2
			If edit_check <> "  " then 
				msgbox client_PMI & " at excel row: " & excel_row & " cannot be updated. Log the PMI number and update manually."
			ELSE 
				'Clears SWKR address
				EMWriteScreen "___________________________________", 6, 32	
				EMWriteScreen "______________________", 8, 32
				EMWriteScreen "______________________", 9, 32
				EMWriteScreen "_______________", 10, 32
				EMWriteScreen "__", 10, 54
				EMWriteScreen "_______", 10, 63
				'Clears SWKR phone number
				EMWriteScreen "___", 12, 34
				EMWriteScreen "___", 12, 40
				EMWriteScreen "___", 12, 44
				EMWriteScreen "____", 12, 54
				
				'Writes in the  SWKR address
				EMWriteScreen first_name & " " & last_name & "/People Inc.", 6, 32	
				EMWriteScreen "1170 15th Ave SE", 8, 32
				EMWriteScreen "Minneapolis", 10, 32
				EMWriteScreen "MN", 10, 54
				EMWriteScreen "55414", 10, 63
				'Clears SWKR phone number
				EMWriteScreen "612", 12, 34
				EMWriteScreen "230", 12, 40
				EMWriteScreen "6270", 12, 44
				EMWriteScreen phone_ext, 12, 54
				EMWriteScreen "Y", 15, 63			'coding notices to be sent to SWKR
				'msgbox "confirm case: " & client_PMI & " at excel row " & excel_row
				Transmit
				Transmit
				transmit
				PF3
			END IF
		END if 
	END IF
	back_to_self
	excel_row = excel_row + 1
	client_PMI = ""
	MAXIS_case_number = ""
	STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
Loop until excel_row = 280

IF privileged_array <> "" THEN 
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a CASE NOTE for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
msgbox STATS_counter
script_end_procedure("Success! The addresses have been updated!")
