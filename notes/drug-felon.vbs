'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DRUG FELON.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 150           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'FUNCTIONS-------------------------------------------------------------------------------------------------------------------------------
Function Generate_Client_List(list_for_dropdown)

	memb_row = 5

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do 
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do 
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20
		
	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next

End Function

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/29/2016", "Adding functionality to case note the return of required documenttion when a drug felon match is reported.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-------------------------------------------------------------------------------------------------------------------------------
BeginDialog dfln_case_number_dialog, 0, 0, 186, 100, "Case Number and Information"
  EditBox 75, 5, 70, 15, MAXIS_case_number
  DropListBox 75, 30, 105, 45, "Select one..."+chr(9)+"Active Family Cash"+chr(9)+"Active Adult Cash"+chr(9)+"Active but NO Cash"+chr(9)+"Closed", case_status_dropdown
  DropListBox 75, 55, 105, 45, "Select one..."+chr(9)+"Initial Information Received"+chr(9)+"Initial Information Not Received"+chr(9)+"Testing Follow Up", action_dropdown
  ButtonGroup ButtonPressed
    OkButton 75, 80, 50, 15
    CancelButton 130, 80, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 35, 45, 10, "Case Status:"
  Text 10, 60, 60, 10, "Action to Process:"
EndDialog

BeginDialog dfln_testing_dialog, 0, 0, 271, 175, "Drug Felon Testing"
  EditBox 65, 5, 60, 15, conviction_date
  EditBox 65, 25, 135, 15, probation_officer
  CheckBox 10, 45, 145, 10, "Check here if the authorization is on file:", authorization_on_file_check
  CheckBox 10, 60, 130, 10, "Check here if client complied with UA:", complied_with_UA_check
  EditBox 40, 75, 45, 15, UA_date
  DropListBox 145, 75, 65, 15, "select one..."+chr(9)+"Positive"+chr(9)+"Negative"+chr(9)+"Refused", UA_results
  EditBox 75, 95, 55, 15, date_of_1st_offense
  EditBox 210, 95, 55, 15, date_of_2nd_offense
  EditBox 60, 115, 205, 15, actions_taken
  EditBox 80, 135, 185, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 155, 50, 15
    CancelButton 215, 155, 50, 15
  Text 140, 100, 70, 10, "Date of 2nd Offense:"
  Text 5, 10, 55, 10, "Conviction Date:"
  Text 5, 120, 50, 10, "Actions Taken:"
  Text 100, 80, 40, 10, "UA Results:"
  Text 5, 140, 70, 15, "Sign your Case Note:"
  Text 5, 100, 65, 10, "Date of 1st Offense:"
  Text 5, 30, 60, 10, "Probation Officer:"
  Text 5, 80, 30, 10, "UA Date:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabbing case number & month/year
EMConnect "" 
CALL MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Setting some variables for easier to read code
PACT_Updated = FALSE
SNAP_Active = FALSE
Case_inactive = FALSE
Family_case = FALSE
Adult_case = FALSE

'If the script was able to get the case number, it is going to attempt to find the active programs to predetermine the actions needed
If MAXIS_case_number <> "" Then 
	navigate_to_MAXIS_screen "CASE", "CURR"
	
	Dim search_fields_array (5)			'Creates an array of different progeam options to loop through
	search_fields_array(0) = "Case:"
	search_fields_array(1) = "MFIP:"
	search_fields_array(2) = "DWP:"
	search_fields_array(3) = "GA:"
	search_fields_array(4) = "MSA:"
	search_fields_array(5) = "FS:"
	For each program in search_fields_array		'this will now loop through each of the program options and set a boolean based on the information found.
		prog_status = ""						'clearin the variable
		row = 1
		col = 1
		search = program
		EMSearch search, row, col
		If row <> 0 Then 						'If the search finds that program type on case curr - it will read the status associated with it
			EMReadScreen prog_status, 9, row, 9
		End If 
		prog_status = trim(prog_status)			'Now it set the case types based on the programs and status
		If program = "Case:" AND prog_status = "INACTIVE" Then Case_inactive = TRUE
		If program = "MFIP:" OR program = "DWP:" Then 
		 	If prog_status = "ACTIVE" Then Family_case = TRUE
		End If
		If program = "GA:" OR program = "MSA:" Then 
			If prog_status = "ACTIVE" Then Adult_case = TRUE
		End If
		If program = "FS:" AND prog_status = "ACTIVE" Then SNAP_Active = TRUE
	Next 
		
	'Next the script will use the case type booleans to preselect the case status dropdown for the initial dialog.
	If Case_inactive = TRUE Then 
		case_status_dropdown = "Closed"
	Else 
		If Family_case = FALSE AND Adult_case = FALSE Then case_status_dropdown = "Active but NO Cash"
	End If
	If Family_case = TRUE Then case_status_dropdown = "Active Family Cash"
	If Adult_case = TRUE Then case_status_dropdown = "Active Adult Cash"
End If 

'Running the initial dialog to confirm what type of DFLN note is needed and the specifics about the case
Do
	err_msg = ""
	Dialog dfln_case_number_dialog
	If Buttonpressed = cancel then StopScript
	IF IsNumeric(MAXIS_case_number)= FALSE THEN err_msg = err_msg & vbNewLine & "You must type a valid numeric case number."
	If case_status_dropdown = "Select one..." Then err_msg = err_msg & vbNewLine & "Indicate what the case status is."
	If action_dropdown = "Select one..." Then err_msg = err_msg & vbNewLine & "Chose what process you are noting."
	If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

'Confirming SNAP status
navigate_to_MAXIS_screen "STAT", "PROG"
EMReadScreen snap_status, 4, 10, 74
IF snap_status = "ACTV" Then SNAP_Active = TRUE

'This uses the function above to create a dropdown of all the clients in the HH for the worker to Select
'This script can only run for 1 person at a time, so the checkbox option does not work.
Call Generate_Client_List(HH_Memb_DropDown)

'There are 3 types of actions the worker could have selected, each with their own process and dialog. This wil run the one the worker specified.
Select Case action_dropdown

'This is for when a client has submitted the proofs needed after a DFLN match has been identified
Case "Initial Information Received"

	'Dialog is defined here as the HH dropdown needs to be defined before the dialog is
	BeginDialog info_recvd_dialog, 0, 0, 191, 105, "Update FSS Information from the Status Update"
	  DropListBox 80, 5, 105, 45, "Select One..." & HH_Memb_DropDown, clt_to_update
	  ComboBox 85, 25, 100, 45, ""+chr(9)+"Assesed as not needing drug treatment."+chr(9)+"Currently in drug treatment."+chr(9)+"Successful completion of drug treatment.", docs_dropdown
	  EditBox 75, 45, 110, 15, more_notes
	  EditBox 75, 65, 110, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 115, 85, 35, 15
	    CancelButton 155, 85, 30, 15
	  Text 5, 10, 70, 10, "Household member:"
	  Text 5, 30, 75, 10, "Information Received:"
	  Text 5, 50, 55, 10, "Additional Notes:"
	  Text 5, 70, 60, 10, "Worker Signature:"
	EndDialog

	'Runs the dialog
	Do
		err_msg = ""
		Dialog info_recvd_dialog
		cancel_confirmation
		If clt_to_update = "Select One..." Then err_msg = err_msg & vbNewLine & "Please pick the client you are processing DFLN information for."
		If docs_dropdown = "" Then err_msg = err_msg & vbNewLine & "Please list which type of document was received, if none of the listed documents match, type your own."
		IF case_status_dropdown = "Active Adult Cash" Then
			If docs_dropdown <> "Assesed as not needing drug treatment." AND docs_dropdown <> "Currently in drug treatment." AND docs_dropdown <> "Successful completion of drug treatment." Then
				err_msg = err_msg & vbNewLine & "For adult cash the only acceptable responses are the three preselected in the list, if the client has not provided one of these, they have not completed the requirement."
			End If 
		End If 
		If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = "" 
		
	clt_ref_num = left(clt_to_update, 2)	'Settin the reference number
	
	'Family cash cases with DFLN are subject to vendoring. This Updates PACT for vendoring
	If case_status_dropdown = "Active Family Cash" Then 
		navigate_to_MAXIS_screen "STAT", "PROG"
		EMReadScreen cash1, 4, 6, 74
		EMReadScreen cash2, 4, 7, 74
		IF cash1 = "ACTV" OR cash2 = "ACTV" Then 
			PACT_Updated = TRUE
			navigate_to_MAXIS_screen "STAT", "PACT"
			EMReadScreen pact_version, 1, 2, 73
			If pact_version = "1" Then
				PF9
			Else 
				EMWriteScreen "NN", 20, 79
				transmit
			End If 
			IF cash1 = "ACTV" Then EMWriteScreen "7", 6, 74
			If cash2 = "ACTV" Then EMWriteScreen "7", 8, 74
			transmit
		End If 
	End If 
	
	'case noting
	start_a_blank_CASE_NOTE
	
	CALL write_variable_in_case_note("***Drug Felon***")
	CALL write_variable_in_case_note("* MEMB " & clt_ref_num & " has cooperated with Drug Felon Notice.")
	Call write_bullet_and_variable_in_case_note ("Client has reported/supplied verification", docs_dropdown)
	If PACT_Updated = TRUE Then Call write_variable_in_case_note ("* Updated PACT for vendoring.")
	CALL write_bullet_and_variable_in_case_note ("Notes", more_notes)
	CALL write_variable_in_case_note ("---")
	CALL write_variable_in_case_note (worker_signature)

'This is for when documentation about follow up has been requested but client failed to provide it within 10 days
'This has no actions associated with it as no process was provided at this time. This is a great place for an enhancement
Case "Initial Information Not Received"

	'Dialog is defined here as the HH dropdown needs to be defined before the dialog is
	BeginDialog info_fail_dialog, 0, 0, 191, 105, "Update FSS Information from the Status Update"
	  DropListBox 80, 5, 105, 45, "Select One..." & HH_Memb_DropDown, clt_to_update
	  EditBox 75, 25, 110, 15, action_taken
	  EditBox 75, 45, 110, 15, more_notes
	  EditBox 75, 65, 110, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 115, 85, 35, 15
	    CancelButton 155, 85, 30, 15
	  Text 5, 10, 70, 10, "Household member:"
	  Text 5, 30, 55, 10, "Action Taken"
	  Text 5, 50, 55, 10, "Additional Notes:"
	  Text 5, 70, 60, 10, "Worker Signature:"
	EndDialog
	
	'Running the dialog
	Do
		err_msg = ""
		Dialog info_fail_dialog
		cancel_confirmation
		If clt_to_update = "Select One..." Then err_msg = err_msg & vbNewLine & "Please pick the client you are processing DFLN information for."
		If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = "" 
	
	clt_ref_num = left(clt_to_update, 2)	'Setting the reference number
	
	'Checks MAXIS for password prompt
	Call check_for_MAXIS(FALSE)

	'Writes the case note
	start_a_blank_CASE_NOTE
	
	CALL write_variable_in_case_note("***Drug Felon***")
	CALL write_variable_in_case_note("* MEMB " & clt_ref_num & " has NOT cooperated with Drug Felon Notice.")
	Call write_bullet_and_variable_in_case_note ("Action Taken", action_taken)
	CALL write_bullet_and_variable_in_case_note ("Notes", more_notes)
	CALL write_variable_in_case_note ("---")
	CALL write_variable_in_case_note (worker_signature)
	

Case "Testing Follow Up"

	'Autofilling the conviction date if script can find it
	navigate_to_MAXIS_screen "STAT", "DFLN"
	EMReadScreen convc_dt, 8, 6, 27
	If convc_dt <> "__ __ __" Then 
		convc_dt = replace(convc_dt, " ", "/")
		conviction_date = convc_dt
	End If
	
	DO
		err_msg = ""
		Dialog dfln_testing_dialog
		cancel_confirmation
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "You must sign your case note"
		IF IsNumeric(MAXIS_case_number)= FALSE THEN err_msg = err_msg & vbNewLine & "You must type a valid numeric case number."
		If UA_results = "select one..." THEN err_msg = err_msg & vbNewLine & "You must select 'UA results field'"
		If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""

	'Checks MAXIS for password prompt
	Call check_for_MAXIS(FALSE)

	'Writes the case note
	start_a_blank_CASE_NOTE
	CALL write_variable_in_case_note("***Drug Felon***")
	CALL write_bullet_and_variable_in_case_note("Conviction date", conviction_date)
	CALL write_bullet_and_variable_in_case_note("Probation Officer", po_officer)
	IF authorization_on_file_check = checked THEN CALL write_variable_in_case_note("* Authorization on file.")
	IF complied_with_UA_check = checked THEN CALL write_variable_in_case_note("* Complied with UA.")
	CALL write_bullet_and_variable_in_case_note("UA Date", UA_date)
	CALL write_bullet_and_variable_in_case_note("Date of 1st offence", date_of_1st_offense)
	CALL write_bullet_and_variable_in_case_note("Date of 2nd offence", date_of_2nd_offense)
	IF UA_results <> "select one..." THEN CALL write_bullet_and_variable_in_case_note("UA results", UA_results)
	CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
	CALL write_variable_in_case_note("---")
	CALL write_variable_in_case_note(worker_signature)
	
End Select

script_end_procedure("")