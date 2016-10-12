'Required for statistical purposes===============================================================================
name_of_script = "DAIL - FIN ORIENT MISSING.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 95          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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
'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""
EMReadScreen MAXIS_case_number, 8, 5, 73		'Getting the case number from the DAIL
MAXIS_case_number = trim(MAXIS_case_number)

EMReadScreen name_for_dail, 57, 5, 5			'Reading the name of the client
'This next block will determine the name of the client the message is for
'If the message is for someone other than M01 - the name is writen next to the name of M01
other_person = InStr(name_for_dail, "--(")	'This determines if it for someone other than M01
'This is for if the message is for M01'
If other_person = 0 Then 
	comma_loc = InStr(name_for_dail, ",")  	'Determines the end of the last name
	dash_loc = InStr(name_for_dail, "-")	'Determines the end of the name
	EMReadscreen last_name, comma_loc - 1, 5, 5									'Reading the last name
	EMReadscreen middle_exists, 1, 5, 5 + (dash_loc - 2)						'Determines if clt's middle initial is listed
	If middle_exists = " " Then 												'If not - reads first name
		EMReadscreen first_name, dash_loc - comma_loc - 5, 5, comma_loc + 5
	Else 																		'If so - reads first name
		EMReadScreen first_name, dash_loc - comma_loc - 3, 5, comma_loc + 5
	End If 
'This is for if the message is for a different HH Member
Else 
	end_other = InStr(name_for_dail, ")--")
	comma_loc = InStr(other_person, name_for_dail, ",")
	EMReadscreen last_name, comma_loc - other_person - 3, 5, other_person + 7
	EMReadscreen middle_exists, 1, 5, end_other + 2
	If middle_exists = " " Then 
		EMReadscreen first_name, end_other - comma_loc - 3, 5, comma_loc + 5
	Else 
		EMReadScreen first_name, end_other - comma_loc - 1, 5, comma_loc + 5
	End If 
End If 
client_name = last_name & ", " & first_name		'putting the name into one string

'Goes to INFC/WORK
EMSendKey "i"
transmit

EMSendKey "work"
transmit

EMReadScreen work_panel_check, 4, 2, 51
If work_panel_check = "WORK" Then 
work_maxis_row = 7
	DO
		EMReadScreen work_name, 26, work_maxis_row, 7			'Reads the client name from INFC/WORK'
		work_name = trim(work_name)
		IF client_name = work_name then 
			memb_check = vbYes		'If the name on INFC/WORK exactly matches the name from the DAIL, the script does not need user input and will gather the referrence number
			EMReadScreen ref_numb, 2, work_maxis_row, 3
		ElseIf client_name <> work_name then 	'if name doesn't match the DAIL name the confirmation is required by the user
			memb_check = MsgBox ("DAIL Message is for - " & client_name & vbNewLine & "Name on INFC/WORK - " & work_name & _ 
			  vbNewLine & vbNewLine & "Is this the client you need ES Referral Information about?", vbYesNo + vbQuestion, "Confirm Client using Banked Monhts")
			If memb_check = vbYes Then		'If the user confirms that this is the correct client, the Ref number are gathered'
				EMReadScreen ref_numb, 2, work_maxis_row, 3
			ElseIf memb_check = vbNo Then	'If the user says NO the script will see if there are other clients listed on INFC/WORK and start back at the beginning of the loop to try to match'
				EMReadScreen next_clt, 1, (work_maxis_row + 1), 7
			END IF
		End If 
		work_maxis_row = work_maxis_row + 1		'Increments to read the next row for a new client'
	Loop until next_clt = " " OR memb_check = vbYes	

	'If a match was found, the script will attempt to get an appointment date
	If memb_check = vbYes Then EMReadScreen es_appt_date, 8, 7, 72
	If es_appt_date = "__ __ __" Then es_appt_date = ""
	es_appt_date = replace(es_ref_date, " ", "/")
End If 

PF3			'Back to Dail

'If the script could not find the reference number on WORK - it will go to MEMB to find it.
If ref_numb = "" Then 
	EMSendKey "s"
	transmit

	EMSendKey "memb"
	transmit
	Do 
		EMReadScreen memb_ln, len(last_name), 6, 30				'Reads the name on the MEMB panel
		EMReadScreen memb_fn, len(first_name), 6, 63
		IF memb_fn = first_name AND memb_ln = last_name Then 	'Compares them to the name found on DAIL
			EMReadScreen ref_numb, 2, 4, 33						'If they match, then it reads the reference number
			Exit Do
		End If 
		transmit
		EMReadScreen last_memb, 5, 24, 2 						'Loops until it goes through all the members
	Loop Until last_memb = "ENTER"
End If 
		
If fin_orient_date = "" Then 		'If the user has not already defined the financial orientation date, the script will look in another place
	EMSendKey "s"					'Goes to STAT PROG
	transmit
	
	EMSendKey "prog"
	transmit

	EMReadScreen cash_actv_check, 4, 6, 74			'Finds an interview date for the active cash program
	If cash_actv_check = "ACTV" Then
		EMReadScreen cash_intv_date, 8, 6, 55
		If cash_intv_date = "__ __ __" Then cash_intv_date = ""
		cash_intv_date = replace(cash_intv_date, " ", "/")
	Else 
		EMReadScreen cash_actv_check, 4, 7, 74
		If cash_actv_check = "ACTV" Then
			EMReadScreen cash_intv_date, 8, 6, 55
			If cash_intv_date = "__ __ __" Then cash_intv_date = ""
			cash_intv_date = replace(cash_intv_date, " ", "/")
		End If 
	End If 
	
	PF3		'Back to DAIL
End If 

If fin_orient_date = "" AND ref_numb <> "" Then 		'If no date identified, script tries one more time
	EMSendKey "s"				'Goes to STAT EMPS for this client
	transmit
	
	EMSendKey "emps"
	transmit
	
	EMWriteScreen ref_numb, 20, 76
	transmit

	EMReadScreen es_ref_date, 8, 16, 40				'Finds the ES Referral date
	If es_ref_date = "__ __ __" Then es_ref_date = ""
	es_ref_date = replace(es_ref_date, " ", "/")
	
End If 

'Defining some coordinates for dynamic dialog
dlg_height = 165
grp_height = 95

If es_appt_date = "" Then 
	dlg_height =  dlg_height - 15
	grp_height = grp_height - 15
End If 
If cash_intv_date = "" Then 
	dlg_height =  dlg_height - 15
	grp_height = grp_height - 15
End If 
If es_ref_date = "" Then 
	dlg_height =  dlg_height - 15
	grp_height = grp_height - 15
End If 
y_pos = 55

'Dialog is defined here so it can be dynamic
BeginDialog FIN_ori_dialog, 0, 0, 300, dlg_height, "Financial Orientation dialog"
  EditBox 65, 5, 120, 15, client_name
  EditBox 265, 5, 25, 15, ref_numb
  Text 15, 40, 265, 10, "Check one of the following to have the script update EMPS with the selected date"
  If es_appt_date   <> "" Then 
  	CheckBox 15, y_pos, 175, 10, es_appt_date & " from INFC/WORK - ES Appointment Date", es_appt_checkbox
	y_pos = y_pos + 15
  End If 
  If cash_intv_date <> "" Then 
  	CheckBox 15, y_pos, 175, 10, cash_intv_date & " from STAT/PROG - Cash Interview Date", intv_date_checkbox
	y_pos = y_pos + 15
  End If 
  If es_ref_date    <> "" Then 
  	CheckBox 15, y_pos, 175, 10, es_ref_date & " from STAT/EMPS - ES Referral Date", es_ref_date_checkbox
	y_pos = y_pos + 15
  End If 
  y_pos = y_pos + 5
  CheckBox 15, y_pos, 30, 10, "Other", other_date_checkbox
  EditBox 85, y_pos - 5, 50, 15, fin_orient_date
  EditBox 180, y_pos - 5, 100, 15, other_source
  EditBox 50, y_pos + 20, 240, 15, other_notes
  EditBox 70, y_pos + 40, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, y_pos + 40, 50, 15
    CancelButton 240, y_pos + 40, 50, 15
  Text 5, 10, 60, 10, "Name from DAIL:"
  Text 195, 10, 70, 10, "HH member number:"
  GroupBox 5, 25, 285, grp_height, "Select Financial Orientation Date"
  Text 5, y_pos + 25, 40, 10, "Other notes:"
  Text 55, y_pos, 20, 10, "Date:"
  Text 150, y_pos, 30, 10, "Source:"
  Text 5, y_pos + 45, 60, 10, "Worker signature:"
EndDialog

'Running the dialog to get user information
Do 
	err_msg = ""
	Dialog FIN_ori_dialog
	cancel_confirmation
	If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Sign your case note."
	If es_appt_checkbox = checked Then 
		If intv_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and the Cash Interview date - you must select only one."
		If es_ref_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and the ES Referral date - you must select only one."
		If other_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and another date - you must select only one."
		If ref_numb = "" Then err_msg = err_msg & vbNewLine & "You must enter the client's reference number in order for the EMPS panel to be correctly updated."
	End If 
	If intv_date_checkbox = checked Then 
		If es_appt_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and the Cash Interview date - you must select only one."
		If es_ref_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the Cash Interview date and the ES Referral date - you must select only one."
		If other_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the Cash Interview date and another date - you must select only one."
		If ref_numb = "" Then err_msg = err_msg & vbNewLine & "You must enter the client's reference number in order for the EMPS panel to be correctly updated."
	End If 
	If es_ref_date_checkbox = checked Then 
		If intv_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Referral date and the Cash Interview date - you must select only one."
		If es_appt_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and the ES Referral date - you must select only one."
		If other_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Referral date and another date - you must select only one."
		If ref_numb = "" Then err_msg = err_msg & vbNewLine & "You must enter the client's reference number in order for the EMPS panel to be correctly updated."
	End If
	If other_date_checkbox = checked Then 
		If intv_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the Cash Interview date and another date - you must select only one."
		If es_ref_date_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Referral date and another date - you must select only one."
		If es_appt_checkbox = Checked Then err_msg = err_msg & vbNewLine &"You have selected to use the ES Appointment date and another date - you must select only one."
		If fin_orient_date = "" Then err_msg = err_msg & vbNewLine & "You have selected to choose another date but have not listed the other date, please list it next to the 'Other' checkbox."
		If other_source = "" Then err_msg = err_msg & vbNewLine & "When selecting another date, provide information for case note regarding why that date was selected"
		If ref_numb = "" Then err_msg = err_msg & vbNewLine & "You must enter the client's reference number in order for the EMPS panel to be correctly updated."
		If isdate(fin_orient_date) = FALSE Then err_msg = err_msg & vbNewLine & "You must enter a valid date for the Financial Orientation Date."
	End If 	
	If err_msg <> "" Then MsgBox "Please resolve before you continue." & vbNewLine & err_msg
Loop until err_msg = ""

'Setting the date and reason for case noting and action
If es_appt_checkbox = checked Then 
	fin_orient_date = es_appt_date
	fo_source = "INFC/WORK - ES Appointment Date"
ElseIf intv_date_checkbox = checked Then 
	fin_orient_date = cash_intv_date
	fo_source = "STAT/PROG - Cash Interview Date"
ElseIf es_ref_date_checkbox = checked Then
	fin_orient_date = es_ref_date
	fo_source = "STAT/EMPS - ES Referral Date"
ElseIf other_date_checkbox = checked Then 
	fo_source = other_source
End If 

'Updating EMPS
If fin_orient_date <> "" Then 
	Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
	EMWriteScreen ref_numb, 20, 76
	transmit
	PF9
	ref_month = right("00" & DatePart("m", fin_orient_date), 2)
	ref_date  = right("00" & DatePart("d", fin_orient_date), 2)
	ref_year  = right(DatePart("yyyy", fin_orient_date), 2)
	EMWriteScreen ref_month, 5, 39
	EMWriteScreen ref_date,  5, 42
	EMWriteScreen ref_year,  5, 45
	EMWriteScreen "y", 5, 65
	transmit
	msgbox "Look"
	PF3

	'Check to make sure we are back to our dail 
	EMReadScreen DAIL_check, 4, 2, 48 
	IF DAIL_check <> "DAIL" THEN 
		PF3 'This should bring us back from UNEA or other screens 
		EMReadScreen DAIL_check, 4, 2, 48 
		IF DAIL_check <> "DAIL" THEN 'If we are still not at the dail, try to get there using custom function, this should result in being on the correct dail (but not 100%) 
			call navigate_to_MAXIS_screen("DAIL", "DAIL") 
		END IF 
	END IF 

	'Writing the case note
	EMWriteScreen "n", 6, 3 
	transmit 

	PF9 
	EMReadScreen case_note_mode_check, 7, 20, 3 
	If case_note_mode_check <> "Mode: A" then MsgBox "You are not in a case note on edit mode. You might be in inquiry. Try the script again in production." 
	If case_note_mode_check <> "Mode: A" then stopscript

	Call Write_Variable_in_CASE_NOTE ("DAIL Processed - Financial Orientation Date Updated for Memb " & ref_numb)
	Call Write_Variable_in_CASE_NOTE ("* PEPR message rec'vd indicating that Financial Orientation Date is Missing")
	Call Write_Bullet_and_Variable_in_Case_Note ("Date Entered", fin_orient_date)
	Call Write_Bullet_and_Variable_in_Case_Note ("Date information from", fo_source)
	Call Write_Bullet_and_Variable_in_Case_Note ("Notes", other_notes)
	Call Write_Variable_in_CASE_NOTE ("---")
	Call Write_Variable_in_CASE_NOTE (worker_signature) 
	end_msg = "Success! EMPS has been updated and Case Note Written"
Else 
	end_msg = "You have selected to not have the EMPS panel updated by the script." & vbNewLine & "You will need to process this DAIL manually."

End If 

script_end_procedure(end_msg)
