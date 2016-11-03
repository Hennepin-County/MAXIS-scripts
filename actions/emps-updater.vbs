'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - UPDATE EMPS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 458                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
'Base Dialog 
BeginDialog fss_code_detail, 0, 0, 371, 130, "Update FSS Information from the Status Update"
  Text 5, 10, 45, 10, "Case Number"
  EditBox 55, 5, 50, 15, MAXIS_case_number
  Text 120, 10, 70, 10, "Household member"
  DropListBox 195, 5, 105, 45, "", List2
  ButtonGroup ButtonPressed
    PushButton 315, 5, 50, 10, "Enter", enter_detail_button
  Text 5, 25, 195, 10, "This script can update EMPS for the following proceedures:"
  ButtonGroup ButtonPressed
    PushButton 10, 40, 185, 10, "Code EMPS to get MFIP results instead of DWP", Intake_MFIP_Button
    PushButton 10, 55, 185, 10, "Code EMPS for Child Under 12 Months Exemption", Child_Under_One_Button
    PushButton 10, 70, 185, 10, "Code EMPS to remove FSS", Remove_FSS_Button
  Text 205, 40, 105, 10, "Workaround process for Intake"
  Text 205, 55, 75, 10, "Adding or removing"
  Text 205, 70, 125, 10, "Return Caregiver to Regular MFIP-ES"
  Text 10, 95, 40, 10, "Other Notes"
  EditBox 55, 90, 310, 15, other_notes
  Text 10, 115, 60, 10, "Worker Signature"
  EditBox 80, 110, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 260, 110, 50, 15
    CancelButton 315, 110, 50, 15
EndDialog



'FULLY EXPANDED'
BeginDialog fss_code_detail, 0, 0, 371, 270, "Update FSS Information from the Status Update"
  Text 5, 10, 45, 10, "Case Number"
  EditBox 55, 5, 50, 15, MAXIS_case_number
  Text 120, 10, 70, 10, "Household member"
  DropListBox 195, 5, 105, 45, "", List2
  ButtonGroup ButtonPressed
    PushButton 315, 5, 50, 10, "Enter", enter_detail_button
  Text 5, 65, 195, 10, "This script can update EMPS for the following proceedures:"
  ButtonGroup ButtonPressed
    PushButton 10, 80, 185, 10, "Code EMPS to get MFIP results instead of DWP", Intake_MFIP_Button
    PushButton 10, 115, 185, 10, "Code EMPS for Child Under 12 Months Exemption", Child_Under_One_Button
    PushButton 10, 175, 185, 10, "Code EMPS to remove FSS", Remove_FSS_Button
  Text 205, 80, 105, 10, "Workaround process for Intake"
  Text 205, 115, 75, 10, "Adding or removing"
  Text 205, 175, 125, 10, "Return Caregiver to Regular MFIP-ES"
  Text 5, 235, 40, 10, "Other Notes"
  EditBox 50, 230, 310, 15, other_notes
  Text 5, 255, 60, 10, "Worker Signature"
  EditBox 75, 250, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 250, 50, 15
    CancelButton 310, 250, 50, 15
  Text 10, 100, 65, 10, "Date of Application"
  EditBox 80, 95, 50, 15, date_of_app
  Text 10, 135, 205, 10, "It appears you need to ADD/REMOVE the exemption.  Reason:"
  DropListBox 220, 130, 125, 45, "Select One..."+chr(9)+"Child Age"+chr(9)+"Caregiver request"+chr(9)+"MFIP results approve - complete workaround", List3
  Text 10, 155, 75, 10, "First month to remove"
  EditBox 90, 150, 15, 15, end_month
  EditBox 105, 150, 15, 15, End_Year
  Text 135, 155, 75, 10, "Date of Client request:"
  EditBox 220, 150, 50, 15, client_request_date
  Text 5, 30, 95, 10, "ES Referral Date is Missing:"
  Text 5, 50, 95, 10, "Fin Orient Date is Missing:"
  CheckBox 110, 30, 140, 10, "Check Here to have the script update to ", update_ES_ref_checkbox
  CheckBox 110, 50, 140, 10, "Check Here to have the script update to ", update_fin_orient_checkbox
  EditBox 255, 25, 50, 15, new_es_referral_dt
  EditBox 255, 45, 50, 15, new_fin_oreient_dt
EndDialog

'FUNCTIONS==================================================================================================================
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

'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

UniversalParticipant = FALSE 
ExtensionCase = FALSE 
FSSCase = FALSE 

If MAXIS_case_number <> "" Then 
	navigate_to_MAXIS_screen "STAT", "EMPS"
	EMWriteScreen HH_member, 20, 76
	transmit
	EMReadScreen Fin_Orient_Dt, 8, 5, 39
	EMReadScreen ES_Referral_dt, 8, 16, 40
	If Fin_Orient_Dt = "__ __ __" then 
		Fin_Orient_Dt = ""
		Fin_Orient_Missing = TRUE 
	Else 
		Fin_Orient_Dt = replace(Fin_Orient_Dt, " ", "/")
	End If 
	If ES_Referral_Dt = "__ __ __" Then 
		ES_Referral_Dt = ""
		ES_referral_Missing = TRUE 
	Else 
		ES_Referral_Dt = replace(ES_Referral_Dt, " ", "/")
	End If 
End If 

Call Generate_Client_List(HH_Memb_DropDown)

Do
	err_msg = ""
	'Dialog defined here so the dropdown can be changed
	BeginDialog select_person_dialog, 0, 0, 191, 65, "Update FSS Information from the Status Update"
	  EditBox 55, 5, 50, 15, MAXIS_case_number
	  ButtonGroup ButtonPressed
	    PushButton 135, 10, 50, 10, "search", search_button
	  DropListBox 80, 25, 105, 45, "Select One..." & HH_Memb_DropDown, clt_to_update
	  ButtonGroup ButtonPressed
	    OkButton 115, 45, 35, 15
	    CancelButton 155, 45, 30, 15
	  Text 5, 10, 45, 10, "Case Number"
	  Text 5, 30, 70, 10, "Household member"
	EndDialog
	Dialog select_person_dialog
	If ButtonPressed = cancel Then StopScript
	If ButtonPressed = search_button Then 
		If MAXIS_case_number = "" Then 
			MsgBox "Cannot search without a case number, please try again."
		Else 
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown)
			err_msg = err_msg & "Start Over"
		End If 
	End If 
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a valid case number."
	If clt_to_update = "Select One..." Then err_msg = err_msg & vbNewLine & "Please pick a client whose EMPS panel you need to update."
	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = "" 
	
'FIND Case Number

'Look at EMPS for M01 
'Find if Fin Orient Date is Missing
'Find if ES Referral date is Missing
clt_ref_num = left(clt_to_update, 2)

Fin_Orient_Missing = FALSE 
ES_referral_Missing = FALSE 

Call navigate_to_MAXIS_screen ("STAT", "EMPS")		'Go to EMPS
EMWriteScreen clt_ref_num, 20, 76
transmit
EMReadScreen Fin_Orient_Dt, 8, 5, 39				'Reading and formatting the ES Referral Date and Financial Orientation Date
EMReadScreen ES_Referral_Dt, 8, 16, 40
If Fin_Orient_Dt = "__ __ __" then 
	Fin_Orient_Dt = ""
	Fin_Orient_Missing = TRUE 
Else 
	Fin_Orient_Dt = replace(Fin_Orient_Dt, " ", "/")
End If 
If ES_Referral_Dt = "__ __ __" Then 
	ES_Referral_Dt = ""
	ES_referral_Missing = TRUE 
Else 
	ES_Referral_Dt = replace(ES_Referral_Dt, " ", "/")
End If 

EMReadScreen ES_Status, 2, 15, 40
ES_Status = abs(ES_Status)
If ES_Status = 20 Then 
	UniversalParticipant = TRUE 
ElseIf ES_Status < 20 Then 
	ExtensionCase = TRUE 
Else 
	FSSCase = TRUE 
End If 

baby_on_case = FALSE							'Defaults to false
Do 
	Call Navigate_to_MAXIS_screen ("STAT", "PNLP")
	EMReadScreen nav_check, 4, 2, 53
Loop until nav_check = "PNLP"
maxis_row = 3
Do 
	EMReadScreen panel_name, 4, maxis_row, 5	'Reads the name of each panel listed on PNLP
	If panel_name = "MEMB" Then 				'Looking for MEMB
		EMReadScreen client_age, 2, maxis_row, 71		'Reads the age on the MEMB line
		If client_age = " 0" Then 
			baby_on_case = TRUE	'If a age is listed as 0 then a baby is on the case'
			EMReadScreen Baby_ref_numb, 2, 10, 10
		End If 
	End If 
	If panel_name = "MEMI" Then Exit Do			'Once it gets to a panel named MEMI, there are no additional MEMB panels
	maxis_row = maxis_row + 1					'Go to next row
	If maxis_row = 20 Then 						'If it gets to row 20 it needs to go to the next page
		transmit
		maxis_row = 3
	End If 
Loop until panel_name = "REVW"
If baby_on_case = TRUE Then 		'If there is no baby on the case the script will not update to a child under 12 months exemption - this notifies the worker and unchecks the selector
	Call Navigate_to_MAXIS_screen ("STAT", "MEMB")
	EMWriteScreen Baby_ref_numb, 20, 76
	transmit
	EMReadScreen Baby_DOB, 10, 8, 42
	Baby_DOB = replace(Baby_DOB, " ", "/")
	Baby_is_One = DateAdd("yyyy", 1, Baby_DOB)
	Exemption_Unaavailable = DateAdd("m", 1, Baby_is_One)
	Exemption_End_Month = right("00" & DatePart("m", Exemption_Unaavailable), 2)
	Exemption_End_Year = DatePart("yyyy", Exemption_Unaavailable)
End If 

'Fin_Orient_Missing = TRUE
'ES_referral_Missing = TRUE 

Do 
	err_msg = ""
	dialog_length = 120
	IF ES_referral_Missing = TRUE Then dialog_length = dialog_length + 15
	IF Fin_Orient_Missing = TRUE Then dialog_length = dialog_length + 15
	IF EMPS_Workaround = TRUE Then dialog_length = dialog_length + 20
	IF Child_Under_One = TRUE Then dialog_length = dialog_length + 40
	IF Remove_FSS = TRUE Then dialog_length = dialog_length + 20
	
	y_pos = 25
	BeginDialog fss_code_detail, 0, 0, 370, dialog_length, "Update FSS Information from the Status Update"
	  Text 5, 10, 195, 10, "This script can update EMPS for the following proceedures:"
	  
	  IF ES_referral_Missing = TRUE Then
		  Text 5, y_pos, 95, 10, "ES Referral Date is Missing:"
		  CheckBox 110, y_pos, 140, 10, "Check Here to have the script update to ", update_ES_ref_checkbox
		  EditBox 255, y_pos - 5, 50, 15, new_es_referral_dt
		  y_pos = y_pos + 15
	  End If 
	  IF Fin_Orient_Missing = TRUE Then
		  Text 5, y_pos, 95, 10, "Fin Orient Date is Missing:"
		  CheckBox 110, y_pos, 140, 10, "Check Here to have the script update to ", update_fin_orient_checkbox
		  EditBox 255, y_pos - 5, 50, 15, new_fin_oreient_dt
		  y_pos = y_pos + 15
	  End If
	  
	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS to get MFIP results instead of DWP", Intake_MFIP_Button
	  Text 205, y_pos, 105, 10, "Workaround process for Intake"
	  y_pos = y_pos + 15
	  IF EMPS_Workaround = TRUE Then
		  y_pos = y_pos + 5
		  Text 10, y_pos, 65, 10, "Date of Application"
		  EditBox 80, y_pos - 5, 50, 15, date_of_app
		  y_pos = y_pos + 15
	  End If 
	  
	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS for Child Under 12 Months Exemption", Child_Under_One_Button
	  Text 205, y_pos, 75, 10, "Adding or removing" 
	  y_pos = y_pos + 15
	  IF Child_Under_One = TRUE Then
		  y_pos = y_pos + 5
		  Text 10, y_pos, 205, 10, "It appears you need to ADD/REMOVE the exemption.  Reason:"
		  DropListBox 220, y_pos - 5, 125, 45, "Select One..."+chr(9)+"Child Age"+chr(9)+"Caregiver request"+chr(9)+"MFIP results approve - complete workaround", List3
		  Text 10, y_pos + 20, 75, 10, "First month to remove"
		  EditBox 90, y_pos + 15, 15, 15, end_month
		  EditBox 105, y_pos + 15, 15, 15, End_Year
		  Text 135, y_pos + 20, 75, 10, "Date of Client request:"
		  EditBox 220, y_pos + 15, 50, 15, client_request_date
		  y_pos = y_pos + 40
	  End If 
	  
	  ButtonGroup ButtonPressed
	    PushButton 10, y_pos, 185, 10, "Code EMPS to remove FSS", Remove_FSS_Button
	  Text 205, y_pos, 125, 10, "Return Caregiver to Regular MFIP-ES"
	  y_pos = y_pos + 15
	  IF Remove_FSS = TRUE Then
		  Text 10, y_pos + 5, 165, 10, "First month client should be Universal Participant"
		  EditBox 180, y_pos, 15, 15, UP_month
		  EditBox 195, y_pos, 15, 15, UP_year
		  y_pos = y_pos +15
	  End If 
	  Text 5, y_pos + 15, 40, 10, "Other Notes"
	  EditBox 50, y_pos + 10, 310, 15, other_notes
	  Text 5, y_pos + 35, 60, 10, "Worker Signature"
	  EditBox 75, y_pos + 30, 110, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 255, y_pos + 30, 50, 15
	    CancelButton 310, y_pos + 30, 50, 15
	EndDialog

	Dialog fss_code_detail
	cancel_confirmation
	If ButtonPressed = Intake_MFIP_Button Then 
		err_msg = err_msg & "Start Over"
		EMPS_Workaround = NOT(EMPS_Workaround)
	End If 
	If ButtonPressed = Child_Under_One_Button Then 
		err_msg = err_msg & "Start Over"
		Child_Under_One = NOT(Child_Under_One)
	End If 
	If ButtonPressed = Remove_FSS_Button Then 
		err_msg = err_msg & "Start Over"
		Remove_FSS = NOT(Remove_FSS)
	End If 

	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop Until err_msg = ""
'If either are missing, run code from the 2 dail scrubbers to try to find dates


'Run Dialog - button puhed will expand the Dialog

'GET MFIP 
'Code the child under 12 months exemption
'Set TIKL to change back

'Code Child under 12 months
'Get code from FSS 

'Code EMPS to remove FSS...?

'Case Note