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

'Default member is member 01
HH_member = "01"

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

Do
	err_msg = ""
	Dialog select_person_dialog
	If ButtonPressed = CancelButton Then StopScript
	If ButtonPressed = search_button Then 
		If MAXIS_case_number = "" Then 
			MsgBox "Cannot search without a case number, please try again."
		Else 
			Call Generate_Client_List(HH_Memb_DropDown)
		End If 
	End If 
		
	
Loop until err_msg = "" 
	
'FIND Case Number

'Look at EMPS for M01 
'Find if Fin Orient Date is Missing
'Find if ES Referral date is Missing

'If either are missing, run code from the 2 dail scrubbers to try to find dates


'Run Dialog - button puhed will expand the Dialog

'GET MFIP 
'Code the child under 12 months exemption
'Set TIKL to change back

'Code Child under 12 months
'Get code from FSS 

'Code EMPS to remove FSS...?

'Case Note