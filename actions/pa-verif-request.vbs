'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PA VERIF REQUEST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 294                	'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/01/2016", "Checkbox added with the option to have 'Other Income' not listed on the word document.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
MAXIS_footer_month = datepart("m", next_month)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = datepart("yyyy", next_month)
MAXIS_footer_year = "" & MAXIS_footer_year - 2000

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 151, 70, "PA Verification Request"
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 95, 50, 50, 15
  EditBox 75, 5, 70, 15, MAXIS_case_number
  EditBox 75, 25, 30, 15, MAXIS_footer_month
  EditBox 115, 25, 30, 15, MAXIS_footer_year
  Text 10, 10, 50, 10, "Case Number"
  Text 10, 30, 65, 10, "Footer month/year:"
EndDialog

BeginDialog PA_verif_dialog, 0, 0, 316, 255, "PA Verif Dialog"
  ButtonGroup ButtonPressed
    OkButton 200, 230, 50, 15
    CancelButton 255, 230, 50, 15
  EditBox 40, 15, 30, 15, snap_grant
  EditBox 105, 15, 35, 15, MFIP_food
  EditBox 145, 15, 35, 15, MFIP_cash
  EditBox 40, 35, 30, 15, MSA_Grant
  EditBox 145, 35, 35, 15, MFIP_housing
  EditBox 40, 55, 30, 15, GA_grant
  EditBox 145, 55, 35, 15, DWP_grant
  EditBox 285, 15, 20, 15, cash_members
  CheckBox 285, 40, 25, 10, "Yes", subsidy_check
  EditBox 285, 55, 20, 15, household_members
  EditBox 85, 75, 220, 15, other_income
  EditBox 105, 95, 20, 15, number_of_months
  CheckBox 15, 155, 280, 10, "Check here to have the income and HH information withheld from the word doc.", no_income_checkbox
  EditBox 55, 180, 250, 15, other_notes
  EditBox 55, 205, 90, 15, completed_by
  EditBox 210, 205, 95, 15, worker_phone
  EditBox 120, 230, 75, 15, worker_signature
  CheckBox 10, 100, 95, 10, "Include screenshot of last", inqd_check
  Text 10, 20, 20, 10, "SNAP:"
  Text 110, 60, 20, 10, "DWP:"
  Text 5, 185, 40, 10, "Other notes:"
  Text 10, 60, 20, 10, "GA:"
  Text 10, 40, 20, 10, "MSA:"
  Text 80, 20, 20, 10, "MFIP:"
  Text 80, 40, 50, 10, "MFIP Housing:"
  Text 5, 80, 75, 10, "Other income and type:"
  Text 200, 40, 80, 10, "$50 subsidy deduction?"
  Text 190, 20, 95, 10, "HH members on cash grant:"
  Text 215, 60, 65, 10, "Total HH members:"
  Text 110, 5, 25, 10, "Food:"
  Text 150, 5, 25, 10, "Cash:"
  Text 130, 100, 60, 10, "months' benefits"
  Text 155, 210, 55, 10, "Worker phone #:"
  Text 5, 210, 50, 10, "Completed by:"
  Text 5, 235, 110, 10, "Worker Signature (for case note):"
  Text 15, 130, 280, 20, "Do not share FTI with outside agencies using this form, including information from SSA such as SSI/RSDI amounts."
  GroupBox 5, 120, 300, 50, "Warning!"
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone & grabs the case number and footer month/year
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing case number dialog
Do
	Do
		err_msg = ""
  		Dialog case_number_dialog
  		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine &  "* You need to type a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Jumping to STAT
call navigate_to_MAXIS_screen("stat", "memb")
'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Pulling household and worker info for the letter
call navigate_to_MAXIS_screen("stat", "addr")
EMReadScreen addr_line1, 21, 6, 43
EMReadScreen addr_line2, 21, 7, 43
EMReadScreen addr_city, 14, 8, 43
EMReadScreen addr_state, 2, 8, 66
EMReadScreen addr_zip, 5, 9, 43
hh_address = addr_line1 & " " & addr_line2 'Finding and Formatting household address
hh_address_line2 = addr_city & " " & addr_state & " " & addr_zip
hh_address = replace(hh_address, "_", "") & vbCrLf & replace(hh_address_line2, "_", "")


household_members = UBound(HH_member_array) + 1 'Total members in household
household_members = cStr(household_members)

'Collecting and formatting client name
call navigate_to_MAXIS_screen("stat", "memb")
call find_variable("Last: ", last_name, 24)
call find_variable("First: ", first_name, 11)
client_name = first_name & " " & last_name
client_name = replace(client_name, "_", "")

'Autofilling info for case note
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up info for case note
earned_income = trim(earned_income)
if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
other_income = earned_income & " " & unearned_income


'This function looks for an approved version of elig
Function approved_version
	EMReadScreen version, 2, 2, 12
	For approved = version to 0 Step -1
	EMReadScreen approved_check, 8, 3, 3
	If approved_check = "APPROVED" then Exit Function
	version = version -1
	EMWriteScreen version, 20, 79
	transmit
	Next
End Function


'This finds the number of members on a DWP/MFIP grant
Function cash_members_finder
	call find_variable("Caregivers......", caregivers, 4)
	call find_variable("Children........", children, 4)
	cash_members = cInt(caregivers) + cInt(children)
	cash_members = cStr(cash_members)
End Function

'Pulling the elig amounts for all open progs on case / curr
call navigate_to_MAXIS_screen("case", "curr")
  call find_variable("MFIP: ", MFIP_check, 6)
   If MFIP_check = "ACTIVE" OR MFIP_check = "APP CL" then
		call navigate_to_MAXIS_screen("elig", "mfip")
	  	EMReadScreen version, 1, 2, 12 'Reading the version, the for loop finds most recent approved.
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
        EMWriteScreen "MFB2", 20, 71
        transmit
        EMReadScreen MFIP_cash, 7, 12, 35
        EMReadScreen MFIP_food, 7, 7, 35
		EMReadScreen MFIP_housing, 6, 17, 36
		IF MFIP_housing = "" then MFIP_housing = 0
		'MFIP_cash = (cInt(MFIP_cash) + MFIP_housing)
		'MFIP_cash = cstr(MFIP_cash)
 		'rental subsidy check
		EMWriteScreen "MFB1", 20, 71
		EMReadScreen subsidy, 2, 17, 37
		If subsidy = "50" then subsidy_check = 1
		'Finding the number of members on cash grant
		call cash_members_finder
		Call navigate_to_MAXIS_screen("case", "curr")
	End if
	If MFIP_check = "PENDIN" then msgbox "MFIP is pending, please enter amounts manually to avoid errors."

	call find_variable("FS: ", fs_check, 6)
	If fs_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "fs")
		EMReadScreen version, 2, 2, 12
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen version, 19, 78
			transmit
		Next
		EMWriteScreen version, 19, 78
		transmit
		EMWriteScreen "FSB2", 19, 70
		transmit
		EMReadScreen SNAP_grant, 7, 10, 75
	    call navigate_to_MAXIS_screen ("case", "curr")
	End if
	If fs_check = "APP CL" then msgbox "SNAP is set to close, please enter amounts manually to avoid errors."
	If fs_check = "PENDIN" then msgbox "SNAP is pending, please enter amounts manually to avoid errors."

	call find_variable("DWP: ", DWP_check, 6)
	If DWP_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "dwp")
		EMReadScreen version, 2, 2, 11
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "DWB2", 20, 71
		transmit
		EMReadScreen DWP_grant, 7, 5, 37
	    EMWriteScreen "DWSM", 20, 71
		transmit
		call find_variable("Caregivers....", caregivers, 5)
		call find_variable("Children......", children, 5)
		cash_members = cInt(caregivers) + cInt(children)
		cash_members = cStr(cash_members)
		call navigate_to_MAXIS_screen ("case", "curr")
	 End if
	If DWP_check = "PENDIN" then msgbox "DWP is pending, please enter amounts manually to avoid errors."

	call find_variable("GA: ", GA_check, 6)
	If GA_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "GA")
		EMReadScreen version, 2, 2, 12
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 78
			transmit
		Next
		EMWriteScreen version, 20, 78
		transmit
		EMWriteScreen "GASM", 20, 70
		transmit
		EMReadScreen GA_grant, 7, 9, 73
	    EMReadScreen ga_members, 1, 13, 32 'Reading file unit type to determine members on cash grant
		If ga_members = "1" then cash_members = "1"
		If ga_members = "6" then cash_members = "2"
		call navigate_to_MAXIS_screen ("case", "curr")
	End If
	If GA_check = "APP CL" then msgbox "GA is set to close, please enter amounts manually to avoid errors."
	If GA_check = "PENDIN" then msgbox "GA is pending, please enter amounts manually to avoid errors."

	call find_variable("MSA: ", MSA_check, 6)
	If MSA_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "msa")
		EMReadScreen version, 2, 2, 11
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "MSSM", 20, 71
		transmit
		EMReadScreen MSA_Grant, 7, 11, 74
		EMReadScreen cash_members, 1, 14, 29
		call navigate_to_MAXIS_screen ("case", "curr")
	End If
	If MSA_check = "APP CL" then MsgBox "MSA is set to close, please enter amounts manually to avoid errors."
	If MSA_check = "PENDIN" then MsgBox "MSA is pending, please enter amounts manually to avoid errors."

	call find_variable("Cash: ", cash_check, 6)
	If cash_check = "PENDIN" then MsgBox "Cash is pending for this household, please explain in additional notes."

'calling the main dialog
Do
	Do
		err_msg = ""
		Dialog PA_verif_dialog
		cancel_confirmation
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Please sign your case note."
		If completed_by = "" then err_msg = err_msg & vbNewLine & "* Please fill out the completed by field."
		If worker_phone = "" then err_msg = err_msg & vbNewLine & "* Please fill out the worker phone field."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'****writing the word document
Set objWord = CreateObject("Word.Application")
Const wdDialogFilePrint = 88
Const end_of_doc = 6
objWord.Caption = "PA Verif Request"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "14"
objSelection.TypeText "Your agency requested information about public assistance from "
objSelection.TypeText county_name
objSelection.TypeText " for the following client:"
objSelection.TypeParagraph()
objSelection.TypeText client_name
objSelection.TypeParagraph()
objSelection.TypeText hh_address
objSelection.TypeParagraph()
objSelection.TypeText "The following grant amounts are active for this household:"

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 7, 3
set objTable = objDoc.Tables(1)

objTable.Cell(1, 2).Range.Text = "Cash  "
objTable.Cell(1, 3).Range.Text = "Food Portion"
objTable.Cell(2, 1).Range.Text = "MFIP (MN Family Investment program) "
objTable.Cell(3, 1).Range.Text = "MFIP Housing Grant"
objTable.Cell(4, 1).Range.Text = "GA (General Assistance)"
objTable.Cell(5, 1).Range.Text = "MSA (MN supplemental Aid)"
objTable.Cell(6, 1).Range.Text = "SNAP (Supplemental Nutrition Assistance program)"
objTable.Cell(2, 2).Range.Text = MFIP_cash
objTable.Cell(2, 3).Range.Text = MFIP_food
objTable.Cell(3, 2).Range.Text = MFIP_housing
objTable.Cell(4, 2).Range.Text = GA_grant
objTable.Cell(5, 2).Range.Text = MSA_Grant
objTable.Cell(6, 3).Range.Text = SNAP_grant
objTable.Cell(7, 1).Range.Text = "DWP (Diversionary Work program) "
objTable.Cell(7, 2).Range.Text = DWP_grant

objTable.AutoFormat(16)
If no_income_checkbox = unchecked Then 		'Only adding the detail from stat if the worker leaves the omit income unchecked
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()

	objSelection.TypeText "Other income known to agency: "
	objSelection.TypeText other_income
	objSelection.TypeParagraph()
	objSelection.TypeText "Number of family members on cash grant: "
	objSelection.TypeText cash_members
	objSelection.TypeParagraph()
	objSelection.TypeText "Number of persons in household: "
	objSelection.TypeText household_members
	objSelection.TypeParagraph()
	ObjSelection.TypeText "Other Notes: "
	objSelection.TypeText other_notes
	objSelection.TypeParagraph()
Else 										'If worker requests income from STAT to be omitted, the script only adds the cash grant size and other notes.
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
	
	objSelection.TypeText "Number of family members on cash grant: "
	objSelection.TypeText cash_members
	objSelection.TypeParagraph()
	
	ObjSelection.TypeText "Other Notes: "
	objSelection.TypeText other_notes
	objSelection.TypeParagraph()
End If 

'Writing INQX to the doc if selected
IF inqd_check = checked THEN
	objSelection.TypeText "Benefits Issued for last " & number_of_months & " months:"
	objSelection.TypeParagraph()
	objSelection.TypeText "Issue Date	    Benefit               Amount                            Benefit Period"
	objSelection.TypeParagraph()
	call navigate_to_MAXIS_screen("MONY", "INQX")
	start_date = dateadd("m", - number_of_months, date) 'Converting dates to determine how far back to look
	start_month = datepart("m", start_date)
	IF len(start_month) = 1 THEN start_month = "0" & start_month
	EMWriteScreen start_month, 6, 38
	EMWriteScreen right(datepart("YYYY", start_date), 2), 6, 41
	transmit
	output_array = "" 'resetting array
	row = 6
	DO
	EMReadScreen reading_line, 80, row, 1
	output_array = output_array & reading_line & "UUDDLRLRBA" 'adding the info to the array
	row = row + 1
	IF row = 18 THEN 'Checking for more screens
		EMReadScreen more_check, 1, 19, 9
		IF more_check <> "+" THEN EXIT DO
		PF8
		row = 6
	END IF
	LOOP
	output_array = split(output_array, "UUDDLRLRBA")
	FOR EACH line in output_array 'Type the info from array into word doc
		IF line <> "                                                                                " THEN
			objSelection.TypeText line & Chr(11)
		END IF
	NEXT
	objSelection.TypeParagraph()
	objSelection.TypeText "**********PROGRAM KEY**********"
	objSelection.TypeParagraph()
	objSelection.TypeText "DW = DWP (Diversionary Work program"
	objSelection.TypeParagraph()
	objSelection.TypeText "EA = Emergency Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "EG = Emergency General Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "FS = SNAP (Supplemental Nutrition)"
	objSelection.TypeParagraph()
	objSelection.TypeText "GA = General Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "HG = MFIP Housing Grant"
	objSelection.TypeParagraph()
	objSelection.TypeText "MF-MF = MFIP (MN Family Investment program, cash portion)"
	objSelection.TypeParagraph()
	objSelection.TypeText "MF-FS = MFIP SNAP (food portion)"
	objSelection.TypeParagraph()
	objSelection.TypeText "MS = MSA (MN Supplemental Aid)"
	objSelection.TypeParagraph()
	objSelection.TypeText "RC = RCA (Refugee Cash Assistance)"
	objSelection.TypeParagraph()
	objSelection.TypeText "GR = Group Residential Housing"
	objSelection.TypeParagraph()
	objSelection.TypeText "SA = Special Needs/Diet"
	objSelection.TypeParagraph()
	objSelection.TypeText "SM = Special Needs MSA (MN Supplemental Aid)"
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
END IF

objSelection.TypeText "Completed By: "
objSelection.TypeText completed_by
objSelection.TypeParagraph()
objSelection.TypeText "Worker phone: "
objSelection.TypeText worker_phone

'Enters the case note
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("PA verification request completed and sent to requesting agency.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Starts the print dialog
objword.dialogs(wdDialogFilePrint).Show

script_end_procedure("")
