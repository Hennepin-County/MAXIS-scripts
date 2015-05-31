'Created by Ilse Ferris from Hennepin County, Gay Sikkink from Sterns County, and Charles Potter from Anoka County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CS Disregard FIAT.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'DIALOG
'===========================================================================================================================
BeginDialog CSD_FIAT_dlg, 0, 0, 161, 90, "Child Support Disregard FIATer"
  EditBox 60, 5, 90, 15, case_number
  EditBox 60, 25, 20, 15, footer_month
  EditBox 130, 25, 20, 15, footer_year
  EditBox 75, 45, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 65, 50, 15
    CancelButton 90, 65, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 30, 50, 10, "Footer Month:"
  Text 85, 30, 45, 10, "Footer Year:"
  Text 10, 50, 60, 10, "Worker Signature:"
EndDialog

'============================================================================================================================

EMConnect ""

'Finds the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
CALL find_variable("Benefit Period (MM YY): ", footer_month, 2)
IF footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & footer_month & " ", footer_year, 2)
ELSE
CALL find_variable("Month: ", footer_month, 2)
IF footer_month <> "" THEN CALL find_variable("Month: " & footer_month & " ", footer_year, 2)
END IF

'Warning/instruction box
MsgBox "This script is intended for use for Family Cash cases (DWP/MFIP)." & vbnewline & vbNewLine &_
		vbTab & "- If this case has adult caregivers and minor caregivers in it's " & vbNewline &_
		vbTab & "   household composition please refer to CM0017.15.03 for child and" & vbNewLine &_
		vbtab & "   spousal support income budgeting" & vbNewLine &_
		vbTab & "- If the case is involved in a sanction please process manually" & vbNewLine & vbNewLine &_
		"The script will now display all household members on the case. Please check" & vbNewLine &_
		"the boxes next to the eligible DWP/MFIP assistance unit members."

check_for_maxis(true)

DO
	DO
		DO
			DO
				Dialog CSD_FIAT_dlg
				IF buttonpressed = cancel THEN stopscript
				If case_number = "" then MsgBox "You must have a case number to continue."
			LOOP until case_number <> ""
			If worker_signature = "" then Msgbox "Please sign your case note"
		LOOP until worker_signature <> ""
		If footer_month = "" then MsgBox "You must have a starting footer month to continue."
	LOOP until footer_month <> ""
	If footer_year = "" then MsgBox "You must have a starting footer year to continue."
Loop until footer_year <> ""

back_to_self

'starting at requested month
EMwritescreen footer_month, 20, 43
EMwritescreen footer_year, 20, 46

check_for_maxis(true)

'Checking for family cash
CALL navigate_to_MAXIS_screen("CASE", "CURR")

call find_variable("DWP: ", DWP_cash_status, 7)
call find_variable("MFIP: ", MFIP_cash_status, 7)
IF DWP_cash_status = "" AND MFIP_cash_status = "" THEN script_end_procedure("This case does not seem to have MFIP or DWP open or pending. Please review case number or prog and try again.")


'Creating custom dialog to determine counted children in Household
CALL hh_member_custom_dialog(hh_member_array)

'checking unea panels to see if selected HH members have one of the 3 types of CS UNEA panels. 
FOR each HH_member in hh_member_array
	CALL navigate_to_MAXIS_screen("STAT", "UNEA")
	EMwritescreen HH_member, 20, 76
	transmit
	DO
		EMReadScreen current_panel, 1, 2, 73
		EMReadScreen total_panels, 1, 2, 78
		EmReadScreen UNEA_type, 2, 5, 37
		IF UNEA_type = "08" or UNEA_type = "36" or UNEA_type = "39" THEN
			CS_found = True
			Exit do
		END IF
		Transmit
	LOOP until current_panel = total_panels
Next

'if no CS UNEA is found the script ends
IF CS_found <> True THEN script_end_procedure("A child support UNEA panel was not found for requested HH members. Please check to make sure the panel is assigned to the correct person and the correct HH members have been selected.")

back_to_self

IF DWP_cash_status <> "" Then
	CALL navigate_to_MAXIS_screen("ELIG", "DWP")
	PF9
	EMwritescreen "04", 10, 41
	transmit
	
END If

IF MFIP_cash_status <> "" Then
	CALL navigate_to_MAXIS_screen("FIAT", "")
	EMwritescreen "03", 4, 34
	EMwritescreen "x", 9, 22
	transmit
END If


