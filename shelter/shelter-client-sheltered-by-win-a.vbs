'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-CLIENT SHELTERED BY WINDOW A.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
''CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("09/20/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("09/23/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 271, 220, "Client Sheltered"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  DropListBox 170, 5, 90, 15, "Select one..."+chr(9)+"FMF"+chr(9)+"PSP"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
  EditBox 80, 25, 45, 15, voucher_date
  EditBox 240, 25, 20, 15, nights_housed
  EditBox 105, 45, 20, 15, adults_vouchered
  EditBox 240, 45, 20, 15, children_vouchered
  CheckBox 5, 65, 130, 10, "Check here if any adults are pregnant", PX_check
  EditBox 105, 80, 155, 15, reason_for_homelessness
  EditBox 55, 115, 205, 15, name_of_person_verifying
  EditBox 55, 135, 95, 15, relationship
  EditBox 210, 135, 50, 15, phone_number
  Text 5, 205, 60, 10, "Worker signature:"
  EditBox 55, 180, 210, 15, other_notes
  EditBox 65, 200, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 200, 45, 15
    CancelButton 220, 200, 45, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 120, 10, 45, 10, "Shelter name:"
  Text 5, 30, 70, 10, "Shelter voucher date:"
  Text 175, 30, 60, 10, "Number of nights:"
  Text 5, 50, 95, 10, "Number of adults vouchered:"
  Text 135, 50, 105, 10, "Number of Children vouchered:"
  Text 5, 85, 90, 10, "Reason for homelessness:"
  GroupBox 5, 100, 260, 55, "Homelessness verified by contacting:"
  Text 10, 120, 25, 10, "Name:"
  Text 10, 140, 45, 10, "Relationship:"
  Text 155, 140, 50, 10, "Phone number:"
  CheckBox 10, 160, 250, 10, "Informed client that they will need to see Rapid ReHousing Screener first,", informed_client_checkbox
  Text 20, 170, 185, 10, "then see the Shelter team for interview and revoucher."
  Text 5, 185, 40, 10, "Other notes:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If shelter_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the name of the shelter where client(s) housed"
		If isdate(voucher_date) = false then err_msg = err_msg & vbNewLine & "* Enter a valid shelter voucher date"
		If nights_housed = "" then err_msg = err_msg & vbNewLine & "* Enter the number of nights clients housed"
		If isnumeric(adults_vouchered) = false then err_msg = err_msg & vbNewLine & "* Enter the number of adults vouchered"
		If isnumeric(children_vouchered) = false then err_msg = err_msg & vbNewLine & "* Enter the number of children vouchered"
		If reason_for_homelessness = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for client's homelessness"
		If name_of_person_verifying = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the person who verified client's homelessness"
		If relationship = "" then err_msg = err_msg & vbNewLine & "* Enter the relationship to the client of the person who verified client's homelessness"
		If phone_number = "" then err_msg = err_msg & vbNewLine & "* Enter the phone number of the person who verified client's homelessness"
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'adding the case number
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'creating date variable to add to header
exit_date = dateadd("d", nights_housed, voucher_date)
header_date = voucher_date & " - " & exit_date

Household_comp = ""
If PX_check = 1 then
	adults_vouchered = adults_vouchered - 1
	Household_comp = Household_comp & "1 PX, "
END IF
If adults_vouchered <> "0" or adults_vouchered <> "" then Household_comp = Household_comp & adults_vouchered & "A, "
If children_vouchered <> "0" or children_vouchered <> "" then Household_comp = Household_comp & children_vouchered & "C, "

Household_comp = trim(Household_comp)
'takes the last comma off of Household_comp when autofilled into dialog if more more than one app date is found and additional app is selected
If right(Household_comp, 1) = "," THEN Household_comp = left(Household_comp, len(Household_comp) - 1)

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### App'd shelter at " & shelter_droplist & " for " & header_date & " for " & nights_housed & " nights###")
Call write_bullet_and_variable_in_CASE_NOTE("Voucher keyed for", Household_comp)
Call write_bullet_and_variable_in_CASE_NOTE("* Reason for client's homelessness", reason_for_homelessness)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("* Homelessness verified by contacting:")
Call write_bullet_and_variable_in_CASE_NOTE("   Name", name_of_person_verifying)
Call write_bullet_and_variable_in_CASE_NOTE("   Relationship to the client", relationship)
Call write_bullet_and_variable_in_CASE_NOTE("   Phone Number of person verifying client's homelessness", phone_number)
If informed_client_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Client will need to see Rapid ReHousing Screener first, then see shelter team for interview and revoucher. ")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
