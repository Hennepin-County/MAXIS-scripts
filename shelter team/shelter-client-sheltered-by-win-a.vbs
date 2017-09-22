'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT SHELTERED BY WINDOW A.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog client_sheltered_window_A, 0, 0, 301, 245, "Client Sheltered Window A"
  DropListBox 200, 5, 90, 15, "Select one..."+chr(9)+"FMF"+chr(9)+"PSP"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
  EditBox 75, 25, 65, 15, voucher_date
  EditBox 260, 25, 30, 15, nights_housed
  EditBox 110, 45, 30, 15, adults_vouchered
  EditBox 260, 45, 30, 15, children_vouchered
  CheckBox 10, 65, 130, 10, "Check here if any adults are pregnant", PX_check
  EditBox 110, 80, 180, 15, reason_for_homelessness
  EditBox 60, 115, 230, 15, name_of_person_verifying
  EditBox 60, 135, 230, 15, relationship
  EditBox 60, 155, 230, 15, phone_number
  CheckBox 15, 180, 280, 10, "Informed client that they will need to see Rapid ReHousing Screener first, then see ", informed_client_checkbox
  EditBox 70, 205, 220, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 225, 50, 15
    CancelButton 240, 225, 50, 15
  Text 35, 120, 25, 10, "Name:"
  Text 15, 140, 45, 10, "Relationship:"
  Text 10, 160, 50, 10, "Phone Number:"
  Text 25, 210, 40, 10, "Other notes:"
  Text 10, 50, 100, 10, "Number of Aduilts vouchered:"
  Text 195, 30, 60, 10, "How many nights?"
  Text 155, 50, 105, 10, "Number of Children vouchered:"
  GroupBox 5, 100, 290, 75, "Homelessness verified by contacting:"
  Text 15, 85, 90, 10, "Reason for homelessness:"
  Text 25, 190, 150, 10, "the Shelter team for interview and revoucher."
  Text 150, 10, 45, 10, "Shelter name:"
  Text 25, 10, 45, 10, "Case number:"
  EditBox 75, 5, 65, 15, MAXIS_case_number
  Text 5, 30, 70, 10, "Shelter voucher date:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog client_sheltered_window_A
		cancel_confirmation
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