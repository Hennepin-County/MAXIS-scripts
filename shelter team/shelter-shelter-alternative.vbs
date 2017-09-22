'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SHELTER ALTERATIVE.vbs"
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
BeginDialog shelter_alternative_dialog, 0, 0, 301, 180, "Shelter Alternative"
  EditBox 55, 5, 60, 15, MAXIS_case_number
  EditBox 210, 5, 20, 15, number_of_adults_sheltered
  EditBox 255, 5, 20, 15, number_of_children_sheltered
  EditBox 55, 30, 225, 15, reason_not_authorized
  EditBox 30, 65, 250, 15, needed_one
  EditBox 30, 85, 250, 15, needed_two
  EditBox 30, 105, 250, 15, needed_three
  EditBox 30, 125, 250, 15, needed_four
  EditBox 45, 155, 135, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 155, 50, 15
    CancelButton 240, 155, 50, 15
  Text 20, 35, 35, 10, "Situation:"
  Text 125, 10, 85, 10, "Client seeking shelter for"
  Text 235, 10, 20, 10, "A and"
  Text 5, 10, 45, 10, "Case number:"
  Text 15, 70, 10, 10, "1."
  GroupBox 5, 50, 285, 100, "What is needed for shelter?"
  Text 15, 90, 10, 10, "2."
  Text 280, 10, 10, 10, "C"
  Text 15, 110, 10, 10, "3."
  Text 5, 160, 40, 10, "Comments:"
  Text 15, 130, 10, 10, "4."
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog shelter_alternative_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If IsNumeric(number_of_adults_sheltered) = False then err_msg = err_msg & vbNewLine & "* Enter the nubmer of adults sheltered"
		If IsNumeric(number_of_children_sheltered) = False then err_msg = err_msg & vbNewLine & "* Enter the number of children sheltered"
		If reason_not_authorized = "" then err_msg = err_msg & vbNewLine & "* Enter reason not authorized"	
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter N/A in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 	
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### SHELTER ALTERNATIVE ###")
Call write_variable_in_CASE_NOTE("* Client seeking shelter for " & number_of_adults_sheltered & "A and " & number_of_children_sheltered & "C")
Call write_bullet_and_variable_in_CASE_NOTE("Situation", reason_not_authorized)
If trim(needed_one) <> "" or needed_two <> "" or needed_three <> "" or needed_four <> "" then Call write_variable_in_CASE_NOTE("--What is needed for shelter?--")
If trim(needed_one) <>   "" then Call write_variable_in_CASE_NOTE("1. " & needed_one)
If trim(needed_two) <>   "" then Call write_variable_in_CASE_NOTE("2. " & needed_two)
If trim(needed_three) <> "" then Call write_variable_in_CASE_NOTE("3. " & needed_three)
If trim(needed_four) <>  "" then Call write_variable_in_CASE_NOTE("4. " & needed_four)
Call write_bullet_and_variable_in_CASE_NOTE("Comments", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team") 

script_end_procedure("")