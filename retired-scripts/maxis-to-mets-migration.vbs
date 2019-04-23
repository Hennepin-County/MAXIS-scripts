'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MAXIS TO METS MIGRATION.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 320                     'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("03/15/2019", "Updated MEMO to allow for length of client name.", "MiKayla Handley, Hennepin County")
Call changelog_update("03/11/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG
BeginDialog MAXIS_to_METS_dialog, 0, 0, 196, 120, "MAXIS to METS Migration"
  EditBox 70, 80, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 85, 100, 50, 15
    CancelButton 140, 100, 50, 15
  Text 15, 20, 170, 35, "This script will case note and send a SPEC/MEMO to the selected member with specific verbiage about how to apply in METS for continued health care coverage."
  GroupBox 10, 5, 180, 55, "Using this script:"
  Text 35, 65, 120, 10, "Case Number:"
  Text 5, 85, 60, 10, "Worker Signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'grabbing case number & connecting to MAXIS
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
member_number = "01"

'Main dialog: user will input case number and member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog main_dialog				'main dialog
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid case number."		'mandatory field
		IF len(member_number) <> 2 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid 2-digit member number."		'mandatory field
        IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Gathering member information for notice, sending the memo & case noting
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then
    script_end_procedure("PRIV case, cannot access/update. The script will now end.")
Else
    Call write_value_and_transmit(member_number, 20, 76)
    EmReadscreen first_name, 12, 6, 63
    EMReadScreen last_name, 24, 6, 30

    first_name = trim(replace(first_name, "_", ""))
    last_name = trim(replace(last_name, "_", ""))
    Call fix_case(first_name, 0)
    Call fix_case(last_name, 0)

    Client_name = trim(first_name) & " " & trim(last_name)

    'logic to add closing date in the SPEC/MEMO for the client
    next_month = DateAdd("M", 1, date)
    next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
    last_day_of_month = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string

    'THE MEMO----------------------------------------------------------------------------------------------------
    Call start_a_new_spec_memo

    Call write_variable_in_SPEC_MEMO (Client_name & "'s Medical Assistance will end at the end of the day on " & last_day_of_month & ". It will end because our records show that you need to complete application in MNsure so we can redetermine your eligibility for health care coverage.")
    Call write_variable_in_SPEC_MEMO ("(Code of Federal Regulations, title 42, section 435.916, and Minnesota Statutes, section 256B.056, subdivision 7a)")
    'Call write_variable_in_SPEC_MEMO ("")
    Call write_variable_in_SPEC_MEMO ("You can still apply for health care coverage. To apply, you must go to http://www.mnsure.org and complete an online application. If you cannot apply online, you can complete a paper application.")
    'Call write_variable_in_SPEC_MEMO ("")
    Call write_variable_in_SPEC_MEMO ("NOTE: If you already applied for coverage for this person through MNsure or your county human services agency and got an approval notice, you do not have to apply again.")
    'page 2 of MEMO
    Call write_variable_in_SPEC_MEMO ("If you have questions or want to ask for a paper application, call your county human services agency at 612-596-1300. You can also call the DHS Minnesota Health Care Programs (MHCP) Member Help Desk at 651-431-2670 or 800-657-3739. Or call using your preferred relay service.")
    'Call write_variable_in_SPEC_MEMO ("")
    Call write_variable_in_SPEC_MEMO ("You can also get help through a navigator. To find one, go to http://www.mnsure.org. Click the ""Get Help"" tab on the home page. Then click the ""Find an assister"" link and use the assister directory to find a navigator near you. Your county human services agency can also help you find a navigator in your area.")
    Call write_variable_in_SPEC_MEMO ("You have the right to appeal. Visit this website for more information: https://www.hennepin.us/residents/health-medical/health-care-assistance")
    PF4

    'THE CASE NOTE----------------------------------------------------------------------------------------------------
    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("---Closed HC " & CM_plus_1_mo & "/" & CM_plus_1_yr & " for MEMB " & member_number & "---")
    Call write_variable_in_CASE_NOTE("* This case was identified by DHS as requiring conversion to the METS system.")
    Call write_variable_in_CASE_NOTE("* No associated METS case exists for MEMB " & member_number & ": " & client_name)
    Call write_variable_in_CASE_NOTE("* Informational notice generated via SPEC/MEMO to client regarding applying through mnsure.org.")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
	PF3
End if

script_end_procedure("")
