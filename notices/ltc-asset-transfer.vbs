'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - LTC - ASSET TRANSFER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 70                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
CALL changelog_update("10/20/2021", "Added CASE:NOTE option, mandatory fields and updated design of dialog.", "Ilse Ferris, Hennepin County")
CALL changelog_update("03/19/2018", "Updated text regarding client's name.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script------------------------
'connecting to MAXIS
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
get_county_code

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 306, 100, "LTC asset transfer dialog"
  Text 30, 10, 45, 10, "Case number:"
  EditBox 80, 5, 50, 15, MAXIS_case_number
  Text 165, 10, 60, 10, "ER date (MM/YY):"
  EditBox 230, 5, 20, 15, renewal_footer_month
  EditBox 255, 5, 20, 15, renewal_footer_year
  Text 10, 35, 70, 10, "Resident First Name:"
  EditBox 80, 30, 70, 15, client
  Text 165, 35, 65, 10, "Spouse First Name:"
  EditBox 230, 30, 70, 15, spouse
  Text 15, 60, 60, 10, "Worker Signature:"
  EditBox 80, 55, 220, 15, worker_signature
  CheckBox 65, 80, 125, 15, "Check here to CASE:NOTE actions.", case_note_checkbox
  ButtonGroup ButtonPressed
    OkButton 195, 80, 50, 15
    CancelButton 250, 80, 50, 15
EndDialog

Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
        If trim(renewal_footer_month) = "" or len(renewal_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a two-digit renewal month."
        If trim(renewal_footer_year) = "" or len(renewal_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a two-digit renewal year."
        If trim(client) = "" then err_msg = err_msg & vbNewLine & "* Enter the resident's first name."
        If trim(spouse) = "" then err_msg = err_msg & vbNewLine & "* Enter the spouse's first name."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

renewal_date = renewal_footer_month & "/" & renewal_footer_year 'Creating renewal date
'Ensureing the client/spouse's namesa are the correct case in the MEMO
Call fix_case_for_name(client)
Call fix_case_for_name(spouse)

'Ensuring we're in MAXIS, the case is not PRIV and it's in-county.
Call check_for_MAXIS(False)
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")
EmReadscreen county_code, 4, 21, 14
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE:NOTE. The script will now stop.")

Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode

Call write_variable_in_SPEC_MEMO("The ownership of " & client & "'s assets must be transferred to " & spouse & " to avoid having them counted in future eligibility determinations. You are encouraged to do this as soon as possible. This transfer of assets must be done before " & client & "'s first annual renewal for " & renewal_date & ". Verification of the transfer can be provided at any time.")
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("At the first annual renewal in " & renewal_date & ", the value of all assets that list " & client & " as an owner or co-owner will be applied towards the Medical Assistance Asset limit of $3,000.00. If the total value of all countable assets for " & client & " is more than $3,000.00, Medical Assistance may be closed for " & renewal_date & ".")

If case_note_checkbox = 1 then
    stats_counter = stats_counter + 1
    PF4 'saving notice
    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("-- Asset Tranfer SPEC/MEMO Sent --")

    'creating a list of who the memo was sent to for the case note
    notice_recip = "Resident, "
    If forms_to_arep = "Y" then notice_recip = notice_recip & "AREP, "
    If forms_to_swkr = "Y" then notice_recip = notice_recip & "SWKR, "
    If send_to_other = "Y" then notice_recip = notice_recip & "Other, "

    notice_recip = trim(notice_recip)  'trims excess spaces of notice_recip
    If right(notice_recip, 1) = "," THEN notice_recip = left(notice_recip, len(notice_recip) - 1)

    Call write_bullet_and_variable_in_CASE_NOTE("MEMO sent to", notice_recip)
    Call write_variable_in_CASE_NOTE("Content of the MEMO:")
    Call write_variable_in_CASE_NOTE("The ownership of" & client & "'s assets must be transferred to" & spouse & " to avoid having them counted in future eligibility determinations. You are encouraged to do this as soon as possible. This transfer of assets must be done before" & client & "'s first annual renewal for " & renewal_date & ". Verification of the transfer can be provided at any time.")
    Call write_variable_in_CASE_NOTE("At the first annual renewal in " & renewal_date & ", the value of all assets that list" & client & " as an owner or co-owner will be applied towards the Medical Assistance Asset limit of $3,000.00. If the total value of all countable assets for" & client & " is more than $3,000.00, Medical Assistance may be closed for " & renewal_date & ".")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
End if

script_end_procedure_with_error_report("Success! Please review your MEMO and/or CASE:NOTE for accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/20/2021
'--Tab orders reviewed & confirmed----------------------------------------------10/20/2021
'--Mandatory fields all present & Reviewed--------------------------------------10/20/2021
'--All variables in dialog match mandatory fields-------------------------------10/20/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/20/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------10/20/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/20/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/20/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------10/20/2021
'--PRIV Case handling reviewed -------------------------------------------------10/20/2021
'--Out-of-County handling reviewed----------------------------------------------10/20/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/20/2021
'--BULK - review output of statistics and run time/count (if applicable)--------10/20/2021----------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/20/2021
'--Incrementors reviewed (if necessary)-----------------------------------------10/20/2021
'--Denomination reviewed -------------------------------------------------------10/20/2021
'--Script name reviewed---------------------------------------------------------10/20/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/20/2021----------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------10/20/2021
'--comment Code-----------------------------------------------------------------10/20/2021
'--Update Changelog for release/update------------------------------------------10/20/2021
'--Remove testing message boxes-------------------------------------------------10/20/2021
'--Remove testing code/unnecessary code-----------------------------------------10/20/2021
'--Review/update SharePoint instructions----------------------------------------10/20/2021
'--Review Best Practices using BZS page ----------------------------------------10/20/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/20/2021
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/20/2021
'--Complete misc. documentation (if applicable)---------------------------------10/20/2021
'--Update project team/issue contact (if applicable)----------------------------10/20/2021
