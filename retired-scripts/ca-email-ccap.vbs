'Required for statistical purposes==========================================================================================
name_of_script = "CA - EMAIL CCAP.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("05/25/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
Call check_for_MAXIS(False)                         'Confirming MAXIS is logged in

CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number

'Dialog to get the case number and confirmation number
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 226, 120, "CCAP Request Received"
  EditBox 130, 10, 50, 15, MAXIS_case_number
  EditBox 130, 30, 90, 15, confirmation_number
  ButtonGroup ButtonPressed
    OkButton 115, 100, 50, 15
    CancelButton 170, 100, 50, 15
  Text 75, 15, 50, 10, "Case Number:"
  Text 15, 35, 115, 10, "MNbenefits Confirmation Number:"
  Text 10, 60, 210, 40, "This script will send an email to the CCAP Team to pend new a Child Care Request. This script is used for when a MNbenefits application has been received for an active MAXIS program and CCAP. If the MAXIS program needs to be pended, use CA - Application Recived."
EndDialog

Do
    Do
        err_msg = ""                        'resetting the error message

        dialog Dialog1                      'calling the dialog
        cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")       'both fields are mandatory - confirming the varibales
        If trim(confirmation_number) = "" Then err_msg = err_msg & vbNewLine & "* Enter the conformation for the MNbenefits application."

        'displaying the err_msg to confirm the information has been entered into the dialog.
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)      'making sure we are not passworded out
Loop until are_we_passworded_out = False

Call navigate_to_MAXIS_screen("STAT", "SUMM")           'going to SUMM to read the case numbe from MAXIS
EMReadScreen case_name, 20, 21, 46                      'the case name is listed at the bottom of SUMM
case_name = trim(case_name)                             'trimming the case_name variable for email formatting

Call find_user_name(the_person_running_the_script)      'Reading the name of the person running the script from Outlook

'Creating the email text
ccap_email_subject = "ES - CA - Financial Case " & MAXIS_case_number & " - " & case_name & " also requests CCAP"
ccap_email_body = "MNbenefits Application received requesting CCAP with other financial programs."
ccap_email_body = ccap_email_body & vbCr & "Confirmation Number: " & confirmation_number
ccap_email_body = ccap_email_body & vbCr & vbCr & "MAXIS Case Number: " & MAXIS_case_number
ccap_email_body = ccap_email_body & vbCr & "MAXIS Case Name: " & case_name
ccap_email_body = ccap_email_body & vbCr & vbCr & "Application form can be found in the ECF case file for this case."
ccap_email_body = ccap_email_body & vbCr & vbCr & "Thank you,"
ccap_email_body = ccap_email_body & vbCr & the_person_running_the_script
ccap_email_body = ccap_email_body & vbCr & "Case Assignment Team"

'sending the email
Call create_outlook_email("HSPH.ResourcesCCAP@hennepin.us", "", ccap_email_subject, ccap_email_body, "", True)

'This is a UTILITY script. No CASE/NOTE is needed.

Call script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/25/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/25/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/25/2022
'--All variables in dialog match mandatory fields-------------------------------05/25/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/25/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/25/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/25/2022
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------05/25/2022
'--Script name reviewed---------------------------------------------------------05/25/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/25/2022
'--comment Code-----------------------------------------------------------------05/25/2022
'--Update Changelog for release/update------------------------------------------05/25/2022
'--Remove testing message boxes-------------------------------------------------05/25/2022
'--Remove testing code/unnecessary code-----------------------------------------05/25/2022
'--Review/update SharePoint instructions----------------------------------------05/25/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
