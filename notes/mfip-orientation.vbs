'STATS GATHERING=============================================================================================================
name_of_script = "NOTES - MFIP ORIENTATION.vbs"
start_time = timer
STATS_counter = 0               'sets the stats counter at one
STATS_manualtime = 300           'manual run time in seconds
STATS_denomination = "M"        'M is for member

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
call changelog_update("09/06/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================

'DECLARATIONS================================================================================================================
'constants for the HH_MEMB_ARRAY array
const ref_number				= 0
const full_name_const			= 1
const age						= 2
const memb_is_caregiver			= 3
const cash_request_const		= 4
const hours_per_week_const		= 5
const exempt_from_ed_const		= 6
const comply_with_ed_const		= 7
const orientation_needed_const	= 8
const orientation_done_const	= 9
const orientation_exempt_const	= 10
const exemption_reason_const	= 11
const emps_exemption_code_const	= 12
const choice_form_done_const	= 13
const orientation_notes			= 14
const last_const				= 15

Dim HH_MEMB_ARRAY()						'defining this array as an adjustible multi-dimensional array.
ReDim HH_MEMB_ARRAY(last_const, 0)
'============================================================================================================================

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_mo								'we will always run this in Current Month
MAXIS_footer_year = CM_yr

'Inital dialog to capture the case number and worker signature
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 110, "MFIP Orientation"
  EditBox 65, 50, 55, 15, MAXIS_case_number
  EditBox 65, 70, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 110, 90, 50, 15
    CancelButton 165, 90, 50, 15
    PushButton 90, 30, 125, 15, "MFIP Orientation Script Instructions", mfip_orientation_instructions_btn
  Text 5, 10, 205, 20, "This script will facilitate the MFIP Orientation, guiding through all of the information needed during the MFIP Orientation."
  Text 15, 55, 50, 10, "Case Number"
  Text 5, 75, 60, 10, "Worker Signature"
EndDialog

Do
	DO
		err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
		Dialog Dialog1                               'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
		cancel_without_confirmation

	    'Handling for error messaging (in the case of mandatory fields or fields requiring a specific format)-----------------------------------
		Call validate_MAXIS_case_number(err_msg, "*")																	'case number is mandatory here
		IF worker_signature = ""           THEN err_msg = err_msg & vbNewLine & "* You must sign your case note!"       'worker_signature is usually also a mandatory field

		If ButtonPressed = mfip_orientation_instructions_btn Then				'This button will open the instructions and then reshow the dialog
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20MFIP%20ORIENTATION.docx"
		End If
	    'If the error message isn't blank or if the instructions button wasn't pressed, it'll pop up a message telling you what to do!
		IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
	LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
	call check_for_password(are_we_passworded_out)		'ensuring we did not become passworded out while the dialog was up
Loop until are_we_passworded_out = False

Call back_to_SELF
EMWriteScreen MAXIS_case_number, 18, 43
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT MFIP Orientation Script cancelled as it was run in inquiry.")
End If

'QUESTION - should we check to ensure MFIP is active or pending, or at least Cash is pending?

'We are not using the function here because we do not need to select the members to look at.
'The function has the ability to select which members of the household are caregivers in the next dialog.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)			'navigating to stat memb to gather the list of all ref number on the case
If is_this_priv = True Then Call script_end_procedure("It appears that this case is PRIVILEGED and you do not currently have access to it. Check the Case Number and, if needed, request access. Run the script agian once you confirm you have access to the case. The script will now end.")
EMWriteScreen "01", 20, 76
transmit

EMReadScreen case_pw_county, 2, 21, 23											'check to see if the case is in another county
If case_pw_county <> "27" Then Call script_end_procedure("This case is not in Hennepin County and the script cannot take action on a case in another county. The script will now end.")

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		EMWaitReady 0, 0
	End If
	If client_array <> "" Then client_array = client_array & "|" & ref_nbr
	If client_array = "" Then client_array = client_array & ref_nbr
	transmit      'Going to the next MEMB panel
	Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
	member_count = member_count + 1
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = split(client_array, "|")		'make this list an array we can loop through

'Looping through all of the reference nummbers on the case to gather client name and age
'Age is needed to complete the assessment for MFIP Orientation exemptions
clt_count = 0
For each hh_clt in client_array

	'now we are resizing the multi-dimensional array that will store the client information
	'and details about the orientation process that happens during the script run'
	ReDim Preserve HH_MEMB_ARRAY(last_const, clt_count)
	HH_MEMB_ARRAY(ref_number, clt_count) = hh_clt								'setting from the original client array made just before this loop'

	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   							'navigating to the MEMB panel for each reference number
	EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
	transmit

	EMReadScreen access_denied_check, 13, 24, 2         						'Sometimes MEMB gets this access denied issue and we have to work around it.
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		EMWaitReady 0, 0
		memb_last_name_const = "UNABLE TO FIND"
		memb_first_name_const = "Access Denied"
	Else
		EMReadScreen HH_MEMB_ARRAY(age, clt_count), 3, 8, 76					'Reading the name and age if there was not 'Access Denied' issue
		EMReadscreen memb_last_name_const, 25, 6, 30
		EMReadscreen memb_first_name_const, 12, 6, 63
		memb_last_name_const = trim(replace(memb_last_name_const, "_", ""))
		memb_first_name_const = trim(replace(memb_first_name_const, "_", ""))
	End If

	HH_MEMB_ARRAY(age, clt_count) = trim(HH_MEMB_ARRAY(age, clt_count))			'formatting the age and name information.
	If HH_MEMB_ARRAY(age, clt_count) = "" Then HH_MEMB_ARRAY(age, clt_count) = 0
	HH_MEMB_ARRAY(age, clt_count) = HH_MEMB_ARRAY(age, clt_count) * 1
	HH_MEMB_ARRAY(full_name_const, clt_count) = memb_first_name_const & " " & memb_last_name_const

	clt_count = clt_count + 1
Next

family_cash_program = "MFIP"			'defaulting to MFIP as the program selection.

'this iswhere the main functionality of this script is called.
'We are using a function because this needs to match the experiance in other scripts.
'This function will call dialogs and enter CASE/NOTEs - eventually it may update EMPS panels
Call complete_MFIP_orientation(HH_MEMB_ARRAY, ref_number, full_name_const, age, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes, family_cash_program)

'Now that the CASE/NOTES are completed the script will gather information for the end_msg report out MsgBox
'This next block is ONLY to fill the end_msg
If family_cash_program = "DWP" Then
	STATS_counter = 1
	STATS_manualtime = 60		'if DWP - the manual time is changed becuase we didn't complete an orientation
	end_msg = "The NOTES - MFIP Orientation script has completed without taking any action." & vbCr
	end_msg = end_msg & "You have indicated that the family cash program is DWP." & vbCr & vbCr
	end_msg = end_msg & "This script does not have support for the financial orientation and information on DWP cases. This functionality is built to specifically support MFIP cases and MFIP caregivers."
Else
	end_msg = "NOTES - MFIP Orientation script run completed." & vbCr
	for each_clt = 0 to UBound(HH_MEMB_ARRAY, 2)

		If HH_MEMB_ARRAY(memb_is_caregiver, each_clt) = True Then
			STATS_counter = STATS_counter + 1
			caregiver_detail = HH_MEMB_ARRAY(full_name_const, each_clt) & " is a caregiver on this case." & vbCr
			If HH_MEMB_ARRAY(orientation_needed_const, each_clt) = True Then caregiver_detail = caregiver_detail & " - An MFIP Orientation is needed for this caregiver. " & vbCr
			If HH_MEMB_ARRAY(orientation_needed_const, each_clt) = False Then caregiver_detail = caregiver_detail & " - An MFIP Orientation is NOT needed for this caregiver." & vbCr
			If HH_MEMB_ARRAY(orientation_exempt_const, each_clt) = True Then
				caregiver_detail = caregiver_detail & " - This caregiver is exemmpt from needing an MFIP Orientation." & vbCr
				caregiver_detail = caregiver_detail & "   Exemption Reason: " & HH_MEMB_ARRAY(exemption_reason_const, each_clt) & vbCr
			End If
			If HH_MEMB_ARRAY(orientation_done_const, each_clt) = True Then  caregiver_detail = caregiver_detail & " * The orientation was completed during this script run and a CASE/NOTE has been entered." & vbCr
			If HH_MEMB_ARRAY(orientation_done_const, each_clt) = False Then  caregiver_detail = caregiver_detail & " * MFIP ORIENTATION NOT COMPLETED AND STILL NEEDED FOR " & HH_MEMB_ARRAY(full_name_const, each_clt) & "." & vbCr

			end_msg = end_msg & vbCr & caregiver_detail
		End If
	next
End If
end_msg = end_msg & vbCr & "CASE/NOTEs have been made by the script. Updates to EMPS should have been completed manually during the script run. If that is still needed, go back and update STAT/EMPS now."

'End the script.
script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/31/2022
'--Tab orders reviewed & confirmed----------------------------------------------08/31/2022
'--Mandatory fields all present & Reviewed--------------------------------------08/31/2022
'--All variables in dialog match mandatory fields-------------------------------08/31/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/31/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------08/31/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function:
'    confirm that proper punctuation is used -----------------------------------09/08/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/31/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------08/31/2022
'--PRIV Case handling reviewed -------------------------------------------------08/31/2022
'--Out-of-County handling reviewed----------------------------------------------08/31/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/31/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---08/31/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/31/2022
'--Incrementors reviewed (if necessary)-----------------------------------------08/31/2022
'--Denomination reviewed -------------------------------------------------------08/31/2022
'--Script name reviewed---------------------------------------------------------08/31/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/31/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------08/31/2022
'--comment Code-----------------------------------------------------------------08/31/2022
'--Update Changelog for release/update------------------------------------------08/31/2022
'--Remove testing message boxes-------------------------------------------------08/31/2022
'--Remove testing code/unnecessary code-----------------------------------------08/31/2022					There is still some testing code in the function - this will behandled when moved to FuncLib
'--Review/update SharePoint instructions----------------------------------------08/31/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------							TODO - Once initial testing is done - add feedback to add the script to the HSR manual page
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------TODO
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
