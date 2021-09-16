'GATHERING STATS===========================================================================================
name_of_script = "UTILITIES - UC VERIFICATION REQUEST.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "M"
'END OF STATS BLOCK===========================================================================================

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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/30/2021", "Inital Version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "Request for Unemployment Insurance Verification email has been sent." 'setting up closing_message UCriable for possible additions later based on conditions
'----------------------------------------------------------------------------------------------------Initial dialog
initial_help_text = "*** Unemployment Compensation ***" & vbNewLine & vbNewLine & "For residents under the age of 18 please see" & vbNewLine & "CM0010.18.01 - MANDATORY VERIFICATIONS - CASH" & vbNewLine & "CM0010.18.02 - MANDATORY VERIFICATIONS - SNAP" & vbNewLine & "Unemployment Insurance benefits are considered countable unearned income for all programs." & vbNewLine & vbNewLine &"NOTE: UC(Unemployment Compensation) and UI(Unemployment Income) are used interchangeably by workers."

'---------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 311, 85, "Request for Unemployment Insurance"
  EditBox 55, 5, 40, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 175, 5, 15, 15, "!", initial_help_button
    PushButton 200, 5, 105, 15, "Unemployment Insurance", HSR_manual_button
  DropListBox 245, 30, 60, 15, "Select One:"+chr(9)+"1800 Chicago"+chr(9)+"Central/NE"+chr(9)+"HC in METs"+chr(9)+"QI"+chr(9)+"Northwest"+chr(9)+"South", team_email_dropdown
  CheckBox 10, 35, 25, 10, "CCA", cca_checkbox
  CheckBox 10, 50, 80, 10, "Other (please specify)", other_checkbox
  EditBox 95, 45, 45, 15, other_check_editbox
  ButtonGroup ButtonPressed
    OkButton 195, 50, 55, 15
    CancelButton 250, 50, 55, 15
  Text 5, 10, 50, 10, "Case Number:"
  GroupBox 5, 25, 140, 40, "Department (if outside ES)"
  Text 5, 70, 305, 10, "Overpaypayment, tax, child/spousal support deductions will be reviewed for this individual."
  Text 170, 35, 70, 10, "Select a team/region:"
EndDialog

DO
    DO
        err_msg = ""
        DO
            Dialog Dialog1
            cancel_without_confirmation
            IF ButtonPressed = initial_help_button then
                tips_tricks_msg = MsgBox(initial_help_text, vbInformation, "Tips and Tricks") 'see initial_help_text above for details of the text
            End if
            IF buttonpressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Unemployment_Insurance.aspx") 'HSR manual policy page
        LOOP until ButtonPressed = -1
        IF team_email_dropdown = "Select One:" THEN err_msg = err_msg & vbCr & "* Specify what team/region you want to send your email to."
      IF other_checkbox = CHECKED and trim(other_check_editbox) = "" THEN err_msg = err_msg & vbCr & "* Specify what your department you are in."
      IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & "Resolve for the following for the script to continue." & vbcr & err_msg
        LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to  assword back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

CALL check_for_MAXIS(False)

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv) 'navigating to stat memb to gather the ref number and name.
If is_this_priv = TRUE then script_end_procedure("PRIV case, cannot access/update. The script will now end.")
DO
    CALL HH_member_custom_dialog(HH_member_array)
    IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

back_to_SELF
EMReadScreen county_code, 4, 21, 14  'Out of county cases from STAT
If county_code <> "X127" then script_end_procedure("Out of County case, cannot access/update. The script will now end.")

'--------------------------------------------------------------------------------Gathering the MEMB/ALIA information

'Establishing array
uc_membs = 0       'incrementor for array
DIM uc_members_array()  'Declaring the array this is what this list is
ReDim uc_members_array(client_alia_ssn_const, 0)  'Resizing the array 'redimmed to the size of the last constant  'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const maxis_case_number_const  	 	= 0 '=  Maxis'
const member_number_const   		= 1 '=  Member Number
const client_first_name_const       = 2 '=  First Name MEMB
const client_last_name_const        = 3 '=  Last Name MEMB
const client_mid_name_const    	    = 4 '=  Middle initial MEMB
const client_DOB_const   		    = 5 '=  Date of Birth MEMB
const client_ssn_const		        = 6 '=  SSN
const client_age_const	            = 7 '=  age MEMB
const client_alia_name_const	    = 8 '=  ALIA Name
const client_alia_ssn_const		    = 9 '=  ALIA SSN

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
FOR EACH person IN HH_member_array
    CALL write_value_and_transmit(person, 20, 76) 'reads the reference number, last name, first name, and THEN puts it into an array YOU HAVENT defined the uc_members_array yet
    EMReadscreen ref_nbr, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen MEMB_number, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EMReadscreen client_SSN, 11, 7, 42
    If client_ssn = "___ __ ____" THEN client_ssn = ""
    last_name = trim(replace(last_name, "_", "")) & " "
    first_name = trim(replace(first_name, "_", "")) & " "
    mid_initial = replace(mid_initial, "_", "")
    EMReadScreen client_age, 2, 8, 76
    IF client_age = "  " THEN client_age = 0
    client_age = client_age * 1
    ReDim Preserve uc_members_array(client_alia_ssn_const, uc_membs)  'redimmed to the size of the last constant
    uc_members_array(member_number_const,     uc_membs) = ref_nbr
    uc_members_array(client_first_name_const, uc_membs) = first_name
    uc_members_array(client_last_name_const,  uc_membs) = last_name
    uc_members_array(client_mid_name_const,   uc_membs) = mid_initial
    uc_members_array(client_DOB_const,        uc_membs) = client_DOB
    uc_members_array(client_ssn_const,        uc_membs) = client_SSN
    uc_members_array(client_age_const,        uc_membs) = client_age
    uc_membs = uc_membs + 1
    STATS_counter = STATS_counter + 1
NEXT
'--------------------------------------------------------------------------------ALIA
CALL navigate_to_MAXIS_screen("STAT", "ALIA")
    FOR uc_membs = 0 to uBound(uc_members_array, 2)
    	CALL write_value_and_transmit(uc_members_array(member_number_const,     uc_membs), 20, 76)
        row = 7
        DO
            EMReadscreen alia_ref_num, 02, 04, 33
    	    EMReadScreen alia_last_name, 17, row, 26
            alia_last_name = replace(alia_last_name, "_", "")
            alia_last_name = trim(alia_last_name)
            EMReadScreen alia_first_name, 12, row, 53
            alia_first_name = replace(alia_first_name, "_", "")
            alia_first_name = trim(alia_first_name)
            EMReadScreen mid_initial, 1, row, 75
            IF alia_last_name = "" THEN EXIT DO
            row = row + 1
            uc_members_array(client_alia_name_const, uc_membs) = uc_members_array(client_alia_name_const, uc_membs) & ", " & alia_first_name & " " & alia_last_name
        Loop until row = 13

        row = 15
        Do
            EMReadScreen alia_client_SSN, 11, row, 28
            If alia_client_SSN = "___ __ ____" then
                alia_client_SSN = ""
                EXIT DO
            ELSE
                alia_client_SSN = trim(alia_client_SSN)
            END IF
            uc_members_array(client_alia_ssn_const, uc_membs) = uc_members_array(client_alia_ssn_const, uc_membs) & ", " & alia_client_SSN 'adding the varibale without writing over it each time i go throught the loop'
            EMReadScreen alia_client_SSN_II, 11, row, 53
            If alia_client_SSN_II = "___ __ ____" then
                alia_client_SSN_II = ""
                EXIT DO
            END IF
            row = row + 1
            uc_members_array(client_alia_ssn_const, uc_membs) = uc_members_array(client_alia_ssn_const, uc_membs) & ", " & alia_client_SSN_II
        Loop until row = 18
        IF left(uc_members_array(client_alia_name_const, uc_membs), 1) = "," THEN uc_members_array(client_alia_name_const, uc_membs) = right(uc_members_array(client_alia_name_const, uc_membs), len(uc_members_array(client_alia_name_const, uc_membs)) - 1)
        uc_members_array(client_alia_name_const, uc_membs) = trim(uc_members_array(client_alia_name_const, uc_membs)) 'once I have added everything to the array THEN i can format'
        IF left(uc_members_array(client_alia_ssn_const, uc_membs), 1) = "," THEN uc_members_array(client_alia_ssn_const, uc_membs) = right(uc_members_array(client_alia_ssn_const, uc_membs), len(uc_members_array(client_alia_ssn_const, uc_membs)) - 1)
        uc_members_array(client_alia_ssn_const, uc_membs) = trim(uc_members_array(client_alia_ssn_const, uc_membs))
        alia_first_name = ""
        alia_last_name = ""
        alia_client_SSN = ""
        alia_client_SSN_II = ""
    NEXT

FOR uc_membs = 0 to Ubound(uc_members_array, 2) 'start at the zero person and go to each of the selected people '
    member_info = member_info & "-------------"  & vbNewLine  & "Name of Resident: " & uc_members_array(client_first_name_const, uc_membs) & " " & uc_members_array(client_mid_name_const, uc_membs) & " " & uc_members_array(client_last_name_const, uc_membs)  &  vbCr & "DOB: " & uc_members_array(client_DOB_const,  uc_membs) &  vbcr & "SSN of Resident: " & uc_members_array(client_ssn_const,  uc_membs)
    IF trim(uc_members_array(client_alia_name_const, uc_membs)) <> "" THEN member_info = member_info & vbNewLine &  "ALIA Name: " & uc_members_array(client_alia_name_const, uc_membs)
    If trim(uc_members_array(client_alia_ssn_const,  uc_membs)) <> "" THEN member_info = member_info & vbNewLine & "ALIA SSN: " & uc_members_array(client_alia_ssn_const,  uc_membs)
    member_info = member_info & vbNewLine & "Please review deductions and withholdings for this individual. " & vbNewLine
NEXT

IF team_email_dropdown = "Central/NE" THEN team_email_dropdown = "HSPH.ES.DEED"
IF team_email_dropdown = "West" THEN team_email = "HSPH.ES.DEED"
IF team_email_dropdown = "Onboarding" THEN team_email = "HSPH.ES.DEED"
IF team_email_dropdown = "North" THEN team_email = "HSPH.ES.DEED"
IF team_email_dropdown = "South Suburban" THEN team_email = "HSPH.ES.DEED"
IF team_email_dropdown = "HC in METs" THEN team_email = "James.Berka@Hennepin.us; diane.beauchamp@hennepin.us"
IF team_email_dropdown = "Northwest" THEN team_email = "Shamikka.Lenear@Hennepin.us; Samantha.Haw@Hennepin.us"
IF team_email_dropdown = "QI" THEN team_email= "Mandora.Young@Hennepin.us; Laurie.Hennen@Hennepin.us"
IF team_email_dropdown = "South" THEN team_email = "Faduma.Abdi@Hennepin.us; Lindsey.Remus@Hennepin.us"
IF team_email_dropdown = "1800 Chicago" THEN team_email= "Jennifer.Moses@Hennepin.us"

IF cca_checkbox = CHECKED THEN member_info = "CCA Request" & vbNewLine & member_info
IF other_checkbox = CHECKED and other_check_editbox <> "" THEN member_info = "Other Request: " & other_check_editbox & vbNewLine & member_info

CALL find_user_name(the_person_running_the_script)' this is for the signature in the email'

'Creating the email
'Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachmentsend_email)
Call create_outlook_email(team_email, "", "UC Request for Case #" & MAXIS_case_number, member_info & vbNewLine & vbNewLine & "Submitted By: " & vbNewLine & the_person_running_the_script, "", TRUE)   'will create email, will send.

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/30/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/30/2021
'--Mandatory fields all present & Reviewed--------------------------------------08/30/2021
'--All variables in dialog match mandatory fields-------------------------------08/30/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/09/21
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------08/30/2021
'--Out-of-County handling reviewed----------------------------------------------08/30/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/09/21
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------09/09/21
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------08/30/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/30/2021
'--Comment Code-----------------------------------------------------------------09/09/21
'--Update Changelog for release/update------------------------------------------09/09/21
'--Remove testing message boxes-------------------------------------------------09/09/21
'--Remove testing code/unnecessary code-----------------------------------------09/09/21
'--Review/update SharePoint instructions----------------------------------------09/09/21
'--Review Best Practices using BZS page ----------------------------------------09/09/21
'--Review script information on SharePoint BZ Script List-----------------------09/09/21
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/09/21
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/09/21
'--Complete misc. documentation (if applicable)---------------------------------09/09/21
'--Update project team/issue contact (if applicable)----------------------------09/09/21
