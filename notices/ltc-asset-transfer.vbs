'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - LTC - ASSET TRANSFER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 130                      'manual run time in seconds
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
CALL changelog_update("05/01/2023", "Updated dialog with policy resources, added dialog autofill capability and automatic case noting.", "Megan Geissler, Hennepin County")
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
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
get_county_code


'Listing button numbers for links
onesource_URL = 101
epm_URL = 102
button_fill = 103

Do
    Do
      err_msg = ""
      fill_err_msg = ""

      Dialog1 = ""    
      BeginDialog Dialog1, 0, 0, 301, 140, "LTC asset transfer dialog"
        EditBox 75, 5, 50, 15, MAXIS_case_number
        ButtonGroup ButtonPressed
          PushButton 135, 5, 160, 15, "Press HERE to autofill fields below.", button_fill
        EditBox 75, 30, 20, 15, renewal_footer_month
        EditBox 100, 30, 20, 15, renewal_footer_year
        EditBox 75, 55, 70, 15, client
        EditBox 225, 55, 70, 15, spouse
        EditBox 75, 80, 220, 15, worker_signature
        ButtonGroup ButtonPressed
          OkButton 185, 120, 50, 15
          CancelButton 245, 120, 50, 15
          PushButton 20, 115, 40, 15, "OneSource", onesource_URL
          PushButton 70, 115, 40, 15, "EPM", epm_URL
        Text 30, 10, 45, 10, "Case number:"
        Text 15, 35, 60, 10, "ER date (MM/YY):"
        Text 5, 60, 70, 10, "  Resident First Name:"
        Text 160, 60, 65, 10, "Spouse First Name:"
        Text 15, 85, 60, 10, " Worker Signature:"
        GroupBox 15, 105, 100, 30, "LTC Asset Transfer Policies"
      EndDialog

      Dialog Dialog1
      cancel_without_confirmation
  
      If ButtonPressed > 100 Then
        err_msg = "Loop"

        If ButtonPressed = onesource_URL Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-170126"
        If ButtonPressed = epm_URL Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://hcopub.dhs.state.mn.us/epm/2_4_2_1_2.htm?rhhlterm=12%20community%20spouse&rhsearch=12%20months%20community%20spouse"
        If ButtonPressed = button_fill Then 
          Call Check_for_MAXIS(false)                         'Ensuring we are not passworded out  
          Call validate_MAXIS_case_number(fill_err_msg, "*")
          If fill_err_msg <> "" Then 
            MsgBox "Enter the required information specified below before selecting autofill:" & fill_err_msg
          Else
            Back_to_SELF
            Call navigate_to_MAXIS_screen("STAT","REVW")
            CALL write_value_and_transmit("X", 5, 71)
            EMReadScreen renewal_footer_month, 2, 9, 27
            Renewal_footer_month = Replace (renewal_footer_month, "_","")
            EMReadScreen renewal_footer_year, 2, 9, 33
            Renewal_footer_year = Replace (renewal_footer_year, "_","")
            PF3
            Call navigate_to_MAXIS_screen("STAT","MEMB")   
            EMReadScreen client, 12, 6 ,63 'Read client first name & input into dialog
            Client = Replace (Client, "_","")
            Call fix_case_for_name(client)
            Call navigate_to_MAXIS_screen("STAT","MEMI")
            EMReadScreen ref, 2, 9, 49
            If ref <> "__" Then
              Call navigate_to_MAXIS_screen("STAT","MEMB")
              CALL write_value_and_transmit(ref, 20, 76)
              EMReadScreen spouse, 12, 6, 63
              Spouse = Replace (Spouse, "_","")
              Call fix_case_for_name(spouse)
            End If
          End If
        End If
      Else  
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(renewal_footer_month, renewal_footer_year, err_msg, "*")
        If trim(client) = "" then err_msg = err_msg & vbNewLine & "* Enter the resident's first name."
        If trim(spouse) = "" then err_msg = err_msg & vbNewLine & "* Enter the spouse's first name."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      End If
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    Call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
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

script_end_procedure_with_error_report("**Review your MEMO and/or CASE:NOTE for accuracy.**")


'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/01/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/01/2023
'--Mandatory fields all present & Reviewed--------------------------------------05/01/2023
'--All variables in dialog match mandatory fields-------------------------------05/01/2023
'Review dialog names for content and content fit in dialog----------------------05/01/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/01/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------05/01/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/01/2023
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------05/01/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/01/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------05/01/2023
'--PRIV Case handling reviewed -------------------------------------------------05/01/2023
'--Out-of-County handling reviewed----------------------------------------------05/01/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/01/2023
'--BULK - review output of statistics and run time/count (if applicable)-------------------------------------------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/01/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/01/2023
'--Incrementors reviewed (if necessary)-----------------------------------------05/01/2023
'--Denomination reviewed -------------------------------------------------------05/01/2023
'--Script name reviewed---------------------------------------------------------05/01/2023
'--BULK - remove 1 incrementor at end of script reviewed-----------------------------------------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/01/2023
'--comment Code-----------------------------------------------------------------05/01/2023
'--Update Changelog for release/update------------------------------------------05/01/2023
'--Remove testing message boxes-------------------------------------------------05/01/2023
'--Remove testing code/unnecessary code-----------------------------------------05/01/2023
'--Review/update SharePoint instructions----------------------------------------05/01/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/01/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/01/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/01/2023
'--Complete misc. documentation (if applicable)---------------------------------05/01/2023
'--Update project team/issue contact (if applicable)----------------------------05/01/2023