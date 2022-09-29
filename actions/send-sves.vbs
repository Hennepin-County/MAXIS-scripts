'Required for STATistical purposes==========================================================================================
name_of_script = "ACTIONS - SEND SVES.vbs"
start_time = timer
STATS_counter = 1                     	'sets the STATs counter at one
STATS_manualtime = 60                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF STATs block=========================================================================================================
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
		IF req.STATus = 200 THEN									'200 means great success
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
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley, Hennepin County") '#493
call changelog_update("02/22/2018", "Added option to send QURY for covered quarters.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Makes sure we're in MAXIS
Call check_for_MAXIS(False)

'Grabs case number
call MAXIS_case_number_finder(MAXIS_case_number)
member_number = "01"

'Shows and defines the initial dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 271, 105, "Send SVES Dialog"
  EditBox 90, 5, 60, 15, MAXIS_case_number
  EditBox 125, 25, 25, 15, member_number
  CheckBox 5, 50, 165, 10, "Check here to case note that a QURY was sent.", case_note_checkbox
  CheckBox 5, 65, 190, 10, "Check here to submit a query for covered quarters only.", quarters_check
  OptionGroup RadioGroup1
    RadioButton 190, 15, 60, 15, "SSN? (default)", SSN_radiobutton
    RadioButton 190, 30, 75, 15, "Claim # on UNEA?", UNEA_radiobutton
    RadioButton 190, 45, 75, 15, "Claim # on BNDX?", BNDX_radiobutton
  ButtonGroup ButtonPressed
    OkButton 160, 80, 50, 15
    CancelButton 215, 80, 50, 15
  EditBox 80, 80, 70, 15, worker_signature
  Text 5, 30, 120, 10, "Enter your member number (ex: 01): "
  Text 5, 85, 70, 10, "Sign your case note:"
  GroupBox 185, 5, 80, 55, "Number to use?"
  Text 5, 10, 80, 10, "Enter your case number:"
EndDialog

Do
    Do
        err_msg = ""
        Dialog Dialog1  					'Calling a dialog without a assigned variable will call the most recently defined dialog
        cancel_without_confirmation
        member_number = trim(member_number)
        call validate_MAXIS_case_number(err_msg, "*")
        If member_number = "" Then err_msg = err_msg & vbNewLine & "* Please enter the member number that needs a SVES sent."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        If err_msg <> "" Then MsgBox "*** Please Resolve to Continue: " & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Goes to MEMB to get info
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")

'Goes to the right HH member
EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
transmit 'This transmits to STAT/MEMB for the client indicated.
'If this member does not exist, this will stop the script from continuing.
EMReadScreen no_MEMB, 13, 8, 22
If no_MEMB = "Arrival Date:" then script_end_procedure("Error! This HH member does not exist.")
'Reads the SSN pieces and the PMI
EMReadScreen SSN1, 3, 7, 42
EMReadScreen SSN2, 2, 7, 46
EMReadScreen SSN3, 4, 7, 49
EMReadScreen PMI, 8, 4, 46

If SSN_radiobutton = 1 then
    call navigate_to_MAXIS_screen("INFC", "SVES")
    'checking for IRS non-disclosure agreement.
    EMReadScreen agreement_check, 9, 2, 24
    IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the systems and review the agreement.")
    EMWriteScreen SSN1,  4, 68
    EMWriteScreen SSN2,  4, 71
    EMWriteScreen SSN3,  4, 73
    EMWriteScreen PMI,  5, 68
    EMWriteScreen "QURY",  20, 70
    transmit 'Now we will enter the QURY screen to type the case number.
ElseIf UNEA_radiobutton = 1 then
    call navigate_to_MAXIS_screen("STAT", "UNEA")
    EMWriteScreen "UNEA", 20, 71 'It does this to move past error prone cases
    EMWriteScreen member_number, 20, 76 'It does this to make sure that it navigates to the right HH member.
    transmit 'This transmits to STAT/UNEA for the client indicated.
    EMReadScreen PMI, 8, 4, 71
    PMI = trim(PMI)
    'If there is no PMI, then there is no UNEA panel entered for the script to work off of.
    If PMI = "" then script_end_procedure("This HH member does not exist, or does not have a UNEA panel made.")
    EMReadScreen amt_of_UNEA_panels, 1, 2, 78
    If amt_of_UNEA_panels <> "1" then
        dialog_size_variable = (15 * cint(amt_of_UNEA_panels)) + 20
        Do
            EMReadScreen UNEA_type, 2, 5, 37
            EMReadScreen UNEA_claim_number, 15, 6, 37
            UNEA_claim_number = replace(UNEA_claim_number, "_", "")
            If UNEA_type_01 = "" then
                UNEA_type_01 = UNEA_type
                UNEA_claim_number_01 = UNEA_claim_number
            ElseIf UNEA_type_02 = "" then
                UNEA_type_02 = UNEA_type
                UNEA_claim_number_02 = UNEA_claim_number
            ElseIf UNEA_type_03 = "" then
                UNEA_type_03 = UNEA_type
                UNEA_claim_number_03 = UNEA_claim_number
            ElseIf UNEA_type_04 = "" then
                UNEA_type_04 = UNEA_type
                UNEA_claim_number_04 = UNEA_claim_number
            ElseIf UNEA_type_05 = "" then
                UNEA_type_05 = UNEA_type
                UNEA_claim_number_05 = UNEA_claim_number
            End if
            EMReadScreen current_UNEA_panel, 1, 2, 73
            If current_UNEA_panel <> amt_of_UNEA_panels then transmit
        Loop until current_UNEA_panel = amt_of_UNEA_panels
    	'The dialog the UNEA Claim option is defined here and then displayed
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 240, dialog_size_variable, "UNEA claim dialog"
          ButtonGroup ButtonPressed
            OkButton 185, 10, 50, 15
            CancelButton 185, 30, 50, 15
          Text 10, 5, 105, 10, "UNEA types to look at:"
          OptionGroup RadioGroup1
          If UNEA_type_01 <> "" then RadioButton 10, 20, 160, 10, "Type " & UNEA_type_01 & ", claim number " & UNEA_claim_number_01, UNEA_type_01_radiobutton
          If UNEA_type_02 <> "" then RadioButton 10, 35, 160, 10, "Type " & UNEA_type_02 & ", claim number " & UNEA_claim_number_02, UNEA_type_02_radiobutton
          If UNEA_type_03 <> "" then RadioButton 10, 50, 160, 10, "Type " & UNEA_type_03 & ", claim number " & UNEA_claim_number_03, UNEA_type_03_radiobutton
          If UNEA_type_04 <> "" then RadioButton 10, 65, 160, 10, "Type " & UNEA_type_04 & ", claim number " & UNEA_claim_number_04, UNEA_type_04_radiobutton
          If UNEA_type_05 <> "" then RadioButton 10, 80, 160, 10, "Type " & UNEA_type_05 & ", claim number " & UNEA_claim_number_05, UNEA_type_05_radiobutton
        EndDialog
        Dialog Dialog1  					'Calling a dialog without a assigned variable will call the most recently defined dialog
        cancel_without_confirmation
        If UNEA_type_01_radiobutton = 1 then
            claim_number = UNEA_claim_number_01
        ElseIf UNEA_type_02_radiobutton = 1 then
            claim_number = UNEA_claim_number_02
        ElseIf UNEA_type_03_radiobutton = 1 then
            claim_number = UNEA_claim_number_03
        ElseIf UNEA_type_04_radiobutton = 1 then
            claim_number = UNEA_claim_number_04
        ElseIf UNEA_type_05_radiobutton = 1 then
            claim_number = UNEA_claim_number_05
        End if
    Else
        EMReadScreen claim_number, 15, 6, 37
        claim_number = replace(claim_number, "_", "")
    End if

    call navigate_to_MAXIS_screen("INFC", "SVES")
    'checking for IRS non-disclosure agreement.
    EMReadScreen agreement_check, 9, 2, 24
    IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the systems and review the agreement.")
    EMWriteScreen PMI,  5, 68
    EMWriteScreen "QURY",  20, 70
    transmit 'Now we will enter the QURY screen to type the claim number.

    EMWriteScreen claim_number, 7, 38
ElseIf BNDX_radiobutton = 1 then
    call navigate_to_MAXIS_screen ("INFC", "____")
    EMWriteScreen SSN1,  4, 63
    EMWriteScreen SSN2,  4, 66
    EMWriteScreen SSN3,  4, 68
    EMWriteScreen "BNDX", 20, 71

	'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
	EMReadScreen agreement_check, 9, 2, 24
	IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
    transmit

    EMReadScreen BNDX_claim_number_01, 13, 5, 12
    EMReadScreen BNDX_claim_number_02, 13, 5, 38
    EMReadScreen BNDX_claim_number_03, 13, 5, 64
    If BNDX_claim_number_01 = "             " then script_end_procedure("BNDX claim number is not found. Was there a BNDX message for this client? Try sending using SSN or UNEA claim number.")
    If BNDX_claim_number_02 = "             " and BNDX_claim_number_03 = "             " then
        claim_number = replace(BNDX_claim_number_01, " ", "")
    Else
        'The dialog for the BNDX Claim option is defined here then displayed'
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 240, 70, "BNDX claim dialog"
          ButtonGroup ButtonPressed
            OkButton 185, 10, 50, 15
            CancelButton 185, 30, 50, 15
          Text 10, 5, 105, 10, "BNDX claims to look at:"
          OptionGroup RadioGroup1
          If BNDX_claim_number_01 <> "" then RadioButton 10, 20, 160, 10, BNDX_claim_number_01, BNDX_claim_number_01_radiobutton
          If BNDX_claim_number_02 <> "" then RadioButton 10, 35, 160, 10, BNDX_claim_number_02, BNDX_claim_number_02_radiobutton
          If BNDX_claim_number_03 <> "" then RadioButton 10, 50, 160, 10, BNDX_claim_number_03, BNDX_claim_number_03_radiobutton
        EndDialog

        Dialog Dialog1  					'Calling a dialog without a assigned variable will call the most recently defined dialog
        cancel_without_confirmation
        If BNDX_claim_number_01_radiobutton = 1 then
            claim_number = replace(BNDX_claim_number_01, " ", "")
        ElseIf BNDX_claim_number_02_radiobutton = 1 then
            claim_number = replace(BNDX_claim_number_02, " ", "")
        ElseIf BNDX_claim_number_03_radiobutton = 1 then
            claim_number = replace(BNDX_claim_number_03, " ", "")
        End if
    End if

    PF3
    EMWriteScreen "SVES", 20, 71
    transmit
    EMWriteScreen "QURY", 20, 70
    transmit
    'checking for IRS non-disclosure agreement.
    EMReadScreen agreement_check, 9, 2, 24
    IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the systems and review the agreement.")
    EMWriteScreen "_________", 5, 38
    EMWriteScreen claim_number, 7, 38
    EMWriteScreen "________", 9, 38
    EMWriteScreen PMI, 9, 38
End if
EMWriteScreen MAXIS_case_number, 11, 38
If quarters_check = checked then
    EMWriteScreen "y", 16, 38
else
    EMWriteScreen "y", 14, 38
End if

'Shuts down here if the user does not want to case note
If case_note_checkbox = unchecked then script_end_procedure("")
'Now it sends the SVES.
transmit

'Now it case notes
start_a_blank_CASE_NOTE
call write_variable_in_case_note("~~~SVES/QURY sent for MEMB " & member_number & "~~~")
If SSN_radiobutton = 1 then
	call write_variable_in_case_note("* Used SSN for QURY.")
ElseIf UNEA_radiobutton = 1 then
	call write_variable_in_case_note("* Used claim number on UNEA for QURY.")
ElseIf BNDX_radiobutton = 1 then
	call write_variable_in_case_note("* Used claim number on BNDX for QURY.")
End If
If quarters_check = checked then write_variable_in_case_note("* Query sent for covered quarters only.")
call write_variable_in_case_note("* QURY sent using script.")
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/25/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/25/2022
'--Mandatory fields all present & reviewed--------------------------------------09/16/2022
'--All variables in dialog match mandatory fields-------------------------------05/25/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/25/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/25/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/25/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/25/2022-------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------05/25/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/25/2022
'--Out-of-County handling reviewed----------------------------------------------05/25/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/25/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/25/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/25/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/25/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/25/2022
'--Denomination reviewed -------------------------------------------------------05/25/2022
'--Script name reviewed---------------------------------------------------------05/25/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/25/2022-------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/25/2022
'--comment Code-----------------------------------------------------------------05/25/2022
'--Update Changelog for release/update------------------------------------------05/16/2022
'--Remove testing message boxes-------------------------------------------------05/25/2022
'--Remove testing code/unnecessary code-----------------------------------------05/25/2022
'--Review/update SharePoint instructions----------------------------------------05/25/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/25/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/25/2022
'--Complete misc. documentation (if applicable)---------------------------------05/25/2022
'--Update project team/issue contact (if applicable)----------------------------05/25/2022
