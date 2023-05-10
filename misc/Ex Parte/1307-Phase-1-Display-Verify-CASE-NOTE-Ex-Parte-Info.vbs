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


'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabs the MAXIS case number automatically

Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.

BeginDialog Dialog1, 0, 0, 556, 385, "Phase 1 - Ex Parte Determination"
    GroupBox 10, 10, 220, 45, "Person 1 - Case Information"
        Text 15, 25, 50, 10, "Person Name:"
        Text 65, 25, 75, 10, name_01
        Text 145, 25, 20, 10, "PMI:"
        Text 165, 25, 45, 10, PMI_01
        Text 15, 40, 50, 10, "Case Number:"
        Text 65, 40, 60, 10, case_number_01
        Text 145, 40, 50, 10, "Review Month:"
        Text 195, 40, 25, 10, review_month_01
    GroupBox 10, 65, 220, 60, "Person 1 - TPQY Information"
        Text 15, 80, 50, 10, "Claim Number:"
        Text 65, 80, 50, 10, claim_number_01
        Text 25, 95, 50, 10, "Sent Date:"
        Text 65, 95, 50, 10, sent_date_01
        Text 20, 110, 50, 10, "Return Date:"
        Text 65, 110, 50, 10, return_date_01
        Text 120, 80, 50, 10, "SDXS Amount"
        Text 170, 80, 45, 10, sdxs_amount_01
        Text 120, 95, 50, 10, "BNDX Amount:"
        Text 170, 95, 45, 10, bndx_amount_01
        Text 135, 110, 35, 10, "MEDI Info: "
        Text 170, 110, 45, 10, medi_info_01
    ButtonGroup ButtonPressed
        Text 480, 5, 70, 10, "--- INSTRUCTIONS ---"
        PushButton 490, 20, 55, 15, instructions, Button7
        Text 495, 40, 45, 10, "--- POLICY ---"
        PushButton 490, 55, 55, 15, policy_1, Button5
        PushButton 490, 70, 55, 15, Policy_2, Button9
        PushButton 490, 85, 55, 15, Policy_3, Button8
        Text 490, 105, 55, 10, "--- NAVIGATE ---"
        PushButton 495, 120, 25, 15, REVW, Button3
        PushButton 520, 120, 25, 15, REVW, Button30
        PushButton 495, 135, 25, 15, REVW, Button31
        PushButton 520, 135, 25, 15, REVW, Button32
        PushButton 495, 150, 25, 15, REVW, Button33
        PushButton 520, 150, 25, 15, REVW, Button34
        PushButton 495, 165, 25, 15, REVW, Button35
        PushButton 520, 165, 25, 15, REVW, Button36
    GroupBox 10, 130, 220, 180, "Person 1 - Add'l Information"
        Text 15, 140, 105, 10, "Supplemental Security Income:"
        Text 120, 140, 80, 10, SSI_01
        Text 50, 155, 65, 10, "Other UNEA Types:"
        Text 120, 155, 80, 10, other_UNEA_types_01
        Text 70, 170, 50, 10, "JOBS Exists:"
        Text 120, 170, 80, 10, JOBS_01
        Text 55, 185, 60, 10, "MAXIS MA Basis:"
        Text 120, 185, 80, 10, MAXIS_MA_basis_01
        Text 55, 200, 60, 10, "MAXIS MSP Prog:"
        Text 120, 200, 80, 10, MAXIS_msp_prog_01
        Text 50, 215, 65, 10, "MAXIS MSP Basis:"
        Text 120, 215, 80, 10, MAXIS_msp_basis_01
        Text 55, 230, 55, 10, "MMIS MA Basis:"
        Text 120, 230, 80, 10, MMIS_ma_basis_01
        Text 55, 245, 60, 10, "MMIS MSP Prog:"
        Text 120, 245, 80, 10, MMIS_msp_prog_01
        Text 50, 260, 60, 10, "MMIS MSP Basis:"
        Text 120, 260, 80, 10, MMIS_msp_basis_01
        Text 40, 275, 70, 10, "MEDI - Part A Exists:"
        Text 120, 275, 80, 10, MEDI_part_a_01
        Text 40, 290, 70, 10, "MEDI - Part B Exists:"
        Text 120, 290, 80, 10, MEDI_part_b_01
    GroupBox 245, 10, 220, 45, "Person 2 - Case Information"
        Text 250, 25, 50, 10, "Person Name:"
        Text 300, 25, 75, 10, name_02
        Text 380, 25, 20, 10, "PMI:"
        Text 400, 25, 45, 10, PMI_02
        Text 250, 40, 50, 10, "Case Number:"
        Text 300, 40, 60, 10, case_number_02
        Text 380, 40, 50, 10, "Review Month:"
        Text 430, 40, 25, 10, review_month_02
    GroupBox 245, 65, 220, 60, "Person 2 - TPQY Information"
        Text 250, 80, 50, 10, "Claim Number:"
        Text 300, 80, 50, 10, claim_number_02
        Text 260, 95, 50, 10, "Sent Date:"
        Text 300, 95, 50, 10, sent_date_02
        Text 255, 110, 50, 10, "Return Date:"
        Text 300, 110, 50, 10, return_date_02
        Text 355, 80, 50, 10, "SDXS Amount"
        Text 405, 80, 45, 10, sdxs_amount_02
        Text 355, 95, 50, 10, "BNDX Amount:"
        Text 405, 95, 45, 10, bndx_amount_02
        Text 370, 110, 35, 10, "MEDI Info: "
        Text 405, 110, 45, 10, medi_info_02
    GroupBox 245, 130, 220, 180, "Person 2 - Add'l Information"
        Text 250, 140, 105, 10, "Supplemental Security Income:"
        Text 355, 140, 80, 10, SSI_02
        Text 285, 155, 65, 10, "Other UNEA Types:"
        Text 355, 155, 80, 10, other_UNEA_types_02
        Text 305, 170, 50, 10, "JOBS Exists:"
        Text 355, 170, 80, 10, JOBS_02
        Text 290, 185, 60, 10, "MAXIS MA Basis:"
        Text 355, 185, 80, 10, MAXIS_MA_basis_02
        Text 290, 200, 60, 10, "MAXIS MSP Prog:"
        Text 355, 200, 80, 10, MAXIS_msp_prog_02
        Text 285, 215, 65, 10, "MAXIS MSP Basis:"
        Text 355, 215, 80, 10, MAXIS_msp_basis_02
        Text 290, 230, 55, 10, "MMIS MA Basis:"
        Text 355, 230, 80, 10, MMIS_ma_basis_02
        Text 290, 245, 60, 10, "MMIS MSP Prog:"
        Text 355, 245, 80, 10, MMIS_msp_prog_02
        Text 285, 260, 60, 10, "MMIS MSP Basis:"
        Text 355, 260, 80, 10, MMIS_msp_basis_02
        Text 275, 275, 70, 10, "MEDI - Part A Exists:"
        Text 355, 275, 80, 10, MEDI_part_a_02
        Text 275, 290, 70, 10, "MEDI - Part B Exists:"
        Text 355, 290, 80, 10, MEDI_part_b_02
    ButtonGroup ButtonPressed
        OkButton 440, 365, 50, 15
        CancelButton 500, 365, 50, 15
EndDialog

'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation
        'Add in all of your mandatory field handling from your dialog here.
        Call validate_MAXIS_case_number(err_msg, "*") ' IF NEEDED
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")   'IF NEEDED
        'The rest of the mandatory handling here
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note." 'IF NEEDED
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

'Checks to see if in MAXIS
Call check_for_MAXIS(False)

'Do you need to check for PRIV status
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB")

'Do you need to check to see if case is out of county? Add Out-of-County handling here:
'All your other navigation, data catpure and logic here. any other logic or pre case noting actions here.

Call MAXIS_background_check 'IF NEEDED: meaning if you send it through background. Move this to where it makes sense.

'Do you need to set a TIKL?
Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)

'Now it navigates to a blank case note
Call start_a_blank_case_note

'...and enters a title (replace variables with your own content)...
CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")

'...some editboxes or droplistboxes (replace variables with your own content)...
CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )

'...checkbox responses (replace variables with your own content)...
If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )

'...and a worker signature.
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)
'leave the case note open and in edit mode unless you have a business reason not to (BULK scripts, multiple case notes, etc.)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")

'Add your closing issue documentation here. Make sure it's the most up-to-date version (date on file).