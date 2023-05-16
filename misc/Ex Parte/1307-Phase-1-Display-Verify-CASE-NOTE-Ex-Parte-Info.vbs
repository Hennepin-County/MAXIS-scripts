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
    GroupBox 10, 5, 220, 45, "Person 1 - Case Information"
        Text 15, 20, 50, 10, "Person Name:"
        Text 65, 20, 75, 10, name_01
        Text 145, 20, 20, 10, "PMI:"
        Text 165, 20, 45, 10, PMI_01
        Text 15, 35, 50, 10, "Case Number:"
        Text 65, 35, 60, 10, case_number_01
        Text 145, 35, 50, 10, "Review Month:"
        Text 195, 35, 25, 10, review_month_01
    GroupBox 10, 60, 220, 60, "Person 1 - TPQY Information"
        Text 15, 75, 50, 10, "Claim Number:"
        Text 65, 75, 50, 10, claim_number_01
        Text 15, 90, 35, 10, "Sent Date:"
        Text 65, 90, 50, 10, sent_date_01
        Text 15, 105, 45, 10, "Return Date:"
        Text 65, 105, 50, 10, return_date_01
        Text 120, 75, 50, 10, "SDXS Amount:"
        Text 170, 75, 45, 10, sdxs_amount_01
        Text 120, 90, 50, 10, "BNDX Amount:"
        Text 170, 90, 45, 10, bndx_amount_01
        Text 120, 105, 35, 10, "MEDI Info: "
        Text 170, 105, 45, 10, medi_info_01
    ButtonGroup ButtonPressed
        Text 480, 5, 70, 10, "--- INSTRUCTIONS ---"
        PushButton 490, 15, 55, 15, "instructions", instructions_button
        Text 495, 40, 45, 10, "--- POLICY ---"
        PushButton 490, 50, 55, 15, policy_1, policy_1_button
        PushButton 490, 65, 55, 15, policy_2, policy_2_button
        PushButton 490, 80, 55, 15, policy_3, policy_3_button
        Text 490, 105, 55, 10, "--- NAVIGATE ---"
        PushButton 490, 115, 25, 15, "ACCI", acci_button
        PushButton 515, 115, 25, 15, "BILS", bils_button
        PushButton 490, 130, 25, 15, "BUDG", budg_button
        PushButton 515, 130, 25, 15, "BUSI", busi_button
        PushButton 490, 145, 25, 15, "DISA", disa_button
        PushButton 515, 145, 25, 15, "EMMA", emma_button
        PushButton 490, 160, 25, 15, "FACI", faci_button
        PushButton 515, 160, 25, 15, "HCMI", hcmi_button
        PushButton 490, 175, 25, 15, "IMIG", imig_button
        PushButton 515, 175, 25, 15, "INSA", insa_button
        PushButton 490, 190, 25, 15, "JOBS", jobs_button
        PushButton 515, 190, 25, 15, "LUMP", lump_button
        PushButton 490, 205, 25, 15, "MEDI", medi_button
        PushButton 515, 205, 25, 15, "MEMB", memb_button
        PushButton 490, 220, 25, 15, "MEMI", memi_button
        PushButton 515, 220, 25, 15, "PBEN", pben_button
        PushButton 490, 235, 25, 15, "PDED", pded_button
        PushButton 515, 235, 25, 15, "REVW", revw_button
        PushButton 490, 250, 25, 15, "SPON", spon_button
        PushButton 515, 250, 25, 15, "STWK", stwk_button
        PushButton 490, 265, 25, 15, "UNEA", unea_button
    GroupBox 10, 125, 220, 180, "Person 1 - Add'l Information"
        Text 15, 135, 15, 10, "SSI:"
        Text 90, 135, 80, 10, SSI_01
        Text 15, 150, 65, 10, "Other UNEA Types:"
        Text 90, 150, 80, 10, other_UNEA_types_01
        Text 15, 165, 50, 10, "JOBS Exists:"
        Text 90, 165, 80, 10, JOBS_01
        Text 15, 180, 60, 10, "MAXIS MA Basis:"
        Text 90, 180, 80, 10, MAXIS_MA_basis_01
        Text 15, 195, 60, 10, "MAXIS MSP Prog:"
        Text 90, 195, 80, 10, MAXIS_msp_prog_01
        Text 15, 210, 65, 10, "MAXIS MSP Basis:"
        Text 90, 210, 80, 10, MAXIS_msp_basis_01
        Text 15, 225, 55, 10, "MMIS MA Basis:"
        Text 90, 225, 80, 10, MMIS_ma_basis_01
        Text 15, 240, 60, 10, "MMIS MSP Prog:"
        Text 90, 240, 80, 10, MMIS_msp_prog_01
        Text 15, 255, 60, 10, "MMIS MSP Basis:"
        Text 90, 255, 80, 10, MMIS_msp_basis_01
        Text 15, 270, 70, 10, "MEDI - Part A Exists:"
        Text 90, 270, 80, 10, MEDI_part_a_01
        Text 15, 285, 70, 10, "MEDI - Part B Exists:"
        Text 90, 285, 80, 10, MEDI_part_b_01
    GroupBox 245, 5, 220, 45, "Person 2 - Case Information"
        Text 250, 20, 50, 10, "Person Name:"
        Text 300, 20, 75, 10, name_02
        Text 380, 20, 20, 10, "PMI:"
        Text 400, 20, 45, 10, PMI_02
        Text 250, 35, 50, 10, "Case Number:"
        Text 300, 35, 60, 10, case_number_02
        Text 380, 35, 50, 10, "Review Month:"
        Text 430, 35, 25, 10, review_month_02
    GroupBox 245, 60, 220, 60, "Person 2 - TPQY Information"
        Text 250, 75, 50, 10, "Claim Number:"
        Text 300, 75, 50, 10, claim_number_02
        Text 250, 90, 50, 10, "Sent Date:"
        Text 300, 90, 50, 10, sent_date_02
        Text 250, 105, 50, 10, "Return Date:"
        Text 300, 105, 50, 10, return_date_02
        Text 355, 75, 50, 10, "SDXS Amount:"
        Text 405, 75, 45, 10, sdxs_amount_02
        Text 355, 90, 50, 10, "BNDX Amount:"
        Text 405, 90, 45, 10, bndx_amount_02
        Text 355, 105, 35, 10, "MEDI Info: "
        Text 405, 105, 45, 10, medi_info_02
    GroupBox 245, 125, 220, 180, "Person 2 - Add'l Information"
        Text 255, 135, 15, 10, "SSI:"
        Text 330, 135, 80, 10, SSI_02
        Text 255, 150, 65, 10, "Other UNEA Types:"
        Text 330, 150, 80, 10, other_UNEA_types_02
        Text 255, 165, 50, 10, "JOBS Exists:"
        Text 330, 165, 80, 10, JOBS_02
        Text 255, 180, 60, 10, "MAXIS MA Basis:"
        Text 330, 180, 80, 10, MAXIS_MA_basis_02
        Text 255, 195, 60, 10, "MAXIS MSP Prog:"
        Text 330, 195, 80, 10, MAXIS_msp_prog_02
        Text 255, 210, 65, 10, "MAXIS MSP Basis:"
        Text 330, 210, 80, 10, MAXIS_msp_basis_02
        Text 255, 225, 55, 10, "MMIS MA Basis:"
        Text 330, 225, 80, 10, MMIS_ma_basis_02
        Text 255, 240, 60, 10, "MMIS MSP Prog:"
        Text 330, 240, 80, 10, MMIS_msp_prog_02
        Text 255, 255, 60, 10, "MMIS MSP Basis:"
        Text 330, 255, 80, 10, MMIS_msp_basis_02
        Text 255, 270, 70, 10, "MEDI - Part A Exists:"
        Text 330, 270, 80, 10, MEDI_part_a_02
        Text 255, 285, 70, 10, "MEDI - Part B Exists:"
        Text 330, 285, 80, 10, MEDI_part_b_02
    GroupBox 10, 310, 455, 50, "Ex Parte Determination"
        Text 15, 325, 85, 10, "Additional Notes:"
        Text 15, 340, 85, 10, "Ex Parte Determination:"
        EditBox 105, 320, 290, 15, additional_notes
        DropListBox 105, 340, 110, 50, ""+chr(9)+"Does Qualify for Ex Parte"+chr(9)+"Does Not Qualify for Ex Parte", ex_parte_determination
    Text 15, 365, 70, 10, "Worker Signature:"
    EditBox 80, 360, 110, 15, worker_signature
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
        'Function belows creates navigation to STAT panels for navigation buttons
        MAXIS_dialog_navigation

        'Add placeholder link to script instructions - To DO - update with correct link
        If ButtonPressed = instructions_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
        
        'Add placeholder links for policy buttons - TO DO - update with correct links
        If ButtonPressed = policy_1_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
        If ButtonPressed = policy_2_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
        If ButtonPressed = policy_3_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
        
        
        'Add validation to ensure ex parte determination is made
        If ex_parte_determination = "" THEN err_msg = err_msg & vbCr & "* You must make an ex parte determination." 

        'Add validation to ensure worker signature is not blank
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please include your worker signature."

        'Call validate_MAXIS_case_number(err_msg, "*") ' IF NEEDED
        'Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")   'IF NEEDED
        'The rest of the mandatory handling here
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = "" AND ButtonPressed = -1
    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

'Checks to see if in MAXIS
Call check_for_MAXIS(False)

'Do you need to check for PRIV status
'Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB")

'Do you need to check to see if case is out of county? Add Out-of-County handling here:
'All your other navigation, data catpure and logic here. any other logic or pre case noting actions here.

'Call MAXIS_background_check 'IF NEEDED: meaning if you send it through background. Move this to where it makes sense.

'Do you need to set a TIKL?
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)

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