'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - PA VERIF REQUEST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 294                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("03/02/2021", "BUG FIX - error for cases with a Significant Change detail in the budget. Added a fix to move past it.", "Casey Love, Hennepin County")
call changelog_update("11/12/2020", "Updated HSR Manual link for Data Privacy due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("07/29/2020", "Removed the 'PRINT' default of the document at the end of the script run because we are not currently in the office.", "Casey Love, Hennepin County")
call changelog_update("07/29/2020", "Removed the option to include income information from MAXIS in the document. The official policy and process needs to be followed for this type of information. Added a button to open the HSR Manual page.", "Casey Love, Hennepin County")
call changelog_update("01/15/2019", "Updated to accomodate benefits larger than $1,000 for SNAP, MFIP, and DWP.", "Casey Love, Hennepin County")
call changelog_update("12/01/2016", "Checkbox added with the option to have 'Other Income' not listed on the word document.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


function check_if_mmis_in_session(mmis_running, maxis_region)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		mmis_running = True
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			mmis_running = True
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				mmis_running = False
			ELSE
				mmis_running = True
			END IF
		END IF
	END IF
	If maxis_region = "PRODUCTION" Then EMWriteScreen "1", 2, 15
	If maxis_region = "INQUIRY DB" Then EMWriteScreen "2", 2, 15
	If maxis_region = "TRAINING" Then EMWriteScreen "3", 2, 15
	transmit
end function


Function Create_List_Of_Notices(notice_panel, notices_array, selected_const, information_const, WCOM_row_const, no_notices, specific_prog)
	Erase notices_array
	no_notices = FALSE
	If notice_panel = "WCOM" Then
		wcom_row = 7
		array_counter = 0
		Do
			EMReadScreen notice_prog, 2,  wcom_row, 26
			save_this_notc = True
			If specific_prog <> "" Then
				If notice_prog <> specific_prog Then save_this_notc = False
			End If
			If save_this_notc = True Then
				ReDim Preserve notices_array(3, array_counter)
				EMReadScreen notice_date, 8,  wcom_row, 16
				EMReadScreen notice_info, 31, wcom_row, 30
				EMReadScreen notice_stat, 8,  wcom_row, 71

				notice_date = trim(notice_date)
				notice_prog = trim(notice_prog)
				notice_info = trim(notice_info)
				notice_stat = trim(notice_stat)

				notices_array(selected_const,    array_counter) = unchecked
				notices_array(information_const, array_counter) = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
				notices_array(WCOM_row_const,    array_counter) = wcom_row

				array_counter = array_counter + 1
			End If
			wcom_row = wcom_row + 1

			EMReadScreen next_notice, 4, wcom_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If

	If notice_panel = "MEMO" Then
		memo_row = 7
		array_counter = 0
		Do
			ReDim Preserve notices_array(3, array_counter)
			EMReadScreen notice_date, 8,  memo_row, 19
			EMReadScreen notice_info, 31, memo_row, 29
			EMReadScreen notice_stat, 8,  memo_row, 67

			notice_date = trim(notice_date)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			notices_array(selected_const,    array_counter) = unchecked
			notices_array(information_const, array_counter) = notice_info & " - " & notice_date & " - Status: " & notice_stat
			notices_array(WCOM_row_const,    array_counter) = memo_row

			array_counter = array_counter + 1
			memo_row = memo_row + 1

			EMReadScreen next_notice, 4, memo_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If
	If array_counter = 0 Then no_notices = TRUE
End Function

function leave_notice_text(ask_first)
	EMReadScreen notice_indicator, 6, 2, 72

	If notice_indicator = "FMINFO" Then
		If ask_first = True Then ask_to_leave_msg = MsgBox("It appears we are in a notice text, in order to contine, we must leave th notice text." & vbCr & vbCr & "Is it alright to leave to notice text now?", vbQuestion + vbYesNo, "Leave Notice Text")
		If ask_to_leave_msg = vbYes OR ask_first = False Then  PF3
	End If
end function

function sort_dates(dates_array)
'--- Takes an array of dates and reorders them to be chronological.
'~~~~~ dates_array: an array of dates only
'===== Keywords: MAXIS, date, order, list, array
    dim ordered_dates ()
    redim ordered_dates(0)
    original_array_items_used = "~"
    days =  0
    do

        prev_date = ""
        original_array_index = 0
        for each thing in dates_array
            check_this_date = TRUE
            new_array_index = 0
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                new_array_index = new_array_index + 1
                ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "check this date - " & check_this_date
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                    index_used = original_array_index
                Else
                    if DateDiff("d", prev_date, thing) < 0 then
                        prev_date = thing
                        index_used = original_array_index
                    end if
                end if
            end if
            original_array_index = original_array_index + 1
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            original_array_items_used = original_array_items_used & index_used & "~"
            days = days + 1
        end if
        counter = 0
        For each thing in dates_array
            If InStr(original_array_items_used, "~" & counter & "~") = 0 Then
                For each new_date_thing in ordered_dates
                    If thing = new_date_thing Then
                        original_array_items_used = original_array_items_used & counter & "~"
                        days = days + 1
                    End If
                Next
            End If
            counter = counter + 1
        Next
        ' MsgBox "Ordered Dates array - " & join(ordered_dates, ", ") & vbCR & "days - " & days & vbCR & "Ubound - " & UBOUND(dates_array) & vbCR & "used list - " & original_array_items_used
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

function Select_New_WCOM(notices_array, selected_const, information_const, WCOM_row_const, case_number_known, allow_wcom, allow_memo, notc_month, notc_year, no_notices, specific_prog, allow_multiple_notc, allow_cancel)
	If allow_wcom = True AND allow_memo = True Then
		notice_panel = "Select One..."
	ElseIf allow_wcom = True Then
		notice_panel = "WCOM"
	ElseIf allow_memo = True Then
		notice_panel = "MEMO"
	End If
	Do
	    Do
	    	err_msg = ""

	    	dlg_y_pos = 85
	    	dlg_length = 145
			If no_notices = False Then dlg_length = dlg_length + (UBound(notices_array, 2) * 20)

	        Dialog1 = ""
	    	BeginDialog Dialog1, 0, 0, 205, dlg_length, "Notices to Print"
	    	  Text 5, 10, 50, 10, "Case Number"
	    	  If case_number_known = False Then EditBox 65, 5, 50, 15, MAXIS_case_number
			  If case_number_known = True Then Text 65, 10, 50, 15, MAXIS_case_number
	    	  Text 5, 30, 130, 10, "Where is the notice you want to print?"
	    	  If allow_wcom = True AND allow_memo = True Then
			      DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"WCOM"+chr(9)+"MEMO", notice_panel
			  ElseIf allow_wcom = True Then
			  	  DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"WCOM", notice_panel
			  ElseIf allow_memo = True Then
			  	  DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"MEMO", notice_panel
			  End If
	    	  Text 5, 50, 120, 10, "In which month was the notice sent?"
	    	  EditBox 140, 45, 20, 15, notc_month
	    	  EditBox 165, 45, 20, 15, notc_year
	    	  ButtonGroup ButtonPressed
	    	    PushButton 60, 70, 50, 10, "Find Notices", find_notices_button
	    	  If no_notices = FALSE Then
	    		  For notices_listed = 0 to UBound(notices_array, 2)
	    		  	CheckBox 10, dlg_y_pos, 185, 10, notices_array(information_const, notices_listed), notices_array(selected_const, notices_listed)
	    			dlg_y_pos = dlg_y_pos + 15
	    		  Next
	    	  Else
	    	  	  Text 10, dlg_y_pos, 185, 10, "**No Notices could be found here.**"
	    		  dlg_y_pos = dlg_y_pos + 15
	    	  End If
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  If case_number_known = False Then EditBox 75, dlg_y_pos, 125, 15, worker_signature
			  If case_number_known = True Then Text 80, dlg_y_pos, 125, 15, worker_signature
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  Text 5, dlg_y_pos, 60, 10, "Worker Signature:"
	    	  dlg_y_pos = dlg_y_pos + 15
	    	  If allow_cancel = True Then
				  ButtonGroup ButtonPressed
		    	    OkButton 100, dlg_y_pos, 50, 15
		    	    CancelButton 150, dlg_y_pos, 50, 15
			  Else
				  ButtonGroup ButtonPressed
					OkButton 150, dlg_y_pos, 50, 15
			  End If
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  If case_number_known = False Then CheckBox 5, dlg_y_pos, 90, 10, "Check here to case note.", case_note_check
	    	EndDialog

	    	Dialog Dialog1
	    	If allow_cancel = True Then cancel_confirmation

	    	notice_selected = FALSE
			If no_notices = False Then
		    	For notice_to_print = 0 to UBound(notices_array, 2)
		    		If notices_array(selected_const, notice_to_print) = checked Then
						If allow_multiple_notc = False AND notice_selected = TRUE AND InStr(err_msg, "One one NOTICE can be selected.") = 0 Then err_msg = err_msg & vbNewLine &  "- One one NOTICE can be selected."
						notice_selected = TRUE
					End If
		    	Next
			End If

			If case_number_known = False Then
	    		If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
			End If
	    	If notice_panel = "Select One..." Then err_msg = err_msg & vbNewLine & "- Select where the notice to print is."
	    	If notc_month = "" or notc_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
	    	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice to be copied to a Word Document."


	    	If ButtonPressed = find_notices_button then
	    		If notice_panel <> "Select One..." AND MAXIS_case_number <> "" AND notc_month <> "" AND notc_year <> "" Then
	    			Call navigate_to_MAXIS_screen ("SPEC", notice_panel)
	    			If notice_panel = "MEMO" then
	    				EMWriteScreen notc_month, 3, 48
	    				EMWriteScreen MAXIS_footer_year, 3, 53
	    			ElseIf notice_panel = "WCOM" Then
	    				EMWriteScreen notc_month, 3, 46
	    				EMWriteScreen notc_year, 3, 51
	    			End If
	    			transmit
					Call Create_List_Of_Notices(notice_panel, notices_array, selected_const, information_const, WCOM_row_const, no_notices, specific_prog)
	    			err_msg = "LOOP"
	    		Else
	    			err_msg = err_msg & vbNewLine & "!!! Cannot read a list of notices without a panel selected, a case number entered, and footer month & year entered !!!"
	    		End If
	    	End If

	    	If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg

	    Loop Until err_msg = ""
	    Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
end function

'WHAT DO WE NEED TO START WITH?
'Gather the Case Number
'Determine Contact Type
'Determine which programs are being requested
	'If HC - determine the individual the request is for
'If HC - determine if MAXIS or METS
'Determine if we need Current Benefit or Issuance over a Time period.


'WHAT SHOULD THE SCRIPT DO?
'Resend an ELIGIBILITY NOTICE
'Create Word Doc of the ELIGIBILITY NOTICE
'Create a SPEC/MEMO with information requested
'Create Word Doc of the SPEC/MEMO
'Note the information requested and how provided.


'FIRST DIALOG - CN & Contact type - this will NOT be in a function since if placed in another script - this will already be known
'Ask if the request is from PHA
'Ask if the request is for medical payment history'


'Script will review case to determine:
'program status
'Most recent ELIG NOTICE
'Current benefit amount
'Determine in MMIS is running


'SECOND DIALOG - display the found information
'ask the questions about what is needed and how to outpuut

'CREATE NOTICES OR RESEND

'NOTING

EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
clt_in_person = FALSE
check_for_MAXIS(False)


Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 301, 130, "Verification of Public Assistance"
		  EditBox 85, 50, 60, 15, MAXIS_case_number
		  DropListBox 85, 70, 210, 45, "Resident on the Phone (or AREP)"+chr(9)+"Resident in Person (or AREP)"+chr(9)+"PHA (Public Housing form)"+chr(9)+"Request of Medical Payment History (from Resident or AREP)", contact_type
		  EditBox 85, 90, 210, 15, worker_signature
		  ButtonGroup ButtonPressed
		    OkButton 195, 110, 50, 15
		    CancelButton 245, 110, 50, 15
		    PushButton 120, 30, 175, 15, "HSR Manual for Verification of Public Assistance", verif_pa_hsr_manual_btn
		  Text 10, 10, 290, 20, "PA Verif Request script will assist you in following the procedure for Client Requests of their Public Assistance benefits. Details of this process can be found in the HSR Manual."
		  Text 30, 55, 50, 10, "Case Number:"
		  Text 15, 75, 65, 10, "Source of Request:"
		  Text 20, 95, 60, 10, "Worker Signature:"
		EndDialog

		dialog Dialog1
		cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")
		If ButtonPressed = verif_pa_hsr_manual_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Verification-of-public-assistance.aspx"
			err_msg = "LOOP"
		End If

		If err_msg <> "LOOP" and err_msg <> "" Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

If contact_type = "PHA (Public Housing form)" Then
End If
If contact_type = "Request of Medical Payment History (from Resident or AREP)" Then
End If
If contact_type = "Documents from ECF" Then
End If


Call back_to_SELF
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If
' Call check_if_mmis_in_session(mmis_running, MX_region)

If contact_type = "Resident in Person (or AREP)" Then clt_in_person = True

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)

If ga_status = "ACTIVE" Then

	Call navigate_to_MAXIS_screen("MONY", "INQB")
	inqb_row = 6
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23
		If inqb_program = "GA" Then
			EMReadScreen ga_amount, 10, 6, 38
			ga_amount = trim(ga_amount)
			Exit Do
		End If
		inqb_row = inqb_row + 1
	Loop until inqb_program = "  "

	Call back_to_SELF

 	Call navigate_to_MAXIS_screen("ELIG", "GA")
	EMWriteScreen "99", 20, 78
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""
	EMReadScreen ga_approved_date, 8, status_row, 26
	ga_approved_date = DateAdd("m", 1, ga_approved_date)
	ga_month = right("00" & DatePart("m", ga_approved_date), 2)
	ga_year = right(DatePart("yyyy", ga_approved_date), 2)

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	EMWriteScreen ga_month, 3, 46
	EMWriteScreen ga_year, 3, 51
	transmit

	wcom_row = 7
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "GA" AND InStr(notc_title, "ELIG") <> 0 Then
			ga_wcom_row = wcom_row
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			ga_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR ga_wcom_text <> ""
	If ga_wcom_text = "" Then ga_wcom_text = "NO WCOM Found"

End If

If msa_status = "ACTIVE" Then
	Call navigate_to_MAXIS_screen("MONY", "INQB")
	inqb_row = 6
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23
		If inqb_program = "MS" Then
			EMReadScreen msa_amount, 10, 6, 38
			msa_amount = trim(msa_amount)
			Exit Do
		End If
		inqb_row = inqb_row + 1
	Loop until inqb_program = "  "

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("ELIG", "MSA")
	EMWriteScreen "99", 20, 79
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""
	EMReadScreen msa_approved_date, 8, status_row, 26
	msa_approved_date = DateAdd("m", 1, msa_approved_date)
	msa_month = right("00" & DatePart("m", msa_approved_date), 2)
	msa_year = right(DatePart("yyyy", msa_approved_date), 2)
	transmit

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	EMWriteScreen msa_month, 3, 46
	EMWriteScreen msa_year, 3, 51
	transmit

	wcom_row = 7
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "MS" AND InStr(notc_title, "ELIG") <> 0 Then
			ga_wcom_row = wcom_row
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			msa_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR msa_wcom_text <> ""
	If msa_wcom_text = "" Then msa_wcom_text = "NO WCOM Found"
End If

If mfip_status = "ACTIVE" Then
	Call navigate_to_MAXIS_screen("MONY", "INQB")
	inqb_row = 6
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23
		If inqb_program = "MF" Then
			EMReadScreen mfip_amount, 10, 6, 38
			mfip_amount = trim(mfip_amount)
			Exit Do
		End If
		inqb_row = inqb_row + 1
	Loop until inqb_program = "  "

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("ELIG", "MFIP")
	EMWriteScreen "99", 20, 79
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""
	EMReadScreen mfip_approved_date, 8, status_row, 26
	mfip_approved_date = DateAdd("m", 1, mfip_approved_date)
	mfip_month = right("00" & DatePart("m", mfip_approved_date), 2)
	mfip_year = right(DatePart("yyyy", mfip_approved_date), 2)
	transmit

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	EMWriteScreen mfip_month, 3, 46
	EMWriteScreen mfip_year, 3, 51
	transmit

	wcom_row = 7
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "MF" AND InStr(notc_title, "ELIG") <> 0 Then
			ga_wcom_row = wcom_row
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			mfip_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR mfip_wcom_text <> ""
	If mfip_wcom_text = "" Then mfip_wcom_text = "NO WCOM Found"
End If
If dwp_status = "ACTIVE" Then
End If

'If grh_status,
If snap_status = "ACTIVE" Then
	Call navigate_to_MAXIS_screen("MONY", "INQB")
	inqb_row = 6
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23
		If inqb_program = "FS" Then
			EMReadScreen snap_amount, 10, 6, 38
			snap_amount = trim(snap_amount)
			Exit Do
		End If
		inqb_row = inqb_row + 1
	Loop until inqb_program = "  "

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("ELIG", "FS")
	EMWriteScreen "99", 19, 78
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""
	EMReadScreen snap_approved_date, 8, status_row, 26
	snap_approved_date = DateAdd("m", 1, snap_approved_date)
	snap_month = right("00" & DatePart("m", snap_approved_date), 2)
	snap_year = right(DatePart("yyyy", snap_approved_date), 2)

	Call back_to_SELF

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")
	EMWriteScreen snap_month, 3, 46
	EMWriteScreen snap_year, 3, 51
	transmit

	wcom_row = 7
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "FS" AND InStr(notc_title, "ELIG") <> 0 Then
			snap_wcom_row = wcom_row
			snap_wcom_position = wcom_row - 6
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			snap_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR snap_wcom_text <> ""
	If snap_wcom_text = "" Then snap_wcom_text = "NO WCOM Found"
End If

If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then
End If

snap_prog_history_exists = False
ga_prog_history_exists = False
msa_prog_history_exists = False
mfip_prog_history_exists = False
dwp_prog_history_exists = False
grh_prog_history_exists = False

Call navigate_to_MAXIS_screen("CASE", "CURR")
EMWriteScreen "X", 4, 9
transmit

If snap_status <> "ACTIVE" Then
	EMWriteScreen "FS", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then snap_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If ga_status <> "ACTIVE" Then
	EMWriteScreen "GA", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then ga_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If msa_status <> "ACTIVE" Then
	EMWriteScreen "MS", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then msa_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If mfip_status <> "ACTIVE" Then
	EMWriteScreen "MF", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then mfip_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If dwp_status <> "ACTIVE" Then
	EMWriteScreen "DW", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then dwp_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If grh_status <> "ACTIVE" Then
	EMWriteScreen "GR", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then grh_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
Call back_to_SELF
' MsgBox "GA" & vbCr & "GA Amount - " & ga_amount & vbCr & "GA WCOM row - " & ga_wcom_row & vbCr & "GA WCOM position - "  & ga_wcom_position & vbCr &  "GA WCOM:" & vbCr & ga_wcom_text & vbCr & vbCr &_
	 ' "SNAP" & vbCr & "FS Amount - " & snap_amount & vbCr & "FS WCOM row - " & snap_wcom_row & vbCr & "FS WCOM position - "  & snap_wcom_position & vbCr &  "FS WCOM:" & vbCr & snap_wcom_text

snap_change_wcom_btn = 1010
ga_change_wcom_btn   = 1020
msa_change_wcom_btn  = 1030
mfip_change_wcom_btn = 1040
dwp_change_wcom_btn  = 1050
grh_change_wcom_btn  = 1060
hc_change_wcom_btn   = 1070

snap_wcom_btn = 110
ga_wcom_btn   = 120
msa_wcom_btn  = 130
mfip_wcom_btn = 140
dwp_wcom_btn  = 150
grh_wcom_btn  = 160
hc_wcom_btn   = 170

snap_program_history_button = 51
ga_program_history_button 	= 52
msa_program_history_button 	= 53
mfip_program_history_button = 54
dwp_program_history_button 	= 55
grh_program_history_button 	= 56
hc_program_history_button 	= 57

Dim notices_array()
ReDim notices_array(3,0)

Const selected = 0
Const information = 1
Const WCOM_search_row = 2

 Do
 	Do
 		err_msg = ""
		y_pos = 30

 		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 551, 385, "Verification of Public Assistance"
		  ButtonGroup ButtonPressed
			If snap_status = "ACTIVE" Then
				GroupBox 15, y_pos, 450, 75, "SNAP"
				y_pos = y_pos + 15
				Text 20, y_pos, 120, 10, "SNAP Assistance Verification to be sent via "
				DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of SNAP Needed", snap_verification_method
				y_pos = y_pos + 10
				Text 25, y_pos, 200, 10, "SNAP current benefit amount appears to be $" & snap_amount & "."
				y_pos = y_pos + 10
				Text 25, y_pos, 400, 10, "Most recent SNAP Eligibility Notice appears to have been sent for benefit month: " & snap_month & "/" & snap_year & ". WCOM Information:"
				y_pos = y_pos + 10
				Text 30, y_pos, 200, 10, snap_wcom_text
				PushButton 225, y_pos, 100, 10, "Go To this WCOM", snap_wcom_btn
				PushButton 330, y_pos, 100, 10, "Select Different WCOM", snap_change_wcom_btn
				y_pos = y_pos + 15
				Text 25, y_pos, 105, 10, "Date range of issuance needed:"
				EditBox 130, y_pos - 5, 30, 15, snap_start_month
				Text 160, y_pos, 5, 10, "---"
				EditBox 165, y_pos - 5, 30, 15, snap_end_month
				Text 200, y_pos, 100, 10, "(use mm/yy format)"
				y_pos = y_pos + 15
			End If
			If ga_status = "ACTIVE" Then
				GroupBox 15, y_pos, 450, 75, "GA"
				y_pos = y_pos + 15
				Text 20, y_pos, 120, 10, "GA Assistance Verification to be sent via "
				DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of GA Needed", ga_verification_method
				y_pos = y_pos + 10
				Text 25, y_pos, 200, 10, "GA current benefit amount appears to be $" & ga_amount & "."
				y_pos = y_pos + 10
				Text 25, y_pos, 400, 10, "Most recent GA Eligibility Notice appears to have been sent for benefit month: " & ga_month & "/" & ga_year & ". WCOM Information:"
				y_pos = y_pos + 10
				Text 30, y_pos, 200, 10, ga_wcom_text
				PushButton 225, y_pos, 100, 10, "Go To WCOM", ga_wcom_btn
				PushButton 330, y_pos, 100, 10, "Select Different WCOM", ga_change_wcom_btn
				y_pos = y_pos + 15
				Text 25, y_pos, 105, 10, "Date range of issuance needed:"
				EditBox 130, y_pos - 5, 30, 15, ga_start_month
				Text 160, y_pos, 5, 10, "---"
				EditBox 165, y_pos - 5, 30, 15,ga_end_month
				Text 200, y_pos, 100, 10, "(use mm/yy format)"
				y_pos = y_pos + 15
			End If
			If msa_status = "ACTIVE" Then
				GroupBox 15, y_pos, 450, 75, "MSA"
				y_pos = y_pos + 15
				Text 20, y_pos, 120, 10, "MSA Assistance Verification to be sent via "
				DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of MSA Needed", msa_verification_method
				y_pos = y_pos + 10
				Text 25, y_pos, 200, 10, "MSA current benefit amount appears to be $" & msa_amount & "."
				y_pos = y_pos + 10
				Text 25, y_pos, 400, 10, "Most recent MSA Eligibility Notice appears to have been sent for benefit month: " & msa_month & "/" & msa_year & ". WCOM Information:"
				y_pos = y_pos + 10
				Text 30, y_pos, 200, 10, msa_wcom_text
				PushButton 225, y_pos, 100, 10, "Go To WCOM", msa_wcom_btn
				PushButton 330, y_pos, 100, 10, "Select Different WCOM", msa_change_wcom_btn
				y_pos = y_pos + 15
				Text 25, y_pos, 105, 10, "Date range of issuance needed:"
				EditBox 130, y_pos - 5, 30, 15, msa_start_month
				Text 160, y_pos, 5, 10, "---"
				EditBox 165, y_pos - 5, 30, 15,msa_end_month
				Text 200, y_pos, 100, 10, "(use mm/yy format)"
				y_pos = y_pos + 15
			End If
			If mfip_status = "ACTIVE" Then
				GroupBox 15, y_pos, 450, 75, "MFIP"
				y_pos = y_pos + 15
				Text 20, y_pos, 120, 10, "MFIP Assistance Verification to be sent via "
				DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of MFIP Needed", mfip_verification_method
				y_pos = y_pos + 10
				Text 25, y_pos, 200, 10, "MFIP current benefit amount appears to be $" & mfip_amount & "."
				y_pos = y_pos + 10
				Text 25, y_pos, 400, 10, "Most recent MFIP Eligibility Notice appears to have been sent for benefit month: " & mfip_month & "/" & mfip_year & ". WCOM Information:"
				y_pos = y_pos + 10
				Text 30, y_pos, 200, 10, mfip_wcom_text
				PushButton 225, y_pos, 100, 10, "Go To WCOM", mfip_wcom_btn
				PushButton 330, y_pos, 100, 10, "Select Different WCOM", mfip_change_wcom_btn
				y_pos = y_pos + 15
				Text 25, y_pos, 105, 10, "Date range of issuance needed:"
				EditBox 130, y_pos - 5, 30, 15, mfip_start_month
				Text 160, y_pos, 5, 10, "---"
				EditBox 165, y_pos - 5, 30, 15,mfip_end_month
				Text 200, y_pos, 100, 10, "(use mm/yy format)"
				y_pos = y_pos + 15
			End If
			If dwp_status = "ACTIVE" Then
			End If
			If grh_status = "ACTIVE" Then
			End If
			If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then
			End If
			' Text 20, 300, 200, 10, "Select the method of Notification:"
			' DropListBox 225, 295, 100, 45, "Select One..."+chr(9)+"Resend Eligibility Notices"+chr(9)+"Create new WCOM with Details", verification_method_selection
			y_pos = y_pos + 5

			If snap_status <> "ACTIVE" Then
				If snap_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "SNAP is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View SNAP Program History", snap_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of SNAP benefits issued from ", snap_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, snap_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, snap_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "SNAP is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If ga_status <> "ACTIVE" Then
				If ga_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "GA is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View GA Program History", ga_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of GA benefits issued from ", ga_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, ga_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, ga_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "GA is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If msa_status <> "ACTIVE" Then
				If msa_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "MSA is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View MSA Program History", msa_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of MSA benefits issued from ", msa_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, msa_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, msa_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "MSA is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If mfip_status <> "ACTIVE" Then
				If mfip_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "MFIP is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View MFIP Program History", mfip_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of MFIP benefits issued from ", mfip_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, mfip_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, mfip_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "MFIP is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If dwp_status <> "ACTIVE" Then
				If dwp_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "DWP is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View DWP Program History", dwp_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of DWP benefits issued from ", dwp_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, dwp_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, dwp_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "DWP is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If grh_status <> "ACTIVE" Then
				If grh_prog_history_exists = True Then
					Text 20, y_pos, 100, 10, "GRH is NOT currently Active"
					PushButton 120, y_pos-2, 100, 13, "View GRH Program History", grh_program_history_button
					y_pos = y_pos + 15
					CheckBox 25, y_pos, 210, 10, "Check here to include amounts of GRH benefits issued from ", grh_not_actv_memo_for_old_beneftis_checkbox
					EditBox 235, y_pos - 5, 30, 15, grh_start_month
					Text 265, y_pos, 5, 10, "---"
					EditBox 270, y_pos - 5, 30, 15, grh_end_month
					Text 305, y_pos, 100, 10, "(use mm/yy format)"
					y_pos = y_pos + 15
				Else
					Text 20, y_pos, 300, 10, "GRH is NOT currently Active and there is no ACTIVE Program history for this case."
					y_pos = y_pos + 15
				End If
			End If
			If ma_status <> "ACTIVE" AND msp_status <> "ACTIVE" Then
			End If


			OkButton 445, 365, 50, 15
			CancelButton 495, 365, 50, 15
			PushButton 35, 345, 25, 10, "CURR", CASE_CURR_button
		    PushButton 60, 345, 25, 10, "PERS", CASE_PERS_button
		    PushButton 85, 345, 25, 10, "NOTE", CASE_NOTE_button
		    PushButton 160, 345, 25, 10, "XFER", SPEC_XFER_button
		    PushButton 185, 345, 25, 10, "WCOM", SPEC_WCOM_button
		    PushButton 210, 345, 25, 10, "MEMO", SPEC_MEMO_button
		    PushButton 35, 355, 25, 10, "PROG", PROG_button
		    PushButton 60, 355, 25, 10, "MEMB", MEMB_button
		    PushButton 85, 355, 25, 10, "REVW", REVW_button
		    PushButton 160, 355, 25, 10, "INQB", MONY_INQB_button
		    PushButton 185, 355, 25, 10, "INQD", MONY_INQB_button
		    PushButton 210, 355, 25, 10, "INQX", MONY_INQB_button
		    PushButton 35, 365, 25, 10, "SNAP", ELIG_FS_button
		    PushButton 60, 365, 25, 10, "MFIP", ELIG_MFIP_button
		    PushButton 85, 365, 25, 10, "DWP", ELIG_DWP_button
		    PushButton 110, 365, 25, 10, "GA", ELIG_GA_button
		    PushButton 135, 365, 25, 10, "MSA", ELIG_MSA_button
		    PushButton 160, 365, 25, 10, "GRH", ELIG_GRH_button
		    PushButton 185, 365, 25, 10, "HC", ELIG_HC_button
		    PushButton 210, 365, 25, 10, "SUMM", ELIG_SUMM_button
		    PushButton 235, 365, 25, 10, "DENY", ELIG_DENY_button
		  Text 250, 5, 290, 10, "NOTICE Information for Verification of Public Assistance for Case # " & MAXIS_case_number
		  GroupBox 5, 15, 470, 315, "Details"
		  GroupBox 5, 335, 390, 45, "Navigation"
		  Text 10, 345, 25, 10, "CASE/"
		  Text 135, 345, 25, 10, "SPEC/"
		  Text 10, 355, 25, 10, "STAT/"
		  Text 10, 365, 20, 10, "ELIG/"
		  Text 135, 355, 25, 10, "MONY/"
		EndDialog

		dialog Dialog1
		cancel_confirmation
		MAXIS_dialog_navigation
		Call leave_notice_text(False)

		If ButtonPressed > 1000 Then
			If ButtonPressed = snap_change_wcom_btn Then
				selected_prog = "FS"
				notc_month = snap_month
				notc_year = snap_year
			End If
			If ButtonPressed = ga_change_wcom_btn Then
				selected_prog = "GA"
				notc_month = ga_month
				notc_year = ga_year
			End If
			If ButtonPressed = msa_change_wcom_btn Then
				selected_prog = "MS"
				notc_month = msa_month
				notc_year = msa_year
			End If
			If ButtonPressed = mfip_change_wcom_btn Then
				selected_prog = "MF"
				notc_month = mfip_month
				notc_year = mfip_year
			End If

			Call Create_List_Of_Notices("WCOM", notices_array, selected, information, WCOM_search_row, no_notices, selected_prog)

			Call Select_New_WCOM(notices_array, selected, information, WCOM_search_row, True, 		True, False, notc_month, notc_year, no_notices, selected_prog, False, False)

			for each_notc = 0 to UBound(notices_array, 2)
				If notices_array(selected, each_notc) = checked Then
					If selected_prog = "FS" Then
						snap_month = notc_month
						snap_year = notc_year
						snap_wcom_text = notices_array(information, each_notc)
						snap_wcom_row = notices_array(WCOM_search_row, each_notc)
						snap_wcom_position = snap_wcom_row - 6
					End If
					If selected_prog = "GA" Then
						ga_month = notc_month
						ga_year = notc_year
						ga_wcom_text = notices_array(information, each_notc)
						ga_wcom_row = notices_array(WCOM_search_row, each_notc)
						ga_wcom_position = ga_wcom_row - 6
					End If
					If selected_prog = "MS" Then
						msa_month = notc_month
						msa_year = notc_year
						msa_wcom_text = notices_array(information, each_notc)
						msa_wcom_row = notices_array(WCOM_search_row, each_notc)
						msa_wcom_position = msa_wcom_row - 6
					End If
					If selected_prog = "MF" Then
						mfip_month = notc_month
						mfip_year = notc_year
						mfip_wcom_text = notices_array(information, each_notc)
						mfip_wcom_row = notices_array(WCOM_search_row, each_notc)
						mfip_wcom_position = mfip_wcom_row - 6
					End If
				End If
			next
			err_msg = "LOOP"
		End If
		selected_prog = ""

		If ButtonPressed < 1000 AND ButtonPressed > 100 Then
			If ButtonPressed = snap_wcom_btn Then
				wcom_row_to_open = snap_wcom_row
				wcom_month = snap_month
				wcom_year = snap_year
			End If
			If ButtonPressed = ga_wcom_btn Then
				wcom_row_to_open = ga_wcom_row
				wcom_month = ga_month
				wcom_year = ga_year
			End If
			If ButtonPressed = msa_wcom_btn Then
				wcom_row_to_open = msa_wcom_row
				wcom_month = msa_month
				wcom_year = msa_year
			End If
			If ButtonPressed = mfip_wcom_btn Then
				wcom_row_to_open = mfip_wcom_row
				wcom_month = mfip_month
				wcom_year = mfip_year
			End If

			Call navigate_to_MAXIS_screen("SPEC", "WCOM")
			EMWriteScreen wcom_month, 3, 46
			EMWriteScreen wcom_year, 3, 51
			transmit
			EMWriteScreen "X", wcom_row_to_open, 13
			open_wcom = MsgBox("The WCOM Notice has been selected." & vbCr & vbCr & "Would you like to open the notice?", vbQuestion + vbYesNo, "WCOM selected")
			If open_wcom = vbYes Then
				transmit
			Else
				EMWriteScreen " ", wcom_row_to_open, 13
			End If

			err_msg = "LOOP"
		End If

		If ButtonPressed > 50 AND ButtonPressed < 100 Then
			If ButtonPressed = snap_program_history_button Then prog_to_search = "FS"
			If ButtonPressed = ga_program_history_button Then prog_to_search = "GA"
			If ButtonPressed = msa_program_history_button Then prog_to_search = "MS"
			If ButtonPressed = mfip_program_history_button Then prog_to_search = "MF"
			If ButtonPressed = dwp_program_history_button Then prog_to_search = "DW"
			If ButtonPressed = grh_program_history_button Then prog_to_search = "GR"
			If ButtonPressed = hc_program_history_button Then
				'WAY MORE STUFF GOES HERE
			End If

			Call navigate_to_MAXIS_screen("CASE", "CURR")
			EMWriteScreen "X", 4, 9
			transmit
			EMWriteScreen prog_to_search, 3, 19
			transmit

			err_msg = "LOOP"
		End If

		If err_msg <> "LOOP" Then
			snap_start_month = trim(snap_start_month)
			snap_end_month = trim(snap_end_month)
			ga_start_month = trim(ga_start_month)
			ga_end_month = trim(ga_end_month)
			msa_start_month = trim(msa_start_month)
			msa_end_month = trim(msa_end_month)
			mfip_start_month = trim(mfip_start_month)
			mfip_end_month = trim(mfip_end_month)
			dwp_start_month = trim(dwp_start_month)
			dwp_end_month = trim(dwp_end_month)
			grh_start_month = trim(grh_start_month)
			grh_end_month = trim(grh_end_month)

			If snap_status = "ACTIVE" Then
				If snap_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since SNAP is active, indicate if Verification of SNAP benefits is needed, and if so, which method works best."
				If snap_verification_method = "Resend WCOM - Eligibility Notice" AND snap_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of SNAP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If snap_verification_method = "Create New MEMO with range of Months" Then
				 	If len(snap_start_month) <> 5 OR Mid(snap_start_month, 3, 1) <> "/" OR len(snap_end_month) <> 5 OR Mid(snap_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of SNAP issuance history to be sent as verification of Active SNAP, enter a start and end month in the 'mm/yy' format."
				End If
			End If
			If ga_status = "ACTIVE" Then
				If ga_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since GA is active, indicate if Verification of GA benefits is needed, and if so, which method works best."
				If ga_verification_method = "Resend WCOM - Eligibility Notice" AND ga_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of GA, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If ga_verification_method = "Create New MEMO with range of Months" Then
				 	If len(ga_start_month) <> 5 OR Mid(ga_start_month, 3, 1) <> "/" OR len(ga_end_month) <> 5 OR Mid(ga_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GA issuance history to be sent as verification of Active GA, enter a start and end month in the 'mm/yy' format."
				End If
			End If
			If msa_status = "ACTIVE" Then
				If msa_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since MSA is active, indicate if Verification of MSA benefits is needed, and if so, which method works best."
				If msa_verification_method = "Resend WCOM - Eligibility Notice" AND msa_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of MSA, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If msa_verification_method = "Create New MEMO with range of Months" Then
				 	If len(msa_start_month) <> 5 OR Mid(msa_start_month, 3, 1) <> "/" OR len(msa_end_month) <> 5 OR Mid(msa_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MSA issuance history to be sent as verification of Active MSA, enter a start and end month in the 'mm/yy' format."
				End If
			End If
			If mfip_status = "ACTIVE" Then
				If mfip_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since MFIP is active, indicate if Verification of MFIP benefits is needed, and if so, which method works best."
				If mfip_verification_method = "Resend WCOM - Eligibility Notice" AND mfip_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of MFIP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If mfip_verification_method = "Create New MEMO with range of Months" Then
				 	If len(mfip_start_month) <> 5 OR Mid(mfip_start_month, 3, 1) <> "/" OR len(mfip_end_month) <> 5 OR Mid(mfip_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MFIP issuance history to be sent as verification of Active MFIP, enter a start and end month in the 'mm/yy' format."
				End If
			End If
			If dwp_status = "ACTIVE" Then
			 	If dwp_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since DWP is active, indicate if Verification of DWP benefits is needed, and if so, which method works best."
				If dwp_verification_method = "Resend WCOM - Eligibility Notice" AND dwp_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of DWP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If dwp_verification_method = "Create New MEMO with range of Months" Then
				 	If len(dwp_start_month) <> 5 OR Mid(dwp_start_month, 3, 1) <> "/" OR len(dwp_end_month) <> 5 OR Mid(dwp_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of DWP issuance history to be sent as verification of Active DWP, enter a start and end month in the 'mm/yy' format."
				End If
			End If
			If grh_status = "ACTIVE" Then
			 	If grh_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since GRH is active, indicate if Verification of GRH benefits is needed, and if so, which method works best."
				If grh_verification_method = "Resend WCOM - Eligibility Notice" AND grh_wcom_text = "No WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of GRH, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
				If grh_verification_method = "Create New MEMO with range of Months" Then
				 	If len(grh_start_month) <> 5 OR Mid(grh_start_month, 3, 1) <> "/" OR len(grh_end_month) <> 5 OR Mid(grh_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GRH issuance history to be sent as verification of Active GRH, enter a start and end month in the 'mm/yy' format."
				End If
			End If

			If snap_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(snap_start_month) <> 5 OR Mid(snap_start_month, 3, 1) <> "/" OR len(snap_end_month) <> 5 OR Mid(snap_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of SNAP issuance history to be sent as verification of Previous SNAP Eligibility, enter a start and end month in the 'mm/yy' format."
			End If
			If ga_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(ga_start_month) <> 5 OR Mid(ga_start_month, 3, 1) <> "/" OR len(ga_end_month) <> 5 OR Mid(ga_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GA issuance history to be sent as verification of Previous GA Eligibility, enter a start and end month in the 'mm/yy' format."
			End If
			If msa_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(msa_start_month) <> 5 OR Mid(msa_start_month, 3, 1) <> "/" OR len(msa_end_month) <> 5 OR Mid(msa_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MSA issuance history to be sent as verification of Previous MSA Eligibility, enter a start and end month in the 'mm/yy' format."
			End If
			If mfip_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(mfip_start_month) <> 5 OR Mid(mfip_start_month, 3, 1) <> "/" OR len(mfip_end_month) <> 5 OR Mid(mfip_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MFIP issuance history to be sent as verification of Previous MFIP Eligibility, enter a start and end month in the 'mm/yy' format."
			End If
			If dwp_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(dwp_start_month) <> 5 OR Mid(dwp_start_month, 3, 1) <> "/" OR len(dwp_end_month) <> 5 OR Mid(dwp_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of DWP issuance history to be sent as verification of Previous DWP Eligibility, enter a start and end month in the 'mm/yy' format."
			End If
			If grh_not_actv_memo_for_old_beneftis_checkbox = checked Then
				If len(grh_start_month) <> 5 OR Mid(grh_start_month, 3, 1) <> "/" OR len(grh_end_month) <> 5 OR Mid(grh_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GRH issuance history to be sent as verification of Previous GRH Eligibility, enter a start and end month in the 'mm/yy' format."
			End If

			If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
		End If
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Go look to see if AREP Needs Mail
'Go look if SWKR needs Mail
'Need to add handling to check ADDRESS
'Need to add handling for OTHER and AREP and SWKR to see if they should be included.

create_memo = False
If snap_verification_method = "Create New MEMO with range of Months" Then create_memo = True
If ga_verification_method = "Create New MEMO with range of Months" Then create_memo = True
If msa_verification_method = "Create New MEMO with range of Months" Then create_memo = True
If mfip_verification_method = "Create New MEMO with range of Months" Then create_memo = True
If dwp_verification_method = "Create New MEMO with range of Months" Then create_memo = True
If grh_verification_method = "Create New MEMO with range of Months" Then create_memo = True

resend_wcom = False
If snap_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
If ga_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
If msa_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
If mfip_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
If dwp_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
If grh_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True


snap_resent_wcom = False

If resend_wcom = True Then

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")

	If snap_verification_method = "Resend WCOM - Eligibility Notice" Then
		EMWriteScreen snap_month, 3, 46
		EMWriteScreen snap_year, 3, 51
		transmit
		EMWriteScreen "A", snap_wcom_row, 13
		transmit
		EMWriteScreen "X", 5, 12 		'This is the CLIENT Row
		' EMWriteScreen "X", 6, 12		'This is the OTHER Row - need handling
		' EMWriteScreen "X", ?, 12 		'AREP
		' EMWriteScreen "X", ?, 12		'SWKR
		transmit

		EMReadScreen check_for_resent, 6, snap_wcom_row, 3
		EMReadScreen check_for_waiting, 6, snap_wcom_row, 71

		If check_for_resent = "ReSent" and check_for_waiting = "Waiting" Then snap_resent_wcom = True
	End If
	If ga_verification_method = "Resend WCOM - Eligibility Notice" Then

	End If
	If msa_verification_method = "Resend WCOM - Eligibility Notice" Then

	End If
	If mfip_verification_method = "Resend WCOM - Eligibility Notice" Then

	End If
	If dwp_verification_method = "Resend WCOM - Eligibility Notice" Then

	End If
	If grh_verification_method = "Resend WCOM - Eligibility Notice" Then

	End If
End If

Call back_to_SELF

const grant_amount_const 	= 0
const benefit_month_const	= 1
const note_message_const	= 2
const benefit_month_as_date_const = 3
const last_const			= 4

Dim SNAP_ISSUANCE_ARRAY()
ReDim SNAP_ISSUANCE_ARRAY(last_const, 0)
Dim ga_issuance_array()
Dim msa_issuance_array()
Dim mfip_issuance_array()
Dim dwp_issuance_array()
Dim grh_issuance_array()

If create_memo = True Then
	Call navigate_to_MAXIS_screen("MONY", "INQX")

	If snap_verification_method = "Create New MEMO with range of Months" Then
		first_date_of_range = replace(snap_start_month, "/", "/01/")
		first_date_of_range = DateAdd("d", 0, first_date_of_range)
		last_date_of_range = replace(snap_end_month, "/", "/01/")
		last_date_of_range = DateAdd("d", 0, last_date_of_range)

		EMWriteScreen "X", 9, 5		'This is the SNAP place
		EMWriteScreen left(snap_start_month, 2), 6, 38
		EMWriteScreen right(snap_start_month, 2), 6, 41
		' EMWriteScreen left(snap_end_month, 2), 6, 53
		' EMWriteScreen right(snap_end_month, 2), 6, 56
		EMWriteScreen CM_plus_1_mo, 6, 53
		EMWriteScreen CM_plus_1_yr, 6, 56

		transmit

		inqx_row = 6
		msg_counter = 0
		Do
			EMReadScreen issued_date, 8, inqx_row, 7
			EMReadScreen tran_amount, 8, inqx_row, 38
			EMReadScreen from_month, 2, inqx_row, 62
			EMReadScreen from_year, 2, inqx_row, 68
			EMReadScreen from_date, 8, inqx_row, 62

			issued_date = trim(issued_date)
			tran_amount = trim(tran_amount)

			If issued_date <> "" Then
				from_date = DateAdd("d", 0, from_date)
				If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

					benefit_month = from_month & "/" & from_year
					tran_amount = tran_amount * 1
					ammount_added_in = False
					For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)
						If benefit_month = SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
							SNAP_ISSUANCE_ARRAY(grant_amount_const, each_known_issuance) = SNAP_ISSUANCE_ARRAY(grant_amount_const, each_known_issuance) + tran_amount
							ammount_added_in = True
						End If
					Next
					If ammount_added_in = False Then
						ReDim Preserve SNAP_ISSUANCE_ARRAY(last_const, msg_counter)
						SNAP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
						SNAP_ISSUANCE_ARRAY(grant_amount_const, msg_counter) = tran_amount
						'maybe add the 'from_date' into benefit_month_as_date' so that we don't hvae to do more handling down below
						msg_counter = msg_counter + 1
					End If
				End If
			End If

			inqx_row = inqx_row + 1
			If inqx_row = 18 Then
				PF8
				inqx_row = 6
				EMreadScreen end_of_list, 9, 24, 14
				if end_of_list = "LAST PAGE" Then Exit Do
			End If
			'NEED TO ADD 'INQX' line limit
			'Need to look up display limitations'
		Loop until issued_date = ""
		dates_array = ""
		For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)
			total_amount = left(SNAP_ISSUANCE_ARRAY(grant_amount_const, each_known_issuance) & "        ", 8)
			SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
			SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance) = replace(SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance), "/", "/01/")
			SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance) = DateAdd("d", 0, SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance))
			dates_array = dates_array & "~" & SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)
		Next
		If left(dates_array, 1) = "~" Then dates_array = right(dates_array, len(dates_array) - 1)
		If Instr(dates_array, "~") = 0 Then
			dates_array = Array(dates_array)
		Else
			dates_array = split(dates_array, "~")
		End If
		Call sort_dates(dates_array)
		' MsgBox Join(dates_array, " - ")
		for each ordered_date in dates_array
			For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					msg_display = msg_display & vbCr & SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					' Call write_variable_in_SPEC_MEMO(SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance))
					' Call write_variable_in_CASE_NOTE(SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance))
				End If
			Next
		Next

		MsgBox "This is the list" & msg_display

	End If
	If ga_verification_method = "Create New MEMO with range of Months" Then
		EMWriteScreen "X", 11, 5		'This is the GA place
		EMWriteScreen left(ga_start_month, 2), 6, 38
		EMWriteScreen right(ga_start_month, 2), 6, 41
		EMWriteScreen left(ga_end_month, 2), 6, 53
		EMWriteScreen right(ga_end_month, 2), 6, 56
		transmit

	End If
	If msa_verification_method = "Create New MEMO with range of Months" Then
		EMWriteScreen "X", 13, 50		'This is the MSA place
		EMWriteScreen left(msa_start_month, 2), 6, 38
		EMWriteScreen right(msa_start_month, 2), 6, 41
		EMWriteScreen left(msa_end_month, 2), 6, 53
		EMWriteScreen right(msa_end_month, 2), 6, 56
		transmit

	End If
	If mfip_verification_method = "Create New MEMO with range of Months" Then
		EMWriteScreen "X", 10, 5		'This is the MFIP place
		EMWriteScreen left(mfip_start_month, 2), 6, 38
		EMWriteScreen right(mfip_start_month, 2), 6, 41
		EMWriteScreen left(mfip_end_month, 2), 6, 53
		EMWriteScreen right(mfip_end_month, 2), 6, 56
		transmit

	End If
	If dwp_verification_method = "Create New MEMO with range of Months" Then
		EMWriteScreen "X", 17, 50		'This is the DWO place
		EMWriteScreen left(dwp_start_month, 2), 6, 38
		EMWriteScreen right(dwp_start_month, 2), 6, 41
		EMWriteScreen left(dwp_end_month, 2), 6, 53
		EMWriteScreen right(dwp_end_month, 2), 6, 56
		transmit

	End If
	If grh_verification_method = "Create New MEMO with range of Months" Then
		EMWriteScreen "X", 16, 50		'This is the GRH place
		EMWriteScreen left(grh_start_month, 2), 6, 38
		EMWriteScreen right(grh_start_month, 2), 6, 41
		EMWriteScreen left(grh_end_month, 2), 6, 53
		EMWriteScreen right(grh_end_month, 2), 6, 56
		transmit

	End If
End If




MsgBox "STOP HERE"





















































'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
MAXIS_footer_month = datepart("m", next_month)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = datepart("yyyy", next_month)
MAXIS_footer_year = "" & MAXIS_footer_year - 2000

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone & grabs the case number and footer month/year
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 151, 70, "PA Verification Request"
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 95, 50, 50, 15
  EditBox 75, 5, 70, 15, MAXIS_case_number
  EditBox 75, 25, 30, 15, MAXIS_footer_month
  EditBox 115, 25, 30, 15, MAXIS_footer_year
  Text 10, 10, 50, 10, "Case Number"
  Text 10, 30, 65, 10, "Footer month/year:"
EndDialog

'Showing case number dialog
Do
	Do
		err_msg = ""
  		Dialog Dialog1
  		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine &  "* You need to type a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Jumping to STAT
call navigate_to_MAXIS_screen("stat", "memb")
'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Pulling household and worker info for the letter
call navigate_to_MAXIS_screen("stat", "addr")
EMReadScreen addr_line1, 21, 6, 43
EMReadScreen addr_line2, 21, 7, 43
EMReadScreen addr_city, 14, 8, 43
EMReadScreen addr_state, 2, 8, 66
EMReadScreen addr_zip, 5, 9, 43
hh_address = addr_line1 & " " & addr_line2 'Finding and Formatting household address
hh_address_line2 = addr_city & " " & addr_state & " " & addr_zip
hh_address = replace(hh_address, "_", "") & vbCrLf & replace(hh_address_line2, "_", "")


household_members = UBound(HH_member_array) + 1 'Total members in household
household_members = cStr(household_members)

'Collecting and formatting client name
call navigate_to_MAXIS_screen("stat", "memb")
call find_variable("Last: ", last_name, 24)
call find_variable("First: ", first_name, 11)
client_name = first_name & " " & last_name
client_name = replace(client_name, "_", "")


'This function looks for an approved version of elig
Function approved_version
	EMReadScreen version, 2, 2, 12
	For approved = version to 0 Step -1
	EMReadScreen approved_check, 8, 3, 3
	If approved_check = "APPROVED" then Exit Function
	version = version -1
	EMWriteScreen version, 20, 79
	transmit
	Next
End Function


'This finds the number of members on a DWP/MFIP grant
Function cash_members_finder
	call find_variable("Caregivers......", caregivers, 4)
	call find_variable("Children........", children, 4)
	cash_members = cInt(caregivers) + cInt(children)
	cash_members = cStr(cash_members)
End Function

'Pulling the elig amounts for all open progs on case / curr
call navigate_to_MAXIS_screen("case", "curr")
  call find_variable("MFIP: ", MFIP_check, 6)
   If MFIP_check = "ACTIVE" OR MFIP_check = "APP CL" then
		call navigate_to_MAXIS_screen("elig", "mfip")
		EMReadScreen are_we_at_sig_change, 4, 3, 38			'If we have a SIG Change budget listed - the version thing doesn't work the same - need to see if we are and transmit past
		If are_we_at_sig_change = "MFSC" Then transmit
	  	EMReadScreen version, 1, 2, 12 'Reading the version, the for loop finds most recent approved.
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
        EMWriteScreen "MFB2", 20, 71
        transmit
        EMReadScreen MFIP_cash, 8, 12, 34
        EMReadScreen MFIP_food, 8, 7, 34
		EMReadScreen MFIP_housing, 6, 17, 36
        MFIP_cash = trim(MFIP_cash)
        MFIP_food = trim(MFIP_food)
		IF MFIP_housing = "" then MFIP_housing = 0
		'MFIP_cash = (cInt(MFIP_cash) + MFIP_housing)
		'MFIP_cash = cstr(MFIP_cash)
 		'rental subsidy check
		EMWriteScreen "MFB1", 20, 71
		EMReadScreen subsidy, 2, 17, 37
		If subsidy = "50" then subsidy_check = 1
		'Finding the number of members on cash grant
		call cash_members_finder
		Call navigate_to_MAXIS_screen("case", "curr")
	End if
	If MFIP_check = "PENDIN" then msgbox "MFIP is pending, please enter amounts manually to avoid errors."

	call find_variable("FS: ", fs_check, 6)
	If fs_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "fs")
		EMReadScreen version, 2, 2, 12
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen version, 19, 78
			transmit
		Next
		EMWriteScreen version, 19, 78
		transmit
		EMWriteScreen "FSB2", 19, 70
		transmit
		EMReadScreen SNAP_grant, 8, 10, 73
        SNAP_grant = trim(SNAP_grant)
	    call navigate_to_MAXIS_screen ("case", "curr")
	End if
	If fs_check = "APP CL" then msgbox "SNAP is set to close, please enter amounts manually to avoid errors."
	If fs_check = "PENDIN" then msgbox "SNAP is pending, please enter amounts manually to avoid errors."

	call find_variable("DWP: ", DWP_check, 6)
	If DWP_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "dwp")
		EMReadScreen version, 2, 2, 11
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "DWB2", 20, 71
		transmit
		EMReadScreen DWP_grant, 8, 5, 36
        DWP_grant = trim(DWP_grant)
	    EMWriteScreen "DWSM", 20, 71
		transmit
		call find_variable("Caregivers....", caregivers, 5)
		call find_variable("Children......", children, 5)
		cash_members = cInt(caregivers) + cInt(children)
		cash_members = cStr(cash_members)
		call navigate_to_MAXIS_screen ("case", "curr")
	 End if
	If DWP_check = "PENDIN" then msgbox "DWP is pending, please enter amounts manually to avoid errors."

	call find_variable("GA: ", GA_check, 6)
	If GA_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "GA")
		EMReadScreen version, 2, 2, 12
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 78
			transmit
		Next
		EMWriteScreen version, 20, 78
		transmit
		EMWriteScreen "GASM", 20, 70
		transmit
		EMReadScreen GA_grant, 7, 9, 73
	    EMReadScreen ga_members, 1, 13, 32 'Reading file unit type to determine members on cash grant
		If ga_members = "1" then cash_members = "1"
		If ga_members = "6" then cash_members = "2"
		call navigate_to_MAXIS_screen ("case", "curr")
	End If
	If GA_check = "APP CL" then msgbox "GA is set to close, please enter amounts manually to avoid errors."
	If GA_check = "PENDIN" then msgbox "GA is pending, please enter amounts manually to avoid errors."

	call find_variable("MSA: ", MSA_check, 6)
	If MSA_check = "ACTIVE" then
		call navigate_to_MAXIS_screen("elig", "msa")
		EMReadScreen version, 2, 2, 11
		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen "0" & version, 20, 79
			transmit
		Next
		EMWriteScreen version, 20, 79
		transmit
		EMWriteScreen "MSSM", 20, 71
		transmit
		EMReadScreen MSA_Grant, 7, 11, 74
		EMReadScreen cash_members, 1, 14, 29
		call navigate_to_MAXIS_screen ("case", "curr")
	End If
	If MSA_check = "APP CL" then MsgBox "MSA is set to close, please enter amounts manually to avoid errors."
	If MSA_check = "PENDIN" then MsgBox "MSA is pending, please enter amounts manually to avoid errors."

	call find_variable("Cash: ", cash_check, 6)
	If cash_check = "PENDIN" then MsgBox "Cash is pending for this household, please explain in additional notes."

'calling the main dialog

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 316, 260, "PA Verif Dialog"
  ButtonGroup ButtonPressed
    OkButton 200, 240, 50, 15
    CancelButton 255, 240, 50, 15
  EditBox 40, 15, 30, 15, snap_grant
  EditBox 105, 15, 35, 15, MFIP_food
  EditBox 145, 15, 35, 15, MFIP_cash
  EditBox 40, 35, 30, 15, MSA_Grant
  EditBox 145, 35, 35, 15, MFIP_housing
  EditBox 40, 55, 30, 15, GA_grant
  EditBox 145, 55, 35, 15, DWP_grant
  EditBox 285, 15, 20, 15, cash_members
  CheckBox 285, 40, 25, 10, "Yes", subsidy_check
  EditBox 285, 55, 20, 15, household_members
  EditBox 85, 75, 220, 15, other_income
  EditBox 105, 95, 20, 15, number_of_months
  CheckBox 15, 145, 280, 10, "Check here to have the HH information withheld from the word doc.", no_hh_info_checkbox
  EditBox 55, 195, 250, 15, other_notes
  EditBox 55, 220, 90, 15, completed_by
  EditBox 210, 220, 95, 15, worker_phone
  EditBox 120, 240, 75, 15, worker_signature
  CheckBox 10, 100, 95, 10, "Include screenshot of last", inqd_check
  Text 10, 20, 20, 10, "SNAP:"
  Text 110, 60, 20, 10, "DWP:"
  Text 5, 200, 40, 10, "Other notes:"
  Text 10, 60, 20, 10, "GA:"
  Text 10, 40, 20, 10, "MSA:"
  Text 80, 20, 20, 10, "MFIP:"
  Text 80, 40, 50, 10, "MFIP Housing:"
  Text 5, 80, 75, 10, "Other income and type:"
  Text 200, 40, 80, 10, "$50 subsidy deduction?"
  Text 190, 20, 95, 10, "HH members on cash grant:"
  Text 215, 60, 65, 10, "Total HH members:"
  Text 110, 5, 25, 10, "Food:"
  Text 150, 5, 25, 10, "Cash:"
  Text 130, 100, 60, 10, "months' benefits"
  Text 155, 225, 55, 10, "Worker phone #:"
  Text 5, 225, 50, 10, "Completed by:"
  Text 5, 245, 110, 10, "Worker Signature (for case note):"
  Text 15, 125, 280, 20, "We cannot provide information about budgeted or known income as we cannot verify anything other that Pulic Assistance Income."
  GroupBox 5, 115, 300, 75, "Warning!"
  Text 15, 160, 275, 20, "If a resident needs information from their file, they must request it through the ROI Team (Release of Information)."
  ButtonGroup ButtonPressed
	PushButton 175, 172, 120, 13, "HSR Manual - Data Privacy", data_privacy_btn
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Please sign your case note."
		If completed_by = "" then err_msg = err_msg & vbNewLine & "* Please fill out the completed by field."
		If worker_phone = "" then err_msg = err_msg & vbNewLine & "* Please fill out the worker phone field."

		if ButtonPressed = data_privacy_btn Then
			Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Data_Privacy.aspx")
			err_msg = "LOOP" & err_msg
		Else
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		End If
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'****writing the word document
Set objWord = CreateObject("Word.Application")
Const wdDialogFilePrint = 88
Const end_of_doc = 6
objWord.Caption = "PA Verif Request"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "14"
objSelection.TypeText "Your agency requested information about public assistance from "
objSelection.TypeText county_name
objSelection.TypeText " for the following client:"
objSelection.TypeParagraph()
objSelection.TypeText client_name
objSelection.TypeParagraph()
objSelection.TypeText hh_address
objSelection.TypeParagraph()
objSelection.TypeText "The following grant amounts are active for this household:"

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 7, 3
set objTable = objDoc.Tables(1)

objTable.Cell(1, 2).Range.Text = "Cash  "
objTable.Cell(1, 3).Range.Text = "Food Portion"
objTable.Cell(2, 1).Range.Text = "MFIP (MN Family Investment program) "
objTable.Cell(3, 1).Range.Text = "MFIP Housing Grant"
objTable.Cell(4, 1).Range.Text = "GA (General Assistance)"
objTable.Cell(5, 1).Range.Text = "MSA (MN supplemental Aid)"
objTable.Cell(6, 1).Range.Text = "SNAP (Supplemental Nutrition Assistance program)"
objTable.Cell(2, 2).Range.Text = MFIP_cash
objTable.Cell(2, 3).Range.Text = MFIP_food
objTable.Cell(3, 2).Range.Text = MFIP_housing
objTable.Cell(4, 2).Range.Text = GA_grant
objTable.Cell(5, 2).Range.Text = MSA_Grant
objTable.Cell(6, 3).Range.Text = SNAP_grant
objTable.Cell(7, 1).Range.Text = "DWP (Diversionary Work program) "
objTable.Cell(7, 2).Range.Text = DWP_grant

objTable.AutoFormat(16)

objSelection.EndKey end_of_doc
objSelection.TypeParagraph()

objSelection.TypeText "Number of family members on cash grant: "
objSelection.TypeText cash_members
objSelection.TypeParagraph()

If no_hh_info_checkbox = unchecked Then 		'Only adding the detail from stat if the worker leaves the omit income unchecked
	objSelection.TypeText "Number of persons in household: "
	objSelection.TypeText household_members
	objSelection.TypeParagraph()
End If

ObjSelection.TypeText "Other Notes: "
objSelection.TypeText other_notes
objSelection.TypeParagraph()

'Writing INQX to the doc if selected
IF inqd_check = checked THEN
	objSelection.TypeText "Benefits Issued for last " & number_of_months & " months:"
	objSelection.TypeParagraph()
	objSelection.TypeText "Issue Date	    Benefit               Amount                            Benefit Period"
	objSelection.TypeParagraph()
	call navigate_to_MAXIS_screen("MONY", "INQX")
	start_date = dateadd("m", - number_of_months, date) 'Converting dates to determine how far back to look
	start_month = datepart("m", start_date)
	IF len(start_month) = 1 THEN start_month = "0" & start_month
	EMWriteScreen start_month, 6, 38
	EMWriteScreen right(datepart("YYYY", start_date), 2), 6, 41
	transmit
	output_array = "" 'resetting array
	row = 6
	DO
	EMReadScreen reading_line, 80, row, 1
	output_array = output_array & reading_line & "UUDDLRLRBA" 'adding the info to the array
	row = row + 1
	IF row = 18 THEN 'Checking for more screens
		EMReadScreen more_check, 1, 19, 9
		IF more_check <> "+" THEN EXIT DO
		PF8
		row = 6
	END IF
	LOOP
	output_array = split(output_array, "UUDDLRLRBA")
	FOR EACH line in output_array 'Type the info from array into word doc
		IF line <> "                                                                                " THEN
			objSelection.TypeText line & Chr(11)
		END IF
	NEXT
	objSelection.TypeParagraph()
	objSelection.TypeText "**********PROGRAM KEY**********"
	objSelection.TypeParagraph()
	objSelection.TypeText "DW = DWP (Diversionary Work program"
	objSelection.TypeParagraph()
	objSelection.TypeText "EA = Emergency Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "EG = Emergency General Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "FS = SNAP (Supplemental Nutrition)"
	objSelection.TypeParagraph()
	objSelection.TypeText "GA = General Assistance"
	objSelection.TypeParagraph()
	objSelection.TypeText "HG = MFIP Housing Grant"
	objSelection.TypeParagraph()
	objSelection.TypeText "MF-MF = MFIP (MN Family Investment program, cash portion)"
	objSelection.TypeParagraph()
	objSelection.TypeText "MF-FS = MFIP SNAP (food portion)"
	objSelection.TypeParagraph()
	objSelection.TypeText "MS = MSA (MN Supplemental Aid)"
	objSelection.TypeParagraph()
	objSelection.TypeText "RC = RCA (Refugee Cash Assistance)"
	objSelection.TypeParagraph()
	objSelection.TypeText "GR = Group Residential Housing"
	objSelection.TypeParagraph()
	objSelection.TypeText "SA = Special Needs/Diet"
	objSelection.TypeParagraph()
	objSelection.TypeText "SM = Special Needs MSA (MN Supplemental Aid)"
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
END IF

objSelection.TypeText "Completed By: "
objSelection.TypeText completed_by
objSelection.TypeParagraph()
objSelection.TypeText "Worker phone: "
objSelection.TypeText worker_phone

'Enters the case note
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("PA verification request completed and sent to requesting agency.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

'removal of print option during the COVID-19 PEACETIME STATE OF EMERGENCY
'Starts the print dialog
' objword.dialogs(wdDialogFilePrint).Show

script_end_procedure("")
