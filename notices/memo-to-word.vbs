'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MEMO TO WORD.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 100         'manual run time in seconds
STATS_denomination = "I"        'I is for item (notices)
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

'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/09/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function Create_List_Of_Notices
	Erase notices_array
	no_notices = FALSE
	If notice_panel = "WCOM" Then
		wcom_row = 7
		array_counter = 0
		Do
			ReDim Preserve notices_array(3, array_counter)
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE

			notices_array(selected,    array_counter) = unchecked
			notices_array(information, array_counter) = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
			notices_array(MAXIS_row,   array_counter) = wcom_row

			array_counter = array_counter + 1
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

			If array_counter = 0 AND notice_date = "" Then no_notices = TRUE

			notices_array(selected,    array_counter) = unchecked
			notices_array(information, array_counter) = notice_info & " - " & notice_date & " - Status: " & notice_stat
			notices_array(MAXIS_row,   array_counter) = memo_row

			array_counter = array_counter + 1
			memo_row = memo_row + 1

			EMReadScreen next_notice, 4, memo_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If
End Function


EMConnect ""

Dim notices_array()
ReDim notices_array(3,0)

Const selected = 0
Const information = 1
Const MAXIS_row = 2

Call check_for_MAXIS(False)

'Finds MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

EMReadScreen which_panel, 4, 2, 47
If which_panel = "WCOM" then
	notice_panel = "WCOM"
	at_notices = True
ElseIf which_panel = "MEMO" Then
	notice_panel = "MEMO"
	at_notices = True
Else
	at_notices = false
End If

If at_notices = True then
	If notice_panel = "WCOM" Then
		EMReadScreen MAXIS_footer_month, 2, 3, 46
		EMReadScreen MAXIS_footer_year,  2, 3, 51
	ElseIf notice_panel = "MEMO" Then
		EMReadScreen MAXIS_footer_month, 2, 3, 48
		EMReadScreen MAXIS_footer_year,  2, 3, 53
	End If

	Create_List_Of_Notices

End If

Do
    Do
    	err_msg = ""

    	dlg_y_pos = 85
    	dlg_length = 145 + (UBound(notices_array, 2) * 20)

        Dialog1 = ""
    	BeginDialog Dialog1, 0, 0, 205, dlg_length, "Notices to Print"
    	  Text 5, 10, 50, 10, "Case Number"
    	  EditBox 65, 5, 50, 15, MAXIS_case_number
    	  Text 5, 30, 130, 10, "Where is the notice you want to print?"
    	  DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"WCOM"+chr(9)+"MEMO", notice_panel
    	  Text 35, 50, 95, 10, "Enter the month of the notice:"
    	  EditBox 140, 45, 20, 15, MAXIS_footer_month
    	  EditBox 165, 45, 20, 15, MAXIS_footer_year
    	  ButtonGroup ButtonPressed
    	    PushButton 60, 70, 50, 10, "Find Notices", find_notices_button
    	  If no_notices = FALSE Then
    		  For notices_listed = 0 to UBound(notices_array, 2)
    		  	CheckBox 10, dlg_y_pos, 185, 10, notices_array(information, notices_listed), notices_array(selected, notices_listed)
    			dlg_y_pos = dlg_y_pos + 15
    		  Next
    	  Else
    	  	  Text 10, dlg_y_pos, 185, 10, "**No Notices could be found here.**"
    		  dlg_y_pos = dlg_y_pos + 15
    	  End If
    	  dlg_y_pos = dlg_y_pos + 5
    	  EditBox 75, dlg_y_pos, 125, 15, worker_signature
    	  dlg_y_pos = dlg_y_pos + 5
    	  Text 5, dlg_y_pos, 60, 10, "Worker Signature:"
    	  dlg_y_pos = dlg_y_pos + 15
    	  ButtonGroup ButtonPressed
    	    OkButton 100, dlg_y_pos, 50, 15
    	    CancelButton 150, dlg_y_pos, 50, 15
    	  dlg_y_pos = dlg_y_pos + 5
    	  CheckBox 5, dlg_y_pos, 90, 10, "Check here to case note.", case_note_check
    	EndDialog

    	Dialog Dialog1
    	cancel_confirmation

    	notice_selected = FALSE
    	For notice_to_print = 0 to UBound(notices_array, 2)
    		If notices_array(selected, notice_to_print) = checked Then notice_selected = TRUE
    	Next

    	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
    	If notice_panel = "Select One..." Then err_msg = err_msg & vbNewLine & "- Select where the notice to print is."
    	If MAXIS_footer_month = "" or MAXIS_footer_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
    	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice to be copied to a Word Document."

    	If ButtonPressed = find_notices_button then
    		If notice_panel <> "Select One..." AND MAXIS_case_number <> "" AND MAXIS_footer_month <> "" AND MAXIS_footer_year <> "" Then
    			Call navigate_to_MAXIS_screen ("SPEC", notice_panel)
    			If notice_panel = "MEMO" then
    				EMWriteScreen MAXIS_footer_month, 3, 48
    				EMWriteScreen MAXIS_footer_year, 3, 53
    			ElseIf notice_panel = "WCOM" Then
    				EMWriteScreen MAXIS_footer_month, 3, 46
    				EMWriteScreen MAXIS_footer_year, 3, 51
    			End If
    			transmit
    			Create_List_Of_Notices
    			err_msg = "LOOP"
    		Else
    			err_msg = err_msg & vbNewLine & "!!! Cannot read a list of notices without a panel selected, a case number entered, and footer month & year entered !!!"
    		End If
    	End If

    	If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg

    Loop Until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call navigate_to_MAXIS_screen ("SPEC", notice_panel)

If notice_panel = "MEMO" then
	EMWriteScreen MAXIS_footer_month, 3, 48
	EMWriteScreen MAXIS_footer_year, 3, 53
ElseIf notice_panel = "WCOM" Then
	EMWriteScreen MAXIS_footer_month, 3, 46
	EMWriteScreen MAXIS_footer_year, 3, 51
End If
transmit

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

For notice_to_print = 0 to UBound(notices_array, 2)
	client_notice = ""
	If notices_array(selected, notice_to_print) = checked Then
		STATS_counter = STATS_counter + 1
		If notice_panel = "WCOM" Then EMWriteScreen "X", notices_array(MAXIS_row, notice_to_print), 13
		If notice_panel = "MEMO" Then EMWriteScreen "X", notices_array(MAXIS_row, notice_to_print), 16
		transmit

		notice_length = 0
        page_nbr = 2
		Do
			For notice_row = 2 to 21
				EMReadScreen notice_line, 73, notice_row, 8
                If notice_row = 3 Then first_line = notice_line
                'MsgBox notice_line
				if right(trim(notice_line),9) = "FMINFO___" Then notice_line = ""
                If right(trim(notice_line),4) = "Page" Then
                    notice_line = trim(notice_line) & " " & page_nbr
                    page_nbr = page_nbr + 1
                End If
				client_notice = client_notice & notice_line & vbcr
				If left(trim(notice_line), 7) = "WORKER:" Then Exit For
				notice_line = ""
			Next
            PF8
            EMReadScreen notice_end, 9, 24,14
			If notice_end = "LAST PAGE" Then
                EMReadScreen top_of_page, 73, 3, 8
                If top_of_page = first_line Then Exit Do
            End If
            notice_length = notice_length + 1
		Loop until notice_length = 20

		Set objDoc = objWord.Documents.Add()
		objWord.Caption = notices_array(information, notice_to_print)
		Set objSelection = objWord.Selection
		objSelection.PageSetup.LeftMargin = 50
		objSelection.PageSetup.RightMargin = 50
		objSelection.PageSetup.TopMargin = 30
		objSelection.PageSetup.BottomMargin = 25
		objSelection.Font.Name = "Courier New"
		objSelection.Font.Size = "10"
		objSelection.ParagraphFormat.SpaceAfter = 0

		objSelection.TypeText client_notice

		pf3
	End If
Next

If case_note_check = checked Then

	start_a_blank_CASE_NOTE
	Call Write_variable_in_case_note ("System Notice Reprinted in Office for Client")
	Call Write_variable_in_case_note ("Clt in office, requested copy of system generated notice.")
	Call Write_variable_in_case_note ("Notices Printed:")
	For printed_notice = 0 to UBound(notices_array, 2)
		If notices_array(selected, printed_notice) = checked Then Call Write_variable_in_case_note ("* SPEC/" & notice_panel & " - " & notices_array(information, printed_notice))
	Next
	Call Write_variable_in_case_note ("---")
	Call Write_variable_in_case_note (worker_signature)
End If

STATS_counter = STATS_counter - 1

script_end_procedure("Success! The script has generated a Word Document of the MAXIS Notice(s) requested.")
