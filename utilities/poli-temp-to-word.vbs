'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - POLI TEMP to Word.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                     'manual run time in seconds
STATS_denomination = "I"                   'C is for each CASE
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
call changelog_update("06/06/2019", "Bug fixed when the POLI//TEMP reference is more than 9 pages long.", "Casey Love, Hennepin County")
call changelog_update("01/09/2019", "Added date created to bottom of Word Document.", "Casey Love, Hennepin County")
call changelog_update("01/08/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""        'Connects to BlueZone
EMReadScreen in_poli_temp, 26, 1, 29

If in_poli_temp = "R E V I E W    M A N U A L" Then
    EMReadScreen current_poli_page, 18, 3, 54
    current_poli_page = trim(current_poli_page)
    array_of_codes = split(current_poli_page, ".")

    temp_one = array_of_codes(0)
    temp_two = array_of_codes(1)
    If UBOUND(array_of_codes) > 1 Then temp_three = array_of_codes(2)
    If UBOUND(array_of_codes) > 2 Then temp_four = array_of_codes(3)
End If

'Displays dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 211, 90, "POLI/TEMP dialog"
  EditBox 40, 40, 20, 15, temp_one
  EditBox 65, 40, 20, 15, temp_two
  EditBox 90, 40, 20, 15, temp_three
  EditBox 115, 40, 20, 15, temp_four
  ButtonGroup ButtonPressed
    OkButton 95, 65, 50, 15
    CancelButton 155, 65, 50, 15
  Text 5, 10, 155, 10, "What policy of POLI/TEMP you want to print?"
  Text 5, 25, 70, 10, "POLI/TEMP source:"
  Text 10, 45, 25, 10, "TABLE:"
  Text 110, 45, 5, 10, ""
  Text 60, 45, 5, 10, ""
  Text 85, 45, 5, 10, ""
EndDialog
Do
    Do
        err_msg = ""
        Dialog Dialog1 
        Cancel_without_confirmation
        If trim(temp_one) <> "" AND trim(temp_two) = "" Then err_msg = err_msg & vbNewLine & "* TEMP Table Codes have at least two reference positions."
        If trim(temp_three) = "" AND trim(temp_four) <> "" Then err_msg = err_msg & vbNewLine & "* If there is a code in the 4th position, there needs to be one in the third."
        If err_msg <> "" Then MsgBox "**Please Resolve to Continue **" & vbNewLine & err_msg
    Loop Until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

temp_one = trim(temp_one)
temp_two = trim(temp_two)
temp_three = trim(temp_three)
temp_four = trim(temp_four)
index_topic = trim(index_topic)

'Determines which POLI/TEMP section to go to, using the dropdown list outcome to decide
If index_topic <> "" Then panel_title = "INDEX"
If temp_one <> "" Then panel_title = "TABLE"

'call screen back to SELF screen to proceed onward with POLI
'navigating back to SELF menu, since back_to_SELF does not work in POLI function
DO
	PF3
	EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
EMWriteScreen "TEMP", 5, 40     'Writes TEMP

'Writes the panel_title selection
Call write_value_and_transmit(panel_title, 21, 71)
policy_info = "POLI/TEMP"

If index_topic <> "" Then
    panel_title = "INDEX"
    policy_info = policy_info & "INDEX: " & index_topic

    EMWriteScreen index_topic, 3, 21
    transmit

    EMWriteScreen "X", 6, 4
    transmit
End If

If temp_one <> "" Then
    panel_title = "TABLE"

    If temp_one <> "" Then temp_one = right("00" & temp_one, 2)
    If len(temp_two) = 1 Then temp_two = right("00" & temp_two, 2)
    If len(temp_three) = 1 Then temp_three = right("00" & temp_three, 2)
    If len(temp_four) = 1 Then temp_four = right("00" & temp_four, 2)

    total_code = "TE" & temp_one & "." & temp_two
    If temp_three <> "" Then total_code = total_code & "." & temp_three
    If temp_four <> "" Then total_code = total_code & "." & temp_four

    EMWriteScreen total_code, 3, 21
    transmit

    EMReadScreen section_found, 18, 6, 54
    section_found = trim(section_found)
    'MsgBox "Section Found: " & section_found & vbNewLine & "Total Code: " & total_code
    If section_found = total_code Then
        'MsgBox "HERE"
        EMReadScreen poli_title, 46, 6, 8
        poli_title = trim(poli_title)
        EMWriteScreen "X", 6, 4
        transmit

    Else
        end_msg = "The POLI/TEMP table reference: " & total_code & " could not be found. Please check the reference and try again."
        script_end_procedure(end_msg)
    End If
    policy_info = policy_info & ": " & total_code
End If

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

notice_length = 0
page_nbr = 2

EMReadScreen end_of_poli, 2, 3, 79
end_of_poli = trim(end_of_poli)
Do
    For notice_row = 4 to 21
        EMReadScreen poli_line, 74, notice_row, 6
        poli_line = rtrim(poli_line)
        If notice_row = 3 Then first_line = poli_line
        'MsgBox poli_line
        if right(trim(poli_line),9) = "FMINFO___" Then poli_line = ""
        If right(trim(poli_line),4) = "Page" Then
            poli_line = trim(poli_line) & " " & page_nbr
            page_nbr = page_nbr + 1
        End If
        poli_wording = poli_wording & poli_line & vbcr
        If left(trim(poli_line), 7) = "WORKER:" Then Exit For
        poli_line = ""
    Next
    EMReadScreen current_page, 2, 3, 72
    current_page = trim(current_page)
    PF8
    notice_length = notice_length + 1
Loop until current_page = end_of_poli

Set objDoc = objWord.Documents.Add()
objWord.Caption = policy_info
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 30
objSelection.PageSetup.BottomMargin = 25
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "14"
objSelection.TypeText poli_title & " - "
objSelection.TypeText policy_info
objSelection.TypeParagraph()
objSelection.Font.Size = "10"
objSelection.ParagraphFormat.SpaceAfter = 0

objSelection.TypeText poli_wording

objSelection.TypeParagraph()
objSelection.TypeText "POLI/TEMP Information up-to-date as of: " & date & " (date Word Document created)"
objSelection.TypeParagraph()

script_end_procedure("")