'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - HOSPICE FORM RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("06/14/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'This function creates the HH Member dropdown for a number of different dialogs
function Generate_Client_List(list_for_dropdown)

	memb_row = 5       'setting the row to look at the list of members on the left hand side of the panel

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")         'go to MEMB
	Do                                                     'this loop transmits to each MEMB panel to read information for each member
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do           'this is the end of the list of members
		EMWriteScreen ref_numb, 20, 76            'writing the reference number in the command line to go to each MEMB panel
		transmit
		EMReadScreen first_name, 12, 6, 63        'reading the name on the panel
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")     'adding each client information to a string
		memb_row = memb_row + 1                   'going to the next member
	Loop until memb_row = 20

    If memb_row = 6 Then        'If the row is only 6, then there is only one person in the HH
        list_for_dropdown = right(client_info, len(client_info) - 1)    'taking the '~' off of the string
    Else
    	client_info = right(client_info, len(client_info) - 1)             'taking the left most '~' off
    	client_list_array = split(client_info, "~")                        'making this an array

    	For each person in client_list_array                               'creating the string to be added to the dialog code to fill the dropdown
    		list_for_dropdown = list_for_dropdown & chr(9) & person
    	Next
    End If

end function
'DIALOGS----------------------------------------------------------------------------------------------------
'Dialog to gather the case number and footer month and year
BeginDialog case_number_dialog, 0, 0, 156, 70, "Case number dialog"
  EditBox 60, 5, 90, 15, MAXIS_case_number
  EditBox 60, 25, 30, 15, MAXIS_footer_month
  EditBox 120, 25, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 100, 50, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 95, 30, 20, 10, "Year:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Searching for case number.
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog
Do
    Dialog case_number_dialog
    cancel_confirmation
    If MAXIS_case_number = "" then MsgBox "You must type a case number!"
Loop until MAXIS_case_number <> ""

'Now it checks to make sure MAXIS is running on this screen.
Call check_for_MAXIS(True)

'Looking for a previous case note to autofill some information as this script may be run twice on the same case.
Call navigate_to_MAXIS_screen("CASE", "NOTE")

note_row = 5                                'beginning of listed case notes
one_year_ago = DateAdd("yyyy", -1, date)    'we will look back 1 year
Do
    EMReadScreen note_date, 8, note_row, 6      'reading the date
    EMReadScreen note_title, 55, note_row, 25   'reading the header
    note_title = trim(note_title)

    If left(note_title, 41) = "*** HOSPICE TRANSACTION FORM RECEIVED ***" Then      'if the note is for a Hospice form
        EmWriteScreen "X", note_row, 3      'open the note
        transmit

        this_row = 5            'this MAXIS is the top of the note body
        Do
            EMReadScreen note_line, 78, this_row, 3     'reading each line
            note_line = trim(note_line)                 'Each of the lines will have the header look at to see if we can autofill information

            If  left(note_line, 9) = "* Client:" Then
                client_in_hospice = right(note_line, len(note_line) - 9)
                client_in_hospice = trim(client_in_hospice)

            ElseIf left(note_line, 15) = "* Hospice Name:" Then
                hospice_name = right(note_line, len(note_line) - 15)
                hospice_name = trim(hospice_name)

            ElseIf left(note_line, 13) = "* NPI Number:" Then
                npi_number = right(note_line, len(note_line) - 13)
                npi_number = trim(npi_number)

            ElseIf left(note_line, 16) = "* Date of Entry:" Then
                hospice_entry_date = right(note_line, len(note_line) - 16)
                hospice_entry_date = trim(hospice_entry_date)

            ElseIf left(note_line, 12) = "* Exit Date:" Then
                hospice_exit_date = right(note_line, len(note_line) - 12)
                hospice_exit_date = trim(hospice_exit_date)

            ElseIf left(note_line, 26) = "* MMIS not updated due to:" Then
                reason_not_updated = right(note_line, len(note_line) - 26)
                reason_not_updated = trim(reason_not_updated)

            End If
            If this_row = 18 Then       'this is the bottom of the note, will go to the next page if possible
                PF8
                EMReadScreen check_for_end, 9, 24, 14   'if we try to PF8 and it doesn't go down, a message happens at the bottom
                If check_for_end = "LAST PAGE" Then
                    PF3             'leaving the note
                    Exit Do         'don't need to look at any more of the note
                End If
                this_row = 4        'if the message isn't there reset the row to the top
            End If
            this_row = this_row + 1     'go to the next row
            If note_line = "" Then PF3  'if it is blank - the note is over and we need to leave the note
        Loop until note_line = ""

        Exit Do     'if a HOSPICE note is found, we don't need to look at more notes
    End If
    IF note_date = "        " then Exit Do      'if the end of the list is reached we leave the loop
    note_row = note_row + 1
    IF note_row = 19 THEN       'going to the next page of notes
        PF8
        note_row = 5
    END IF
    EMReadScreen next_note_date, 8, note_row, 6
    IF next_note_date = "        " then Exit Do
Loop until datevalue(next_note_date) < one_year_ago 'looking ahead at the next case note kicking out the dates before app'

If hospice_exit_date <> "" Then     'if there is an exit date in the note found then we don't want to use the information from that note
    client_in_hospice = ""          'since if they exited already - the HOSPICE will be different - resetting these variables to NOT fill
    hospice_name = ""
    npi_number = ""
    hospice_entry_date = ""
    hospice_exit_date = ""
    reason_not_updated = ""
End If

Call navigate_to_MAXIS_screen ("STAT", "MEMB")      'Going to MEMB for M01 to see if there is a date of death - as that would be the exit date
EMReadScreen date_of_death, 10, 19, 42
date_of_death = replace(date_of_death, " ", "/")
If IsDate(date_of_death) = TRUE Then hospice_exit_date = date_of_death

Call Generate_Client_List(HH_Memb_DropDown)         'filling the dropdown with ALL of the household members

'Next dialog - here so that the dropdown can be filled with case information
BeginDialog hospice_info_dlg, 0, 0, 291, 240, "Hospice Form Received"
  DropListBox 80, 25, 160, 45, HH_Memb_DropDown, client_in_hospice
  EditBox 80, 45, 205, 15, hospice_name
  EditBox 80, 65, 80, 15, npi_number
  EditBox 80, 85, 50, 15, hospice_entry_date
  EditBox 185, 85, 50, 15, hospice_exit_date
  EditBox 80, 105, 50, 15, mmis_updated_date
  EditBox 10, 140, 275, 15, reason_not_updated
  EditBox 10, 170, 275, 15, other_notes
  EditBox 80, 190, 205, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 215, 50, 15
    CancelButton 235, 215, 50, 15
  Text 15, 10, 140, 10, "Enter information from the Hospice Form"
  Text 30, 30, 45, 10, "Client Name:"
  Text 15, 50, 60, 10, "Name of Hospice:"
  Text 35, 70, 40, 10, "NPI Numbe:"
  Text 35, 90, 40, 10, "Entry Date:"
  Text 150, 90, 35, 10, "Exit Date:"
  Text 10, 110, 70, 10, "MMIS Updated as of "
  Text 10, 130, 165, 10, "If MMIS has not yet been updated, explain reason:"
  Text 10, 160, 50, 10, "Other Notes:"
  Text 10, 195, 60, 10, "Worker Signature:"
EndDialog

'DIALOG with a field for reason for exit - this may be added later
' BeginDialog hospice_info_dlg, 0, 0, 291, 255, "Hospice Form Received"
'   DropListBox 80, 25, 160, 45, "HH_Memb_DropDown", client_in_hospice
'   EditBox 80, 45, 205, 15, hospice_name
'   EditBox 80, 65, 80, 15, npi_number
'   EditBox 80, 85, 50, 15, hospice_entry_date
'   EditBox 185, 85, 50, 15, hospice_exit_date
'   EditBox 80, 105, 205, 15, exit_cause
'   EditBox 80, 125, 50, 15, mmis_updated_date
'   EditBox 10, 160, 275, 15, reason_not_updated
'   EditBox 10, 190, 275, 15, other_notes
'   EditBox 80, 210, 205, 15, worker_signature
'   ButtonGroup ButtonPressed
'     OkButton 180, 235, 50, 15
'     CancelButton 235, 235, 50, 15
'   Text 15, 10, 140, 10, "Enter information from the Hospice Form"
'   Text 30, 30, 45, 10, "Client Name:"
'   Text 15, 50, 60, 10, "Name of Hospice:"
'   Text 35, 70, 40, 10, "NPI Numbe:"
'   Text 35, 90, 40, 10, "Entry Date:"
'   Text 35, 110, 40, 10, "Exit due to:"
'   Text 10, 130, 70, 10, "MMIS Updated as of "
'   Text 10, 150, 165, 10, "If MMIS has not yet been updated, explain reason:"
'   Text 10, 180, 50, 10, "Other Notes:"
'   Text 10, 215, 60, 10, "Worker Signature:"
'   Text 150, 90, 35, 10, "Exit Date:"
' EndDialog

'showing the dialog
Do
    err_msg = ""

    Dialog hospice_info_dlg
    cancel_confirmation

    If trim(hospice_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the Hospice the client entered."       'hospice name required
    If IsDate(hospice_entry_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valide date for the Hospice Entry."   'entry date also required

    If err_msg <> "" Then MsgBox "Please resolve the following to conitune:" & vbNewLine & err_msg
Loop until err_msg = ""

'case noting the information from the dialog.
Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE("*** HOSPICE TRANSACTION FORM RECEIVED ***")
Call write_bullet_and_variable_in_CASE_NOTE("Client", client_in_hospice)
Call write_bullet_and_variable_in_CASE_NOTE("Hospice Name", hospice_name)
Call write_bullet_and_variable_in_CASE_NOTE("NPI Number", npi_number)
Call write_bullet_and_variable_in_CASE_NOTE("Date of Entry", hospice_entry_date)
Call write_bullet_and_variable_in_CASE_NOTE("Exit Date", hospice_exit_date)
'Call write_bullet_and_variable_in_MMIS_NOTE("Exit due to", exit_cause)         'This field is not currently in use so commented out - workers are testing, may add it back in
Call write_bullet_and_variable_in_CASE_NOTE("MMIS updated as of", mmis_updated_date)
Call write_bullet_and_variable_in_CASE_NOTE("MMIS not updated due to", reason_not_updated)
Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
