'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - Insert MBI from MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                      'manual run time in seconds
STATS_denomination = "M"                   'M is for Member
'END OF stats block=========================================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DELARACTIONS ==============================================================================================================
'Setting up the information for an array of all the clients.
const clt_full_name_const   = 0
const clt_pmi_const         = 1
const clt_ref_nbr_const     = 2
const clt_medi_mbi          = 3
const mbi_on_medi_const     = 4
const search_checkbox       = 5
const clt_mbi_found_on_RMCR = 6
const clt_rmcr_mbi          = 7
const clt_medi_exists       = 8
const clt_notes_const       = 9

Dim CLIENT_LIST_ARRAY()
ReDim CLIENT_LIST_ARRAY(clt_notes_const, 0)

end_msg = ""
'===========================================================================================================================

'THE SCRIPT ================================================================================================================
'connecting to MAXIS & grabbing the case number & footer month and year
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
If IsNumeric(MAXIS_footer_month) = FALSE Then       'If the footer month and year are not found, the script will default to the current month and year
    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr
End If

'Shows and defines the case number dialog
BeginDialog , 0, 0, 161, 65, "Case number and footer month"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  EditBox 95, 25, 25, 15, MAXIS_footer_month
  EditBox 130, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 95, 45, 30, 15
    CancelButton 125, 45, 30, 15
  Text 5, 10, 85, 10, "Enter your case number:"
  Text 10, 30, 75, 10, "Footer month and year:"
  Text 125, 30, 5, 10, "/"
EndDialog

Do
	Dialog 					'Calling a dialog without a assigned variable will call the most recently defined dialog
	cancel_confirmation
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks footer month and year. If footer month and year do not match the worker entry, it'll back out and get there manually.
Call MAXIS_footer_month_confirmation

'Setting up a list of all the clients on the case to gather actions and details.
memb_row = 5
client_counter = 0
Call navigate_to_MAXIS_screen ("STAT", "MEMB")      'Go to MEMB
Do
    EMReadScreen ref_numb, 2, memb_row, 3           'Reading the reference numbers on the case to navigate to MEMB for each member
    If ref_numb = "  " Then Exit Do                 'Once we read a blank reference number in the list we reach the end of the members
    ReDim Preserve CLIENT_LIST_ARRAY(clt_notes_const, client_counter)       'resizing the array to add the next member
    EMWriteScreen ref_numb, 20, 76                  'navigating to the MEMB panel
    transmit
    EMReadScreen first_name, 12, 6, 63              'Reading the client information from MEMB
    EMReadScreen last_name, 25, 6, 30
    EMReadScreen the_pmi, 8, 4, 46

    CLIENT_LIST_ARRAY(clt_ref_nbr_const, client_counter) = ref_numb         'Saving the client information to the Array
    CLIENT_LIST_ARRAY(clt_full_name_const, client_counter) = replace(first_name, "_", "") & " " & replace(last_name, "_", "")
    CLIENT_LIST_ARRAY(clt_pmi_const, client_counter) = trim(the_pmi)

    memb_row = memb_row + 1                         'Going to the next member and adding more room to the array
    client_counter = client_counter + 1
Loop until memb_row = 20

'Go to STAT:MEDI for each member in the household
'Read if the MBI exists on MEDI - if NOT - select this person by default to search for MBI
MEDI_exists_on_case = FALSE                         'setting this variable to default to false, we will look for MEDI panels and if one is on the case it will reset to true
Call navigate_to_MAXIS_screen ("STAT", "SUMM")
For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)      'looking through each of the members
    ref_nbr = CLIENT_LIST_ARRAY(clt_ref_nbr_const, person)      'going to MEDI for each person
    EMWriteScreen "MEDI", 20, 71
    EMWriteScreen ref_nbr, 20, 76
    transmit
    EMReadScreen panel_exists, 14, 24, 13                       'checking to see if a MEDI panel exists for this member

    If panel_exists = "DOES NOT EXIST" Then
        CLIENT_LIST_ARRAY(clt_medi_exists, person) = FALSE      'saving if the panel deos not exist for this person
    Else
        MEDI_exists_on_case = TRUE                              'resetting to know there is at least one MEDI panel
        CLIENT_LIST_ARRAY(clt_medi_exists, person) = True       'saving that the panel exists for this member

        EMReadScreen mbi_listed_on_medi, 13, 5, 38              'reading the field for MBI on this panel and aving the information to the array
        If mbi_listed_on_medi = "____ ___ ____" Then
            CLIENT_LIST_ARRAY(mbi_on_medi_const, person) = FALSE
            CLIENT_LIST_ARRAY(search_checkbox, person) = checked
        Else
            CLIENT_LIST_ARRAY(mbi_on_medi_const, person) = TRUE
        End If
        CLIENT_LIST_ARRAY(clt_medi_mbi, person) = mbi_listed_on_medi
    End If
Next
'Ending the script run if there are no MEDI panels because there is nothing the script can do.
If MEDI_exists_on_case = FALSE Then script_end_procedure_with_error_report("The script had ended as no member on this case has a MEDI panel. " & vbNewLine & vbNewLine & "This script does not have functionality to create a MEDI panel and the MBI must be entered on the MEDI panel.")

'Show a dialog asking which members to check for MBI
Do
    Do
        err_msg = ""

        dlg_len = 95 + 15 * UBound(CLIENT_LIST_ARRAY, 2)
        y_pos = 30

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 456, dlg_len, "Select Clients to have MBI filled"
          Text 10, 10, 60, 10, "Ref - Client Name"
          Text 165, 10, 15, 10, "PMI"
          Text 220, 10, 65, 10, "MBI Field on MEDI"
          Text 315, 10, 30, 10, "Update"
          For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)
              Text 10, y_pos, 145, 10, CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person)
              Text 165, y_pos, 40, 10, CLIENT_LIST_ARRAY(clt_pmi_const, person)
              Text 220, y_pos, 80, 10, CLIENT_LIST_ARRAY(clt_medi_mbi, person)
              CheckBox 315, y_pos, 140, 10, "Check here to look for the MBI in MMIS", CLIENT_LIST_ARRAY(search_checkbox, person)

            Y_pos = y_pos + 15
          Next
          y_pos = y_pos + 5
          ButtonGroup ButtonPressed
            OkButton 345, dlg_len - 20, 50, 15
            CancelButton 400, dlg_len - 20, 50, 15
          Text 10, y_pos, 440, 20, "Once you press 'OK' on this dialog, the script will go in to MMIS to attempt to find the MBI number on the RMCR panel for any client with the checkbox to the right of their name checked. If found, it will enter the information in the STAT:MEDI panel."
          y_pos = y_pos +25
          Text 10, y_pos, 230, 15, "At the end of the script run a message box will display actions taken. No CASE:NOTE is required or entered for this action."
        EndDialog

        dialog Dialog1
        cancel_without_confirmation

        For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)
            If CLIENT_LIST_ARRAY(search_checkbox, person) = checked AND CLIENT_LIST_ARRAY(clt_medi_exists, person) = FALSE then err_msg = err_msg & "* MEMB " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & " is selected to have MBI updated but no MEDI panel exists."
        Next
        If err_msg <> "" Then MsgBox "--- Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'checking to be sure we have someone checked
members_to_update = ""
members_to_ignore = ""
client_to_check = FALSE
For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)
    If CLIENT_LIST_ARRAY(search_checkbox, person) = checked then
        client_to_check = TRUE
        members_to_update= members_to_update & " - Memb " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & vbNewLine
    Else
        members_to_ignore = members_to_ignore & " - Memb " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & vbNewLine
    End If
Next

end_msg = end_msg & vbNewLine & "Members selected to have the script update: " & vbNewLine & members_to_update & vbNewLine & "Members to NOT update: " & vbNewLine & members_to_ignore & vbNewLine &_
                                "------- ACTIONS ---------" & vbNewLine

'Any member that is checked, use the PMI to go to MMIS and read for MBI on RMCR
If client_to_check = TRUE Then

    Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area
    For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)          'going through each person
        If CLIENT_LIST_ARRAY(search_checkbox, person) = checked then
            CLIENT_LIST_ARRAY(clt_pmi_const, person) = right("00000000" & CLIENT_LIST_ARRAY(clt_pmi_const, person), 8)
            the_pmi = CLIENT_LIST_ARRAY(clt_pmi_const, person)
            'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
            EMWriteScreen "I", 2, 19    'enter into case in MMIS in INQUIRY mode
            EMWriteScreen CLIENT_LIST_ARRAY(clt_pmi_const, person), 4, 19
            transmit

            EMWriteScreen "RMCR", 1, 8          'Going to RMCR to get the MBI
            transmit

            EMReadScreen rmcr_mbi, 11, 3, 56
            CLIENT_LIST_ARRAY(clt_rmcr_mbi, person) = trim(rmcr_mbi)
            If CLIENT_LIST_ARRAY(clt_rmcr_mbi, person) = "" Then CLIENT_LIST_ARRAY(clt_mbi_found_on_RMCR, person) = FALSE
            If CLIENT_LIST_ARRAY(clt_rmcr_mbi, person) <> "" Then CLIENT_LIST_ARRAY(clt_mbi_found_on_RMCR, person) = TRUE

            PF3
        End If
    Next

    Call navigate_to_MAXIS("PRODUCTION")                'Going back to MAXIS
    For person = 0 to UBound(CLIENT_LIST_ARRAY, 2)      'Going through each person
        If CLIENT_LIST_ARRAY(search_checkbox, person) = checked then            'If selected in the dialog to update for this person
            If CLIENT_LIST_ARRAY(clt_mbi_found_on_RMCR, person) = TRUE Then     'If the script actually found the MBI on RMCR

                call navigate_to_MAXIS_screen("STAT", "MEDI")
                EMWriteScreen ref_nbr, 20, 76
                transmit
                PF9

                MBI_Number = CLIENT_LIST_ARRAY(clt_rmcr_mbi, person)
                MBI_one = left(MBI_Number, 4)
                MBI_two = left(right(MBI_Number, 7), 3)
                MBI_three = right(MBI_Number, 4)

                'enter each of the sections on to the MEDI panel
                EMWriteScreen MBI_one, 5, 38
                EMWriteScreen MBI_two, 5, 43
                EMWriteScreen MBI_three, 5, 47

                EMWriteScreen "M", 5, 64        'entering the source of the MBI information - which came from MMIS for this process
                transmit            'transmit to save the information to the panel

                EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
                look_for_error = trim(look_for_error)

                If look_for_error = "WARNING:" Then     'we can transmit past warning messages and then look again
                    transmit
                    EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
                    look_for_error = trim(look_for_error)
                End If

                If look_for_error <> "" Then        'if there is anything here - assume an error
                    PF10                            'blank out the work

                    end_msg = end_msg & vbNewLine & "The MEDI panel could NOT be updated for " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & " as the MAXIS panel had an error when trying to save the update." & vbNewLine &_
                                                    "The MBI for this person is: " & MBI_one & " " & MBI_two & " " & MBI_three & vbNewLine
                    'MsgBox "ERROR"
                Else                                'if no error the number should have saved
                    'Read the MBI number to be sure it succeeded.
                    EMReadScreen Check_MBI_one, 4, 5, 38
                    EMReadScreen Check_MBI_two, 3, 5, 43
                    EMReadScreen Check_MBI_three, 4, 5, 47

                    EMReadScreen Check_source, 1, 5, 64             'reading the saved source code and reformatting for the escel file
                    If Check_source = "_" Then Check_source = ""

                    CHECK_MBI = Check_MBI_one & Check_MBI_two & Check_MBI_three

                    If CHECK_MBI = MBI_Number Then          'If it succeeded then enter 'DONE' to the action column.
                        end_msg = end_msg & vbNewLine & "Memb " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & " Updated" & vbNewLine & "MBI: " & MBI_one & " " & MBI_two & " " & MBI_three &_
                                                        " has been added to MEDI and coded the source from MMIS." & vbNewLine
                        'MsgBox "DONE"
                    Else            'If it did not succeed then enter 'FAILED' to the action column.
                        end_msg = end_msg & vbNewLine & "FAILED - Memb " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & " update FAILED - " & vbNewLine & "MBI: " & MBI_one & " " & MBI_two & " " & MBI_three &_
                                                        " has NOT been added to MEDI and coded the source from MMIS." & vbNewLine
                        'MsgBox "FAILED"
                    End If
                End If
            Else
                end_msg = end_msg & vbNewLine & "No MBI for Memb " & CLIENT_LIST_ARRAY(clt_ref_nbr_const, person) & " - " & CLIENT_LIST_ARRAY(clt_full_name_const, person) & " found in MMIS on RMCR. Script could not update MEDI." & vbNewLine
            End If

        End If
    Next
    end_msg = "Script run Complete" & vbNewLine & vbNewLine & end_msg
Else
    end_msg = "Script run Complete" & vbNewLine & vbNewLine & "There were no clients checked to have the script pull MBI information from MMIS. The script has taken no action to update any STAT panel." & vbNewLine & end_msg & "NONE"
End If
Call back_to_SELF       'This saves the updates and send the case through background
Call navigate_to_MAXIS_screen("STAT", "MEDI")

call script_end_procedure_with_error_report(end_msg)
