'**********THIS IS A RAMSEY SPECIFIC SCRIPT.  IF YOU REVERSE ENGINEER THIS SCRIPT, JUST BE CAREFUL.************
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MHC ENROLLMENT.vbs"
start_time = timer

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
call changelog_update("09/11/2020", "This script now contains the functionality for Open Enrollment and for any other enrollment.##~## ##~##The seperate script for Open Enrollment will no longer be available.##~##", "Casey Love, Hennepin County")
call changelog_update("12/19/2019", "Added IM 12 as an option for contract codes.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Added the ability for the script to delete the current enrollment plan if the beginning date for the current plan is the same as the new enrollment date.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Added the ability for the script to delete the Delayed Decision exclusion if the start date is the same as the enrollment date.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Changed the script coding on REFM screen to enter 'N' if enrollment information source was NOT the Paper Enrollment Form.", "Casey Love, Hennepin County")
call changelog_update("04/17/2019", "Resolving a BUG for METS cases enrolling for the first time, no exclusion code is defaulted.", "Casey Love, Hennepin County")
call changelog_update("04/16/2019", "BUG when disenrolling and reenrolling in a different plan. Functionality should work to disenroll and renroll in the same run - specific to issues discovered with NT option.", "Casey Love, Hennepin County")
call changelog_update("04/02/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


Function get_to_RKEY()
    EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
    IF MMIS_panel_check <> "RKEY" THEN
        attempt = 1
        DO
            If MMIS_case_number = "" Then Call MMIS_case_number_finder(MMIS_case_number)
            PF6
            EMReadScreen MMIS_panel_check, 4, 1, 52
            attempt = attempt + 1
            If attempt = 15 Then Exit Do
        Loop Until MMIS_panel_check = "RKEY"
    End If
    EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
    IF MMIS_panel_check <> "RKEY" THEN
    	DO
    		PF6
    		EMReadScreen session_terminated_check, 18, 1, 7
    	LOOP until session_terminated_check = "SESSION TERMINATED"

        'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
        EMWriteScreen "MW00", 1, 2
        transmit
        transmit

        EMReadScreen MMIS_menu, 24, 3, 30
	    If MMIS_menu = "GROUP SECURITY SELECTION" Then
            row = 1
            col = 1
            EMSearch " C3", row, col
            If row <> 0 Then
                EMWriteScreen "x", row, 4
                transmit
            Else
                row = 1
                col = 1
                EMSearch " C4", row, col
                If row <> 0 Then
                    EMWriteScreen "x", row, 4
                    transmit
                Else
                    script_end_procedure_with_error_report("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
                End If
            End If

            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "x", row, col - 3
            transmit
        Else
            'Now it finds the recipient file application feature and selects it.
            row = 1
            col = 1
            EMSearch "RECIPIENT FILE APPLICATION", row, col
            EMWriteScreen "x", row, col - 3
            transmit
        End If
    END IF
End Function

Function write_variable_in_MMIS_NOTE(variable)
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
			If character_test <> " " or noting_row >= 20 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 20 then
                    PF11
                    noting_row = 5
				End if
			End if
		Loop until character_test = " "

        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array
            word = trim(word)
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 8
			End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 20 then
                PF11
                noting_row = 5
			End if

            'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
    End if
End Function

Function write_bullet_and_variable_in_MMIS_NOTE(bullet, variable)
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
        Do
            EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            If character_test <> " " or noting_row >= 20 then
                noting_row = noting_row + 1

                'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
                If noting_row >= 20 then
                    PF11
                    noting_row = 5
                End if
            End if
        Loop until character_test = " "

        'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
        If len(bullet) >= 14 then
            indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
        Else
            indent_length = len(bullet) + 4 'It's four more for the reason explained above.
        End if

        'Writes the bullet
        EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

        'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
        noting_col = noting_col + (len(bullet) + 4)

        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array

            'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
            If len(word) + noting_col > 80 then
                noting_row = noting_row + 1
                noting_col = 8
            End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
            If noting_row >= 20 then
                PF11
                noting_row = 5
            End if

            'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
            If noting_col = 8 then
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Writes the word and a space using EMWriteScreen
            EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

            'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
            If right(word, 1) = ";" then
                noting_row = noting_row + 1
                noting_col = 8
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Increases noting_col the length of the word + 1 (for the space)
            noting_col = noting_col + (len(word) + 1)
        Next

        'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
        EMSetCursor noting_row + 1, 3
    End if
End Function

Function MMIS_case_number_finder(MMIS_case_number)
    row = 1
    col = 1
    EMSearch "CASE NUMBER:", row, col
    If row <> 0 Then
        EMReadScreen MMIS_case_number, 8, row, col + 13
        MMIS_case_number = trim(MMIS_case_number)
    End If
    If MMIS_case_number = "" Then
        row = 1
        col = 1
        EMSearch "CASE NBR:", row, col
        If row <> 0 Then
            EMReadScreen MMIS_case_number, 8, row, col + 10
            MMIS_case_number = trim(MMIS_case_number)
        End If
    End If
    If MMIS_case_number = "" Then
        row = 1
        col = 1
        EMSearch "CASE:", row, col
        If row <> 0 Then
            EMReadScreen MMIS_case_number, 8, row, col + 6
            MMIS_case_number = trim(MMIS_case_number)
        End If
    End If
End Function

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
' testing_run = TRUE
'call check_for_MMIS(True) 'Sending MMIS back to the beginning screen and checking for a password prompt
Call MMIS_case_number_finder(MMIS_case_number)

Call get_to_RKEY

'grabs the PMI number if one is listed on RKEY
If MMIS_case_number = "" Then
    EMReadscreen MMIS_case_number, 8, 9, 19
    MMIS_case_number= trim(MMIS_case_number)
End If

open_enrollment_case = FALSE
If Month(date) = 10 OR Month(date) = 12 OR Month(date) = 11 Then
	ask_if_open_enrollment = MsgBox("Are you processing an Open Enrollment?", vbQuestion + vbYesNo, "Open Enrollment?")
	If ask_if_open_enrollment = vbYes Then
		enrollment_month = "01"
		enrollment_year = "21"
		open_enrollment_case = TRUE
		case_open_enrollment_yn = "Yes"
	End If
End If

IF open_enrollment_case = FALSE Then
	enrollment_month = CM_plus_1_mo
	enrollment_year = CM_plus_1_yr

	this_month = monthname(month(date))
	Select Case this_month
	    Case "January"
	        cut_off_date = #01/22/2020#
	    Case "February"
	        cut_off_date = #2/19/2020#
	    Case "March"
	        cut_off_date = #3/20/2020#
	    Case "April"
	        cut_off_date = #4/21/2020#
	    Case "May"
	        cut_off_date = #5/19/2020#
	    Case "June"
	        cut_off_date = #6/19/2020#
	    Case "July"
	        cut_off_date = #7/22/2020#
	    Case "August"
	        cut_off_date = #8/20/2020#
	    Case "September"
	        cut_off_date = #9/21/2020#
	    Case "October"
	        cut_off_date = #10/21/2020#
	    Case "November"
	        cut_off_date = #11/17/2020#
	    Case "December"
	        cut_off_date = #12/21/2020#
	End Select
	'MsgBox cut_off_date
	If cut_off_date <> "" Then
	    If DateDiff("d", date, cut_off_date) < 0 Then
	        'MsgBox DateDiff("d", date, cut_off_date)
	        enrollment_month = CM_plus_2_mo
	        enrollment_year = CM_plus_2_yr
	    End If
	End If
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 206, 180, "Enrollment Information"
  EditBox 90, 25, 60, 15, MMIS_case_number
  EditBox 90, 45, 20, 15, enrollment_month
  EditBox 115, 45, 20, 15, enrollment_year
  DropListBox 55, 75, 95, 15, "Select one..."+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare", Health_plan
  CheckBox 120, 95, 25, 10, "Yes", Insurance_yes
  CheckBox 120, 105, 25, 10, "Yes", foster_care_yes
  DropListBox 110, 120, 90, 45, "Select One..."+chr(9)+"Phone"+chr(9)+"Paper Enrollment Form"+chr(9)+"Morning Letters", enrollment_source
  DropListBox 110, 140, 50, 45, "No"+chr(9)+"Yes", case_open_enrollment_yn
  ButtonGroup ButtonPressed
    OkButton 95, 160, 50, 15
    CancelButton 150, 160, 50, 15
  GroupBox 5, 10, 150, 55, "Leading zeros not needed"
  Text 10, 30, 50, 10, "Case Number:"
  Text 10, 50, 80, 10, "Enrollment Month/Year:"
  Text 10, 80, 40, 10, "Health plan:"
  Text 10, 95, 100, 10, "Other Insurance for this case?"
  Text 10, 105, 50, 10, "Foster Care?"
  Text 10, 125, 100, 10, "Enrollment was requested via"
  Text 20, 145, 85, 10, "Is this Open Enrollment?"
EndDialog

'do the dialog here
Do
    err_msg = ""

	Dialog Dialog1
	cancel_without_confirmation

    MMIS_case_number = trim(MMIS_case_number)

    If MMIS_case_number = "" then err_msg = err_msg & vbNewLine & "* Enter the case number."
	If enrollment_source = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate where the request for the enrollment came from (phone call or enrollment form)."
	If case_open_enrollment_yn = "Yes" Then
		enrollment_month = "01"
		enrollment_year = "21"
		open_enrollment_case = TRUE
	Else
		If enrollment_month = "" OR enrollment_year = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year enrollment is effective."
	End If
    If health_plan = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select a health plan."

    If err_msg <> "" Then MsgBOx "Please resolve to continue: " & vbNewLine & err_msg
Loop until err_msg = ""
If case_open_enrollment_yn = "No" Then open_enrollment_case = FALSE
If open_enrollment_case = TRUE then testing_run = TRUE
MAXIS_case_number = MMIS_case_number

If Insurance_yes = checked then
	insurance_yn = "Y"
Else
	insurance_yn = "N"
End If

If foster_care_yes = checked Then
	foster_care_yn = "Y"
Else
	foster_care_yn = "N"
End if

'checking for an active MMIS session
Call check_for_MMIS(True)
Call get_to_RKEY

'formatting variables----------------------------------------------------------------------------------------------------
If len(enrollment_month) = 1 THEN enrollment_month = "0" & enrollment_month
IF len(enrollment_year) <> 2 THEN enrollment_year = right(enrollment_year, 2)

MNSURE_Case = False
If len(MMIS_case_number) = 8 AND left(MMIS_case_number, 1) <> 0 THEN MNSURE_Case = TRUE
MMIS_case_number = right("00000000" & MMIS_case_number, 8)

'MsgBox "MNSure Case? " & MNSURE_Case & vbNewLine & MMIS_case_number
enrollment_date = enrollment_month & "/01/" & enrollment_year

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "i", 2, 19
Call clear_line_of_text(4, 19)		'Clearing all of the search options used on RKEY as we must ONLY enter a case number
Call clear_line_of_text(5, 19)
Call clear_line_of_text(5, 48)
Call clear_line_of_text(6, 19)
Call clear_line_of_text(6, 48)
Call clear_line_of_text(6, 69)
Call clear_line_of_text(9, 19)
Call clear_line_of_text(9, 48)
Call clear_line_of_text(9, 69)

EMWriteScreen MMIS_case_number, 9, 19
transmit
transmit
transmit
EMReadscreen RCIN_check, 4, 1, 49
If RCIN_check <> "RCIN" then script_end_procedure_with_error_report("The listed Case number was not found. Check your Case number and try again.")

Dim listed_clients_array
ReDim listed_clients_array (0)


rcin_row = 11
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen pmi_nbr, 8, rcin_row, 4
	EMReadscreen last_name, 17, rcin_row, 24
	EMReadscreen first_name, 9, rcin_row, 42
	last_name = trim(last_name)
	first_name = trim(first_name)
	client_string = pmi_nbr & " - " & last_name & ", " & first_name
	client_array = client_array & client_string & "|"
	rcin_row = rcin_row + 1
	If rcin_row = 21 Then
		PF8
		EMReadScreen end_rcin, 6, 24, 2
		If end_rcin = "CANNOT" then Exit Do
		rcin_row = 11
	End If
	Emreadscreen last_clt_check, 8, rcin_row, 4
LOOP until last_clt_check = "        "			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

DIM all_client_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	PF7
	PF7
	Do
		Search_pmi = left(Interim_array(x), 8)
		row = 1
		col = 1
		EMSearch Search_pmi, row, col
		If row = 0 then
			PF8
		Else
			EMReadScreen hc_status, 1, row, 76
			If hc_status = "A" Then all_clients_array(x, 1) = 1
			Exit Do
		End If
		EMReadScreen end_rcin, 6, 24, 2
	Loop until end_rcin = "CANNOT"
NEXT

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 250, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 105, 10, "Household members to look at:"
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 175, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
	ButtonGroup ButtonPressed
	OkButton 195, 10, 50, 15
	CancelButton 195, 30, 50, 15
EndDialog

'runs the dialog that has been dynamically created. Streamlined with new functions.
Dialog Dialog1
If buttonpressed = 0 then stopscript

HH_member_array = ""

FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			'msgbox all_clients_
			HH_member_array = HH_member_array & left(all_clients_array(i, 0), 8) & " "
		END IF
	END IF
NEXT

HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
HH_member_array = SPLIT(HH_member_array, " ")

const client_name  = 0
const client_pmi   = 1
const current_plan = 2
const new_plan     = 3
const change_rsn   = 4
const disenrol_rsn = 5
const med_code     = 6
const dent_code    = 7
const contr_code   = 8
const preg_yes 	   = 9
const interp_code  = 10
const enrol_sucs   = 11

Dim MMIS_clients_array
ReDim MMIS_clients_array (12, 0)

EMReadScreen RCIN_check, 4, 1, 49
If RCIN_check = "RCIN" Then PF6
Call get_to_RKEY

item = 0

For each member in HH_member_array
	ReDim Preserve MMIS_clients_array(12, item)
	EMWriteScreen "I", 2, 19
	EMWriteScreen member, 4, 19
	EMWriteScreen "        ", 9, 19
	transmit
	MMIS_clients_array (client_pmi, item) = member
	EMReadScreen last_name, 18, 3, 2
	EMReadScreen first_name, 12, 3, 20
	last_name = trim(last_name)
	first_name = trim(first_name)
	MMIS_clients_array (client_name, item) = last_name & ", " & first_name

	'check RPOL to see if there is other insurance available, if so worker processes manually
	'EMWriteScreen "X", 11, 2
	'Transmit
	EMWriteScreen "RPOL", 1, 8
	transmit
	'making sure script got to right panel
	EMReadScreen RPOL_check, 4, 1, 52
	If RPOL_check <> "RPOL" then script_end_procedure_with_error_report("The script was unable to navigate to RPOL process manually if needed.")

	EMreadscreen policy_number, 1, 7, 8
    If policy_number <> " " then

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
          CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
          ButtonGroup ButtonPressed
            OkButton 105, 125, 50, 15
          Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
          GroupBox 10, 45, 145, 70, "REVIEW RPOL"
          Text 50, 60, 65, 10, "*** Check RPOL ***"
          Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
        EndDialog

        dialog Dialog1

        If rpol_ended_checkbox = unchecked Then
            PF6
    		script_end_procedure_with_error_report ("This case has spans on RPOL. Please evaluate manually at this time.")
        End If
    End If

	PF6

	EMWriteScreen "RPPH", 1, 8
	transmit
	row = 1
	col = 1
	EMSearch "99/99/99", row, col
	IF row < 10 Then
		If col = 18 Then
			EMReadScreen excl_code, 2, row, 2
		ElseIf col = 45 Then
			EMReadScreen excl_code, 2, row, 29
		ElseIf col = 72 Then
			EMReadScreen excl_code, 2, row, 56
		End If
		If excl_code = "AA" Then MMIS_clients_array(current_plan, item) = "XCL - Adoption Assistance"
		If excl_code = "AB" Then MMIS_clients_array(current_plan, item) = "XCL - Part A or B Only"
		If excl_code = "BB" Then MMIS_clients_array(current_plan, item) = "XCL - Blind/Disabled under 65 years"
		If excl_code = "CC" Then MMIS_clients_array(current_plan, item) = "XCL - Child Protection Case"
		If excl_code = "CD" Then MMIS_clients_array(current_plan, item) = "XCL - Chemical Dependant Pilot"
		If excl_code = "CS" Then MMIS_clients_array(current_plan, item) = "XCL - Condumer Support Grant"
		If excl_code = "CV" Then MMIS_clients_array(current_plan, item) = "XCL - Center for Victims of Torture"
		If excl_code = "DD" Then MMIS_clients_array(current_plan, item) = "XCL - Communicable Disease"
		If excl_code = "DO" Then MMIS_clients_array(current_plan, item) = "XCL - Diability Opt Out"
		If excl_code = "EE" Then MMIS_clients_array(current_plan, item) = "XCL - SED/SPMI"
		If excl_code = "FF" Then MMIS_clients_array(current_plan, item) = "XCL - Child in Foster Care"
		If excl_code = "GG" THen MMIS_clients_array(current_plan, item) = "XCL - Geographic Exclusion"
		If excl_code = "HH" Then MMIS_clients_array(current_plan, item) = "XCL - Private HMO Coverage"
		If excl_code = "II" Then MMIS_clients_array(current_plan, item) = "XCL - Breast/Cervical Cancer"
		If excl_code = "IP" THen MMIS_clients_array(current_plan, item) = "XCL - Insurance Pending"
		If excl_code = "KK" Then MMIS_clients_array(current_plan, item) = "XCL - Elderly Waiver"
		If excl_code = "LL" Then MMIS_clients_array(current_plan, item) = "XCL - Personal Care Attendent"
		If excl_code = "MD" Then MMIS_clients_array(current_plan, item) = "XCL - MA Delay"
		If excl_code = "MM" Then MMIS_clients_array(current_plan, item) = "XCL - Native American on Reservation"
		If excl_code = "MS" Then MMIS_clients_array(current_plan, item) = "XCL - MNSURE Tracking"
		If excl_code = "PC" Then MMIS_clients_array(current_plan, item) = "XCL - Payment County"
		If excl_code = "QQ" Then MMIS_clients_array(current_plan, item) = "XCL - QMB/SLMB Eligibility"
		If excl_code = "RR" Then MMIS_clients_array(current_plan, item) = "XCL - Refugee/EMA/EGA"
		If excl_code = "SS" Then MMIS_clients_array(current_plan, item) = "XCL - Medical Spenddown"
		If excl_code = "TT" Then MMIS_clients_array(current_plan, item) = "XCL - Terminal Illness"
		If excl_code = "UU" Then MMIS_clients_array(current_plan, item) = "XCL - Limited Disability"
		If excl_code = "WW" Then MMIS_clients_array(current_plan, item) = "XCL - Delayed Nursing Home"
		If excl_code = "YY" Then MMIS_clients_array(current_plan, item) = "XCL - Delayed Decision"
		If excl_code = "ZZ" Then MMIS_clients_array(current_plan, item) = "XCL - RTC/IMD Resident"
	Else
		EMReadScreen hp_code, 10, row, 23

		If hp_code = "A585713900" then MMIS_clients_array(current_plan, item) = "Health Partners"
		If hp_code = "A565813600" then MMIS_clients_array(current_plan, item) = "Ucare"
		If hp_code = "A405713900" then MMIS_clients_array(current_plan, item) = "Medica"
		If hp_code = "A065813800" then MMIS_clients_array(current_plan, item) = "Blue Plus"
		If hp_code = "A836618200" then MMIS_clients_array(current_plan, item) = "Hennepin Health PMAP"
		If hp_code = "A965713400" then MMIS_clients_array(current_plan, item) = "Hennepin Health SNBC"
	End If
	MMIS_clients_array(new_plan,     item) = health_plan
	MMIS_clients_array(change_rsn,   item) = change_reason
	MMIS_clients_array(disenrol_rsn, item) = disenrollment_reason
	PF6
	EMWaitReady 0, 0
	item = item + 1
Next

x = 0
max = Ubound(MMIS_clients_array, 2)
dlg_len = 60
If enrollment_source = "Phone" Then
    dlg_len = dlg_len + 20
End If
If enrollment_source = "Paper Enrollment Form" Then
	dlg_len = dlg_len + 15
End If

name_list = ""
For person = 0 to Ubound(MMIS_clients_array, 2)
    name_list = name_list & +chr(9)+MMIS_clients_array(first_name_ini, person)
    dlg_len = dlg_len + 20
Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 750, dlg_len, "Enrollment Information"
  Text 5, 5, 25, 10, "Name"
  Text 100, 5, 15, 10, "PMI"
  Text 145, 5, 75, 10, "Current Plan/Exclusion"
  Text 250, 5, 50, 10, "Medical Clinic"
  Text 310, 5, 45, 10, "Dental Clinic"
  Text 370, 5, 40, 10, "Health plan:"
  Text 440, 5, 55, 10, "Contract Code:"
  Text 500, 5, 55, 10, "Change reason:"
  Text 565, 5, 60, 10, "Disenroll reason:"
  Text 640, 5, 35, 10, "Pregnant?"
  Text 695, 5, 55, 10, "Interpreter Code"

  For person = 0 to Ubound(MMIS_clients_array, 2)
    If enrollment_source = "Morning Letters" Then MMIS_clients_array(change_rsn, person) = "Reenrollment"
    If MMIS_clients_array(new_plan, person) = "Medica" Then MMIS_clients_array(contr_code, person) = "MA 30"
	If open_enrollment_case = TRUE Then
		MMIS_clients_array(change_rsn, person) = "Open enrollment"
		MMIS_clients_array(disenrol_rsn, person) = "Open Enrollment"
	End If
  	Text 5, (x * 20) + 25, 95, 10, MMIS_clients_array(client_name, person)
  	Text 100, (x * 20) + 25, 35, 10, MMIS_clients_array(client_pmi, person)
  	Text 145, (x * 20) + 25, 95, 10, MMIS_clients_array(current_plan, person)
  	EditBox 250, (x * 20) + 20, 55, 15, MMIS_clients_array(med_code, person)
  	EditBox 310, (x * 20) + 20, 50, 15, MMIS_clients_array(dent_code, person)
    DropListBox 370, (x * 20) + 20, 60, 15, " "+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare", MMIS_clients_array(new_plan, person)
  	DropListBox 440, (x * 20) + 20, 50, 15, "MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35"+chr(9)+"IM 12", MMIS_clients_array(contr_code, person)
	DropListBox 500, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn, person)
  	DropListBox 565, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", MMIS_clients_array(disenrol_rsn, person)
  	CheckBox 645, (x * 20) + 20, 25, 10, "Yes", MMIS_clients_array(preg_yes, person)
	EditBox 700, (x * 20) + 20, 25, 15, MMIS_clients_array(interp_code, person)
	x = x + 1
  Next

  Text 5, (x * 20) + 25, 45, 10, "Other Notes:"
  EditBox 55, (x * 20) + 20, 690, 15, other_notes

  If enrollment_source = "Phone" Then
      GroupBox 5, (x * 20) + 40, 410, 35, "Phone Call Information"
      Text 10, (x * 20) + 60, 40, 10, "Caller name"
      ComboBox 55, (x * 20) + 55, 120, 45, " " & name_list, caller_name
      Text 180, (x * 20) + 60, 40, 10, ", who is the"
      ComboBox 225, (x * 20) + 55, 80, 45, "Client"+chr(9)+"AREP", caller_rela
      CheckBox 340, (x * 20) + 55, 65, 10, "Used Interpreter", used_interpreter_checkbox
      x = x + 1
  End If
  If enrollment_source = "Paper Enrollment Form" Then
	  GroupBox 5, (x * 20) + 40, 180, 30, "Paper Form Information"
	  Text 10, (x * 20) + 55, 80, 10, "Form Received Date:"
	  EditBox 95, (x * 20) + 50, 80, 15, form_received_date
  End If

  Text 445, dlg_len - 15, 60, 10, "Worker Signature"
  EditBox 510, dlg_len - 20, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 640, dlg_len - 20, 50, 15
    CancelButton 695, dlg_len - 20, 50, 15
EndDialog

Do
    err_msg = ""

	Dialog Dialog1
	cancel_confirmation

    For person = 0 to Ubound(MMIS_clients_array, 2)
        If left(MMIS_clients_array(current_plan, person), 3) <> "XCL" AND trim(MMIS_clients_array(current_plan, person)) <> "" Then
            If MMIS_clients_array(disenrol_rsn, person) = "Select one..." Then err_msg = err_msg & vbNewLine & "* Since " & MMIS_clients_array(client_name, person) & " is currently on a health plan, please select a disenrollment reason for the " & MMIS_clients_array(current_plan, person) & " plan."
        End If
        If MMIS_clients_array(change_rsn, person) = "Select one..." Then err_msg = err_msg & vbNewLine & "* Select a reason to enroll  " & MMIS_clients_array(client_name, person) & " into a new plan."
    Next

    If enrollment_source = "Phone" Then

        If trim(caller_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the caller."
        If trim(caller_rela) = "" Then err_msg = err_msg & vbNewLine & "* Select who is calling (typically Client or AREP)."

    End If

    If worker_signature = "" THen err_msg = err_msg & vbNewLine & "* Enter your name for the case note signature."

    If err_msg <> "" THen MsgBox "Please resovle to continue:" & vbNewLine & err_msg
	If ButtonPressed = 0 Then err_msg = "LOOP"

Loop Until err_msg = ""

process_manually_message = ""

If MNSURE_Case = TRUE Then
	For member = 0 to Ubound(MMIS_clients_array, 2)
		'MMIS Codes
		Enrollment_date = Enrollment_month & "/01/" & enrollment_year
		'change reasons
		If MMIS_clients_array(change_rsn, member) = "First year change option" 	then change_reason = "FY"
		If MMIS_clients_array(change_rsn, member) = "Health plan contract end" 	then change_reason = "HP"
		If MMIS_clients_array(change_rsn, member) = "Initial enrollment"       	then change_reason = "IN"
		If MMIS_clients_array(change_rsn, member) = "Move"                     	then change_reason = "MV"
		If MMIS_clients_array(change_rsn, member) = "Ninety Day change option" 	then change_reason = "NT"
		If MMIS_clients_array(change_rsn, member) = "Open enrollment"    	  	then change_reason = "OE"
		If MMIS_clients_array(change_rsn, member) = "PMI merge" 				then change_reason = "PM"
		If MMIS_clients_array(change_rsn, member) = "Reenrollment" 			  	then change_reason = "RE"
		If MMIS_clients_array(change_rsn, member) = "Select one..." 			then change_reason = ""

		'Disenrollment reasons
		If MMIS_clients_array(disenrol_rsn, member) = "Eligibility ended"        then disenrollment_reason = "EE"
		If MMIS_clients_array(disenrol_rsn, member) = "Exclusion"                then disenrollment_reason = "EX"
		If MMIS_clients_array(disenrol_rsn, member) = "First year change option" then disenrollment_reason = "FY"
		If MMIS_clients_array(disenrol_rsn, member) = "Health plan contract end" then disenrollment_reason = "HP"
		If MMIS_clients_array(disenrol_rsn, member) = "Jail - Incarceration"     then disenrollment_reason = "JL"
		If MMIS_clients_array(disenrol_rsn, member) = "Move"                     then disenrollment_reason = "MV"
		If MMIS_clients_array(disenrol_rsn, member) = "Loss of disability"       then disenrollment_reason = "ND"
		If MMIS_clients_array(disenrol_rsn, member) = "Ninety Day change option" then disenrollment_reason = "NT"
		If MMIS_clients_array(disenrol_rsn, member) = "Open Enrollment"          then disenrollment_reason = "OE"
		If MMIS_clients_array(disenrol_rsn, member) = "PMI merge"                then disenrollment_reason = "PM"
		If MMIS_clients_array(disenrol_rsn, member) = "Voluntary"                then disenrollment_reason = "VL"
		If MMIS_clients_array(disenrol_rsn, member) = "Select one..."            then disenrollment_reason = ""

		'REFM Codes
		If MMIS_clients_array(preg_yes, member) = checked Then
			pregnant_yn = "Y"
		Else
			pregnant_yn = "N"
		End If

		If MMIS_clients_array(interp_code, member) = "" Then
			interpreter_yn = "N"
		Else
			interpreter_yn = "Y"
		End If

		EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
		IF MMIS_panel_check <> "RKEY" THEN
			DO
				PF6
				EMReadScreen session_terminated_check, 18, 1, 7
			LOOP until session_terminated_check = "SESSION TERMINATED"
			'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
			EMWriteScreen "mw00", 1, 2
			transmit
			transmit
			EMWriteScreen "x", 8, 3
			transmit
		END IF
		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
		EMWriteScreen "c", 2, 19
		EMWriteScreen "        ", 9, 19
		EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19
		transmit
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine

		DO
			'check RPOL to see if there is other insurance available, if so worker processes manually
			EMWriteScreen "rpol", 1, 8
			transmit
			'making sure script got to right panel
			EMReadScreen RPOL_check, 4, 1, 52
			If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
			EMreadscreen policy_number, 1, 7, 8
            If policy_number <> " " then

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
                  CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
                  ButtonGroup ButtonPressed
                    OkButton 105, 125, 50, 15
                  Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
                  GroupBox 10, 45, 145, 70, "REVIEW RPOL"
                  Text 50, 60, 65, 10, "*** Check RPOL ***"
                  Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
                EndDialog

                dialog Dialog1

                If rpol_ended_checkbox = unchecked Then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
            End If
			'nav to RPPH
			EMWriteScreen "rpph", 1, 8
			transmit

			'making sure script got to right panel
			EMReadScreen RPPH_check, 4, 1, 52
			If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

			Enrollment_date = Enrollment_month & "/01/" & enrollment_year
			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year
			' msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" Then
					row = 1
					col = 1
					EMSearch "99/99/99", row, col
                    If col <> 0 Then
                        EMReadScreen beg_of_excl, 8, row, 9
                        IF beg_of_excl = enrollment_date Then
                            EMWriteScreen "...", row, 2
                        Else
                            EMWriteScreen xcl_end_date, row, col
                        End if
                    End If
				Else
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If
			End If
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
			' msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"

			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)

            If process_manually_message = "" Then
    			'enter disenrollment reason
                If disenrollment_reason <> "" Then
                    EMReadScreen beg_of_curr_span, 8, 13, 5
                    If beg_of_curr_span = enrollment_date Then
                        EMWriteScreen "...", 13, 5
                    Else
                        EMWriteScreen xcl_end_date, 13, 14
                        EMWriteScreen disenrollment_reason, 13, 75
                    End If
                End If

    			'resets to bottom of the span list.
    			pf11

    			'enter enrollment date
    			EMWriteScreen enrollment_date, 13, 5
    			'enter managed care plan code
    			EMWriteScreen health_plan_code, 13, 23
    			'enter contract code
    			EMWriteScreen contract_code_part_one, 13, 34
    			EMWriteScreen contract_code_part_two, 13, 37
    			'enter change reason
    			EMWriteScreen change_reason, 13, 71

    			EMWaitReady 0, 0

    			EMReadScreen false_end, 8, 14, 14
    			If false_end = "99/99/99" Then
    				EMReadScreen double_check, 2, 14, 5
    				If double_check = "  " Then EMWriteScreen "...", 14, 5
    			End If

    			' msgbox "RPPH updated"

    			'REFM screen
    			EMWriteScreen "refm", 1, 8
    			transmit
    			EMReadScreen RPPH_error_check, 10, 24, 2
    			If trim(RPPH_error_check) = "EXCLSN END" then
    				Do
                        Dialog1 = ""
                        BeginDialog Dialog1, 0, 0, 191, 45, "Exclusion Code Error"
                          ButtonGroup ButtonPressed
                            OkButton 85, 25, 50, 15
                            CancelButton 135, 25, 50, 15
                          Text 15, 10, 155, 10, "Update the exclusion code field, then press OK."
                        EndDialog

    					Dialog Dialog1
    					cancel_confirmation
    					transmit
    					EMReadScreen RPPH_error_check, 10, 24, 2
    				Loop until trim(RPPH_error_check) <> "EXCLSN END"
    				' Msgbox "Updated the exclusion code field, then press OK."
    				' transmit
    			ELSEIF trim(RPPH_error_check) <> "" then

                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 236, 110, "RPPH error detected"
                      DropListBox 70, 50, 160, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", change_reason
                      DropListBox 70, 65, 160, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", disenrollment_reason
                      ButtonGroup ButtonPressed
                        OkButton 125, 85, 50, 15
                        CancelButton 180, 85, 50, 15
                      Text 10, 55, 55, 10, "Change reason:"
                      Text 10, 70, 60, 10, "Disenroll reason:"
                      ButtonGroup ButtonPressed
                        OkButton 155, 330, 50, 15
                      Text 15, 20, 210, 10, "* Initial enrollment is selected, but has been enrolled previously"
                      GroupBox 5, 5, 225, 40, "An error occurred on in RPPH. Typical errors include:"
                      Text 15, 30, 210, 10, "* Exclusion code may be the same as the enrollment date"
                    EndDialog

    				dialog Dialog1
    				If buttonpressed = 0 then script_end_procedure_with_error_report("Error message was not resolved. Please review enrollment information before trying the script again.")
    				EMWriteScreen "...", 13, 5
    				EMReadScreen false_end, 8, 14, 14
    				If false_end = "99/99/99" Then
    					EMReadScreen double_check, 2, 14, 5
    					If double_check = "??" Then EMWriteScreen "...", 14, 5
    				End If
    			END IF
            ELSE
                'REFM screen
                EMWriteScreen "refm", 1, 8
                transmit
            End If

	        'blanking out varibles if the other option is selected
	        If change_reason = "Select one..." then change_reason = ""
	        If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			EMReadScreen REFM_check, 4, 1, 52
			If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until REFM_check = "REFM"

        If enrollment_source = "Paper Enrollment Form" Then
    		'form rec'd
    		EMsetcursor 10, 16
    		EMSendkey "y"
    		'other insurance y/n
    		EMsetcursor 11, 18
    		EMsendkey insurance_yn
    		'preg y/n
    		EMsetcursor 12, 19
    		EMsendkey pregnant_yn
    		'interpreter y/n
    		EMsetcursor 13, 29
    		EMsendkey interpreter_yn
    		'interpreter type
    		if MMIS_clients_array(interp_code, member) <> "" then
    			EMsetcursor 13, 52
    			EMsendKey MMIS_clients_array(interp_code, member)
    		end if
    		'medical clinic code
    		EMsetcursor 19, 4
    		EMsendkey MMIS_clients_array(med_code, member)
    		'dental clinic code if applicable
    		EMsetcursor 19, 24
    		EMsendkey MMIS_clients_array(dent_code, member)
    		'foster care y/n
    		EMsetcursor 21, 15
    		EMsendkey foster_care_yn
		    ' msgbox "REFM updated"
        Else
            'form rec'd
            EMsetcursor 10, 16
            EMSendkey "n"
        End If
		PF9

		'Save and case note
		pf3

        EMReadScreen look_for_RKEY, 4, 1, 52
        ' MsgBox "Look for RKEY - " & look_for_RKEY
        If look_for_RKEY <> "RKEY" Then
            'We are going to try again to save the information
            PF3
            EMReadScreen REFM_error_check, 79, 24, 2 'checks for an inhibiting edit
            REFM_error_check = trim(REFM_error_check)
            ' MsgBox "REFM error - " & REFM_error_check
            If REFM_error_check <> "ACTION COMPLETED" Then
                process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & REFM_error_check & vbNewLine & vbNewLine
                PF6
            End If
        End If

		EMWriteScreen "c", 2, 19
		transmit
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52
		' MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
			' pf4
			' pf11
			' EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
			' pf3
		Else
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE
		End If
		' MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""
	Next
Else
	For member = 0 to Ubound(MMIS_clients_array, 2)
		'MMIS Codes
		'change reasons
		If MMIS_clients_array(change_rsn, member) = "First year change option" 	then change_reason = "FY"
		If MMIS_clients_array(change_rsn, member) = "Health plan contract end" 	then change_reason = "HP"
		If MMIS_clients_array(change_rsn, member) = "Initial enrollment"       	then change_reason = "IN"
		If MMIS_clients_array(change_rsn, member) = "Move"                     	then change_reason = "MV"
		If MMIS_clients_array(change_rsn, member) = "Ninety Day change option" 	then change_reason = "NT"
		If MMIS_clients_array(change_rsn, member) = "Open enrollment"    	  	then change_reason = "OE"
		If MMIS_clients_array(change_rsn, member) = "PMI merge" 				then change_reason = "PM"
		If MMIS_clients_array(change_rsn, member) = "Reenrollment" 			  	then change_reason = "RE"
		If MMIS_clients_array(change_rsn, member) = "Select one..." 			then change_reason = ""

		'Disenrollment reasons
		If MMIS_clients_array(disenrol_rsn, member) = "Eligibility ended"        then disenrollment_reason = "EE"
		If MMIS_clients_array(disenrol_rsn, member) = "Exclusion"                then disenrollment_reason = "EX"
		If MMIS_clients_array(disenrol_rsn, member) = "First year change option" then disenrollment_reason = "FY"
		If MMIS_clients_array(disenrol_rsn, member) = "Health plan contract end" then disenrollment_reason = "HP"
		If MMIS_clients_array(disenrol_rsn, member) = "Jail - Incarceration"     then disenrollment_reason = "JL"
		If MMIS_clients_array(disenrol_rsn, member) = "Move"                     then disenrollment_reason = "MV"
		If MMIS_clients_array(disenrol_rsn, member) = "Loss of disability"       then disenrollment_reason = "ND"
		If MMIS_clients_array(disenrol_rsn, member) = "Ninety Day change option" then disenrollment_reason = "NT"
		If MMIS_clients_array(disenrol_rsn, member) = "Open Enrollment"          then disenrollment_reason = "OE"
		If MMIS_clients_array(disenrol_rsn, member) = "PMI merge"                then disenrollment_reason = "PM"
		If MMIS_clients_array(disenrol_rsn, member) = "Voluntary"                then disenrollment_reason = "VL"
		If MMIS_clients_array(disenrol_rsn, member) = "Select one..."            then disenrollment_reason = ""

		'REFM Codes
		If MMIS_clients_array(preg_yes, member) = checked Then
			pregnant_yn = "Y"
		Else
			pregnant_yn = "N"
		End If

		If MMIS_clients_array(interp_code, member) = "" Then
			interpreter_yn = "N"
		Else
			interpreter_yn = "Y"
		End If

		EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
		IF MMIS_panel_check <> "RKEY" THEN
			DO
				PF6
				EMReadScreen session_terminated_check, 18, 1, 7
			LOOP until session_terminated_check = "SESSION TERMINATED"
			'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
			EMWriteScreen "mw00", 1, 2
			transmit
			transmit
			EMWriteScreen "x", 8, 3
			transmit
		END IF
		' msgbox "At RKEY"
		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
		EMWriteScreen "c", 2, 19
		EMWriteScreen "        ", 4, 19
		EMWriteScreen MMIS_case_number, 9, 19
		transmit
		transmit
		transmit
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
		Do
			row = 1
			col = 1
			EMSearch MMIS_clients_array(client_pmi, member), row, col
			If row = 0 Then
				PF8
				EMReadScreen end_of_clts, 6, 24, 2
				If end_of_clts = "CANNOT" Then
					process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be found on this case. The enrollment for " &  MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
					Exit Do
				End If
			End If
		Loop until row <> 0
		EMWriteScreen "X", row, 2
		' msgbox "person selected"
		transmit
		' msgbox "at RSUM"
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
		' msgbox process_manually_message

		DO
			'check RPOL to see if there is other insurance available, if so worker processes manually
			EMWriteScreen "rpol", 1, 8
			transmit
			'making sure script got to right panel
			EMReadScreen RPOL_check, 4, 1, 52
			If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
			EMreadscreen policy_number, 1, 7, 8
			If policy_number <> " " then

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
                  CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
                  ButtonGroup ButtonPressed
                    OkButton 105, 125, 50, 15
                  Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
                  GroupBox 10, 45, 145, 70, "REVIEW RPOL"
                  Text 50, 60, 65, 10, "*** Check RPOL ***"
                  Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
                EndDialog

                dialog Dialog1

                If rpol_ended_checkbox = unchecked Then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
            End If

			'nav to RPPH
			EMWriteScreen "rpph", 1, 8
			transmit

			'making sure script got to right panel
			EMReadScreen RPPH_check, 4, 1, 52
			If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

			Enrollment_date = Enrollment_month & "/01/" & enrollment_year
			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year
			' msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" Then
					row = 1
					col = 1
					EMSearch "99/99/99", row, col
                    If col <> 0 Then
                        EMReadScreen beg_of_excl, 8, row, 9
                        IF beg_of_excl = enrollment_date Then
                            EMWriteScreen "...", row, 2
                        Else
                            EMWriteScreen xcl_end_date, row, col
                        End if
                    End If
				Else
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If
			End If
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
			' msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"

			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)

            If process_manually_message = "" Then
    			'enter disenrollment reason
                If disenrollment_reason <> "" Then
                    EMReadScreen beg_of_curr_span, 8, 13, 5
                    If beg_of_curr_span = enrollment_date Then
                        EMWriteScreen "...", 13, 5
                    Else
                        EMWriteScreen xcl_end_date, 13, 14
                        EMWriteScreen disenrollment_reason, 13, 75
                    End If
                End If

    			'resets to bottom of the span list.
    			pf11

    			'enter enrollment date
    			EMWriteScreen enrollment_date, 13, 5
    			'enter managed care plan code
    			EMWriteScreen health_plan_code, 13, 23
    			'enter contract code
    			EMWriteScreen contract_code_part_one, 13, 34
    			EMWriteScreen contract_code_part_two, 13, 37
    			'enter change reason
    			EMWriteScreen change_reason, 13, 71

    			EMWaitReady 0, 0

    			EMReadScreen false_end, 8, 14, 14
    			If false_end = "99/99/99" Then
    				EMReadScreen double_check, 2, 14, 5
    				If double_check = "  " Then EMWriteScreen "...", 14, 5
    			End If
    			'msgbox "RPPH updated"

    			'REFM screen
    			EMWriteScreen "refm", 1, 8
    			transmit
    			EMReadScreen RPPH_error_check, 10, 24, 2
    			If trim(RPPH_error_check) = "EXCLSN END" then
    				Do
                        Dialog1 = ""
                        BeginDialog Dialog1, 0, 0, 191, 45, "Exclusion Code Error"
                          ButtonGroup ButtonPressed
                            OkButton 85, 25, 50, 15
                            CancelButton 135, 25, 50, 15
                          Text 15, 10, 155, 10, "Update the exclusion code field, then press OK."
                        EndDialog

    					Dialog Dialog1
    					cancel_confirmation
    					transmit
    					EMReadScreen RPPH_error_check, 10, 24, 2
    				Loop until trim(RPPH_error_check) <> "EXCLSN END"
    				' Msgbox "Updated the exclusion code field, then press OK."
    				' transmit
    			ELSEIF trim(RPPH_error_check) <> "" then
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 236, 110, "RPPH error detected"
                      DropListBox 70, 50, 160, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", change_reason
                      DropListBox 70, 65, 160, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", disenrollment_reason
                      ButtonGroup ButtonPressed
                        OkButton 125, 85, 50, 15
                        CancelButton 180, 85, 50, 15
                      Text 10, 55, 55, 10, "Change reason:"
                      Text 10, 70, 60, 10, "Disenroll reason:"
                      ButtonGroup ButtonPressed
                        OkButton 155, 330, 50, 15
                      Text 15, 20, 210, 10, "* Initial enrollment is selected, but has been enrolled previously"
                      GroupBox 5, 5, 225, 40, "An error occurred on in RPPH. Typical errors include:"
                      Text 15, 30, 210, 10, "* Exclusion code may be the same as the enrollment date"
                    EndDialog

    				dialog Dialog1
    				If buttonpressed = 0 then script_end_procedure_with_error_report("Error message was not resolved. Please review enrollment information before trying the script again.")
    				EMWriteScreen "...", 13, 5
    			END IF
            ELSE
                'REFM screen
                EMWriteScreen "refm", 1, 8
                transmit
            End If

			'blanking out varibles if the other option is selected
			If change_reason = "Select one..." then change_reason = ""
			If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			EMReadScreen REFM_check, 4, 1, 52
			If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until REFM_check = "REFM"

        If enrollment_source = "Paper Enrollment Form" Then
    		'form rec'd
    		EMsetcursor 10, 16
    		EMSendkey "y"
    		'other insurance y/n
    		EMsetcursor 11, 18
    		EMsendkey insurance_yn
    		'preg y/n
    		EMsetcursor 12, 19
    		EMsendkey pregnant_yn
    		'interpreter y/n
    		EMsetcursor 13, 29
    		EMsendkey interpreter_yn
    		'interpreter type
    		if MMIS_clients_array(interp_code, member) <> "" then
    			EMsetcursor 13, 52
    			EMsendKey MMIS_clients_array(interp_code, member)
    		end if
    		'medical clinic code
    		EMsetcursor 19, 4
    		EMsendkey MMIS_clients_array(med_code, member)
    		'dental clinic code if applicable
    		EMsetcursor 19, 24
    		EMsendkey MMIS_clients_array(dent_code, member)
    		'foster care y/n
    		EMsetcursor 21, 15
    		EMsendkey foster_care_yn
		    ' msgbox "REFM updated"
        Else
            'form rec'd
            EMsetcursor 10, 16
            EMSendkey "n"
        End If
		PF9

		'error handling to ensure that enrollment date and exclusion dates don't conflict
		EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
		IF REFM_error_check <> "                   " then
            IF REFM_error_check <> "INVALID KEY ENTERED" AND REFM_error_check <> "INVALID KEY PRESSED" then
                EMReadScreen full_error_msg, 79, 24, 2
                full_error_msg = trim(full_error_msg)
			    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
		    END IF
        END IF

		' msgbox "all updated - see casenote code"
		'Save and case note
		pf3
		EMWriteScreen "i", 2, 19
        EMWriteScreen "        ", 9, 19
        EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19
		transmit
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52
		' MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
		''			pf4
		''			pf11
		''			EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
		''			pf3
		Else
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE
		End If
		' MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""
	Next
End If

name_of_script = "ACTIONS - MHC ENROLLMENT - " & left(enrollment_source, 5) & ".vbs"
If caller_rela = "" Then caller_rela = "Client"

EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF

'Case Noting - goes into RSUM for the first client to do the case note
EMWriteScreen "c", 2, 19
EMWriteScreen "        ", 9, 19
EMWriteScreen MMIS_clients_array(client_pmi, 0), 4, 19
transmit
pf4
pf11		'Starts a new case note'

' CALL write_variable_in_MMIS_NOTE ("***Hennepin MHC note*** Household enrollment updated for " & Enrollment_date & " per enrollment form")
If open_enrollment_case = TRUE Then
	CALL write_variable_in_MMIS_NOTE ("AHPS request processed for 2020 selection")
Else
	If enrollment_source = "Morning Letters" Then
	    CALL write_variable_in_MMIS_NOTE ("Re-enrollment processed effective: " & enrollment_date)
	    CALL write_variable_in_MMIS_NOTE ("Following clients had PMAP under duplicate PMI(s) in the last 12 months:")
	Else
	    CALL write_variable_in_MMIS_NOTE ("Enrollment effective: " & enrollment_date & " requested by " & caller_rela & " via " & enrollment_source)
	End If
End If
If enrollment_source = "Phone" Then CALL write_variable_in_MMIS_NOTE("Call completed " & now & " with " & caller_name)
If used_interpreter_checkbox = checked then CALL write_variable_in_MMIS_NOTE("Interpreter used for phone call.")
If trim(form_received_date) <> "" Then CALL write_variable_in_MMIS_NOTE("Enrollment requested via Form received on " & form_received_date & ".")
For member = 0 to Ubound(MMIS_clients_array, 2)
	If MMIS_clients_array(enrol_sucs, member) = TRUE Then
        If enrollment_source = "Morning Letters" Then
            CALL write_variable_in_MMIS_NOTE ("- Re-enrolled " & MMIS_clients_array(client_name, member) & " in " & MMIS_clients_array(new_plan, member))
        Else
		    CALL write_variable_in_MMIS_NOTE ("- " & MMIS_clients_array(client_name, member) & " enrolled into " & MMIS_clients_array(new_plan, member))
        End If
	End If
Next
CALL write_bullet_and_variable_in_MMIS_NOTE ("Notes", other_notes)

CALL write_variable_in_MMIS_NOTE ("Processed by " & worker_signature)
CALL write_variable_in_MMIS_NOTE ("*************************************************************************")
pf3
pf3
IF REFM_error_check = "WARNING: MA12,01/16" Then
	PF3
END IF


failed_enrollment_message = "The script is complete. Enrollment has been updated and case noted." & vbNewLine & "There may be some clients enrollments that could not be processed by the script for some reason, they will be listed below:" & vbNewLine & "*****" & vbNewLine & vbNewLine & failed_enrollment_message

script_end_procedure_with_error_report (failed_enrollment_message)
