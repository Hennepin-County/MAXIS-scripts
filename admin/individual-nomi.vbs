'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - NOMI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 276                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
call changelog_update("05/01/2022", "Updated the NOMI to have information for residents about in person support.", "Casey Love, Hennepin County")
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
call changelog_update("08/01/2021", "Changed the notices with updated verbiage on how to submit documents to Hennepin.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2020", "Update to the notice wording. Information and direction for in-person interview option removed. County offices are not currently open due to the COVID-19 Peacetime Emergency.", "Casey Love, Hennepin County")
call changelog_update("01/30/2019", "Added statistics tracking.", "Casey Love, Hennepin County")
call changelog_update("07/20/2018", "Updated verbiage of notice.", "Casey Love, Hennepin County")
call changelog_update("06/22/2018", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'logic to autofill the 'last_day_for_recert' field
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_for_recert = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 131, 45, "Case Number"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 55, 25, 35, 15
    CancelButton 95, 25, 30, 15
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

Do
    err_msg = ""

    Dialog Dialog1
    If buttonpressed = Cancel Then script_end_procedure_with_error_report("")

    If len(MAXIS_case_number) >8 Then err_msg = err_msg & vbNewLine & "* Case numbers should not be more than 8 numbers long."
    If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* Check the case number, it appears to be invalid."

    If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

Call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)

Call navigate_to_MAXIS_screen("CASE", "NOTE")
note_row = 5            'resetting the variables on the loop

day_before_app = DateAdd("d", -1, application_date)

Do
    EMReadScreen note_date, 8, note_row, 6      'reading the note date
    EMReadScreen note_title, 55, note_row, 25   'reading the note header
    note_title = trim(note_title)

    IF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
        EMReadScreen appt_date, 10, note_row, 63
        appt_date = replace(appt_date, "~", "")
        appt_date = trim(appt_date)
        'MsgBox ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)

        Exit Do
    END IF

    IF note_date = "        " then Exit Do
    note_row = note_row + 1
    IF note_row = 19 THEN
        PF8
        note_row = 5
    END IF
    EMReadScreen next_note_date, 8, note_row, 6
    IF next_note_date = "        " then Exit Do
Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

application_date = application_date & ""
appt_date = appt_date & ""
'creates interview date for 7 calendar days from the CAF date
' interview_date = dateadd("d", 7, application_date)
' If interview_date <= date then interview_date = dateadd("d", 7, date)
' interview_date = interview_date & ""		'turns interview date into string for variable

'DIALOGS----------------------------------------------------------------------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 75, "NOMI"
  EditBox 70, 5, 50, 15, application_date
  EditBox 70, 25, 50, 15, appt_date
  ButtonGroup ButtonPressed
    OkButton 15, 50, 50, 15
    CancelButton 70, 50, 50, 15
  Text 10, 30, 60, 10, "Missed Interview date:"
  Text 5, 10, 55, 10, "Application date:"
EndDialog

Do
	Do
		err_msg = ""
		dialog Dialog1
		cancel_confirmation
        If IsDate(application_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date of application."
        If IsDate(appt_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid missed interview date."

		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

last_contact_day = dateadd("d", 30, application_date)
If DateDiff("d", appt_date, last_contact_day) < 1 then last_contact_day = appt_date
CALL start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)

Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & appt_date & ".")
Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
Call write_variable_in_SPEC_MEMO(" ")
CALL write_variable_in_SPEC_MEMO("All interviews are completed via phone. If you do not have a phone, go to one of our Digital Access Spaces at any Hennepin County Library or Service Center. No processing, no interviews are completed at these sites. Some Options:")
CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
CALL write_variable_in_SPEC_MEMO(" More detail can be found at hennepin.us/economic-supports")
CALL write_variable_in_SPEC_MEMO("")
CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at service centers(listed above)")
PF4
PF3
'Writes the case note for the NOMI
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent ~ ")
Call write_variable_in_CASE_NOTE("A notice was previously sent to client with detail about completing an interview. ")
Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
PF3

statistics_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Applications Statistics\2019 Statistics Tracking.xlsx"
call excel_open(statistics_excel_file_path, False,  False, ObjStatsExcel, objStatsWorkbook)

'Now we need to open the right worksheet
'Select Case MonthName(Month(#2/15/19#))
Select Case MonthName(Month(date))

    Case "January"
        sheet_selection = "January 2019"
    Case "February"
        sheet_selection = "February 2019"
    Case "March"
        sheet_selection = "March 2019"
    Case "April"
        sheet_selection = "April 2019"
    Case "May"
        sheet_selection = "May 2019"
    Case "June"
        sheet_selection = "June 2019"
    Case "July"
        sheet_selection = "July 2019"
    Case "August"
        sheet_selection = "August 2019"
    Case "September"
        sheet_selection = "September 2019"
    Case "October"
        sheet_selection = "October 2019"
    Case "November"
        sheet_selection = "November 2019"
    Case "December"
        sheet_selection = "December 2019"

End Select
'Activates worksheet based on user selection
ObjStatsExcel.worksheets(sheet_selection).Activate

stats_excel_nomi_row = 3
Do
    this_entry = ObjStatsExcel.Cells(stats_excel_nomi_row, 1).Value
    this_entry = trim(this_entry)
    If this_entry <> "" Then stats_excel_nomi_row = stats_excel_nomi_row + 1
Loop until this_entry = ""

'Here we add the NOMI to the statistics
ObjStatsExcel.Cells(stats_excel_nomi_row, 1).Value = MAXIS_case_number      'Adding the case number to the statistics sheet
ObjStatsExcel.Cells(stats_excel_nomi_row, 2).Value = application_date       'Adding the date of application to the statistics sheet
ObjStatsExcel.Cells(stats_excel_nomi_row, 3).Value = date                   'Adding today's date of the NOMI date for the stats sheet
ObjStatsExcel.Cells(stats_excel_nomi_row, 4).Value = 1                      'Need to count - this is always 1

objStatsWorkbook.Save
ObjStatsExcel.Quit

script_end_procedure_with_error_report("Success! The NOMI has been sent.")
