'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - BULK - ADJUST REVW MONTHS.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("11/16/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' const case_number_const 	= 0		'Case Number
' const mfip_status_const 	= 1		'MFIP Status
' const ga_status_const 		= 2		'GA Status
' const msa_status_const 		= 3		'MSA Status
' const grh_status_const		= 4		'GRH Status
' const snap_status_const		= 5		'SNAP Status
' const curr_cash_ER_const	= 6		'Current Cash ER
' const curr_cash_SR_const	= 7		'Current Cash SR
' const curr_snap_ER_const 	= 8		'Current SNAP ER
' const curr_snap_SR_const 	= 9		'Current SNAP SR
' const appear_24_mo_const 	= 10	'Does this appear 24 months?
' const new_cash_ER_const 	= 11	'New Cash ER
' const new_cash_SR_const 	= 12	'New Cash SR
' const new_snap_ER_const		= 13	'New SNAP ER
' const new_snao_SR_const 	= 14	'New SNAP SR
' const update_success_const	= 15	'Update Successful
' const cnote_entered_const	= 16	'Case Note Entered
' const excel_row_const 		= 17	'Excel Row
' const case_active_const		= 18
' const case_county_const 	= 19
' const notes_const 			= 20	'Notes
'
' Dim REVW_CASES_ARRAY()
' ReDim REVW_CASES_ARRAY(notes_const, 0)

EMConnect ""

If DatePart("m", date) = 11 Then er_month_to_adjust = "October 2020"
If DatePart("m", date) = 12 Then er_month_to_adjust = "November 2020"
If DatePart("m", date) = 1 Then er_month_to_adjust = "December 2020"
If DatePart("m", date) = 2 Then er_month_to_adjust = "January 2021"
If DatePart("m", date) = 3 Then er_month_to_adjust = "February 2021"

excel_row_to_start = "2"

'DIALOG to collect the correct ER month and worker signature - ADD a restart option
BeginDialog Dialog1, 0, 0, 311, 75, "Dialog"
  DropListBox 125, 10, 130, 45, "Select One..."+chr(9)+"October 2020"+chr(9)+"November 2020"+chr(9)+"December 2020"+chr(9)+"January 2021"+chr(9)+"February 2021", er_month_to_adjust
  EditBox 125, 30, 180, 15, worker_signature
  EditBox 125, 50, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 200, 50, 50, 15
    CancelButton 255, 50, 50, 15
  Text 10, 15, 110, 10, "Select the Recertification Month:"
  Text 55, 35, 60, 10, "Worker Signature:"
  Text 55, 55, 65, 10, "Excel Row to Start:"
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1

		worker_signature = trim(worker_signature)
		excel_row_to_start = trim(excel_row_to_start)

		If er_month_to_adjust = "Select One..." Then err_msg = err_mag & vbNewLine & "* Select the month to work on."
		If worker_signature = "" Then err_msg = err_mag & vbNewLine & "* Sign your case notes."
		If IsNumeric(excel_row_to_start) = FALSE  Then err_msg = err_mag & vbNewLine & "* The excel row "

		If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'Open the correct Excel File from
excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Adjusted ER Cases\" & er_month_to_adjust & " Recertification Report.xlsx"
call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

'Activate the working sheet in the file
objExcel.worksheets("Work List").Activate

'Create an array and add all of the cases in the list to the array
excel_row = excel_row_to_start * 1
line_count = 0

If er_month_to_adjust = "October 2020" Then
	new_er_month = "04"
	new_sr_month = "10"
	new_er_year = "21"
	new_sr_year = "21"

	cash_revw_new_SR_date = #10/1/2021#
	cash_revw_new_ER_date = #4/1/2021#

	snap_revw_new_SR_date = #10/1/2021#
	snap_revw_new_ER_date = #4/1/2021#
End If

If er_month_to_adjust = "November 2020" Then
	new_er_month = "05"
	new_sr_month = "11"
	new_er_year = "21"
	new_sr_year = "21"

	cash_revw_new_SR_date = #11/1/2021#
	cash_revw_new_ER_date = #5/1/2021#

	snap_revw_new_SR_date = #11/1/2021#
	snap_revw_new_ER_date = #5/1/2021#
End If
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Do
	read_case_number = ObjExcel.Cells(excel_row, 1).Value
	read_case_number = trim(read_case_number)
	If read_case_number <> "" Then
		MAXIS_case_number = read_case_number
		Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
		If is_this_priv = FALSE Then
			CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
			EMReadScreen current_pw_county, 2, 21, 16

			ObjExcel.Cells(excel_row, 2).Value = case_active
			ObjExcel.Cells(excel_row, 3).Value = current_pw_county
			ObjExcel.Cells(excel_row, 4).Value = mfip_case
			ObjExcel.Cells(excel_row, 5).Value = ga_case
			ObjExcel.Cells(excel_row, 6).Value = msa_case
			ObjExcel.Cells(excel_row, 7).Value = grh_case
			ObjExcel.Cells(excel_row, 8).Value = snap_case

			attempt_to_update = TRUE
			cash_sr_needs_update = FALSE
			cash_er_needs_update = FALSE
			snap_sr_needs_update = FALSE
			snap_er_needs_update = FALSE
			If case_active = FALSE Then attempt_to_update = FALSE
			If current_pw_county <> "27" Then attempt_to_update = FALSE

			curr_cash_sr_month = ""
			curr_cash_sr_year = ""
			curr_cash_er_month = ""
			curr_cash_er_year = ""
			curr_snap_sr_month = ""
			curr_snap_sr_year = ""
			curr_snap_er_month = ""
			curr_snap_er_year = ""
			cash_updated_SR_date = ""
			cash_updated_ER_date = ""
			snap_updated_SR_date = ""
			snap_updated_ER_date = ""

			If attempt_to_update = TRUE Then
				Call navigate_to_MAXIS_screen("STAT", "REVW")

				If mfip_case = TRUE OR ga_case = TRUE OR msa_case = TRUE OR grh_case = TRUE Then
					EMWriteScreen "X", 5, 35
					Transmit
					EMReadScreen curr_cash_sr_month, 2, 9, 26
					EMReadScreen curr_cash_sr_year, 2, 9, 32
					EMReadScreen curr_cash_er_month, 2, 9, 64
					EMReadScreen curr_cash_er_year, 2, 9, 70
					PF3

					If curr_cash_sr_month = "__" Then curr_cash_sr_month = ""
					If curr_cash_sr_year = "__" Then curr_cash_sr_year = ""
					If curr_cash_er_month = "__" Then curr_cash_er_month = ""
					If curr_cash_er_year = "__" Then curr_cash_er_year = ""

					If curr_cash_sr_month <> new_sr_month or curr_cash_sr_year <> new_sr_year Then cash_sr_needs_update = TRUE
					If curr_cash_er_month <> new_er_month or curr_cash_er_year <> new_er_year Then cash_er_needs_update = TRUE
					If curr_cash_sr_month = "" Then cash_sr_needs_update = FALSE
					If curr_cash_er_month = "" Then cash_er_needs_update = FALSE

				End If

				If snap_case = TRUE Then
					EMWriteScreen "X", 5, 58
					Transmit
					EMReadScreen curr_snap_sr_month, 2, 9, 26
					EMReadScreen curr_snap_sr_year, 2, 9, 32
					EMReadScreen curr_snap_er_month, 2, 9, 64
					EMReadScreen curr_snap_er_year, 2, 9, 70
					PF3

					If curr_snap_sr_month = "__" Then curr_snap_sr_month = ""
					If curr_snap_sr_year = "__" Then curr_snap_sr_year = ""
					If curr_snap_er_month = "__" Then curr_snap_er_month = ""
					If curr_snap_er_year = "__" Then curr_snap_er_year = ""

					If curr_snap_sr_month <> new_sr_month or curr_snap_sr_year <> new_sr_year Then snap_sr_needs_update = TRUE
					If curr_snap_er_month <> new_er_month or curr_snap_er_year <> new_er_year Then snap_er_needs_update = TRUE
					If curr_snap_sr_month = "" Then snap_sr_needs_update = FALSE
					If curr_snap_er_month = "" Then snap_er_needs_update = FALSE

				End If

				for col = 9 to 22
					ObjExcel.Cells(excel_row, col).NumberFormat = "@"
				next

				ObjExcel.Cells(excel_row, 9).Value = curr_cash_er_month & "-" & curr_cash_er_year
				ObjExcel.Cells(excel_row, 10).Value = cash_er_needs_update
				ObjExcel.Cells(excel_row, 11).Value = curr_cash_sr_month & "-" & curr_cash_sr_year
				ObjExcel.Cells(excel_row, 12).Value = cash_sr_needs_update

				ObjExcel.Cells(excel_row, 15).Value = curr_snap_er_month & "-" & curr_snap_er_year
				ObjExcel.Cells(excel_row, 16).Value = snap_er_needs_update
				ObjExcel.Cells(excel_row, 17).Value = curr_snap_sr_month & "-" & curr_snap_sr_year
				ObjExcel.Cells(excel_row, 18).Value = snap_sr_needs_update

				attempt_to_update = FALSE
				If cash_sr_needs_update = TRUE then attempt_to_update = TRUE
				If cash_er_needs_update = TRUE then attempt_to_update = TRUE
				If snap_sr_needs_update = TRUE then attempt_to_update = TRUE
				If snap_er_needs_update = TRUE then attempt_to_update = TRUE

				ObjExcel.Cells(excel_row, 22).Value = "Attempt to update - " & attempt_to_update
			End If

			Call Back_to_self

			If attempt_to_update = TRUE Then
				revw_panel_updated = TRUE
				Call navigate_to_MAXIS_screen("STAT", "REVW")
				PF9

				If cash_sr_needs_update = TRUE OR cash_er_needs_update = TRUE Then
					Call create_mainframe_friendly_date(cash_revw_new_ER_date, 9, 37, "YY")

					EMWriteScreen "X", 5, 35
					transmit

					Call create_mainframe_friendly_date(cash_revw_new_SR_date, 9, 26, "YY")
					Call create_mainframe_friendly_date(cash_revw_new_ER_date, 9, 64, "YY")
					transmit
				End If

				If snap_sr_needs_update = TRUE OR snap_er_needs_update = TRUE then
					EMWriteScreen "X", 5, 58
					transmit

					Call create_mainframe_friendly_date(snap_revw_new_SR_date, 9, 26, "YY")
					Call create_mainframe_friendly_date(snap_revw_new_ER_date, 9, 64, "YY")
					transmit
				End If

				attempt_count = 1

		        Do
		            transmit
		            EMReadScreen actually_saved, 7, 24, 2
		            attempt_count = attempt_count + 1
		            If attempt_count = 20 Then
		                PF10
		                revw_panel_updated = FALSE
		                Exit Do
		            End If
		        Loop until actually_saved = "ENTER A"

				If cash_sr_needs_update = TRUE OR cash_er_needs_update = TRUE Then
					'READ'
					Call create_mainframe_friendly_date(cash_revw_new_ER_date, 9, 37, "YY")

					EMWriteScreen "X", 5, 35
					transmit

					EMReadScreen cash_updated_SR_date, 8, 9, 26
					EMReadScreen cash_updated_ER_date, 8, 9, 64

					cash_updated_SR_date = replace(replace(cash_updated_SR_date, "_", ""), " 01 ", "-")
					cash_updated_ER_date = replace(replace(cash_updated_ER_date, "_", ""), " 01 ", "-")

					transmit
				End If

				If snap_sr_needs_update = TRUE OR snap_er_needs_update = TRUE then
					EMWriteScreen "X", 5, 58
					transmit

					EMReadScreen snap_updated_SR_date, 8, 9, 26
					EMReadScreen snap_updated_ER_date, 8, 9, 64

					snap_updated_SR_date = replace(replace(snap_updated_SR_date, "_", ""), " 01 ", "-")
					snap_updated_ER_date = replace(replace(snap_updated_ER_date, "_", ""), " 01 ", "-")
					transmit
				End If

				ObjExcel.Cells(excel_row, 13).Value = cash_updated_ER_date
				ObjExcel.Cells(excel_row, 14).Value = cash_updated_SR_date
				ObjExcel.Cells(excel_row, 19).Value = snap_updated_ER_date
				ObjExcel.Cells(excel_row, 20).Value = snap_updated_SR_date





				'
				' Call Navigate_to_MAXIS_screen("STAT", "REVW")
				' revw_panel_updated = TRUE
				' PF9
				' EMReadScreen read_only_access, 16, 24, 11
				' If read_only_access = "READ ONLY ACCESS" Then revw_panel_updated = FALSE
				'
				' If cash_dates_to_update = TRUE Then
				' 	EMWriteScreen cash_revw_status_code, 7, 40
				' 	Call create_mainframe_friendly_date(cash_revw_new_ER_date, 9, 37, "YY")
				'
				' 	EMWriteScreen "X", 5, 35
				' 	transmit
				'
				' 	EMWriteScreen snap_revw_new_status_code, 7, 64
				' 	Call create_mainframe_friendly_date(CAF_datestamp, 7, 26, "YY")
				' 	Call create_mainframe_friendly_date(cash_revw_new_SR_date, 9, 26, "YY")
				' 	Call create_mainframe_friendly_date(cash_revw_new_ER_date, 9, 64, "YY")
				'
				' 	transmit
				' End If
				'
				' If snap_dates_to_update = TRUE Then
				' 	EMWriteScreen snap_revw_new_status_code, 7, 60
				'
				' 	EMWriteScreen "X", 5, 58
				' 	transmit
				'
				' 	EMWriteScreen snap_revw_new_status_code, 7, 64
				' 	Call create_mainframe_friendly_date(CAF_datestamp, 7, 26, "YY")
				' 	Call create_mainframe_friendly_date(snap_revw_new_SR_date, 9, 26, "YY")
				' 	Call create_mainframe_friendly_date(snap_revw_new_ER_date, 9, 64, "YY")
				'
				' 	transmit
				' End If





			End If
		Else
			ObjExcel.Cells(excel_row, 22).Value = "PRIV"
		End If
	End If
	excel_row = excel_row + 1
Loop until read_case_number = ""

script_end_procedure("This is all done.")

'Go through each item in the array/excel line by line
'Read Case information and update
    'Gather case details
        'Is case active
        'Is case in Hennepin County
        'What Programs are active
        'Current REVW Dates
        'Is case 24 Month
        'Identify app dates for active programs
    'Identify if needs to update
        'Inactive cases – NO update
        'Out of County – NO update
        'REVW Dates already updated – NO update
        'If App date is after ‘waived’ month – NO update???
    'Update Cases
    'Case Note
'
