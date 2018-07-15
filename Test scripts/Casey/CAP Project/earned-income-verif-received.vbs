'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PAYSTUBS RECEIVED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 473                	'manual run time in seconds
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
CALL changelog_update("04/23/2018", "Fixed bug in which the lines of the PIC were dupicated in the case note.", "Casey Love, Hennepin County")
CALL changelog_update("12/07/2017", "Removed condition to allow paystubs dated with the current date to be accepted. Updated code to write JOBS verification code in.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/11/2017", "The script has been updated to write to the GRH PIC and to case note that the GRH PIC has been updated.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

' 'TESTING ELEMENT REMOVAL'
' Dim TEST_ARRAY()
' ReDim TEST_ARRAY(0)
'
' Full = 6
' For the_thing = 0 to full
'     ReDim Preserve TEST_ARRAY(the_thing)
'     TEST_ARRAY(the_thing) = the_thing * the_thing
' Next
'
' For each square in TEST_ARRAY
'     MsgBox square
' Next
'
' ReDim Preserve TEST_ARRAY(5)
'
' For each square in TEST_ARRAY
'     MsgBox square
' Next
'
' MsgBox "That's it"

'Find case number and footer month - set a variable with the initially found footer month for the default for every loop
'The footer month may be different for EVERY income source. NEED to add handling to identify if there is a begin date for updating MAXIS (app date that activates the case)
    'A client may apply in april and bring in checks from March but we cannot update MAXIS in March


'DIALOG TO GET CASE NUMBER
'Possibly add worker signature here and take it out of the following dialogs
BeginDialog Dialog1, 0, 0, 191, 175, "Case Number"
  EditBox 90, 5, 70, 15, MAXIS_case_number
  EditBox 5, 35, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 155, 50, 15
    CancelButton 135, 155, 50, 15
  Text 5, 10, 85, 10, "Enter your case number:"
  Text 5, 25, 65, 10, "Worker Signature:"
  GroupBox 5, 55, 180, 95, "INSTRUCTIONS - PLEASE READ!!!"
  Text 10, 70, 170, 25, "This script is to help in correctly budgeting EARNED income on JOBS, BUSI, or RBIC. It will update MAXIS and CASE/NOTE the information provided. "
  Text 10, 105, 170, 40, "If a JOBS panel or BUSI panel needs to be added to MAXIS for a client or income source, the script will ask for any panels that need to be added first. Review the case now to ensure that the correct action will be taken in the correct order."
EndDialog

Do
    err_msg = ""
    dialog Dialog1

    If IsNumeric(MAXIS_case_number) = FALSE or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
    If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature for your case notes."

    If err_msg <> "" Then MsgBox "-- Please resolve the following to continue --" & vbNewLine & err_msg
Loop until err_msg = ""

CASH_case = FALSE
SNAP_case = FALSE
HC_case = FALSE

Call Navigate_to_MAXIS_screen("STAT", "PROG")

EMReadScreen cash_one_status, 4, 6, 74
EMReadScreen cash_two_status, 4, 7, 74
EMReadScreen snap_status, 4, 10, 74
EMReadScreen hc_status, 4, 12, 74

If cash_one_status = "ACTV" OR cash_one_status = "PEND" Then CASH_case = TRUE
If cash_two_status = "ACTV" OR cash_two_status = "PEND" Then CASH_case = TRUE
If snap_status = "ACTV" OR snap_status = "PEND" Then SNAP_case = TRUE
If hc_status = "ACTV" OR hc_status = "PEND" Then HC_case = TRUE

'Add functionality ask if worker needs to add a new JOBS or BUSI panel.'

const panel_type        = 0
const panel_member      = 1
const panel_instance    = 2
const employer          = 3
const income_type       = 4
const income_verif      = 5
const hourly_wage       = 6
const income_start_dt   = 7
const income_end_dt     = 8
const income_list_indct = 9
const pay_freq          = 10
const date_of_calc      = 11
const hrs_per_wk        = 12
const pay_per_hr        = 13
const ave_hrs_per_pay   = 14
const ave_inc_per_pay   = 15
const SNAP_mo_inc       = 16
const reg_non_monthly   = 17
const numb_months       = 18
const self_emp_mthd     = 19
const method_date       = 20
const reptd_hours       = 21
const apply_to_SNAP     = 22
const apply_to_CASH     = 23
const apply_to_HC       = 24
const pay_weekday       = 25
const income_received   = 26
const verif_date        = 27
const verif_explain     = 28
const old_verif         = 29

const spoke_to          = 30
const convo_detail      = 31

Dim EARNED_INCOME_PANELS_ARRAY()
ReDim EARNED_INCOME_PANELS_ARRAY(convo_detail, 0)

the_panel = 0
all_ei_panels_found = FALSE

call HH_member_custom_dialog(HH_member_array)

Call navigate_to_MAXIS_screen("STAT", "JOBS")
For each member in HH_member_array
    EMWriteScreen member, 20, 76
    'EMWriteScreen "01", 20, 79
    Transmit

    EMReadScreen number_of_jobs_panels, 1, 2, 78

    If number_of_jobs_panels <> "0" Then
        number_of_jobs_panels = number_of_jobs_panels * 1

        For panel = 1 to number_of_jobs_panels
            EMWriteScreen "0" & panel, 20, 79
            transmit

            ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)

            EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "JOBS"
            EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member
            EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel
            EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
            If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked
            If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
            If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked

            EMReadScreen type_of_job, 1, 5, 34
            EMReadScreen job_verif, 25, 6, 34
            EMReadScreen listed_hrly_wage, 6, 6, 75
            EMReadScreen employer_name, 30, 7, 42
            EMReadScreen start_date, 8, 9, 35
            EMReadScreen end_date, 8, 9, 49
            EMReadScreen frequency, 1, 18, 35
            EMReadScreen current_verif, 27, 6, 34

            If type_of_job = "J" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "J - WIOA"
            If type_of_job = "W" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "W - Wages"
            If type_of_job = "E" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "E - EITC"
            If type_of_job = "G" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "G - Experience Works"
            If type_of_job = "F" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "F - Federal Work Study"
            If type_of_job = "S" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "S - State Work Study"
            If type_of_job = "O" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "O - Other"
            If type_of_job = "C" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "C - Contract Income"
            If type_of_job = "T" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "T - Training Program"
            If type_of_job = "P" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "P - Service Program"
            If type_of_job = "R" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "R - Rehab Program"

            EARNED_INCOME_PANELS_ARRAY(income_verif, the_panel) = trim(job_verif)
            EARNED_INCOME_PANELS_ARRAY(employer, the_panel) = replace(employer_name, "_", "")
            EARNED_INCOME_PANELS_ARRAY(hourly_wage, the_panel) = trim(listed_hrly_wage)
            EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
            EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
            EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = trim(current_verif)
            If frequency = "1" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "1 - One Time Per Month"
            If frequency = "2" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "2 - Two Times Per Month"
            If frequency = "3" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "3 - Every Other Week"
            If frequency = "4" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "4 - Every Week"
            If frequency = "5" Then EARNED_INCOME_PANELS_ARRAY(pay_freq, the_panel) = "5 - Other"

            EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"
            'EARNED_INCOME_PANELS_ARRAY(, the_panel) =

            the_panel = the_panel + 1
        Next
    End If
Next

Call navigate_to_MAXIS_screen("STAT", "BUSI")
For each member in HH_member_array
    EMWriteScreen member, 20, 76
    'EMWriteScreen "01", 20, 79
    Transmit

    EMReadScreen number_of_busi_panels, 1, 2, 78

    If number_of_busi_panels <> "0" Then
        number_of_busi_panels = number_of_busi_panels * 1

        For panel = 1 to number_of_busi_panels
            EMWriteScreen "0" & panel, 20, 79
            transmit

            ReDim Preserve EARNED_INCOME_PANELS_ARRAY(convo_detail, the_panel)

            EARNED_INCOME_PANELS_ARRAY(panel_type, the_panel) = "BUSI"
            EARNED_INCOME_PANELS_ARRAY(panel_member, the_panel) = member
            EARNED_INCOME_PANELS_ARRAY(panel_instance, the_panel) = "0" & panel
            EARNED_INCOME_PANELS_ARRAY(income_received, the_panel) = FALSE
            If CASH_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, the_panel) = checked
            If SNAP_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, the_panel) = checked
            If HC_case = TRUE Then EARNED_INCOME_PANELS_ARRAY(apply_to_HC, the_panel) = checked

            EMReadScreen type_of_busi, 2, 5, 37
            EMReadScreen start_date, 8, 5, 55
            EMReadScreen end_date, 8, 5, 72
            EMReadScreen listed_method, 2, 16, 53
            EMReadScreen lst_mthd_date, 8, 16, 63

            If type_of_busi = "01" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "01 - Farming"
            If type_of_busi = "02" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "02 - Real Estate"
            If type_of_busi = "03" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "03 - Home Product Sales"
            If type_of_busi = "04" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "04 - Other Sales"
            If type_of_busi = "05" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "05 - Personal Services"
            If type_of_busi = "06" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "06 - Paper Route"
            If type_of_busi = "07" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "07 - In Home Daycare"
            If type_of_busi = "08" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "08 - Rental Income"
            If type_of_busi = "09" Then EARNED_INCOME_PANELS_ARRAY(income_type, the_panel) = "09 - Other"
            EARNED_INCOME_PANELS_ARRAY(income_start_dt, the_panel) = replace(start_date, " ", "/")
            EARNED_INCOME_PANELS_ARRAY(income_end_dt, the_panel) = replace(end_date, " ", "/")
            If listed_method = "01" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "01 - 50% Gross Inc"
            If listed_method = "02" Then EARNED_INCOME_PANELS_ARRAY(self_emp_mthd, the_panel) = "02 - Tax Forms"
            EARNED_INCOME_PANELS_ARRAY(method_date, the_panel) = replace(lst_mthd_date, " ", "/")

            EmWriteScreen "X", 6, 26
            transmit

            For busi_row = 9 to 19
                EMReadScreen busi_verif, 1, busi_row, 73
                If busi_verif <> "_" Then
                    If busi_verif = "1" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "1 - Income Tax Returns"
                    If busi_verif = "2" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "2 - Receipts of Sales/Purchases"
                    If busi_verif = "3" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "3 - Client BUSI Records/Ledger"
                    If busi_verif = "6" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "6 - Other Document"
                    If busi_verif = "N" Then EARNED_INCOME_PANELS_ARRAY(old_verif, the_panel) = "N - NO Verif Provided"
                    Exit For
                End If
            Next
            PF3

            EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"

            the_panel = the_panel + 1
        Next
    End If
Next

Do
    y_pos = 25
    dlg_len = 15 * UBOUND(EARNED_INCOME_PANELS_ARRAY, 2) + 15 * UBOUND(HH_member_array) + 125
    BeginDialog Dialog1, 0, 0, 330, dlg_len, "Case Number"

      Text 5, 10, 105, 10, "Known JOBS and BUSI panels:"

      For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
        earned_income_panel_detail = EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel) & " " & EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel) & " - "
        If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then
            earned_income_panel_detail = earned_income_panel_detail & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel)
        ElseIf EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "BUSI" Then
            earned_income_panel_detail = earned_income_panel_detail & "TYPE: " & EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
        End If
        earned_income_panel_detail = earned_income_panel_detail & " - Income Start: " & EARNED_INCOME_PANELS_ARRAY(income_start_dt, ei_panel) & " - Verif: " & EARNED_INCOME_PANELS_ARRAY(old_verif, ei_panel)
        Text 10, y_pos, 305, 10, earned_income_panel_detail
        'Text 10, 25, 295, 10, "JOBS 01 01 - EMPLOYER - Income Start: mm/dd/yy - Verif: N"
        'Text 10, 40, 295, 10, "BUSI 01 01 - TYPE: 04 - Other Sales - Income Start: mm/dd/yy - Verif: N"
        y_pos = y_pos + 15
      Next
      y_pos = y_pos + 5
      Text 5, y_pos, 295, 10, "These are all the panels that are currently known in MAXIS for these Household Members:"
      y_pos = y_pos + 15
      For each member in HH_member_array
        Text 10, y_pos, 45, 10, member
        y_pos = y_pos + 15
        'Text 10, 75, 45, 10, "MEMBER 01"
      Next
      y_pos = y_pos - 10
      Text 80, y_pos, 160, 10, "Do you need to add a new JOBS or BUSI panel?"
      ButtonGroup ButtonPressed
        PushButton 85, y_pos + 15, 140, 20, "Yes - Add a new Earned Income panel", add_new_panel_button
        PushButton 85, y_pos + 40, 140, 10, "No - The panel(s) to update are in MAXIS", continue_to_update_button
    EndDialog
    MsgBOx "Y Position is " & y_pos & vbNewLine & "Dialog length is " & dlg_len

    dialog Dialog1

    If buttonpressed = add_new_panel_button Then
        MsgBox "Add a new panel!"
        '2 different dialogs for JOBS vs BUSI and add here then add to the EARNED_INCOME_PANELS_ARRAY

        BeginDialog Dialog1, 0, 0, 191, 50, "Panel to Add"
          DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS"+chr(9)+"BUSI", panel_to_add
          ButtonGroup ButtonPressed
            OkButton 135, 10, 50, 15
            CancelButton 135, 30, 50, 15
          Text 15, 10, 85, 20, "Which type of panel would you like to add?"
        EndDialog

        Do
            err_msg = ""

            dialog Dialog1
            cancel_confirmation

            If panel_to_add = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate which type of panel needs to be added."

            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

        Loop until err_msg = ""

        Select Case panel_to_add

        Case "JOBS"
        'Start on DIALOG need to keep working on it'
        BeginDialog Dialog1, 0, 0, 436, 250, "Dialog"
          ButtonGroup ButtonPressed
            OkButton 325, 230, 50, 15
            CancelButton 380, 230, 50, 15
          Text 115, 15, 45, 10, "Income Type:"
          DropListBox 165, 10, 60, 45, "W - Wages (Incl Tips)"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", inc_type_code
          Text 10, 15, 65, 10, "Client Ref Number:"
          EditBox 75, 10, 30, 15, jobs_clt_ref_nbr
          Text 240, 15, 85, 10, "Subsidized Income Type:"
          DropListBox 335, 10, 95, 45, "   "+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-The-Job Training"+chr(9)+"04 - AmeriCorps(VISTA/State/National/NCCC)", subsdzd_inc_type
          Text 20, 35, 40, 10, "Verification:"
          DropListBox 70, 30, 60, 45, "   "+chr(9)+"1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", verif_code
          Text 160, 35, 50, 10, "Hourly Wage:"
          EditBox 220, 30, 50, 15, enter_hrly_wage
          Text 15, 55, 35, 10, "Employer:"
          EditBox 55, 50, 300, 15, enter_employer
        EndDialog

        Case "BUSI"

        End Select
    End If

Loop until buttonpressed = continue_to_update_button

const panel_indct           = 0
const pay_date              = 1
const gross_amount          = 2
const hours                 = 3
const budget_in_SNAP_yes    = 4
const budget_in_SNAP_no     = 5
const reason_to_exclude     = 6
const exclude_amount        = 7
const reason_amt_excluded   = 8

Dim LIST_OF_INCOME_ARRAY()
ReDim LIST_OF_INCOME_ARRAY(reason_amt_excluded, 0)
'CREATE ARRAY OF ALL EI panels'
'Put them in a 'FOR-NEXT' to loop through each panel.
'IF all income will be case noted as 1 note then create an ARRAY of all the case note information.


'NAVIGATE TO JOBS for each HH MEMBER and ask if Income information was received for this job.

'This will become dynamic and there will be an array of all the checks listed.
'STILL need some handling for scheduled income with no actual checks or cases where scheduled income is different from actual checks but we get both.
'NEED TO ADD CHECKBOXES FOR PROGRAMS THIS INCOME APPLIES TO - and precheck all the programs that are active on this case'

' HOLDING FOR A VERSION TO BE PUT IN TO DLG EDIT'
' BeginDialog Dialog1, 0, 0, 606, 160, "Enter ALL Paychecks Received"
'   Text 10, 10, 265, 10, "JOBS 01 01 - EMPLOYER"
'   Text 310, 10, 50, 10, "Income Type:"
'   DropListBox 365, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages (Incl Tips)"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", income_type
'   GroupBox 475, 5, 125, 25, "Apply Income to Programs:"
'   CheckBox 485, 15, 30, 10, "SNAP", apply_to_SNAP
'   CheckBox 530, 15, 30, 10, "CASH", apply_to_CASH
'   CheckBox 570, 15, 20, 10, "HC", apply_to_HC
'   Text 5, 40, 60, 10, "JOBS Verif Code:"
'   DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", JOBS_verif_code
'   Text 175, 40, 155, 10, "additional detail of verification received:"
'   EditBox 310, 35, 290, 15, Edit2
'   Text 5, 60, 90, 10, "Date verification received:"
'   EditBox 100, 55, 50, 15, verif_date
'   Text 5, 80, 80, 10, "Pay Date (MM/DD/YY):"
'   Text 90, 80, 50, 10, "Gross Amount:"
'   Text 145, 80, 25, 10, "Hours:"
'   Text 180, 65, 25, 25, "Use in SNAP budget"
'   Text 235, 80, 85, 10, "If not used, explain why:"
'   Text 355, 70, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
'   Text 355, 80, 30, 10, "Amount:"
'   Text 410, 80, 30, 10, "Reason:"
'   EditBox 5, 90, 65, 15, pay_date
'   EditBox 90, 90, 45, 15, gross_amount
'   EditBox 145, 90, 25, 15, hours_on_check
'   OptionGroup RadioGroup1
'     RadioButton 180, 90, 25, 10, "Yes", budget_yes
'     RadioButton 210, 90, 25, 10, "No", budget_no
'   EditBox 235, 90, 115, 15, reason_not_budgeted
'   EditBox 355, 90, 45, 15, not_budgeted_amount
'   EditBox 410, 90, 185, 15, amount_not_budgeted_reason
'   Text 5, 115, 70, 10, "Anticipated Income"
'   Text 5, 130, 50, 10, "Rate of Pay/Hr"
'   Text 75, 130, 35, 10, "Hours/Wk"
'   Text 130, 130, 50, 10, "Pay Frequency"
'   Text 225, 115, 70, 10, "Regular Non-Monthly"
'   Text 225, 130, 25, 10, "Amount"
'   Text 280, 130, 50, 10, "Nbr of Months"
'   EditBox 5, 140, 50, 15, rate_of_pay
'   EditBox 75, 140, 40, 15, hours_per_week
'   DropListBox 130, 140, 85, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", pay_frequency
'   EditBox 225, 140, 40, 15, non_monthly_amt
'   EditBox 280, 140, 30, 15, number_non_reg_months
'   ButtonGroup ButtonPressed
'     PushButton 440, 140, 15, 15, "+", add_another_check
'     PushButton 460, 140, 15, 15, "-", take_a_check_away
'     OkButton 495, 140, 50, 15
'     CancelButton 550, 140, 50, 15
' EndDialog

' BeginDialog Dialog1, 0, 0, 421, 240, "Confirm JOBS Budget"
'   Text 10, 10, 175, 10, "JOBS 01 01 - EMPLOYER"
'   Text 245, 10, 50, 10, "Pay Frequency"
'   DropListBox 305, 5, 95, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week"+chr(9)+"5 - Other", pay_frequency
'   Text 240, 30, 60, 10, "Income Start Date:"
'   EditBox 305, 25, 70, 15, income_start_date
'   GroupBox 5, 40, 410, 105, "SNAP Budget"
'   Text 10, 50, 100, 10, "Paychecks Inclued in Budget:"
'   Text 20, 65, 90, 10, "01/01/2018 - $400 - 40 hrs"
'   Text 20, 75, 90, 10, "01/15/2018- $400 - 40 hrs"
'   Text 10, 95, 130, 10, "Paychecks not included: 12/24/2018"
'   Text 185, 50, 90, 10, "Average hourly rate of pay:"
'   Text 185, 65, 90, 10, "Average weekly hours:"
'   Text 185, 80, 90, 10, "Average paycheck amount:"
'   Text 185, 95, 90, 10, "Monthly Budgeted Income:"
'   CheckBox 10, 110, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
'   Text 10, 130, 60, 10, "Conversation with:"
'   ComboBox 75, 125, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer", converstion_with
'   Text 140, 130, 25, 10, "clarifies"
'   EditBox 170, 125, 235, 15, conversation_detail
'   GroupBox 5, 150, 410, 60, "CASH Budget"
'   Text 15, 165, 110, 10, "Actual Paychecks to add to JOBS:"
'   Text 25, 180, 90, 10, "01/01/2018 - $400 - 40 hrs"
'   Text 25, 190, 90, 10, "01/15/2018- $400 - 40 hrs"
'   ButtonGroup ButtonPressed
'     OkButton 315, 220, 50, 15
'     CancelButton 370, 220, 50, 15
' EndDialog

pay_item = 0
For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)


    If EARNED_INCOME_PANELS_ARRAY(panel_type, ei_panel) = "JOBS" Then

        Call Navigate_to_MAXIS_screen("STAT", "JOBS")
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_member, ei_panel), 20, 76
        EMWriteScreen EARNED_INCOME_PANELS_ARRAY(panel_instance, ei_panel), 20, 79
        transmit

        employer_check = MsgBox("Do you have income verification for this job? Employer name: " & EARNED_INCOME_PANELS_ARRAY(employer, ei_panel), vbYesNo + vbQuestion, "Select Income Panel")

        If employer_check = vbYes Then
            EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE
            Do
                big_err_msg = ""
                dlg_factor = 0

                LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel

                If LIST_OF_INCOME_ARRAY(panel_indct, 0) <> "" Then
                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then dlg_factor = dlg_factor + 1
                    Next
                End If

                dlg_factor = dlg_factor - 1
                Do
                    sm_err_msg = ""

                    MsgBox "Dialog Factor: " & dlg_factor

                    BeginDialog Dialog1, 0, 0, 606, (dlg_factor * 20) + 160, "Enter ALL Paychecks Received"
                      Text 10, 10, 265, 10, "JOBS 01 01 - EMPLOYER"
                      Text 220, 15, 40, 10, "Start Date:"
                      EditBox 255, 10, 50, 15, EARNED_INCOME_PANELS_ARRAY (income_start_dt, ei_panel)
                      Text 315, 15, 50, 10, "Income Type:"
                      DropListBox 365, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", EARNED_INCOME_PANELS_ARRAY(income_type, ei_panel)
                      GroupBox 475, 5, 125, 25, "Apply Income to Programs:"
                      CheckBox 485, 15, 30, 10, "SNAP", EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel)
                      CheckBox 530, 15, 30, 10, "CASH", EARNED_INCOME_PANELS_ARRAY(apply_to_CASH, ei_panel)
                      CheckBox 570, 15, 20, 10, "HC", EARNED_INCOME_PANELS_ARRAY(apply_to_HC, ei_panel)
                      Text 5, 40, 60, 10, "JOBS Verif Code:"
                      DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", EARNED_INCOME_PANELS_ARRAY(income_verif, ei_panel)
                      Text 175, 40, 155, 10, "additional detail of verification received:"
                      EditBox 310, 35, 290, 15, EARNED_INCOME_PANELS_ARRAY(verif_explain, ei_panel)
                      Text 5, 60, 90, 10, "Date verification received:"
                      EditBox 100, 55, 50, 15, EARNED_INCOME_PANELS_ARRAY(verif_date, ei_panel)
                      Text 5, 80, 80, 10, "Pay Date (MM/DD/YY):"
                      Text 90, 80, 50, 10, "Gross Amount:"
                      Text 145, 80, 25, 10, "Hours:"
                      Text 180, 65, 25, 25, "Use in SNAP budget"
                      Text 235, 80, 85, 10, "If not used, explain why:"
                      Text 355, 70, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
                      Text 355, 80, 30, 10, "Amount:"
                      Text 410, 80, 30, 10, "Reason:"

                      y_pos = 0
                      For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                          If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                              EditBox 5, (y_pos * 20) + 90, 65, 15, LIST_OF_INCOME_ARRAY(pay_date, all_income) 'pay_date'
                              EditBox 90, (y_pos * 20) + 90, 45, 15, LIST_OF_INCOME_ARRAY(gross_amount, all_income) 'gross_amount'
                              EditBox 145, (y_pos * 20) + 90, 25, 15, LIST_OF_INCOME_ARRAY(hours, all_income) 'hours_on_check'
                              OptionGroup RadioGroup1
                                If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) <> 1 Then LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = 1
                                RadioButton 180, (y_pos * 20) + 90, 25, 10, "Yes", LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) 'budget_yes'
                                RadioButton 210, (y_pos * 20) + 90, 25, 10, "No", LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) 'budget_no'
                              EditBox 235, (y_pos * 20) + 90, 115, 15, LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income) 'reason_not_budgeted'
                              EditBox 355, (y_pos * 20) + 90, 45, 15, LIST_OF_INCOME_ARRAY(exclude_amount, all_income) 'not_budgeted_amount'
                              EditBox 410, (y_pos * 20) + 90, 185, 15, LIST_OF_INCOME_ARRAY(reason_amt_excluded, all_income) 'amount_not_budgeted_reason'
                              y_pos = y_pos + 1
                          End If
                      Next


                      Text 5, (dlg_factor * 20) + 115, 70, 10, "Anticipated Income"
                      Text 5, (dlg_factor * 20) + 130, 50, 10, "Rate of Pay/Hr"
                      Text 75, (dlg_factor * 20) + 130, 35, 10, "Hours/Wk"
                      Text 130, (dlg_factor * 20) + 130, 50, 10, "Pay Frequency"
                      Text 225, (dlg_factor * 20) + 115, 70, 10, "Regular Non-Monthly"
                      Text 225, (dlg_factor * 20) + 130, 25, 10, "Amount"
                      Text 280, (dlg_factor * 20) + 130, 50, 10, "Nbr of Months"
                      EditBox 5, (dlg_factor * 20) + 140, 50, 15, EARNED_INCOME_PANELS_ARRAY(pay_per_hr, ei_panel)
                      EditBox 75, (dlg_factor * 20) + 140, 40, 15, EARNED_INCOME_PANELS_ARRAY(hrs_per_wk, ei_panel)
                      DropListBox 130, (dlg_factor * 20) + 140, 85, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                      EditBox 225, (dlg_factor * 20) + 140, 40, 15, EARNED_INCOME_PANELS_ARRAY(reg_non_monthly, ei_panel)
                      EditBox 280, (dlg_factor * 20) + 140, 30, 15, EARNED_INCOME_PANELS_ARRAY(numb_months, ei_panel)
                      ButtonGroup ButtonPressed
                        PushButton 440, (dlg_factor * 20) + 140, 15, 15, "+", add_another_check
                        PushButton 460, (dlg_factor * 20) + 140, 15, 15, "-", take_a_check_away
                        OkButton 495, (dlg_factor * 20) + 140, 50, 15
                        CancelButton 550, (dlg_factor * 20) + 140, 50, 15
                    EndDialog

                    Dialog Dialog1
                    cancel_confirmation

                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                            'ADD ERROR HANDLING HERE
                            If IsDate(LIST_OF_INCOME_ARRAY(pay_date, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter a valid pay date for all checks."
                            If IsNumeric(LIST_OF_INCOME_ARRAY(gross_amount, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the Gross Amount of the check as a number."
                            If LIST_OF_INCOME_ARRAY(budget_in_SNAP_no, all_income) = 1 AND trim(LIST_OF_INCOME_ARRAY(reason_to_exclude, all_income)) = "" Then err_msg = err_msg & vbNewLine & "* The check on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " is to be excluded, list a reason for excluding this check."
                            If IsNumeric(LIST_OF_INCOME_ARRAY(hours, all_income)) = FALSE Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the number of hours for the paycheck on " & LIST_OF_INCOME_ARRAY(pay_date, all_income) & " as a number."
                            If IsNumeric(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) = FALSE AND trim(LIST_OF_INCOME_ARRAY(exclude_amount, all_income)) <> "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter the amount excluded from the budget as a number."
                        End If
                    Next

                    If ButtonPressed = add_another_check Then
                        pay_item = pay_item + 1
                        ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                        LIST_OF_INCOME_ARRAY(panel_indct, pay_item) = ei_panel
                        dlg_factor = dlg_factor + 1

                        sm_err_msg = "LOOP" & sm_err_msg

                    End If

                    If ButtonPressed = take_a_check_away Then
                        pay_item = pay_item - 1
                        ReDim Preserve LIST_OF_INCOME_ARRAY(reason_amt_excluded, pay_item)
                        dlg_factor = dlg_factor - 1
                        sm_err_msg = "LOOP" & sm_err_msg
                    End If

                    If sm_err_msg <> "" AND left(sm_err_msg, 4) <> "LOOP" then MsgBox "Please resolve before continuing:" & vbNewLine & sm_err_msg

                Loop until sm_err_msg = ""

                total_of_counted_income = 0
                total_of_hours = 0
                number_of_checks_budgeted = 0
                EARNED_INCOME_PANELS_ARRAY(pay_weekday, ei_panel) = ""
                list_of_excluded_pay_dates = ""

                For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                    If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                        EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) & "~" & all_income

                        'determining the pay frequency.
                        If EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel) = "" Then

                        End If

                        If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                            If LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = "" Then LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = 0
                            LIST_OF_INCOME_ARRAY(exclude_amount, all_income) = LIST_OF_INCOME_ARRAY(exclude_amount, all_income) * 1
                            LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) * 1
                            net_amount = LIST_OF_INCOME_ARRAY(gross_amount, all_income) - LIST_OF_INCOME_ARRAY(exclude_amount, all_income)
                            total_of_counted_income = total_of_counted_income + net_amount
                            number_of_checks_budgeted = number_of_checks_budgeted + 1

                            LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) * 1
                            total_of_hours = total_of_hours + LIST_OF_INCOME_ARRAY(hours, all_income)
                        Else
                            list_of_excluded_pay_dates = list_of_excluded_pay_dates & ", " & LIST_OF_INCOME_ARRAY(pay_date, all_income)
                        End If
                    End If
                Next

                'ONCE PAY FREQUENCY IS DETERMINED, write to assess if paystubs consititute 30 days and if not, force clarification of income.
                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = total_of_hours / number_of_checks_budgeted
                EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel), 2,,0)

                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = total_of_counted_income / number_of_checks_budgeted
                EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel),2,,0)

                EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = total_of_counted_income / total_of_hours
                EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel) = FormatNumber(EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel),2,,0)

                If list_of_excluded_pay_dates <> "" Then list_of_excluded_pay_dates = right(list_of_excluded_pay_dates, len(list_of_excluded_pay_dates) - 2)

                EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel) = right(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel), len(EARNED_INCOME_PANELS_ARRAY(income_list_indct, ei_panel))-1)
                'Script will determine pay frequency and potentially 1st check (if not listed on JOBS)
                'Script will determine the initial footer month to change by the pay dates listed.
                'Script will create a budget based on the program this income applies to
                'Dialog the budget and have the worker confirm - if they decline - pull the check list dialog back up and have them adjust it there.
                BeginDialog Dialog1, 0, 0, 421, 240, "Confirm JOBS Budget"
                  Text 10, 10, 175, 10, "JOBS 01 01 - EMPLOYER"
                  Text 245, 10, 50, 10, "Pay Frequency"
                  DropListBox 305, 5, 95, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week"+chr(9)+"5 - Other", EARNED_INCOME_PANELS_ARRAY(pay_freq, ei_panel)
                  ' Text 240, 30, 60, 10, "Income Start Date:"
                  ' EditBox 305, 25, 70, 15, income_start_date
                  GroupBox 5, 40, 410, 105, "SNAP Budget"
                  Text 10, 50, 100, 10, "Paychecks Inclued in Budget:"
                  y_pos = 0
                  For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                      If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                        If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then
                          Text 20, (y_pos * 10) + 65, 90, 10, LIST_OF_INCOME_ARRAY(pay_date, all_income) & " - $" & LIST_OF_INCOME_ARRAY(gross_amount, all_income) & " - " & LIST_OF_INCOME_ARRAY(hours, all_income) & "hrs."
                          y_pos = y_pos + 1
                          'Text 20, 65, 90, 10, "01/01/2018 - $400 - 40 hrs"
                          'Text 20, 75, 90, 10, "01/15/2018- $400 - 40 hrs"
                        End If
                      End If
                  Next
                  Text 10, 95, 130, 10, "Paychecks not included: " & list_of_excluded_pay_dates
                  Text 185, 50, 200, 10, "Average hourly rate of pay: $" & EARNED_INCOME_PANELS_ARRAY(hourly_wage, ei_panel)
                  Text 185, 65, 200, 10, "Average weekly hours: " & EARNED_INCOME_PANELS_ARRAY(ave_hrs_per_pay, ei_panel)
                  Text 185, 80, 200, 10, "Average paycheck amount: $" & EARNED_INCOME_PANELS_ARRAY(ave_inc_per_pay, ei_panel)
                  Text 185, 95, 200, 10, "Monthly Budgeted Income: $" & EARNED_INCOME_PANELS_ARRAY(SNAP_mo_inc, ei_panel)
                  CheckBox 10, 110, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
                  Text 10, 130, 60, 10, "Conversation with:"
                  ComboBox 75, 125, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer",  EARNED_INCOME_PANELS_ARRAY(spoke_with, ei_panel)
                  Text 140, 130, 25, 10, "clarifies"
                  EditBox 170, 125, 235, 15, EARNED_INCOME_PANELS_ARRAY(convo_detail, ei_panel)
                  GroupBox 5, 150, 410, 60, "CASH Budget"
                  Text 15, 165, 110, 10, "Actual Paychecks to add to JOBS:"
                  Text 25, 180, 90, 10, "01/01/2018 - $400 - 40 hrs"
                  Text 25, 190, 90, 10, "01/15/2018- $400 - 40 hrs"
                  ButtonGroup ButtonPressed
                    OkButton 315, 220, 50, 15
                    CancelButton 370, 220, 50, 15
                EndDialog


                Dialog Dialog1
                cancel_confirmation

                If confirm_budget_checkbox = unchecked then big_err_msg = big_err_msg & vbNewLine & "*** Since the budget is not confirmed as correct, the ENTER PAY INFORMATION DIALOG will reappear and allow information to be corrected to generate an accurate budget. ***"

                If big_err_msg <> "" Then
                    For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)
                        If LIST_OF_INCOME_ARRAY(panel_indct, all_income) = ei_panel Then
                            If LIST_OF_INCOME_ARRAY(budget_in_SNAP_yes, all_income) = checked Then

                                LIST_OF_INCOME_ARRAY(gross_amount, all_income) = LIST_OF_INCOME_ARRAY(gross_amount, all_income) & ""
                                LIST_OF_INCOME_ARRAY(hours, all_income) = LIST_OF_INCOME_ARRAY(hours, all_income) & ""

                            End If
                        End If
                    Next
                    MsgBox "Review JOBS Pay Information" & vbNewLine & big_err_msg
                End If

            Loop until big_err_msg = ""


            'Add handling to check WREG for correct coding based upon this income information (potentiall update)



            'Worker must confirm the frequency, first pay, and footer month
            'Worker will inicate if future months should be updated - default this to 'yes' as script will update retro and prospective specific to each month
            'SNAP PIC, GRH PIC, HC EI EST will be checked to be updated IF any of these programs are open on the case.

            'NEED to add handling for future/current changes - start or stop work - get policy on this from SNAP refresher - talk to Melissa.
        End If
    End If

    'NAVIGATE to BUSI for each HH MEMBER and ask if Income Information was received for this Self Employment.

    BeginDialog Dialog1, 0, 0, 486, 175, "Enter Self Employment Information"
      Text 10, 10, 180, 10, "BUSI 01 01 - CLIENT NAME"
      Text 200, 10, 80, 10, "Self Employment Type:"
      DropListBox 280, 5, 125, 45, "01 - Farming"+chr(9)+"02 - Real Estate"+chr(9)+"03 - Home Product Sales"+chr(9)+"04 - Other Sales"+chr(9)+"05 - Personal Services"+chr(9)+"06 - Paper Route"+chr(9)+"07 - In Home Daycare"+chr(9)+"08 - Rental Income"+chr(9)+"09 - Other", busi_type
      Text 10, 30, 65, 10, "Verification srouce:"
      DropListBox 90, 25, 75, 45, "1 - Income Tax Returns"+chr(9)+"2 - Receipts of Sales/Purch"+chr(9)+"3 - Client Busi Records/Ledger"+chr(9)+"6 - Other Document"+chr(9)+"N - No Ver Prvd", busi_verif_code
      Text 180, 30, 100, 10, "Amount of Income Information:"
      DropListBox 290, 25, 80, 45, "Select One..."+chr(9)+"A Full Year Totaled"+chr(9)+"Month by Month", amount_income
      Text 10, 50, 120, 10, "Self Employment Budgeting Method"
      DropListBox 135, 45, 85, 45, "01 - 50% Grosss Inc"+chr(9)+"02 - Tax Forms", busi_method
      Text 225, 50, 50, 10, "Selection Date:"
      EditBox 280, 45, 50, 15, method_selection_date
      CheckBox 30, 65, 210, 10, "Check here to confirm this method was discussed with Client.", convo_checkbox
      GroupBox 415, 5, 65, 70, "Apply Income To"
      CheckBox 425, 20, 35, 10, "SNAP", apply_to_SNAP
      CheckBox 425, 35, 35, 10, "CASH", apply_to_CASH
      CheckBox 425, 50, 25, 10, "HC", apply_to_HC
      ButtonGroup ButtonPressed
        PushButton 355, 60, 50, 15, "Ready", open_button
      Text 10, 90, 55, 10, "Month and Year"
      Text 70, 90, 50, 10, "Gross Income"
      Text 130, 80, 90, 10, "Exclude from SNAP Budget"
      Text 130, 90, 30, 10, "Amount"
      Text 190, 90, 30, 10, "Reason"
      EditBox 10, 105, 40, 15, month_year
      EditBox 70, 105, 50, 15, gross_income
      EditBox 130, 105, 50, 15, exclude_from_SNAP
      EditBox 190, 105, 185, 15, exclude_reason
      Text 10, 140, 35, 10, "Tax Year"
      Text 60, 130, 35, 20, "Months in Business"
      Text 110, 140, 30, 10, "Income"
      Text 155, 140, 35, 10, "Expenses"
      EditBox 10, 155, 40, 15, tax_year
      DropListBox 60, 155, 40, 45, "12"+chr(9)+"11"+chr(9)+"10"+chr(9)+"9"+chr(9)+"8"+chr(9)+"7"+chr(9)+"6"+chr(9)+"5"+chr(9)+"4"+chr(9)+"3"+chr(9)+"2"+chr(9)+"1", months_covered
      EditBox 110, 155, 40, 15, tax_income
      EditBox 155, 155, 40, 15, tax_expenses
      ButtonGroup ButtonPressed
        PushButton 320, 155, 15, 15, "+", plus_button
        PushButton 340, 155, 15, 15, "-", minus_button
        OkButton 375, 155, 50, 15
        CancelButton 430, 155, 50, 15
    EndDialog



    'NAVIGATE to RBIC for each HH MEMBER and ask if Income Information was received for this RBIC

Next

For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)
    If EARNED_INCOME_PANELS_ARRAY(income_received, ei_panel) = TRUE Then


    End If
Next
