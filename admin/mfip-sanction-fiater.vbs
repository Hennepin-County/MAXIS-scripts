'STATS GATHERING=============================================================================================================
name_of_script = "ACTIONS - MFIP SANCTION FIATer.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
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
call changelog_update("11/02/2018", "Bug Fix that would stop the script with long names.", "Casey Love, Hennepin County")
call changelog_update("11/30/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================

'funtion block==============================================================================================================='
FUNCTION SANCTION_BUTTONS
	If ButtonPressed = MEMB_number then
		MEMB_function
		HH_member_array = ""
		FOR i = 0 to total_clients
			IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
				IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
					'msgbox all_clients_
					HH_member_array = Right(all_clients_array(i, 0), len(all_clients_array(i, 0)) ) & ", " & HH_member_array
				END IF
			END IF
		NEXT
		hh_size_split = Len(HH_member_array) - Len(Replace(HH_member_array,",",""))
          hh_size = CStr(hh_size_split)
     End If
	'Button sends case to BGTX. Waits for MAXIS comes back from BG. Then brings Post Pay results into dialog variant
	If ButtonPressed = sanc_type then

	g = 0
	dialog_measures = 0
	FOR clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN
			IF all_clients_array(clt_i, 1) = 1 THEN dialog_measures = dialog_measures + 1
		End If
	Next
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 241, (45 + (dialog_measures * 35)), "SELECT TYPE OF SANCTION"
	  FOR clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN
			IF all_clients_array(clt_i, 1) = 1 THEN DropListBox 170, (20 + (g * 35)), 60, 45, "Select?"+chr(9)+"Employment"+chr(9)+"Child Support"+chr(9)+"Both", type_of_sanction(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN GroupBox 5, (10 + (g * 35)), 230, 30, all_clients_array(clt_i, 0)
			IF all_clients_array(clt_i, 1) = 1 THEN Text 20, (25 + (g * 35)), 120, 10, "Please select the Sanction Type:"
			IF all_clients_array(clt_i, 1) = 1 THEN g = g + 1
		End If
	  Next
	  ButtonGroup ButtonPressed
	    IF total_clients <> "" THEN OkButton 95, (20 + (g * 35)), 50, 15
	EndDialog

	Dialog Dialog1 + vbSystemModal
	End If
END FUNCTION

FUNCTION MEMB_function
    Dialog1 = ""
    BEGINDIALOG Dialog1, 0,  0, 256, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
      Text 10, 5, 145, 10, "Who is sanctioned"
      FOR clt_i = 0 to total_clients
    ' For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
    	IF all_clients_array(clt_i, 0) <> "" THEN checkbox 10, (20 + (clt_i * 15)), 150, 10, all_clients_array(clt_i, 0), all_clients_array(clt_i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
      NEXT
      ButtonGroup ButtonPressed
    	OkButton 200, 20, 50, 15
    	'CancelButton 155, 40, 50, 15
    ENDDIALOG
    Dialog Dialog1 + vbSystemModal
End Function

'end of Function block===========================================================================================================

EMConnect ""

Dim type_of_sanction(50)

call check_for_password (are_we_passworded_out)
'Grabbing case number and putting in the month and year entered from dialog box.
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'first dialog'
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 191, 120, "Sanction"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 70, 25, 20, 15, MAXIS_footer_month
  EditBox 95, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 70, 45, 50, 15
    CancelButton 120, 45, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 60, 10, "Sanction MM/YY:"
  GroupBox 5, 60, 180, 55, "Note:"
  Text 15, 70, 170, 25, "Please make sure you update EMPS or ABPS panels correctly before running the FIAT SANCTION script."
  Text 15, 100, 155, 10, " Cancel script if you need to update the panels."
EndDialog

Do
	err_msg = ""
	Dialog Dialog1
	cancel_confirmation
	Call validate_MAXIS_case_number(err_msg, "*")
	If MAXIS_footer_month = "" OR len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbCr & "You must enter a valid month value of: MM"
	If MAXIS_footer_year = "" OR len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbCr & "You must enter a valid year value of: YY"
	'If Cint(MAXIS_footer_month) > Cint(month(date()) + 2) then err_msg = err_msg & vbCr & "You cannot sanction for more than 2 months in the future"
	If err_msg <> "" then Msgbox err_msg
	call check_for_password (are_we_passworded_out) 'adding functionality for MAXIS v.6 Password Out issue'
Loop until err_msg = ""

Navigate_to_MAXIS_screen "ELIG", "MFIP"

case_ready_for_sanction = FALSE
mx_row = 7

Do
	EMReadScreen mf_elig_status, 7, mx_row, 53
	If mf_elig_status = "UNKNOWN" Then
		case_ready_for_sanction = TRUE
		Exit Do
	ElseIf mf_elig_status = "       " Then
		Exit Do
	Else
		mx_row = mx_row + 1
	End If
Loop until mx_row = 20

'DETERMINE HOW THIS WORKS FOR NON Included HH Memb who are sanctioned'
If case_ready_for_sanction = FALSE Then script_end_procedure_with_error_report("The ELIG version of MFIP does not appear ready to FIAT due to Sanction. Please ensure that a sanction is indicated on EMPS or ABPS and that there are no inhibiting edits in STAT.")

'CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

'reads clients'
'DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
''	EMReadscreen ref_nbr, 3, 4, 33
''	EMReadscreen last_name_array, 25, 6, 30
''	EMReadscreen first_name_array, 12, 6, 63
''	last_name_array = replace(last_name_array, "_", "")
''	last_name_array = Lcase(last_name_array)
''	last_name_array = UCase(Left(last_name_array, 1)) &  Mid(last_name_array, 2)
''	first_name_array = replace(first_name_array, "_", "") '& " "
''	first_name_array = Lcase(first_name_array)
''	first_name_array = UCase(Left(first_name_array, 1)) &  Mid(first_name_array, 2)
''	client_string = ref_nbr & " " & first_name_array & " " & last_name_array
''	client_array = client_array & client_string & "|"
''	transmit
''	Emreadscreen edit_check, 7, 24, 2
'LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
'client_array = TRIM(client_array)
'test_array = split(client_array, "|")
'total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
'DIM all_client_array()
'ReDim all_clients_array(total_clients, 1)
'FOR clt_x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
''	Interim_array = split(client_array, "|")
''	all_clients_array(clt_x, 0) = Interim_array(clt_x)
''	all_clients_array(clt_x, 1) = 1
'NEXT
'HH_size = CStr(total_clients)
'Sets checkboxes to blank'
'FOR clt_i = 0 to total_clients
'all_clients_array(clt_i, 1) = 0
'NEXT

Const clt_ref   = 0
Const clt_name  = 1
Const clt_counted = 2
Const clt_elig = 3
Const part_of_MF = 4 'T/F
Const caregiver = 5 'T/F
Const in_sanc = 6 'T/F

Const clt_drpdwn = 7
Const sanc_type = 8
Const numb_clt_sanc = 9
Const prev_month_in_sanc = 10

deeming_needed = FALSE

Dim client_sanction_array()
ReDim client_sanction_array(2, 0)

Dim MFIP_Members_array ()
ReDim MFIP_Members_array (10, 0)
Call Navigate_to_MAXIS_screen ("ELIG", "MFIP")
'MsgBox "ELIG/MFIP" & vbNewLine & "Line 240"
EMWriteScreen "99", 20, 79
transmit

mx_row = 7
Do
	EMReadScreen approval_status, 15, mx_row, 50
	approval_status = trim(approval_status)
	If approval_status = "APPROVED" Then
		EMReadScreen elig_version, 2, mx_row, 22
		elig_version = trim(elig_version)
		EMWriteScreen elig_version, 18, 54
		transmit
		Exit Do
	Else
		mx_row = mx_row + 1
	End If
Loop until mx_row = 18

mx_row = 7
client_in_case = 0

Do
	EMReadScreen reference_number, 2, mx_row, 6
	EMReadScreen elig_name, 21, mx_row, 10
	EMReadScreen member_code, 16, mx_row, 36
	EMReadScreen elig_status, 10, mx_row, 53

    first_name = ""
    last_name = ""

	If reference_number = "  " Then Exit Do

	elig_name = trim(elig_name)
	comma_pos = Instr(elig_name, ",")
    If comma_pos <> 0 Then
	    last_name = left(elig_name, comma_pos - 1)
        If left(right(elig_name, 2), 1) = " " Then
            first_name = right(left(elig_name, len(elig_name)-2), len(elig_name)-comma_pos-3)
            middle_initial = right(elig_name, 1)
            first_name = first_name & " " & middle_initial
    	Else
    		first_name = right(elig_name, len(elig_name)-comma_pos-1)
    	End If
    End If

	ReDim Preserve MFIP_Members_array(10, client_in_case)
	MFIP_Members_array(clt_ref, client_in_case) = reference_number
	MFIP_Members_array(clt_name, client_in_case) = first_name & " " & last_name
    MFIP_Members_array(clt_name, client_in_case) = trim(MFIP_Members_array(clt_name, client_in_case))
	MFIP_Members_array(clt_counted, client_in_case) = trim(member_code)
	MFIP_Members_array(clt_elig, client_in_case) = trim(elig_status)
''	MsgBox MFIP_Members_array(clt_name, client_in_case) & vbNewLine & "MAXIS Row: " & mx_row & vbNewLine & "Line 286"
	client_in_case = client_in_case + 1
	mx_row = mx_row + 1

Loop until mx_row = 20

Navigate_to_MAXIS_screen "STAT", "MEMB"

For each_client = 0 to UBOUND(MFIP_Members_array, 2)
    If MFIP_Members_array(clt_name, each_client) = "" Then
        EmWriteScreen MFIP_Members_array(clt_ref, each_client), 20, 76
        transmit

        EmReadscreen last_name, 25, 6, 30
        EmReadscreen first_name, 12, 6, 63
        last_name = replace(last_name, "_", "")
        first_name = replace(first_name, "_", "")

        MFIP_Members_array(clt_name, each_client) = first_name & " " & last_name
    End If
Next

Navigate_to_MAXIS_screen "STAT", "PARE"
'MsgBox "STAT/PARE" & vbNewLine & "Line 292"

For client_list = 0 to UBOUND(MFIP_Members_array, 2)

	EMWriteScreen MFIP_Members_array(clt_ref, client_list), 20, 76
	transmit

	EMReadScreen panel_instance, 1, 2, 73

	If panel_instance = "0" Then
		MFIP_Members_array(caregiver, client_list) = FALSE
	Else
		counted_code = left(MFIP_Members_array(clt_counted, client_list), 1)
		If counted_code = "A" OR counted_code = "F" OR counted_code = "G" OR counted_code = "H" OR counted_code = "J" Then
			MFIP_Members_array(caregiver, client_list) = TRUE
		Else
			MFIP_Members_array(caregiver, client_list) = FALSE
		End If
		counted_code = ""
	End If
Next

Navigate_to_MAXIS_screen "STAT", "EMPS"
'MsgBox "STAT/EMPS" & vbNewLine & "Line 315"

For client_list = 0 to UBOUND(MFIP_Members_array, 2)
	If MFIP_Members_array(caregiver, client_list) = TRUE Then
		EMWriteScreen MFIP_Members_array(clt_ref, client_list), 20, 76
		transmit
		EMReadScreen fin_orient_sanc_begin, 8, 6, 39
		EMReadScreen fin_orient_sanc_end, 8, 6, 65

		If fin_orient_sanc_begin <> "__ 01 __" AND fin_orient_sanc_end = "__ 01 __" Then
			MFIP_Members_array(in_sanc, client_list) = TRUE
			MFIP_Members_array(in_sanc, client_list) = "Employment Services"
		End If

		EMReadScreen sanc_rsn, 2, 18, 40
		EMReadScreen sanc_begin, 8, 18, 51
		EMReadScreen sanc_end, 8, 18, 70

		If sanc_rsn <> "__" Then
			If sanc_begin <> "__ 01 __" AND sanc_end = "__ 01 __" Then
				MFIP_Members_array(in_sanc, client_list) = TRUE
				MFIP_Members_array(in_sanc, client_list) = "Employment Services"
			End If
		End If

	End If
	counted_code = left(MFIP_Members_array(clt_counted, client_list), 1)
	If counted_code = "F" OR counted_code = "G" OR counted_code = "H" OR counted_code = "J" Then deeming_needed = TRUE
Next

Navigate_to_MAXIS_screen "STAT", "ABPS"
'MsgBox "STAT/ABPS" & vbNewLine & "Line 346"

Do
	For client_list = 0 to UBOUND(MFIP_Members_array, 2)
		If MFIP_Members_array(caregiver, client_list) = TRUE Then
			EMReadScreen caregiver_ref_number, 2, 4, 47
			If caregiver_ref_number = MFIP_Members_array(clt_ref, client_list) Then
				EMReadScreen support_coop, 1, 4, 73
				If support_coop = "N" Then
					MFIP_Members_array(in_sanc, client_list) = TRUE
					If MFIP_Members_array(in_sanc, client_list) = "Employment Services" Then
						MFIP_Members_array(in_sanc, client_list) = "Both"
					Else
						MFIP_Members_array(in_sanc, client_list) = "Child Support"
					End If
				End If
			End If
		End If
	Next
	transmit
	EMReadScreen next_panel_check, 7, 24, 2
Loop until next_panel_check = "ENTER A"


Dim sanc_occurence(50)
'Dim number_sanctions_after(50)
Dim number_sanctions_total(50)

numb_case_sanctions = 0
case_in_sanc_last_month = FALSE
Navigate_to_MAXIS_screen "STAT", "SANC"
'MsgBox "STAT/SANC" & vbNewLine & "Line 377"

For client_list = 0 to UBOUND(MFIP_Members_array, 2)
	If MFIP_Members_array(caregiver, client_list) = TRUE Then
		EMReadScreen number_of_client_sanctions, 2, 16, 43
		number_of_client_sanctions = trim(number_of_client_sanctions)
		If number_of_client_sanctions = "" Then number_of_client_sanctions = 0
		number_of_client_sanctions = number_of_client_sanctions * 1

		EMReadScreen case_sanctions, 2, 17, 43
		case_sanctions = trim(case_sanctions)

		EMReadScreen case_clsd_7th_sanc, 5, 18, 43
		If case_clsd_7th_sanc <> "     " Then case_sanctions = 7

		If case_sanctions <> "" Then
			case_sanctions = case_sanctions * 1
			If case_sanctions > numb_case_sanctions Then numb_case_sanctions = case_sanctions
		Else
			case_sanctions = 0
		End If


		row = 7
		DO
			EMReadScreen sanc_year, 2, row, 5	'searching for footer/year
		  	IF sanc_year = MAXIS_footer_year then
			  	Exit Do
		  	ELSE
			  	row = row + 1
		  	END IF
		Loop until row = 14
		'making sanc column variables to read it correctly'
		If MAXIS_footer_month = "01" then col = "76"
		If MAXIS_footer_month = "02" then col = "10"
		If MAXIS_footer_month = "03" then col = "16"
		If MAXIS_footer_month = "04" then col = "22"
		If MAXIS_footer_month = "05" then col = "28"
		If MAXIS_footer_month = "06" then col = "34"
		If MAXIS_footer_month = "07" then col = "40"
		If MAXIS_footer_month = "08" then col = "46"
		If MAXIS_footer_month = "09" then col = "52"
		If MAXIS_footer_month = "10" then col = "58"
		If MAXIS_footer_month = "11" then col = "64"
		If MAXIS_footer_month = "12" then col = "70"
		'determines consecutive or have gaps from previous month'
		If col = 76 then row = row - 1

		EMreadScreen prev_sanc_month, 2, row, col
		If prev_sanc_month = "__" OR prev_sanc_month = "DD" OR prev_sanc_month = "SR" then
			MFIP_Members_array(prev_month_in_sanc, client_list) = FALSE
		Else
			MFIP_Members_array(prev_month_in_sanc, client_list) = TRUE
			case_in_sanc_last_month = TRUE
		End If
	End If
Next

Const panel_name = 0
Const panel_ref = 1
const panel_inst = 2
Const retro_income = 3
Const prosp_income = 4

Dim income_array()
ReDim income_array(4, 0)

panel_counter = 0
If deeming_needed = TRUE Then

	For client_listed = 0 to UBOUND(MFIP_Members_array, 2)
		Navigate_to_MAXIS_screen "STAT", "SUMM"
		EMWriteScreen "JOBS", 20, 71
		'MsgBox "STAT/JOBS" & vbNewLine & "Line 448"
		EMWriteScreen MFIP_Members_array(clt_ref, client_listed), 20, 76
		transmit

		EMReadScreen check_if_exists, 14, 24, 13
		EMReadScreen check_if_exists2, 14, 24, 7
		If check_if_exists <> "DOES NOT EXIST" AND check_if_exists2 <> "DOES NOT EXIST" Then
			Do
				EMReadScreen panel_numb, 1, 2, 73
				panel_numb = right("00"& panel_numb, 2)

				ReDim Preserve income_array(4, panel_counter)

				income_array(panel_name, panel_counter) = "JOBS"
				income_array(panel_ref, panel_counter) = MFIP_Members_array(clt_ref, client_listed)
				income_array(panel_inst, panel_counter) = panel_numb

				EMReadScreen gross_wage_lf, 8, 17, 38
				EMReadScreen gross_wage_rt, 8, 17, 67

				gross_wage_lf = trim(gross_wage_lf)
				gross_wage_rt = trim(gross_wage_rt)

				If gross_wage_lf = "" Then gross_wage_lf = 0
				If gross_wage_rt = "" Then gross_wage_rt = 0

				gross_wage_lf = gross_wage_lf * 1
				gross_wage_rt = gross_wage_rt * 1

				income_array(retro_income, panel_counter) = gross_wage_lf
				income_array(prosp_income, panel_counter) = gross_wage_rt

				panel_counter = panel_counter + 1

				transmit
				EMReadScreen next_panel, 7, 24, 2
			Loop until next_panel = "ENTER A"
		End If

		EMWriteScreen "BUSI", 20, 71
		'MsgBox "STAT/BUSI" & vbNewLine & "Line 488"
		EMWriteScreen MFIP_Members_array(clt_ref, client_listed), 20, 76
		transmit

		EMReadScreen check_if_exists, 14, 24, 13
		EMReadScreen check_if_exists2, 14, 24, 7
		If check_if_exists <> "DOES NOT EXIST" AND check_if_exists2 <> "DOES NOT EXIST" Then
			Do
				EMReadScreen panel_numb, 1, 2, 73
				panel_numb = right("00"& panel_numb, 2)

				ReDim Preserve income_array(4, panel_counter)

				income_array(panel_name, panel_counter) = "BUSI"
				income_array(panel_ref, panel_counter) = MFIP_Members_array(clt_ref, client_listed)
				income_array(panel_inst, panel_counter) = panel_numb

				EMReadScreen gross_wage_lf, 8, 8, 55
				EMReadScreen gross_wage_rt, 8, 8, 69

				gross_wage_lf = trim(gross_wage_lf)
				gross_wage_rt = trim(gross_wage_rt)

				If gross_wage_lf = "" Then gross_wage_lf = 0
				If gross_wage_rt = "" Then gross_wage_rt = 0

				gross_wage_lf = gross_wage_lf * 1
				gross_wage_rt = gross_wage_rt * 1

				income_array(retro_income, panel_counter) = gross_wage_lf
				income_array(prosp_income, panel_counter) = gross_wage_rt

				panel_counter = panel_counter + 1

				transmit
				EMReadScreen next_panel, 7, 24, 2
			Loop until next_panel = "ENTER A"
		End If

		EMWriteScreen "UNEA", 20, 71
		'MsgBox "STAT/UNEA" & vbNewLine & "Line 528"
		EMWriteScreen MFIP_Members_array(clt_ref, client_listed), 20, 76
		transmit

		EMReadScreen check_if_exists, 14, 24, 13
		EMReadScreen check_if_exists2, 14, 24, 7
		If check_if_exists <> "DOES NOT EXIST" AND check_if_exists2 <> "DOES NOT EXIST" Then
			Do
				EMReadScreen panel_numb, 1, 2, 73
				panel_numb = right("00"& panel_numb, 2)

				ReDim Preserve income_array(4, panel_counter)

				income_array(panel_name, panel_counter) = "UNEA"
				income_array(panel_ref, panel_counter) = MFIP_Members_array(clt_ref, client_listed)
				income_array(panel_inst, panel_counter) = panel_numb

				EMReadScreen gross_wage_lf, 8, 18, 39
				EMReadScreen gross_wage_rt, 8, 18, 68

				gross_wage_lf = trim(gross_wage_lf)
				gross_wage_rt = trim(gross_wage_rt)

				If gross_wage_lf = "" Then gross_wage_lf = 0
				If gross_wage_rt = "" Then gross_wage_rt = 0

				gross_wage_lf = gross_wage_lf * 1
				gross_wage_rt = gross_wage_rt * 1

				income_array(retro_income, panel_counter) = gross_wage_lf
				income_array(prosp_income, panel_counter) = gross_wage_rt

				panel_counter = panel_counter + 1

				transmit
				EMReadScreen next_panel, 7, 24, 2
			Loop until next_panel = "ENTER A"
		End If
	Next
End If



'FOR clt_i = 0 to total_clients
'IF all_clients_array(clt_i, 0) <> "" THEN
''  IF all_clients_array(clt_i, 1) = 1 THEN
''  call navigate_to_MAXIS_screen("stat", "sanc")
''  MsgBox "STAT/SANC" & vbNewLine & "Line 575"
''  EMWriteScreen Left(all_clients_array(clt_i, 0), 2), 20, 76		'member number'
''  transmit
''  'reads sanctions'
''  'EMReadScreen number_sanctions_after(clt_i), 1, 16, 43
''  EMReadScreen number_sanctions_total(clt_i), 1, 17, 43
''  'EMReadScreen closed_seventh_occurence, 5, 18, 43
''  'EMReadScreen closed_post_seventh_occurence, 5, 19, 43
''  row = 7
''  DO
''	EMReadScreen sanc_year, 2, row, 5	'searching for footer/year
''	IF sanc_year = MAXIS_footer_year then
''		Exit Do
''	ELSE
''		row = row + 1
''	END IF
''  Loop until row = 14
''	'making sanc column variables to read it correctly'
''	If MAXIS_footer_month = "01" then col = "76"
''	If MAXIS_footer_month = "02" then col = "10"
''	If MAXIS_footer_month = "03" then col = "16"
''	If MAXIS_footer_month = "04" then col = "22"
''	If MAXIS_footer_month = "05" then col = "28"
''	If MAXIS_footer_month = "06" then col = "34"
''	If MAXIS_footer_month = "07" then col = "40"
''	If MAXIS_footer_month = "08" then col = "46"
''	If MAXIS_footer_month = "09" then col = "52"
''	If MAXIS_footer_month = "10" then col = "58"
''	If MAXIS_footer_month = "11" then col = "64"
''	If MAXIS_footer_month = "12" then col = "70"
''	'determines consecutive or have gaps from previous month'
''	If col = 76 then
''		row = row - 1
''	End if
''	EMreadScreen prev_sanc_month, 2, row, col
''	If prev_sanc_month = "__" OR prev_sanc_month = "DD"  then
''		sanc_occurence(clt_i) = 1
''	Else
''		sanc_occurence(clt_i) = 2
''	End If
''	IF all_clients_array(clt_i, 1) = 1 THEN
''		'msgbox all_clients_
''		sanc_occurence_for_fiat = sanc_occurence(clt_i) & ", " & sanc_occurence_for_fiat
''	END IF
''
''  End If
'End If
'Next


Do
	err_msg = ""
	dlg_ext = 0

    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 305, 75, "Sanction Detail"
	  For entered_sanc = 0 to UBOUND(MFIP_Members_array, 2)
	  	If MFIP_Members_array(caregiver, entered_sanc) = TRUE Then
		  	Text 10, 25 + (20 * dlg_ext), 176, 10, MFIP_Members_array(clt_ref, entered_sanc) & " - " & MFIP_Members_array(clt_name, entered_sanc)
		  	DropListBox 200, 20 + (20 * dlg_ext), 90, 45, "Select One..."+chr(9)+"Employment Services"+chr(9)+"Child Support"+chr(9)+"Both", MFIP_Members_array(sanc_type, entered_sanc)
			dlg_ext = dlg_ext + 1
		End If
	  Next
	  ButtonGroup ButtonPressed
	   ' 'PushButton 255, 25, 10, 15, "+", plus_button
	    OkButton 200, 55, 50, 15
	    CancelButton 250, 55, 50, 15
	  Text 10, 10, 65, 10, "Client in Sanction"
	  Text 200, 10, 60, 10, "Sanction Type"
	  Text 10, 55, 110, 10, "Case currently has " & case_sanctions & " sanctions"
	EndDialog

	Dialog Dialog1
	cancel_confirmation

	sanction_listed = FALSE
	For entered_sanc = 0 to UBOUND(MFIP_Members_array, 2)
		If MFIP_Members_array(caregiver, entered_sanc) = TRUE Then
			If MFIP_Members_array(sanc_type, entered_sanc) <> "Select One..." Then sanction_listed = TRUE
		End If
	Next

	If sanction_listed = FALSE Then err_msg = err_msg & vbNewLine & "One of the caregivers must have a sanction indicated for the script to correctly FIAT. Indicate which caregiver is in sanction."

	If err_msg <> "" Then MsgBox "** Please resolve for the script to continue. **" & vbNewLine & err_msg
Loop until err_msg = ""

case_sanc_type = ""
For client_listed = 0 to UBOUND(MFIP_Members_array, 2)
	If MFIP_Members_array(sanc_type, client_listed) = "" OR MFIP_Members_array(sanc_type, client_listed) = "Select One..." Then
		MFIP_Members_array(in_sanc, client_listed) = FALSE
	Else
		MFIP_Members_array(in_sanc, client_listed) = TRUE
		If MFIP_Members_array(sanc_type, client_listed) = "Employment Services" Then
			If case_sanc_type = "Child Support" OR case_sanc_type = "Both" Then
				case_sanc_type = "Both"
			Else
				case_sanc_type = "Employment Services"
			End If
		ElseIf MFIP_Members_array(sanc_type, client_listed) = "Child Support" Then
			If case_sanc_type = "Employment Services" OR case_sanc_type = "Both" Then
				case_sanc_type = "Both"
			Else
				case_sanc_type = "Child Support"
			End If
		ElseIf MFIP_Members_array(sanc_type, client_listed) = "Both" Then
			case_sanc_type = "Both"
		End If
	End If
Next
'MsgBox "Case Sanc Type: " & case_sanc_type

If case_sanctions > 0 Then
	percent_sanction = "30"
	sanction_vendor = "Y"
	sanc_occurence_for_fiat = "2"
Else
	If case_sanc_type = "Employment Services" Then
		percent_sanction = "10"
	Else
		percent_sanction = "30"
	End If
	sanction_vendor = "N"
	sanc_occurence_for_fiat = "1"
End If

'Going to SHEL to get vendor information
If sanction_vendor = "Y" Then
	Navigate_to_MAXIS_screen "STAT", "SHEL"

	EMReadScreen shel_landlord, 25, 7, 50

	EMReadScreen shel_retro_rent, 8, 11, 37
	EMReadScreen shel_retro_rent_verif, 2, 11, 48
	EMReadScreen shel_prosp_rent, 8, 11, 56
	EMReadScreen shel_prosp_rent_verif, 2, 11, 67

	shel_retro_rent = trim(shel_retro_rent)
	shel_prosp_rent = trim(shel_prosp_rent)
	If shel_retro_rent = "________" Then shel_retro_rent = 0
	If shel_prosp_rent = "________" Then shel_prosp_rent = 0

	shel_retro_rent = shel_retro_rent * 1
	shel_prosp_rent = shel_prosp_rent * 1

	If shel_retro_rent <> shel_prosp_rent Then
		rent_consistent = FALSE
	Else
		rent_consistent = TRUE
	End If

	rent_verified = TRUE
	If shel_retro_rent_verif = "NO" OR shel_retro_rent_verif = "PC" OR shel_retro_rent_verif = "NC" OR shel_retro_rent_verif = "__" Then rent_verified = FALSE
	If shel_prosp_rent_verif = "NO" OR shel_prosp_rent_verif = "PC" OR shel_prosp_rent_verif = "NC" OR shel_prosp_rent_verif = "__" Then rent_verified = FALSE

	EMReadScreen shel_retro_lot_rent, 8, 12, 37
	EMReadScreen shel_retro_lot_rent_verif, 2, 12, 48
	EMReadScreen shel_prosp_lot_rent, 8, 12, 56
	EMReadScreen shel_prosp_lot_rent_verif, 2, 12, 67

	shel_retro_lot_rent = trim(shel_retro_lot_rent)
	shel_prosp_lot_rent = trim(shel_prosp_lot_rent)
	If shel_retro_lot_rent = "________" Then shel_retro_lot_rent = 0
	If shel_prosp_lot_rent = "________" Then shel_prosp_lot_rent = 0

	shel_retro_lot_rent = shel_retro_lot_rent * 1
	shel_prosp_lot_rent = shel_prosp_lot_rent * 1

	If shel_retro_lot_rent <> shel_prosp_lot_rent Then
		lot_rent_consistent = FALSE
	Else
		lot_rent_consistent = TRUE
	End If

	lot_rent_verified = TRUE
	If shel_retro_lot_rent_verif = "NO" OR shel_retro_lot_rent_verif = "PC" OR shel_retro_lot_rent_verif = "NC" OR shel_retro_lot_rent_verif = "__" Then lot_rent_verified = FALSE
	If shel_prosp_lot_rent_verif = "NO" OR shel_prosp_lot_rent_verif = "PC" OR shel_prosp_lot_rent_verif = "NC" OR shel_prosp_lot_rent_verif = "__" Then lot_rent_verified = FALSE

	EMReadScreen shel_retro_mortgage, 8, 13, 37
	EMReadScreen shel_retro_mortgage_verif, 2, 13, 48
	EMReadScreen shel_prosp_mortgage, 8, 13, 56
	EMReadScreen shel_prosp_mortgage_verif, 2, 13, 67

	shel_retro_mortgage = trim(shel_retro_mortgage)
	shel_prosp_mortgage = trim(shel_prosp_mortgage)
	If shel_retro_mortgage = "________" Then shel_retro_mortgage = 0
	If shel_prosp_mortgage = "________" Then shel_prosp_mortgage = 0

	shel_retro_mortgage = shel_retro_mortgage * 1
	shel_prosp_mortgage = shel_prosp_mortgage * 1

	If shel_retro_mortgage <> shel_prosp_mortgage Then
		mortgage_consistent = FALSE
	Else
		mortgage_consistent = TRUE
	End If

	mortgage_verified = TRUE
	If shel_retro_mortgage_verif = "NO" OR shel_retro_mortgage_verif = "PC" OR shel_retro_mortgage_verif = "NC" OR shel_retro_mortgage_verif = "__" Then mortgage_verified = FALSE
	If shel_prosp_mortgage_verif = "NO" OR shel_prosp_mortgage_verif = "PC" OR shel_prosp_mortgage_verif = "NC" OR shel_prosp_mortgage_verif = "__" Then mortgage_verified = FALSE

	EMReadScreen shel_retro_room, 8, 16, 37
	EMReadScreen shel_retro_room_verif, 2, 16, 48
	EMReadScreen shel_prosp_room, 8, 16, 56
	EMReadScreen shel_prosp_room_verif, 2, 16, 67

	shel_retro_room = trim(shel_retro_room)
	shel_prosp_room = trim(shel_prosp_room)
	If shel_retro_room = "________" Then shel_retro_room = 0
	If shel_prosp_room = "________" Then shel_prosp_room = 0

	shel_retro_room = shel_retro_room * 1
	shel_prosp_room = shel_prosp_room * 1

	If shel_retro_room <> shel_prosp_room Then
		room_consistent = FALSE
	Else
		room_consistent = TRUE
	End If

	room_verified = TRUE
	If shel_retro_room_verif = "NO" OR shel_retro_room_verif = "PC" OR shel_retro_room_verif = "NC" OR shel_retro_room_verif = "__" Then room_verified = FALSE
	If shel_prosp_room_verif = "NO" OR shel_prosp_room_verif = "PC" OR shel_prosp_room_verif = "NC" OR shel_prosp_room_verif = "__" Then room_verified = FALSE

	known_retro_shelter_amount = shel_retro_rent + shel_retro_lot_rent + shel_retro_mortgage + shel_retro_room
	known_prosp_shelter_amount = shel_prosp_rent + shel_prosp_lot_rent + shel_prosp_mortgage + shel_prosp_room
	expense_exists = TRUE

	If known_prosp_shelter_amount = 0 AND known_retro_shelter_amount = 0 Then
		expense_exists = FALSE
		vendor_confirmation_needed = MsgBox ("It appears no expenses are listed on SHEL for Rent, Lot Rent, Morgage, or Room expense." & vbNewLine & "This will mean that no cash beneift will allocated to pay client's shelter expense." & vbNewLine & "Do you confirm that there is no shelter expense on this case?", vbYesNo + vbQuestion, "Zero Shelter Expense")
	End If

	If shel_prosp_rent <> 0 Then use_rent_expense = checked
	If shel_prosp_lot_rent <> 0 Then use_lot_rent_expense = checked
	If shel_prosp_mortgage <> 0 Then use_mortgage_expense = checked
	If shel_prosp_room <> 0 Then use_room_expense = checked

	If vendor_confirmation_needed = vbNo OR expense_exists =TRUE Then
        Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 296, 130, "Confirm Vendor Amount"
		  Text 65, 10, 50, 10, "Retrospective"
		  Text 140, 10, 50, 10, "Prospective"
		  Text 5, 30, 30, 10, "Rent"
		  Text 65, 30, 35, 10, "$" & shel_retro_rent
		  Text 110, 30, 10, 10, shel_retro_rent_verif
		  Text 140, 30, 35, 10, "$" & shel_prosp_rent
		  Text 185, 30, 10, 10, shel_prosp_rent_verif
		  CheckBox 225, 30, 60, 10, "Vendor Rent", use_rent_expense
		  Text 5, 50, 35, 10, "Lot Rent"
		  Text 65, 50, 35, 10, "$" & shel_retro_lot_rent
		  Text 110, 50, 10, 10, shel_retro_lot_rent_verif
		  Text 140, 50, 35, 10, "$" & shel_prosp_lot_rent
		  Text 185, 50, 10, 10, shel_prosp_lot_rent_verif
		  CheckBox 225, 50, 65, 10, "Vendor Lot Rent", use_lot_rent_expense
		  Text 5, 70, 35, 10, "Mortgage"
		  Text 65, 70, 35, 10, "$" & shel_retro_mortgage
		  Text 110, 70, 10, 10, shel_retro_mortgage_verif
		  Text 140, 70, 35, 10, "$" & shel_prosp_mortgage
		  Text 185, 70, 10, 10, shel_prosp_mortgage_verif
		  CheckBox 225, 70, 65, 10, "Vendor Mortgage", use_mortgage_expense
		  Text 5, 90, 25, 10, "Room"
		  Text 65, 90, 35, 10, "$" & shel_retro_room
		  Text 110, 90, 10, 10, shel_retro_room_verif
		  Text 140, 90, 35, 10, "$" & shel_prosp_room
		  Text 185, 90, 10, 10, shel_prosp_room_verif
		  CheckBox 225, 90, 55, 10, "Vendor Room", use_room_expense
		  ButtonGroup ButtonPressed
		    OkButton 180, 110, 50, 15
		    CancelButton 240, 110, 50, 15
		EndDialog

		Do
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
		Loop until err_msg = ""

		vendor_rent_error = FALSE
		vendor_lot_rent_error = FALSE
		vendor_mortgage_error = FALSE
		vendor_room_error = FALSE

		total_errors = 0

		'If use_rent_expense = unchecked AND use_lot_rent_expense = unchecked AND use_mortgage_expense = unchecked AND use_room_expense = unchecked Then total_errors = total_errors + 1

		If use_rent_expense = checked Then
			If rent_verified = FALSE Then total_errors = total_errors + 1
			If rent_consistent = FALSE Then total_errors = total_errors + 1
		End If

		If use_lot_rent_expense = checked Then
			If lot_rent_verified = TALSE Then total_errors = total_errors + 1
			If lot_rent_consistent = FALSE Then total_errors = total_errors + 1
		End If

		If use_mortgage_expense = checked Then
			If mortgage_verified = FALSE Then total_errors = total_errors + 1
			If mortgage_consistent = FALSE Then total_errors = total_errors + 1
		End If

		If use_room_expense = checked Then
			If room_verified = FALSE Then total_errors = total_errors + 1
			If room_consistent = FALSE Then total_errors = total_errors + 1
		End If

		If total_errors <> 0 Then

			x_pos = 20
            Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 345, 55 + (30 * total_errors), "Confirm Vendor Information"
			  Text 80, 5, 195, 10, "Please confirm before vendoring can be allocated in FIAT"
			  If use_rent_expense = checked Then
			  	If rent_verified = FALSE Then
					Text 5, x_pos, 220, 10, "The Rent expense has not been verified, this is a requirement."
					Text 30, x_pos + 15, 230, 10, "Is there verification on file OR a Mandatory Vendor Form(DHS 3365)?"
					OptionGroup RadioGroupRentVerif
					  RadioButton 275, x_pos + 15, 25, 10, "No", rent_no_radio
					  RadioButton 305, x_pos + 15, 25, 10, "Yes", rent_yes_radio
					x_pos = x_pos + 35
				End If
				If rent_consistent = FALSE Then
					Text 5, x_pos, 225, 10, "The Rent prospective and retrospective expenses do not match."
					Text 30, x_pos + 15, 150, 10, "Which amount should be used for vendoring?"
					OptionGroup RadioGroupRentDiff
					  RadioButton 195, x_pos + 15, 50, 10, "Prospective", rent_prosp_radio
					  RadioButton 260, x_pos + 15, 60, 10, "Retrospective", rent_retro_radio
					x_pos = x_pos + 35
				End If
			  End If

			  If use_lot_rent_expense = checked Then
				  If lot_rent_verified = FALSE Then
					  Text 5, x_pos, 220, 10, "The Lot Rent expense has not been verified, this is a requirement."
					  Text 30, x_pos + 15, 230, 10, "Is there verification on file OR a Mandatory Vendor Form(DHS 3365)?"
					  OptionGroup RadioGroupRentVerif
						RadioButton 275, x_pos + 15, 25, 10, "No", lot_rent_no_radio
						RadioButton 305, x_pos + 15, 25, 10, "Yes", lot_rent_yes_radio
					  x_pos = x_pos + 35
				  End If
				  If lot_rent_consistent = FALSE Then
					  Text 5, x_pos, 225, 10, "The Lot Rent prospective and retrospective expenses do not match."
					  Text 30, x_pos + 15, 150, 10, "Which amount should be used for vendoring?"
					  OptionGroup RadioGroupRentDiff
						RadioButton 195, x_pos + 15, 50, 10, "Prospective", lot_rent_prosp_radio
						RadioButton 260, x_pos + 15, 60, 10, "Retrospective", lot_rent_retro_radio
					  x_pos = x_pos + 35
				  End If
			  End If

			  If use_mortgage_expense = checked Then
				  If mortgage_verified = FALSE Then
					  Text 5, x_pos, 220, 10, "The Mortgage expense has not been verified, this is a requirement."
					  Text 30, x_pos + 15, 230, 10, "Is there verification on file OR a Mandatory Vendor Form(DHS 3365)?"
					  OptionGroup RadioGroupRentVerif
						RadioButton 275, x_pos + 15, 25, 10, "No", mortgage_no_radio
						RadioButton 305, x_pos + 15, 25, 10, "Yes", mortgage_yes_radio
					  x_pos = x_pos + 35
				  End If
				  If mortgage_consistent = FALSE Then
					  Text 5, x_pos, 225, 10, "The Mortgage prospective and retrospective expenses do not match."
					  Text 30, x_pos + 15, 150, 10, "Which amount should be used for vendoring?"
					  OptionGroup RadioGroupRentDiff
						RadioButton 195, x_pos + 15, 50, 10, "Prospective", mortgage_prosp_radio
						RadioButton 260, x_pos + 15, 60, 10, "Retrospective", mortgage_retro_radio
					  x_pos = x_pos + 35
				  End If
			  End If

			  If use_room_expense = checked Then
				  If room_verified = FALSE Then
				  Text 5, x_pos, 220, 10, "The Room expense has not been verified, this is a requirement."
				  Text 30, x_pos + 15, 230, 10, "Is there verification on file OR a Mandatory Vendor Form(DHS 3365)?"
				  OptionGroup RadioGroupRentVerif
					RadioButton 275, x_pos + 15, 25, 10, "No", room_no_radio
					RadioButton 305, x_pos + 15, 25, 10, "Yes", room_yes_radio
				  x_pos = x_pos + 35
				  End If
				  If room_consistent = FALSE Then
				  Text 5, x_pos, 225, 10, "The Room prospective and retrospective expenses do not match."
				  Text 30, x_pos + 15, 150, 10, "Which amount should be used for vendoring?"
				  OptionGroup RadioGroupRentDiff
					RadioButton 195, x_pos + 15, 50, 10, "Prospective", room_prosp_radio
					RadioButton 260, x_pos + 15, 60, 10, "Retrospective", room_retro_radio
				  x_pos = x_pos + 35
				  End If
			  End If
			  ButtonGroup ButtonPressed
				OkButton 225, x_pos, 50, 15
				CancelButton 285, x_pos, 50, 15
			EndDialog

			Do
				vnderr_err_msg = err_msg
				dialog Dialog1
				cancel_confirmation
			Loop until vnderr_err_msg = ""
		End If


		total_vendor_amount = 0
		known_shelter_amount = shel_prosp_rent + shel_prosp_lot_rent + shel_prosp_mortgage + shel_prosp_room
		shelter_not_verified = ""

		If use_rent_expense = checked Then
			If rent_consistent = FALSE Then
				If rent_prosp_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_prosp_rent
				If rent_retro_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_retro_rent
			Else
				total_vendor_amount = total_vendor_amount + shel_prosp_rent
			End If

			If rent_no_radio = 1 Then shelter_not_verified = "Rent, "
		End If

		If use_lot_rent_expense = checked Then
			If lot_rent_consistent = FALSE Then
				If lot_rent_prosp_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_prosp_lot_rent
				If lot_rent_retro_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_retro_lot_rent
			Else
				total_vendor_amount = total_vendor_amount + shel_prosp_lot_rent
			End If

			If lot_rent_no_radio = 1 Then shelter_not_verified = "Lot Rent, "
		End If

		If use_mortgage_expense = checked Then
			If mortgage_consistent = FALSE Then
				If mortgage_prosp_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_prosp_mortgage
				If mortgage_retro_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_retro_mortgage
			Else
				total_vendor_amount = total_vendor_amount + shel_prosp_mortgage
			End If

			If mortgage_no_radio = 1 Then shelter_not_verified = "Mortgage, "
		End If

		If use_room_expense = checked Then
			If room_consistent = FALSE Then
				If room_prosp_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_prosp_room
				If room_retro_radio = 1 Then total_vendor_amount = total_vendor_amount + shel_retro_room
			Else
				total_vendor_amount = total_vendor_amount + shel_prosp_room
			End If

			If room_no_radio = 1 Then shelter_not_verified = "Room, "
		End If

		If shelter_not_verified <> "" Then script_end_procedure_with_error_report ("FIAT Cancelled. The script will now end. This sanction requires Mandatory Vendoring. Shelter Costs must be verified or a Mandatory Vendor Form (DHS 3365) completed.")
		If total_vendor_amount = 0 Then
			expense_exists = FALSE
			vendor_confirmation_needed = MsgBox ("It appears no expenses are listed on SHEL for Rent, Lot Rent, Morgage, or Room expense." & vbNewLine & "This will mean that no cash beneift will allocated to pay client's shelter expense." & vbNewLine & "Do you confirm that there is no shelter expense on this case?", vbYesNo + vbQuestion, "Zero Shelter Expense")
		Else
			expense_exists = TRUE
		End if
	End If

	If expense_exists = FALSE AND vendor_confirmation_needed = vbNo Then script_end_procedure_with_error_report ("FIAT cancelled. Shelter expense appears to be Zero." & vbNewLine & "Worker did not confirm this to be correct." & vbNewLine & "Script cannot continue as this case has Mandatory Vendoring and the shelter expense is not confirmed.")
End If

'MsgBox "Vendor Amount: " & total_vendor_amount

'jumps to sanc fiat'
Navigate_to_MAXIS_screen "FIAT", ""

EMWritescreen "21", 4, 34
EMWritescreen "X", 9, 22
Transmit

'MsgBox "In FIAT"
mx_row = 9
Do
	'MsgBox mx_row
	EMReadScreen elig_status, 4, mx_row, 55
	If elig_status = "UNKN" Then
		EMWriteScreen "X", mx_row, 4
		transmit
		PF3
	ElseIf elig_status = "INEL" Then
		EMReadScreen memb_code, 1, mx_row, 37
		If memb_code = "H" OR memb_code = "F" OR memb_code = "G" OR memb_code = "J" Then
			EMWriteScreen "X", mx_row, 4
			transmit
			PF3
		End If
	End If
	mx_row = mx_row + 1
Loop until elig_status = "    "

EMWritescreen "X", 16, 4
EMWritescreen "x", 17, 4

transmit

'Bypassing FMCR
PF3

'Entering the Sanction information
For client_listed = 0 to UBOUND(MFIP_Members_array, 2)
	If MFIP_Members_array(in_sanc, client_listed) = TRUE Then
		mx_row = 9
		EMReadScreen fmbf_ref_nbr, 2, mx_row, 4
		'MsgBox fmbf_ref_nbr
		If fmbf_ref_nbr = MFIP_Members_array(clt_ref, client_listed) Then
			EMWriteScreen "X", mx_row, 65
			transmit
			If MFIP_Members_array(sanc_type, client_listed) = "Employment Services" Then
				EMWriteScreen "FAILED", 9, 14
			ElseIf MFIP_Members_array(sanc_type, client_listed) = "Child Support" Then
				EMWriteScreen "FAILED", 7, 14
			ElseIf MFIP_Members_array(sanc_type, client_listed) = "Both" then
				EMWriteScreen "FAILED", 9, 14
				EMWriteScreen "FAILED", 7, 14
			End If
			EMWritescreen sanc_occurence_for_fiat, 13, 24
			EMWritescreen MAXIS_footer_month, 13, 42
			EMWritescreen MAXIS_footer_year, 13, 45
			'MsgBox "Pause"
			PF3
		End If
		mx_row = mx_row + 1
	End If
Next

'bypasses warning vnd signs'
EMWritescreen percent_sanction, 15, 29
EMWritescreen sanction_vendor, 15, 49
Transmit
EMreadScreen warning_vnd, 7, 24, 8
If warning_vnd = "WARNING" then PF3
PF3

''	For client_listed = 0 to UBOUND(MFIP_Members_array, 2)
''		counted_code = left(MFIP_Members_array(clt_counted, client_listed), 1)
''		If counted_code = "F" OR counted_code = "G" OR counted_code = "H" OR counted_code = "J" Then
''			mx_row = 8
''			Do
''				EMReadScreen elig_ref, 2, mx_row, 12
''				If elig_ref = MFIP_Members_array(clt_ref, client_listed) Then
''				 	EMWriteScreen "X", mx_row, 8
''					transmit
''				Else
''					mx_row = mx_row + 1
''				End If
''			Loop until elig_ref = "  "
''		End If
''	Next

	'bypasses warning vnd signs'
''	EMWritescreen percent_sanction, 15, 29
''	EMWritescreen sanction_vendor, 15, 49
''	Transmit
''	EMreadScreen warning_vnd, 7, 24, 8
''	If warning_vnd = "WARNING" then PF3
''	PF3

'going to budget
EMWritescreen "X", 18, 4
transmit
'transmit through all of budget
'NEED TO ADD VENDOR AMOUNT! - will be entered here'
If sanction_vendor = "Y" Then
	EMWritescreen "X", 8, 44
	transmit
	EMReadScreen mf_cash_portion, 8, 10, 56
	mf_cash_portion = trim(mf_cash_portion)
	mf_cash_portion = mf_cash_portion * 1
	If total_vendor_amount > mf_cash_portion Then
		shel_expense_not_vendored = total_vendor_amount - mf_cash_portion
		vendor_allotment = mf_cash_portion
	Else
		vendor_allotment = total_vendor_amount
	End If
	EMWritescreen "        ", 11, 56
	EMWritescreen vendor_allotment, 11, 56
	transmit
	transmit
End If

transmit
transmit
PF3

PF3		'Leaving FIAT
EMWritescreen "Y", 13, 41
Transmit
EMWritescreen "N", 11, 52
Transmit
PF3
If sanction_vendor = "Y" then sanction_vendor = "Yes"
If sanction_vendor = "N" then sanction_vendor = "No"

If shel_expense_not_vendored <> "" Then vendor_allotment = vendor_allotment & ". $" & shel_expense_not_vendored & " of the shelter expense was not vendored as only $" & vendor_allotment & " of cash benefit is due to be issued"
If sanction_vendor = "Yes" Then end_message = "Success! A MFIP Sanction version was created for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "." & vbNewLine &_
 	"The sanction was " & percent_sanction & "% deduction and the vendor was set to: '" & sanction_vendor & "." & vbNewLine &_
	"Shelter expenses known are $" & total_vendor_amount & ". The ammount allocated to vendor in FIAT is $" & vendor_allotment & "." & vbNewLine &_
	"Please review your results and run the NOTES - MFIP SANCTION/DWP Disqualification script if needed."

If sanction_vendor = "No" Then end_message = "Success! A MFIP Sanction version was created for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "." & vbNewLine &_
 	"The sanction was " & percent_sanction & "% deduction and the vendor was set to: '" & sanction_vendor & "." & vbNewLine &_
	"Please review your results and run the NOTES - MFIP SANCTION/DWP Disqualification script if needed."

script_end_procedure_with_error_report(end_message)
