'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - SSA INFORMATION.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 80                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
call changelog_update("01/27/2022", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(false)
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SSI-RSDI\02-2022 Renewals.xlsx"

'----------------------------------Set up code
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'Excel column constants
const MAXIS_case_number_col             = 1
const recip_pmi_col                     = 2
const mfip_status_col                   = 3
const snap_status_col                   = 4
const ga_status_col                     = 5
const grh_status_col                    = 6
const msa_status_col                    = 7
const hc_status_col                     = 8
const ssi_income_col                    = 9
const ssi_amt_col                       = 10
const ssi_claim_number_col              = 11
const rsdi_one_income_col               = 12
const rsdi_one_amt_col                  = 13
const rsdi_one_claim_col                = 14
const rsdi_two_income_col               = 15
const rsdi_two_amt_col                  = 16
const rsdi_two_claim_col                = 17
const priv_case_col                     = 18
const sves_status_col                   = 19
const sves_sent_col                     = 20
const sves_response_col                 = 21
const resident_name_col                 = 22
const reident_ssn_col                   = 23
const claim_num_col                     = 24
const resident_dob_col                  = 25
const ssn_verif_col                     = 26
const rsdi_record_col                   = 27
const ssi_record_col                    = 28
const rsdi_status_code_col              = 29
const rsdi_staus_desc_col               = 30
const rsdi_paydate_col                  = 31
const rsdi_claim_numb_col               = 32
const rsdi_gross_amt_col                = 33
const rsdi_net_amt_col                  = 34
const gross_minus_net_col               = 35
const rsdi_disa_date_col                = 36
const medi_claim_num_col                = 37
const part_a_start_col                  = 38
const part_a_stop_col                   = 39
const part_a_prem_col                   = 40
const part_b_start_col                  = 41
const part_b_stop_col                   = 42
const part_b_prem_col                   = 43
const ssi_claim_numb_col                = 44
const fed_living_col                    = 45
const cit_ind_code_col                  = 46
const cit_ind_desc_col                  = 47
const ssi_recip_code_col                = 48
const ssi_recip_desc_col                = 49
const ssi_pay_code_col                  = 50
const ssi_pay_desc_col                  = 51
const ssi_denial_code_col               = 52
const ssi_denial_desc_col               = 53
const ssi_denial_date_col               = 54
const ssi_disa_date_col                 = 55
const ssi_SSP_elig_date_col             = 56
const ssi_appeals_code_col              = 57
const ssi_appeals_date_col              = 58
const ssi_appeals_dec_code_col          = 59
const ssi_appeals_dec_desc_col          = 60
const ssi_appeals_dec_date_col          = 61
const ssi_disa_pay_code_col             = 62
const ssi_disa_pay_desc_col             = 63
const ssi_pay_date_col                  = 64
const ssi_gross_amt_col                 = 65
const ssi_over_under_code_col           = 66
const ssi_over_under_desc_col           = 67
const ssi_pay_hist_1_date_col           = 68
const ssi_pay_hist_1_amt_col            = 69
const ssi_pay_hist_1_type_col           = 70
const ssi_pay_hist_1_desc_col           = 71
const ssi_pay_hist_2_date_col           = 72
const ssi_pay_hist_2_amt_col            = 73
const ssi_pay_hist_2_type_col           = 74
const ssi_pay_hist_2_desc_col           = 75
const gross_EI_col                      = 76
const net_EI_col                        = 77
const rsdi_income_amt_col               = 78
const pass_exclusion_col                = 79
const inc_inkind_start_col              = 80
const inc_inkind_stop_col               = 81
const rep_payee_col                     = 82
const member_number_col                 = 83
const SSI_PIC_col                       = 84
const RSDI_one_PIC_col                  = 85
const RSDI_two_PIC_col                  = 86
const DISA_disa_start_col               = 87
const DISA_disa_end_col                 = 88
const DISA_1619_col                     = 89
const DISA_discrepancy_col              = 90
const MEDI_medi_claim_col               = 91
const MEDI_medi_claim_discrepancy_col   = 92
const MEDI_part_a_prem_col              = 93
const MEDI_part_b_prem_col              = 94
const MEDI_buyin_begin_col              = 95
const MEDI_buyin_end_col                = 96
const MEDI_apply_premiums_col           = 97
const MEDI_apply_premiums_thru_col      = 98
const MEDI_part_a_start_col             = 99
const MEDI_part_a_end_col               = 100
const MEDI_part_b_start_col             = 101
const MEDI_part_b_end_col               = 102
const FMED_medi_deduction_col           = 103
const FMED_medi_deduction_amt_col       = 104
const FMED_begin_bene_date_col          = 105
const FMED_end_bene_date_col            = 106
const FMED_payment_discrepancy_col      = 107
const PDED_DAC_code_col                 = 108
const PDED_DAC_discrepancy_col          = 109
const PDED_rep_payee_fee_col            = 110
const BUSI_EI_col                       = 111
const JOBS_EI_col                       = 112

'Establishing array
Dim SSA_information_array()   'Declaring the array
ReDim SSA_information_array(JOBS_EI_const, 2)  'Resizing the array

'Now the script adds all the clients on the excel list into an array
excel_row = 2                   're-establishing the row to start based on when Report 1 starts
entry_record = 0                'incrementer for the array and count
Do
    'Reading information from the BOBI report in Excel
    MAXIS_case_number = objExcel.cells(excel_row, 1).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do

    recip_pmi = objExcel.cells(excel_row, 2).Value
    recip_pmi = trim(recip_pmi)

    income_type = objExcel.cells(excel_row, 12).Value
    income_type = trim(income_type)

    income_amt = objExcel.cells(excel_row, 13).Value
    income_amt = trim(income_amt)

    claim_number = objExcel.cells(excel_row, 14).Value
    claim_number = trim(claim_number)


    If add_to_array = True then
        'Adding client information to the array
        ReDim Preserve SSA_information_array(JOBS_EI_const, entry_record)	'This resizes the array based on the number of cases
        SSA_information_array(MAXIS_case_number_const, entry_record) = MAXIS_case_number
        SSA_information_array(recip_pmi_const        , entry_record) = recip_pmi
        If income_type = SSA_information_array(ssi_income_const       , entry_record) =
        SSA_information_array(ssi_amt_const          , entry_record) =
        SSA_information_array(ssi_claim_number_const , entry_record) =
        SSA_information_array(rsdi_one_income_const  , entry_record) =
        SSA_information_array(rsdi_one_amt_const     , entry_record) =
        SSA_information_array(rsdi_one_claim_const   , entry_record) =
        SSA_information_array(rsdi_two_income_const  , entry_record) =
        SSA_information_array(rsdi_two_amt_const     , entry_record) =
        SSA_information_array(rsdi_two_claim_const   , entry_record) =


        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter
        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
    End if
    excel_row = excel_row + 1
Loop
msgbox entry_record

'array constants
const MAXIS_case_number_const             = 1
const recip_pmi_const                     = 2
const mfip_status_const                   = 3
const snap_status_const                   = 4
const ga_status_const                     = 5
const grh_status_const                    = 6
const msa_status_const                    = 7
const hc_status_const                     = 8
const ssi_income_const                    = 9
const ssi_amt_const                       = 10
const ssi_claim_number_const              = 11
const rsdi_one_income_const               = 12
const rsdi_one_amt_const                  = 13
const rsdi_one_claim_const                = 14
const rsdi_two_income_const               = 15
const rsdi_two_amt_const                  = 16
const rsdi_two_claim_const                = 17
const priv_case_const                     = 18
const sves_status_const                   = 19
const sves_sent_const                     = 20
const sves_response_const                 = 21
const resident_name_const                 = 22
const reident_ssn_const                   = 23
const claim_num_const                     = 24
const resident_dob_const                  = 25
const ssn_verif_const                     = 26
const rsdi_record_const                   = 27
const ssi_record_const                    = 28
const rsdi_status_code_const              = 29
const rsdi_staus_desc_const               = 30
const rsdi_paydate_const                  = 31
const rsdi_claim_numb_const               = 32
const rsdi_gross_amt_const                = 33
const rsdi_net_amt_const                  = 34
const gross_minus_net_const               = 35
const rsdi_disa_date_const                = 36
const medi_claim_num_const                = 37
const part_a_start_const                  = 38
const part_a_stop_const                   = 39
const part_a_prem_const                   = 40
const part_b_start_const                  = 41
const part_b_stop_const                   = 42
const part_b_prem_const                   = 43
const ssi_claim_numb_const                = 44
const fed_living_const                    = 45
const cit_ind_code_const                  = 46
const cit_ind_desc_const                  = 47
const ssi_recip_code_const                = 48
const ssi_recip_desc_const                = 49
const ssi_pay_code_const                  = 50
const ssi_pay_desc_const                  = 51
const ssi_denial_code_const               = 52
const ssi_denial_desc_const               = 53
const ssi_denial_date_const               = 54
const ssi_disa_date_const                 = 55
const ssi_SSP_elig_date_const             = 56
const ssi_appeals_code_const              = 57
const ssi_appeals_date_const              = 58
const ssi_appeals_dec_code_const          = 59
const ssi_appeals_dec_desc_const          = 60
const ssi_appeals_dec_date_const          = 61
const ssi_disa_pay_code_const             = 62
const ssi_disa_pay_desc_const             = 63
const ssi_pay_date_const                  = 64
const ssi_gross_amt_const                 = 65
const ssi_over_under_code_const           = 66
const ssi_over_under_desc_const           = 67
const ssi_pay_hist_1_date_const           = 68
const ssi_pay_hist_1_amt_const            = 69
const ssi_pay_hist_1_type_const           = 70
const ssi_pay_hist_1_desc_const           = 71
const ssi_pay_hist_2_date_const           = 72
const ssi_pay_hist_2_amt_const            = 73
const ssi_pay_hist_2_type_const           = 74
const ssi_pay_hist_2_desc_const           = 75
const gross_EI_const                      = 76
const net_EI_const                        = 77
const rsdi_income_amt_const               = 78
const pass_exclusion_const                = 79
const inc_inkind_start_const              = 80
const inc_inkind_stop_const               = 81
const rep_payee_const                     = 82
const member_number_const                 = 83
const SSI_PIC_const                       = 84
const RSDI_one_PIC_const                  = 85
const RSDI_two_PIC_const                  = 86
const DISA_disa_start_const               = 87
const DISA_disa_end_const                 = 88
const DISA_1619_const                     = 89
const DISA_discrepancy_const              = 90
const MEDI_medi_claim_const               = 91
const MEDI_medi_claim_discrepancy_const   = 92
const MEDI_part_a_prem_const              = 93
const MEDI_part_b_prem_const              = 94
const MEDI_buyin_begin_const              = 95
const MEDI_buyin_end_const                = 96
const MEDI_apply_premiums_const           = 97
const MEDI_apply_premiums_thru_const      = 98
const MEDI_part_a_start_const             = 99
const MEDI_part_a_end_const               = 100
const MEDI_part_b_start_const             = 101
const MEDI_part_b_end_const               = 102
const FMED_medi_deduction_const           = 103
const FMED_medi_deduction_amt_const       = 104
const FMED_begin_bene_date_const          = 105
const FMED_end_bene_date_const            = 106
const FMED_payment_discrepancy_const      = 107
const PDED_DAC_code_const                 = 108
const PDED_DAC_discrepancy_const          = 109
const PDED_rep_payee_fee_const            = 110
const BUSI_EI_const                       = 111
const JOBS_EI_const                       = 112


'variable names
MAXIS_case_number
recip_pmi
mfip_status
snap_status
ga_status
grh_status
msa_status
hc_status
ssi_income
ssi_amt
ssi_claim_number
rsdi_one_income
rsdi_one_amt
rsdi_one_claim
rsdi_two_income
rsdi_two_amt
rsdi_two_claim
priv_case
sves_status
sves_sent
sves_response
resident_name
reident_ssn
claim_num
resident_dob
ssn_verif
rsdi_record
ssi_record
rsdi_status_code
rsdi_staus_desc
rsdi_paydate
rsdi_claim_numb
rsdi_gross_amt
rsdi_net_amt
gross_minus_net
rsdi_disa_date
medi_claim_num
part_a_start
part_a_stop
part_a_prem
part_b_start
part_b_stop
part_b_prem
ssi_claim_numb
fed_living
cit_ind_code
cit_ind_desc
ssi_recip_code
ssi_recip_desc
ssi_pay_code
ssi_pay_desc
ssi_denial_code
ssi_denial_desc
ssi_denial_date
ssi_disa_date
ssi_SSP_elig_date
ssi_appeals_code
ssi_appeals_date
ssi_appeals_dec_code
ssi_appeals_dec_desc
ssi_appeals_dec_date
ssi_disa_pay_code
ssi_disa_pay_desc
ssi_pay_date
ssi_gross_amt
ssi_over_under_code
ssi_over_under_desc
ssi_pay_hist_1_date
ssi_pay_hist_1_amt
ssi_pay_hist_1_type
ssi_pay_hist_1_desc
ssi_pay_hist_2_date
ssi_pay_hist_2_amt
ssi_pay_hist_2_type
ssi_pay_hist_2_desc
gross_EI
net_EI
rsdi_income_amt
pass_exclusion
inc_inkind_start
inc_inkind_stop
rep_payee
member_number
SSI_PIC
RSDI_one_PIC
RSDI_two_PIC
DISA_disa_start
DISA_disa_end
DISA_1619
DISA_discrepancy
MEDI_medi_claim
MEDI_medi_claim_discrepancy
MEDI_part_a_prem
MEDI_part_b_prem
MEDI_buyin_begin
MEDI_buyin_end
MEDI_apply_premiums
MEDI_apply_premiums_thru
MEDI_part_a_start
MEDI_part_a_end
MEDI_part_b_start
MEDI_part_b_end
FMED_medi_deduction
FMED_medi_deduction_amt
FMED_begin_bene_date
FMED_end_bene_date
FMED_payment_discrepancy
PDED_DAC_code
PDED_DAC_discrepancy
PDED_rep_payee_fee
BUSI_EI
JOBS_EI



'User interface dialog - There's just one in this script.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 105, "ADMIN - SSA INFORMATION"
  DropListBox 260, 80, 100, 15, "Select One..."+chr(9)+"1. Send SVES/QURY"+chr(9)+"2. SSA Information Report"+chr(9)+"3. Update MAXIS Panels", SSA_Information_action
  ButtonGroup ButtonPressed
    OkButton 365, 80, 50, 15
    CancelButton 420, 80, 50, 15
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
  EditBox 15, 45, 400, 15, file_selection_path
  Text 30, 65, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 95, "Using this script:"
  Text 140, 85, 120, 10, "Select the SSA Information Process:"
  Text 20, 20, 435, 20, "This script should be used when processing and validating SSA Information for recertification accuracy purposes for SSI/RSDI Income cases. This is part of the RAP - Recertification Accuracy Project QI projects."
EndDialog

'Display dialog and dialog DO...Loop for mandatory fields and password prompting
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
        If SSA_Information_action = "Select One..." then err_msg = err_msg & vbcr & "* Select the process you wish to perform."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, HS_status_col).Value   = date & " MAXIS HS Status"   'col 16
ObjExcel.Cells(1, vendor_num_col).Value  = "Vendor #"                  'col 17
ObjExcel.Cells(1, faci_name_col).Value   = "Facility Name"             'col 18
ObjExcel.Cells(1, faci_in_col).Value     = "Faci In Date"              'col 19
ObjExcel.Cells(1, faci_out_col).Value    = "Faci Out Date"             'col 20
ObjExcel.Cells(1, impact_vnd_col).Value  = "Impacted Vendor?"          'col 21
ObjExcel.Cells(1, exempt_code_col).Value = "VND2 Exemption Code"       'col 22
ObjExcel.Cells(1, HDL_one_col).Value     = "VND2 HDL 1 Code"           'col 23
ObjExcel.Cells(1, HDL_two_col).Value     = "VND2 HDL 2 Code"           'col 24
ObjExcel.Cells(1, HDL_three_col).Value   = "VND2 HDL 3 Code"           'col 25
ObjExcel.Cells(1, case_status_col).Value = "Case Status"               'col 26

FOR i = 16 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------MAXIS DATA GATHER
Call check_for_MAXIS(False)             'Ensuring we're actually in MAXIS
Call MAXIS_footer_month_confirmation    'Ensuring we're in the right footer month/year: current footer month/year for this process.

Dim faci_array()                        'Delcaring array
ReDim faci_array(faci_out_const, 0)     'Resizing the array to size of last const

const vendor_number_const   = 0         'creating array constants
const faci_name_const       = 1
const faci_in_const         = 2
const faci_out_const        = 3

excel_row = 2
Do
    client_PMI = trim(objExcel.cells(excel_row, 1).Value)
    If client_PMI = "" then exit do
    'removing preceeding 0's from the client PMI. This is needed to measure the PMI's on CASE/PERS.
    Do
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)   'trimming off left-most 0 from client_PMI
	Loop until left(client_PMI, 1) <> "0"                                                      'Looping until 0's are all removed
    client_PMI = trim(client_PMI)

	MAXIS_case_number = trim(objExcel.cells(excel_row, 2).Value)
    case_status = ""            'defaulting case_status to "" to increment later in certain circumsatnces

    faci_count = 0                          'setting increment for array

    '----------------------------------------------------------------------------------------------------CASE/PERS & PERS Search
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv)
    If is_this_priv = True then
        case_status = "Privileged Case. Unable to access."
    Else
        member_found = False
        Call navigate_to_MAXIS_screen("CASE", "PERS")
        row = 10    'staring row for 1st member
        Do
            EMReadScreen person_PMI, 8, row, 34
            person_PMI = trim(person_PMI)
            If person_PMI = "" then exit do
            If trim(person_PMI) = client_PMI then
                EmReadscreen HS_status, 1, row, 66
                If trim(HS_status) <> "" then
                    EmReadscreen member_number, 2, row, 3
                    member_found = True
                    exit do
                End if
            Else
                row = row + 3			'information is 3 rows apart. Will read for the next member.
                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
            END if
        LOOP
        If trim(member_number) = "" then case_status = "Unable to locate case for member."
    End if

    If trim(case_status) = "" then
    '----------------------------------------------------------------------------------------------------FACI panel determination
	   call navigate_to_MAXIS_screen("STAT", "FACI")
       EmWriteScreen member_number, 20, 76
       Call write_value_and_transmit("01", 20, 79)  'making sure we're on the 1st instance for member
       'Based on how many FACI panels exist will determine if/how the information is read.
	    EMReadScreen FACI_total_check, 1, 2, 78
	    If FACI_total_check = "0" then
	    	case_status = "No FACI panel on this case for member #" & member_number & "."
	    Elseif FACI_total_check = "1" then
            'just looking through a singular faci panel
            EmReadscreen faci_name, 30, 6, 43
            faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores
            EmReadscreen vendor_number, 8, 5, 43
            vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores

        	row = 18
	    	Do
                EMReadScreen faci_out, 10, row, 71      'faci out date
                If faci_out = "__ __ ____" then
                    faci_out = ""                       'blanking out faci out if not a date
                Else
                    faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do.
                End if
                EMReadScreen faci_in, 10, row, 47       'faci in date
                If faci_in = "__ __ ____" then
                    faci_in = ""                        'blanking out faci in if not a date
                Else
                    faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do.
                End if
	    		If faci_out = "" then
					If faci_in = "" then
                        row = row - 1   'no faci info on this row
                    else
                        If faci_in <> "" then exit do    'open ended faci found
                    End if
	    		Elseif faci_out <> "" then
                    If faci_in <> "" then exit do    'most recent faci span identified
	    		End if
            Loop
        Else
            'Evaluate multiple faci panels
            faci_out_dates_string = ""                  'setting up blank string to increment
            current_faci_found = False                  'defaulting to false - this boolean will determine if evaluation of the last date is needed. Will become true statement if open-ended faci panel is detected.
            For item = 1 to FACI_total_check

                Call write_value_and_transmit("0" & item, 20, 79)   'Entering the item's faci panel via direct navigation field on FACI panel.
                row = 18
                Do
                    EMReadScreen faci_out, 10, row, 71      'faci out date
                    If faci_out = "__ __ ____" then
                        faci_out = ""                       'blanking out faci out if not a date
                    Else
                        faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do.
                    End if
                    EMReadScreen faci_in, 10, row, 47       'faci in date
                    If faci_in = "__ __ ____" then
                        faci_in = ""                        'blanking out faci in if not a date
                    Else
                        faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do.
                    End if

                    EmReadscreen faci_name, 30, 6, 43
                    faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores
                    EmReadscreen vendor_number, 8, 5, 43
                    vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores
                    'Reading the faci in and out dates
                    If faci_out = "" then
                        If faci_in = "" then
                            row = row - 1   'no faci info on this row - this is blank
                        else
                            If faci_in <> "" then
                                current_faci_found = True   'Condition is met so date evaluation via FACI_array is not needed.
                                exit do    'open ended faci found
                            End if
                        End if
                    Elseif faci_out <> "" then
                        If faci_in <> "" then
                            faci_out_dates_string = faci_out_dates_string & faci_out & "|"

                            Redim Preserve faci_array(faci_out_const, faci_count)
                            faci_array(vendor_number_const, faci_count) = vendor_number
                            faci_array(faci_name_const,     faci_count) = faci_name
                            faci_array(faci_in_const,       faci_count) = faci_in
                            faci_array(faci_out_const,      faci_count) = faci_out
                            faci_count = faci_count + 1
                            exit do    'most recent faci span identified
                        End if
                    End if
                Loop
                If current_faci_found = True then exit for  'exiting the for since most current FACI has been found
            Next

            'If an open-ended faci is NOT found, then futher evaluation is needed to determine the most recent date.
            If current_faci_found = False then
                faci_out_dates_string = left(faci_out_dates_string, len(faci_out_dates_string) - 1)
                faci_out_dates = split(faci_out_dates_string, "|")
                call sort_dates(faci_out_dates)
                first_date = faci_out_dates(0)                              'setting the first and last check dates
                last_date = faci_out_dates(UBOUND(faci_out_dates))

                'finding the most recent date if none of the dates are open-ended
                For item = 0 to Ubound(faci_array, 2)
                    If faci_array(faci_out_const, item) = last_date then
                        vendor_number   = faci_array(vendor_number_const, item)
                        faci_name       = faci_array(faci_name_const, item)
                        faci_in         = faci_array(faci_in_const, item)
                        faci_out        = faci_array(faci_out_const, item)
                    End if
                Next
            End if
            ReDim faci_array(faci_out_const, 0)     'Resizing the array back to original size
            Erase faci_array                        'then once resized it gets erased.
	    End if

        '----------------------------------------------------------------------------------------------------VNDS/VND2
        Call Navigate_to_MAXIS_screen("MONY", "VNDS")
        Call write_value_and_transmit(vendor_number, 4, 59)
        Call write_value_and_transmit("VND2", 20, 70)
        EMReadScreen VND2_check, 4, 2, 54
        If VND2_check <> "VND2" then
            case_status = "Unable to find MONY/VND2 panel"
        Else
            health_depart_reason = False    'defalthing to false
            exemption_reason = False

            EmReadscreen exemption_code, 2, 9, 69
            If exemption_code = "__" then exemption_code = ""
            EmReadscreen HDL_one, 2, 10, 69
            EmReadscreen HDL_two, 2, 10, 72
            EmReadscreen HDL_three, 2, 10, 75
            If HDL_one = "__" then HDL_one = ""
            If HDL_two = "__" then HDL_two = ""
            If HDL_three = "__" then HDL_three = ""
            HDL_string = HDL_one & "|" & HDL_two & "|" & HDL_three

            HDL_applicable_codes = "08,09,10"
            If HDL_one <> "" then
                If instr(HDL_applicable_codes, HDL_one) then health_depart_reason = True
            End if

            If HDL_two <> "" then
                If instr(HDL_applicable_codes, HDL_two) then health_depart_reason = True
            End if

            If HDL_three <> "" then
                If instr(HDL_applicable_codes, HDL_three) then health_depart_reason = True
            End if

            If exemption_code = "15" or exemption_code = "26" or exemption_code = "28" then
                exemption_reason = True
            Else
                exmption_reason = False
            End if

            If exemption_code = "28" and instr(HDL_string, "10") then
                impacted_vendor = "No"
            Else
                If (exemption_reason = True and health_depart_reason = True) then
                    impacted_vendor = "Yes"
                Else
                    impacted_vendor = "No"
                End if
            End if
        End if
    End if

    'outputting to Excel
    ObjExcel.Cells(excel_row, HS_status_col).Value   = HS_status
    ObjExcel.Cells(excel_row, vendor_num_col).Value  = vendor_number
    ObjExcel.Cells(excel_row, faci_name_col).Value   = faci_name
    ObjExcel.Cells(excel_row, faci_in_col).Value     = faci_in
    ObjExcel.Cells(excel_row, faci_out_col).Value    = faci_out
    ObjExcel.Cells(excel_row, impact_vnd_col).Value  = impacted_vendor
    ObjExcel.Cells(excel_row, exempt_code_col).Value = exemption_code
    ObjExcel.Cells(excel_row, HDL_one_col).Value     = HDL_one
    ObjExcel.Cells(excel_row, HDL_two_col).Value     = HDL_two
    ObjExcel.Cells(excel_row, HDL_three_col).Value   = HDL_three
    ObjExcel.Cells(excel_row, case_status_col).Value = case_status

    'Blanking out variables at the end of the loop
    HS_status = ""
    vendor_number = ""
    faci_name = ""
    faci_in = ""
    faci_out = ""
    impacted_vendor = ""
    exemption_code = ""
    HDL_one = ""
    HDL_two = ""
    HDL_three = ""
    case_status = ""
    excel_row = excel_row + 1 'setting up the script to check the next row.
    stats_counter = stats_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'formatting the cells
FOR i = 1 to 26
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

MAXIS_case_number = ""  'blanking out for statistical purposes. Cannot collect more than one case number.

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your facility data has been created.")
