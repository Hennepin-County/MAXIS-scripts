'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - CASE SAMPLING.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
call changelog_update("02/12/2025", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' DECLARATIONS =============================================================================================================
const case_number_const		= 00
const worker_number_const	= 01
const case_name_const		= 02
const appl_date_const		= 03
const days_pending_const	= 04
const cash_status_const		= 05
const cash_prog_const		= 06
const snap_status_const		= 07
const pending_today_const	= 08
const population_const		= 09
const create_review_file_const	= 10
const on_daily_list_const	= 11
const last_pend_array_const = 15

Dim TODAYS_PENDING_CASES_ARRAY()
ReDim TODAYS_PENDING_CASES_ARRAY(last_pend_array_const, 0)

Dim YESTERDAYS_PENDING_CASES_ARRAY()
ReDim YESTERDAYS_PENDING_CASES_ARRAY(last_pend_array_const, 0)

'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1

set basket_detail = CreateObject("Scripting.Dictionary")

'Team 1 Clifton			'Team 2 Coenen			'Team 3 Garrett			'Team 4 Groves
basket_detail.add "X127EQ9", "Adults"			'OLD BASKET STRUCTURE???
basket_detail.add "X127EK8", "Adults" ' - Pending 1"
basket_detail.add "X127EH1", "Adults" ' - Pending 1"
basket_detail.add "X127EP1", "Adults" ' - Pending 1"
basket_detail.add "X127EP2", "Adults" ' - Pending 2"
basket_detail.add "X127EH8", "Adults" ' - Pending 2"
basket_detail.add "X127EP6", "Adults" ' - Pending 2"
basket_detail.add "X127EP7", "Adults" ' - Pending 3"
basket_detail.add "X127EP8", "Adults" ' - Pending 3"
basket_detail.add "X127EP3", "Adults" ' - Pending 3"
basket_detail.add "X127EH7", "Adults" ' - Pending 4"
basket_detail.add "X127EK3", "Adults" ' - Pending 4"
basket_detail.add "X127EK7", "Adults" ' - Pending 4"
basket_detail.add "X127EQ5", "Adults" ' Active 1"
basket_detail.add "X127EQ6", "Adults" ' Active 1"
basket_detail.add "X127EQ7", "Adults" ' Active 1"
basket_detail.add "X127EQ8", "Adults" ' Active 1"
basket_detail.add "X127EX1", "Adults" ' Active 1"
basket_detail.add "X127EX2", "Adults" ' Active 1"
basket_detail.add "X127EX3", "Adults" ' Active 1"
basket_detail.add "X127EX4", "Adults" ' Active 1"
basket_detail.add "X127EX5", "Adults" ' Active 1"
basket_detail.add "X127EX7", "Adults" ' Active 1"
' basket_detail.add "X127F3H", "Adults" ' Active 1"		'DELETE?
basket_detail.add "X127EL7", "Adults" ' Active 2"
basket_detail.add "X127EL8", "Adults" ' Active 2"
basket_detail.add "X127EL9", "Adults" ' Active 2"
basket_detail.add "X127EN1", "Adults" ' Active 2"
basket_detail.add "X127EN2", "Adults" ' Active 2"
basket_detail.add "X127EN3", "Adults" ' Active 2"
basket_detail.add "X127EN5", "Adults" ' Active 2"
basket_detail.add "X127EN4", "Adults" ' Active 2"
basket_detail.add "X127EN7", "Adults" ' Active 2"
basket_detail.add "X127EN8", "Adults" ' Active 3"
basket_detail.add "X127EN9", "Adults" ' Active 3"
basket_detail.add "X127EQ1", "Adults" ' Active 3"
basket_detail.add "X127EQ2", "Adults" ' Active 3"
basket_detail.add "X127EQ3", "Adults" ' Active 3"
basket_detail.add "X127EQ4", "Adults" ' Active 3"
basket_detail.add "X127EX8", "Adults" ' Active 3"
basket_detail.add "X127EX9", "Adults" ' Active 3"
basket_detail.add "X127EG4", "Adults" ' Active 3"
basket_detail.add "X127ED8", "Adults" ' Active 4"
basket_detail.add "X127EE1", "Adults" ' Active 4"
basket_detail.add "X127EE2", "Adults" ' Active 4"
basket_detail.add "X127EE3", "Adults" ' Active 4"
basket_detail.add "X127EE4", "Adults" ' Active 4"
basket_detail.add "X127EE5", "Adults" ' Active 4"
basket_detail.add "X127EE6", "Adults" ' Active 4"
basket_detail.add "X127EE7", "Adults" ' Active 4"
basket_detail.add "X127EL1", "Adults" ' Active 4"
basket_detail.add "X127EL2", "Adults" ' Active 4"
basket_detail.add "X127EL3", "Adults" ' Active 4"
basket_detail.add "X127EL4", "Adults" ' Active 4"
basket_detail.add "X127EL5", "Adults" ' Active 4"
basket_detail.add "X127EL6", "Adults" ' Active 4"

basket_detail.add "X127ET5", "Families" 		'Active 1"
basket_detail.add "X127ET6", "Families" 		'Active 1"
basket_detail.add "X127ET7", "Families" 		'Active 1"
basket_detail.add "X127ET8", "Families" 		'Active 1"
basket_detail.add "X127ET9", "Families" 		'Active 1"
basket_detail.add "X127EZ1", "Families" 		'Active 1"
basket_detail.add "X127ES1", "Families" 		'Active 2"
basket_detail.add "X127ES2", "Families" 		'Active 2"
basket_detail.add "X127ET1", "Families" 		'Active 2"
' basket_detail.add "X127F4E", "Families" 		'Active 2"		'DELETE?
basket_detail.add "X127EZ7", "Families" 		'Active 2"
basket_detail.add "X127FB7", "Families" 		'Active 2"
basket_detail.add "X127ET2", "Families" 		'Active 3"
basket_detail.add "X127ET3", "Families" 		'Active 3"
basket_detail.add "X127ET4", "Families" 		'Active 3"
basket_detail.add "X127ES3", "Families" 		'Active 4"
basket_detail.add "X127ES4", "Families" 		'Active 4"
basket_detail.add "X127ES5", "Families" 		'Active 4"
basket_detail.add "X127ES6", "Families" 		'Active 4"
basket_detail.add "X127ES7", "Families" 		'Active 4"
basket_detail.add "X127ES8", "Families" 		'Active 4"
basket_detail.add "X127ES9", "Families" 		'Active 4"
basket_detail.add "X127EZ6", "Families" 		'- Pending 1"
basket_detail.add "X127EZ8", "Families" 		'- Pending 1"
basket_detail.add "X127EZ9", "Families" 		'- Pending 2"
basket_detail.add "X127EH4", "Families" 		'- Pending 2"
basket_detail.add "X127EH5", "Families" 		'- Pending 3"
basket_detail.add "X127EH6", "Families" 		'- Pending 3"
basket_detail.add "X127EZ3", "Families" 		'- Pending 4"
basket_detail.add "X127EZ4", "Families" 		'- Pending 4"

basket_detail.add "X127F3P", "Adults"   'MA-EPD Adults Basket
basket_detail.add "X127F3K", "Families"  'MA-EPD FAD Basket
' basket_detail.add "X127F3P", "Families - General"		- this is MAEPD

basket_detail.add "X127FE7", "DWP"
basket_detail.add "X127FE8", "DWP"
basket_detail.add "X127FE9", "DWP"
basket_detail.add "X127EY8", "DWP"
basket_detail.add "X127EY9", "DWP"

basket_detail.add "X127FA5", "YET"
basket_detail.add "X127FA6", "YET"
basket_detail.add "X127FA7", "YET"
basket_detail.add "X127FA8", "YET"
basket_detail.add "X127FB1", "YET"
basket_detail.add "X127FA9", "YET"

basket_detail.add "X127EN6", "TEFRA"
basket_detail.add "X127FG1", "Foster Care / IV-E"
basket_detail.add "X127EW6", "Foster Care / IV-E"
basket_detail.add "X1274EC", "Foster Care / IV-E"
basket_detail.add "X127FG2", "Foster Care / IV-E"
basket_detail.add "X127EW4", "Foster Care / IV-E"

basket_detail.add "X127EM8", "GRH / HS - Adults Pending"
basket_detail.add "X127FE6", "GRH / HS - Adults Pending"
basket_detail.add "X127EZ2", "GRH / HS - Families Pending"
basket_detail.add "X127EM2", "GRH / HS - Maintenance"
basket_detail.add "X127EH9", "GRH / HS - Maintenance"
basket_detail.add "X127EJ4", "GRH / HS - Maintenance"
basket_detail.add "X127EH2", "GRH / HS - Maintenance"
basket_detail.add "X127EP4", "GRH / HS - Maintenance"
basket_detail.add "X127EK5", "GRH / HS - Maintenance"
basket_detail.add "X127EG5", "GRH / HS - Maintenance"

'basket_detail.add "X127EG4", "MIPPA"
basket_detail.add "X127F3D", "MA - BC"

basket_detail.add "X127EF8", "1800"
basket_detail.add "X127EF9", "1800"
basket_detail.add "X127EG9", "1800"
basket_detail.add "X127EG0", "1800"

basket_detail.add "X1275H5", "Privileged Cases"
basket_detail.add "X127FAT", "Privileged Cases"
basket_detail.add "X127F3H", "Privileged Cases"
'Contacted Case Mgt
basket_detail.add "X127FG6", "LTC+"           '"Kristen Kasem"
basket_detail.add "X127FG7", "LTC+"           '"Kristen Kasem"
basket_detail.add "X127EM3", "LTC+"           '"True L. or Gina G."
basket_detail.add "X127EM4", "LTC+"            '"True L. or Gina G."
basket_detail.add "X127EW7", "LTC+"            '"Kimberly Hill"
basket_detail.add "X127EW8", "LTC+"            '"Kimberly Hill"
basket_detail.add "X127FF4", "LTC+"            '"Alyssa Taylor"
basket_detail.add "X127FF5", "LTC+"            '"Alyssa Taylor"
basket_detail.add "X127FF8", "LTC+"				'"Contracted - North Memorial"
basket_detail.add "X127FF6", "LTC+"				'"Contracted - HCMC"
basket_detail.add "X127FF7", "LTC+"				'"Contracted - HCMC"

' basket_detail.add "X127EK4", "LTC+ - General"
' basket_detail.add "X127EK9", "LTC+ - General"
' basket_detail.add "X127EH1", "LTC+"
basket_detail.add "X127EH3", "LTC+"
' basket_detail.add "X127EH4", "LTC+"
' basket_detail.add "X127EH5", "LTC+"
' basket_detail.add "X127EH6", "LTC+"
' basket_detail.add "X127EH7", "LTC+"
basket_detail.add "X127EJ8", "LTC+"
basket_detail.add "X127EK1", "LTC+"
basket_detail.add "X127EK2", "LTC+"
' basket_detail.add "X127EK3", "LTC+"
basket_detail.add "X127EK4", "LTC+"
basket_detail.add "X127EK6", "LTC+"
' basket_detail.add "X127EK7", "LTC+"
' basket_detail.add "X127EK8", "LTC+"
basket_detail.add "X127EK9", "LTC+"
basket_detail.add "X127EM9", "LTC+"
' basket_detail.add "X127EN6", "LTC+"
basket_detail.add "X127EP5", "LTC+"
basket_detail.add "X127EP9", "LTC+"
basket_detail.add "X127EZ5", "LTC+"
basket_detail.add "X127F3F", "LTC+"
basket_detail.add "X127FE5", "LTC+"
basket_detail.add "X127FH4", "LTC+"
basket_detail.add "X127FH5", "LTC+"
basket_detail.add "X127FI2", "LTC+"
basket_detail.add "X127FI7", "LTC+"

basket_detail.add "X127FI1", "METS Retro Request"


'DAILY COMPILATION column
const comp_case_numb_col 			= 01   						'Case #
const comp_case_name_col 			= 02   						'Case Name
const comp_appl_date_col 			= 03   						'Application Date
const comp_appl_date_issue_col 		= 04   						'Are there issue(s) with application date(s)?
const comp_prog_date_align_col 		= 05   						'Were all pending program dates aligned on PROG?
const comp_staff_not_meet_appl_stndrd_col = 06  				'Identify staff not meeting appplication standards(comma seperated if more than 1)
const comp_appl_notes_col 			= 07   						'Application Notes
const comp_progs_col 				= 08   						'Programs applied
const comp_spec_cash_prog_col 		= 09   						'Specific Cash Program(s) - If applicable.
const comp_hh_comp_col 				= 10   						'HH Composition
const comp_hh_comp_correct_col 		= 11   						'HH comp correct?
const comp_staff_not_meet_demo_stndrd_col = 12  				'Identify staff not meeting demographic standards(comma seperated if more than 1)
const comp_prog_hh_comp_notes_col 	= 13  						'Prog/HH Comp Notes
const comp_intvw_date_col 			= 14   						'Interview Date
const comp_intvw_script_used_col 	= 15   						'Was the Interview Script used?
const comp_single_intvw_col 		= 16   						'Was single interview completed for multiple programs?
const comp_mf_orient_complete_col 	= 17   						'If MFIP, was orientation completed?
const comp_staff_not_meet_intvw_stndrd_col = 18 				'Identify staff not meeting interview standards(comma seperated if more than 1)
const comp_intvw_notes_col 			= 19   						'Interview Notes
const comp_verif_req_sent_col 		= 20   						'Was a verification request sent to the case/household?
const comp_verif_req_blank_col 		= 21   						'Was the verificaiton request blank?
const comp_single_verif_req_col 	= 22   						'Were all the verifications requested in a single verification request?
const comp_spec_forms_req_col 		= 23   						'Did the worker ask for specific forms instead of providing options to verify proofs?
const comp_unnec_verfs_col 			= 24   						'Did worker ask for unnecessary verifications?
const comp_staff_not_meet_verif_stndrd_col = 25   				'Identify staff not meeting verification standards(comma seperated if more than 1)
const comp_verif_notes_col 			= 26   						'Verification Notes
const comp_date_app_col 			= 27   						'Date(s) case was APP's (approved/denied)
const comp_snap_det_exp_col 		= 28   						'If SNAP, was the case DETERMINED to be expedited?
const comp_same_day_act_col 		= 29   						'If more than one program pending, were all programs acted on the same day?
const comp_stat_pact_col 			= 30   						'Was STAT/PACT used?
const comp_ecf_docs_accept_col 		= 31    						'Were all ECF Documents accepted?
const comp_residnt_in_office_col 	= 32   						'During the pending period did resident/AREP come into the office?
const comp_staff_not_meet_app_stndrd_col = 33					'Identify staff not meeting approval standards(comma seperated if more than 1)
const comp_app_notes_col 			= 34   						'Approval Notes:
const comp_serve_purpose_col 		= 35   						'Serve a purpose? Are there case notes that are about the task or worker vs. the important case information?
const comp_docs_noted_col 			= 36   						'Were all documents sent and/or received case noted in detail?
const comp_staff_not_meet_note_stndrd_col = 37					'Identify staff not meeting CASE/NOTE standards(comma seperated if more than 1)
const comp_case_note_col 			= 38   						'CASE/NOTE Notes
const comp_reviewer_col 			= 39   						'Reviewer Name
const comp_time_of_case_review_col 	= 40   						'Total time of case review (in minutes)
const comp_repair_required_col 		= 41   						'Case required repair from Cash/SNAP staff
const comp_coaching_col 			= 42   						'If no, is coaching recommeded? (IE: not case noting in detail/using interview script)
const comp_fix_summary_col			= 43						'Summary of repair(s) needed to the case
const comp_file_create_date_col		= 44						'Date the file was created

'Repair report columns
const fix_case_numb_col 			= 01   						'Case #
const fix_appl_date_col 			= 02  						'Application Date
const fix_progs_col 				= 03  						'Programs applied
const fix_spec_cash_prog_col 		= 04  						'Specific Cash Program(s) - If applicable.
const fix_fix_summary_col			= 05				'Summary of repair(s) needed to the case

'simple report columns
const simple_comp_case_numb_col 					= 01   						'Case #
const simple_comp_appl_date_col 					= 02   						'Application Date
const simple_comp_staff_not_meet_appl_stndrd_col 	= 03  				'Identify staff not meeting appplication standards(comma seperated if more than 1)
const simple_comp_appl_notes_col 					= 04   						'Application Notes
const simple_comp_progs_col 						= 05   						'Programs applied
const simple_comp_staff_not_meet_demo_stndrd_col 	= 06  				'Identify staff not meeting demographic standards(comma seperated if more than 1)
const simple_comp_prog_hh_comp_notes_col 			= 07  						'Prog/HH Comp Notes
const simple_comp_intvw_date_col 					= 08   						'Interview Date
const simple_comp_staff_not_meet_intvw_stndrd_col 	= 09 				'Identify staff not meeting interview standards(comma seperated if more than 1)
const simple_comp_intvw_notes_col 					= 10   						'Interview Notes
const simple_comp_verif_req_sent_col 				= 11   						'Was a verification request sent to the case/household?
const simple_comp_staff_not_meet_verif_stndrd_col 	= 12   				'Identify staff not meeting verification standards(comma seperated if more than 1)
const simple_comp_verif_notes_col 					= 13   						'Verification Notes
const simple_comp_snap_det_exp_col 					= 14   						'If SNAP, was the case DETERMINED to be expedited?
const simple_comp_stat_pact_col 					= 15   						'Was STAT/PACT used?
const simple_comp_staff_not_meet_app_stndrd_col 	= 16					'Identify staff not meeting approval standards(comma seperated if more than 1)
const simple_comp_app_notes_col 					= 17   						'Approval Notes:
const simple_comp_serve_purpose_col 				= 18   						'Serve a purpose? Are there case notes that are about the task or worker vs. the important case information?
const simple_comp_docs_noted_col 					= 19   						'Were all documents sent and/or received case noted in detail?
const simple_comp_staff_not_meet_note_stndrd_col 	= 20					'Identify staff not meeting CASE/NOTE standards(comma seperated if more than 1)
const simple_comp_case_note_col 					= 21   						'CASE/NOTE Notes

' ==========================================================================================================================

function random_selection(out_of_number, rand_selected)
	'The out_of_number variable is the chance of selection. For a one in three chance, the out_of_number should be set to 3
	'The selected variable is a boolean of if the option queried should be selected. It will return a one in out_of_number chance of TRUE
	rand_selected = False
	Randomize      		 				'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
	rnd_nbr = rnd						'Create a random number between 0 and 1
	size_up = rnd_nbr * out_of_number	'Multiply by the out-of-number to create a number that is between 0 and the out-of-number (exclusive) - this is a float (decimal number)
	chance_selection = int(size_up)		'Take only the integer of the float from above
	If chance_selection = 0 Then rand_selected = True		'If the integer is 0, (which is a one in out_of_number chance) then the selection is TRUE - we use 0 because there is ALWAYS a 0
end function

' SCRIPT ===================================================================================================================
'Gathering county code for multi-county...
get_county_code

'Connects to BlueZone
EMConnect ""

'Case Sampling Criteria
'Population Selection
select_adults_pop_checkbox = checked
select_families_pop_checkbox = checked
select_1800_pop_checkbox = unchecked
select_DWP_pop_checkbox = unchecked
select_GRH_pop_checkbox = unchecked
select_LTC_pop_checkbox = unchecked
select_YET_pop_checkbox = unchecked

'Days pending criteria
days_pending_limit = "20"

'Review Counts
total_review_count = "60"
select_counts_by_population_checkbox = checked
adult_case_count = ""
families_case_count = ""
eighteen100_case_count = ""
dwp_case_count = ""
grh_case_count = ""
ltc_case_count = ""
yet_case_count = ""

'Program Selection
select_SNAP_checkbox = checked
select_CASH_checkbox = checked
select_MFIP_checkbox = checked
select_DWP_checkbox = unchecked
select_GA_checkbox = checked
select_MSA_checkbox = checked
select_GRH_checkbox = unchecked
select_HC_checkbox = unchecked

'DIALOG HERE
'Dialog should include worker emails who will be processing case sampling. (We can default this)
	'sampling method version selection - default to V2.1
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 276, 225, "ADMIN Pending Case Sampling"
	Text 10, 10, 225, 20, "Case Sampling Selection for Cases That were PENDING yesterday and are No Longer PENDING."
	Text 15, 40, 70, 10, "Functionality to Run:"
	DropListBox 85, 35, 180, 45, "Run ALL Options"+chr(9)+"Compilation Only"+chr(9)+"Case Review Selections Only"+chr(9)+"Make More Review Files"+chr(9)+"Compile Biweekly Report", functionality_choice
	'TODO - add a selection option for the version for the template selection and compilation sheet selection
	Text 10, 55, 125, 10, "Select Cases for Case Sampling"
	GroupBox 10, 70, 75, 120, "Population Selection"
	CheckBox 15, 85, 50, 10, "Adults", select_adults_pop_checkbox
	CheckBox 15, 100, 50, 10, "Families", select_families_pop_checkbox
	' CheckBox 15, 115, 50, 10, "1800", select_1800_pop_checkbox
	' CheckBox 15, 130, 50, 10, "DWP", select_DWP_pop_checkbox
	' CheckBox 15, 145, 50, 10, "GRH", select_GRH_pop_checkbox
	' CheckBox 15, 160, 50, 10, "LTC+", select_LTC_pop_checkbox
	' CheckBox 15, 175, 50, 10, "YET", select_YET_pop_checkbox
	Text 95, 75, 80, 10, "Total Cases to Review:"
	Text 95, 85, 170, 10, "(THIS WILL BE RANDOM AMONG POPULATIONS)"
	EditBox 175, 70, 50, 15, total_review_count
	' CheckBox 95, 100, 165, 10, "Check Here to define counts by population", select_counts_by_population_checkbox
	GroupBox 95, 120, 165, 45, "Program Selection"
	CheckBox 105, 135, 35, 10, "SNAP", select_SNAP_checkbox
	CheckBox 105, 150, 35, 10, "CASH", select_CASH_checkbox
	CheckBox 145, 135, 35, 10, "MFIP", select_MFIP_checkbox
	CheckBox 145, 150, 35, 10, "DWP", select_DWP_checkbox
	CheckBox 185, 135, 35, 10, "GA", select_GA_checkbox
	CheckBox 185, 150, 35, 10, "MSA", select_MSA_checkbox
	' CheckBox 225, 135, 35, 10, "GRH", select_GRH_checkbox
	' CheckBox 225, 150, 35, 10, "HC", select_HC_checkbox
	Text 95, 180, 125, 10, "Review Cases Pending for AT LEAST "
	EditBox 220, 175, 20, 15, days_pending_limit
	Text 245, 180, 25, 10, "days."
	ButtonGroup ButtonPressed
		OkButton 160, 200, 50, 15
		CancelButton 215, 200, 50, 15
		PushButton 10, 200, 85, 15, "Script Information", script_info_btn
EndDialog

'Dialog asks what stats are being pulled
Do
	Do
		err_msg = ""

		Dialog Dialog1
		cancel_without_confirmation


		total_review_count = trim(total_review_count)
		If functionality_choice = "Run ALL Options" OR functionality_choice = "Case Review Selections Only" Then
			If total_review_count <> "" and IsNumeric(total_review_count) = False Then err_msg = err_msg & vbCr & "* Total cases to review needs to be blank or a valid number."
			If days_pending_limit <> "" and IsNumeric(days_pending_limit) = False Then err_msg = err_msg & vbCr & "* Minimum Days pending needs to be blank or a valid number"
		End If
		If ButtonPressed = script_info_btn Then
			err_msg = "LOOP"
			Call word_doc_open(t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Support Documents\Case Sampling Script Information.docx", objWord, objDoc)
		Else
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & "Resolve for the following for the script to continue." & vbcr & err_msg
		End If
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

run_compilation = False
run_review_selection = False
add_more_review_files = False
biweekly_report_run = False
If functionality_choice = "Run ALL Options" OR functionality_choice = "Compilation Only" Then run_compilation = True
If functionality_choice = "Run ALL Options" OR functionality_choice = "Case Review Selections Only" Then run_review_selection = True
If functionality_choice = "Make More Review Files" Then add_more_review_files = True
If functionality_choice = "Compile Biweekly Report" Then biweekly_report_run = True

If run_review_selection = True Then
	If IsNumeric(total_review_count) Then
		total_review_count = total_review_count * 1
	Else
		total_review_count = 0
	End If
	total_pops = 0
	If select_adults_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_families_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_1800_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_DWP_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_GRH_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_LTC_pop_checkbox = checked Then total_pops = total_pops + 1
	If select_YET_pop_checkbox = checked Then total_pops = total_pops + 1


	If select_counts_by_population_checkbox = checked Then
		If total_pops <> 0 and total_review_count <> 0 Then
			reviews_per_pop = ROUND(total_review_count/total_pops) & ""
			If select_adults_pop_checkbox = checked Then adult_case_count = reviews_per_pop
			If select_families_pop_checkbox = checked Then families_case_count = reviews_per_pop
			If select_1800_pop_checkbox = checked Then eighteen100_case_count = reviews_per_pop
			If select_DWP_pop_checkbox = checked Then dwp_case_count = reviews_per_pop
			If select_GRH_pop_checkbox = checked Then grh_case_count = reviews_per_pop
			If select_LTC_pop_checkbox = checked Then ltc_case_count = reviews_per_pop
			If select_YET_pop_checkbox = checked Then yet_case_count = reviews_per_pop
		End If

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 241, 175, "Case Reviews Counts by Population"
			Text 10, 10, 130, 10, "How many Cases for Each Population?"
			If select_adults_pop_checkbox = checked Then
				Text 20, 35, 25, 10, "Adults:"
				EditBox 50, 30, 50, 15, adult_case_count
			End If
			If select_families_pop_checkbox = checked Then
				Text 20, 55, 35, 10, "Families:"
				EditBox 60, 50, 50, 15, families_case_count
			End If
			If select_1800_pop_checkbox = checked Then
				Text 20, 75, 25, 10, "1800:"
				EditBox 45, 70, 50, 15, eighteen100_case_count
			End If
			If select_DWP_pop_checkbox = checked Then
				Text 20, 95, 25, 10, "DWP:"
				EditBox 45, 90, 50, 15, dwp_case_count
			End If
			If select_GRH_pop_checkbox = checked Then
				Text 20, 115, 25, 10, "GRH:"
				EditBox 45, 110, 50, 15, grh_case_count
			End If
			If select_LTC_pop_checkbox = checked Then
				Text 20, 135, 25, 10, "LTC+:"
				EditBox 45, 130, 50, 15, ltc_case_count
			End If
			If select_YET_pop_checkbox = checked Then
				Text 20, 155, 20, 10, "YET:"
				EditBox 45, 150, 50, 15, yet_case_count
			End If
			ButtonGroup ButtonPressed
				OkButton 125, 150, 50, 15
				CancelButton 180, 150, 50, 15
		EndDialog

		Do
			Do
				err_msg = ""

				Dialog Dialog1
				cancel_without_confirmation

				adult_case_count = trim(adult_case_count)
				families_case_count = trim(families_case_count)
				eighteen100_case_count = trim(eighteen100_case_count)
				dwp_case_count = trim(dwp_case_count)
				grh_case_count = trim(grh_case_count)
				ltc_case_count = trim(ltc_case_count)
				yet_case_count = trim(yet_case_count)

				If select_adults_pop_checkbox = checked and IsNumeric(adult_case_count) = False Then err_msg = err_msg & vbCr & "* Adults cases to be reviewed needs to be entered as a valid number."
				If select_families_pop_checkbox = checked and IsNumeric(families_case_count) = False Then err_msg = err_msg & vbCr & "* Families cases to be reviewed needs to be entered as a valid number."
				If select_1800_pop_checkbox = checked and IsNumeric(eighteen100_case_count) = False Then err_msg = err_msg & vbCr & "* 1800 cases to be reviewed needs to be entered as a valid number."
				If select_DWP_pop_checkbox = checked and IsNumeric(dwp_case_count) = False Then err_msg = err_msg & vbCr & "* DWP cases to be reviewed needs to be entered as a valid number."
				If select_GRH_pop_checkbox = checked and IsNumeric(grh_case_count) = False Then err_msg = err_msg & vbCr & "* GRH cases to be reviewed needs to be entered as a valid number."
				If select_LTC_pop_checkbox = checked and IsNumeric(ltc_case_count) = False Then err_msg = err_msg & vbCr & "* LTC+ cases to be reviewed needs to be entered as a valid number."
				If select_YET_pop_checkbox = checked and IsNumeric(yet_case_count) = False Then err_msg = err_msg & vbCr & "* YET cases to be reviewed needs to be entered as a valid number."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & "Resolve for the following for the script to continue." & vbcr & err_msg

			Loop until err_msg = ""
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in

	End If
	If IsNumeric(adult_case_count) = True Then  adult_case_count = adult_case_count*1
	If IsNumeric(families_case_count) = True Then  families_case_count = families_case_count*1
	If IsNumeric(eighteen100_case_count) = True Then  eighteen100_case_count = eighteen100_case_count*1
	If IsNumeric(dwp_case_count) = True Then  dwp_case_count = dwp_case_count*1
	If IsNumeric(grh_case_count) = True Then  grh_case_count = grh_case_count*1
	If IsNumeric(ltc_case_count) = True Then  ltc_case_count = ltc_case_count*1
	If IsNumeric(yet_case_count) = True Then  yet_case_count = yet_case_count*1
End If

const pop_name_const 			= 0
const pop_basket_title_const 	= 1
const pop_selected_const 		= 2
const pop_max_count_const 		= 3
const pop_review_count_const	= 4
const pop_list_count_const		= 5
const pop_table_obj				= 6
const pop_table_name			= 7
const pop_table_style_const		= 8
const highest_const 			= 10

Dim POPULATION_FOR_REVIEWS_ARRAY()
ReDim POPULATION_FOR_REVIEWS_ARRAY(highest_const, 0)
pops_count = 0
ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "Families"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "Families"
If select_families_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_families_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= families_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "FamiliesTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium4"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "Adults"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "Adults"
If select_adults_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_adults_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= adult_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "AdultsTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium2"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "1800"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "1800"
If select_1800_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_1800_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= eighteen100_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "Table1800"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium3"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "DWP"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "DWP"
If select_DWP_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_DWP_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= dwp_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "DWPTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium5"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "Housing Support"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "LTH;FAD GRH;Housing Support"
If select_GRH_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_GRH_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= grh_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "GRHTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium6"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "LTC+"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "LTC+"
If select_LTC_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_LTC_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= ltc_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "LTCTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium7"
pops_count = pops_count + 1

ReDim preserve POPULATION_FOR_REVIEWS_ARRAY(highest_const, pops_count)
POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, pops_count) 			= "YET"
POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, pops_count) 	= "YET"
If select_YET_pop_checkbox = checked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = True
If select_YET_pop_checkbox = unchecked Then POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, pops_count) = False
POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, pops_count) 		= yet_case_count
POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, pops_count)	= 0
POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, pops_count)			= "YETTable"
POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, pops_count)		= "TableStyleMedium1"
pops_count = pops_count + 1

If run_compilation = True Then

	compilation_start_time = timer

	'Open Compilation excel
		'find next empty row
	compilation_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Support Documents\Data Compilation V2 - SNAP Cash Application Sampling.xlsx"
	Call excel_open(compilation_file_path, True, False, ObjExcel, objWorkbook)
	ObjExcel.WorkSheets("V2.1 Cont.").Activate
	'TODO - Make worksheet be a selection in dialog - will need to read the file before the dialog to get the list of worksheets

	excel_row = 2
	Case_numb_string = "*"				'using this string to ensure we don't have any duplicate entries
	Do
		excel_row = excel_row + 1
		compilation_case_numb = trim(ObjExcel.cells(excel_row, 1).Value)
		if compilation_case_numb <> "" Then Case_numb_string = Case_numb_string & compilation_case_numb & "*"
	Loop until compilation_case_numb = ""

	'Loop through each file in case sampling folder with correct name/type
		'open file
		'IF C3 is not Empty
			'enter each item from column C into the next available row on compilation
			'save file path to array so we can move the file
		'IF C3 IS Empty AND it was not created today
			'save file path to array for deletion

	Const reviewer_name_const = 0
	Const reviewer_count_const = 1
	Dim REVIEW_COMPLETE_ARRAY()
	ReDim REVIEW_COMPLETE_ARRAY(1, 0)
	total_reviewers = 0

	case_reviews_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews"

	tally = 0
	files_to_move = ""
	file_name_to_move = ""
	files_to_delete = ""
	file_name_to_delete = ""
	Set objFolder = objFSO.GetFolder(case_reviews_folder)										'Creates an oject of the whole my documents folder
	Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder
	For Each objFile in colFiles																'looping through each file
		move_this_file = False																'Default to NOT delete the file
		this_file_name = objFile.Name															'Grabing the file name
		this_file_type = objFile.Type															'Grabing the file type
		this_file_created_date = objFile.DateCreated											'Reading the date created
		this_file_path = objFile.Path															'Grabing the path for the file

		If this_file_type = "Microsoft Excel Worksheet" and InStr(this_file_name, "Template") = 0 and InStr(this_file_name, "Interview") = 0 and DateDiff("d", this_file_created_date, date) <> 0 Then
			Call excel_open(this_file_path, False, False, ObjREVWExcel, objREVWWorkbook)
			If trim(ObjREVWExcel.cells(3, 3)) = "" Then
				files_to_delete = files_to_delete & this_file_path & "~!~"
				file_name_to_delete = file_name_to_delete & this_file_name & "~!~"
			Else
				files_to_move = files_to_move & this_file_path & "~!~"
				file_name_to_move = file_name_to_move & this_file_name & "~!~"
				this_file_case_numb = "*" & trim(ObjREVWExcel.cells(comp_case_numb_col, 3).Value) & "*"

				If InStr(Case_numb_string, this_file_case_numb) = 0 Then
					ObjExcel.cells(excel_row, comp_case_numb_col).Value 					= ObjREVWExcel.cells(comp_case_numb_col, 3).Value
					ObjExcel.cells(excel_row, comp_case_name_col).Value 					= ObjREVWExcel.cells(comp_case_name_col, 3).Value
					ObjExcel.cells(excel_row, comp_appl_date_col).Value 					= ObjREVWExcel.cells(comp_appl_date_col, 3).Value
					ObjExcel.cells(excel_row, comp_appl_date_issue_col).Value 				= ObjREVWExcel.cells(comp_appl_date_issue_col, 3).Value
					ObjExcel.cells(excel_row, comp_prog_date_align_col).Value 				= ObjREVWExcel.cells(comp_prog_date_align_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_appl_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_appl_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_appl_notes_col).Value 					= ObjREVWExcel.cells(comp_appl_notes_col, 3).Value
					ObjExcel.cells(excel_row, comp_progs_col).Value 						= ObjREVWExcel.cells(comp_progs_col, 3).Value
					ObjExcel.cells(excel_row, comp_spec_cash_prog_col).Value 				= ObjREVWExcel.cells(comp_spec_cash_prog_col, 3).Value
					ObjExcel.cells(excel_row, comp_hh_comp_col).Value 						= ObjREVWExcel.cells(comp_hh_comp_col, 3).Value
					ObjExcel.cells(excel_row, comp_hh_comp_correct_col).Value 				= ObjREVWExcel.cells(comp_hh_comp_correct_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_demo_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_demo_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_prog_hh_comp_notes_col).Value 			= ObjREVWExcel.cells(comp_prog_hh_comp_notes_col, 3).Value
					ObjExcel.cells(excel_row, comp_intvw_date_col).Value 					= ObjREVWExcel.cells(comp_intvw_date_col, 3).Value
					ObjExcel.cells(excel_row, comp_intvw_script_used_col).Value 			= ObjREVWExcel.cells(comp_intvw_script_used_col, 3).Value
					ObjExcel.cells(excel_row, comp_single_intvw_col).Value 					= ObjREVWExcel.cells(comp_single_intvw_col, 3).Value
					ObjExcel.cells(excel_row, comp_mf_orient_complete_col).Value 			= ObjREVWExcel.cells(comp_mf_orient_complete_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_intvw_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_intvw_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_intvw_notes_col).Value 					= ObjREVWExcel.cells(comp_intvw_notes_col, 3).Value
					ObjExcel.cells(excel_row, comp_verif_req_sent_col).Value 				= ObjREVWExcel.cells(comp_verif_req_sent_col, 3).Value
					ObjExcel.cells(excel_row, comp_verif_req_blank_col).Value 				= ObjREVWExcel.cells(comp_verif_req_blank_col, 3).Value
					ObjExcel.cells(excel_row, comp_single_verif_req_col).Value 				= ObjREVWExcel.cells(comp_single_verif_req_col, 3).Value
					ObjExcel.cells(excel_row, comp_spec_forms_req_col).Value 				= ObjREVWExcel.cells(comp_spec_forms_req_col, 3).Value
					ObjExcel.cells(excel_row, comp_unnec_verfs_col).Value 					= ObjREVWExcel.cells(comp_unnec_verfs_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_verif_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_verif_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_verif_notes_col).Value 					= ObjREVWExcel.cells(comp_verif_notes_col, 3).Value
					ObjExcel.cells(excel_row, comp_date_app_col).Value 						= ObjREVWExcel.cells(comp_date_app_col, 3).Value
					ObjExcel.cells(excel_row, comp_snap_det_exp_col).Value 					= ObjREVWExcel.cells(comp_snap_det_exp_col, 3).Value
					ObjExcel.cells(excel_row, comp_same_day_act_col).Value 					= ObjREVWExcel.cells(comp_same_day_act_col, 3).Value
					ObjExcel.cells(excel_row, comp_stat_pact_col).Value 					= ObjREVWExcel.cells(comp_stat_pact_col, 3).Value
					ObjExcel.cells(excel_row, comp_ecf_docs_accept_col).Value 				= ObjREVWExcel.cells(comp_ecf_docs_accept_col, 3).Value
					ObjExcel.cells(excel_row, comp_residnt_in_office_col).Value 			= ObjREVWExcel.cells(comp_residnt_in_office_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_app_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_app_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_app_notes_col).Value 					= ObjREVWExcel.cells(comp_app_notes_col, 3).Value
					ObjExcel.cells(excel_row, comp_serve_purpose_col).Value 				= ObjREVWExcel.cells(comp_serve_purpose_col, 3).Value
					ObjExcel.cells(excel_row, comp_docs_noted_col).Value 					= ObjREVWExcel.cells(comp_docs_noted_col, 3).Value
					ObjExcel.cells(excel_row, comp_staff_not_meet_note_stndrd_col).Value 	= ObjREVWExcel.cells(comp_staff_not_meet_note_stndrd_col, 3).Value
					ObjExcel.cells(excel_row, comp_case_note_col).Value 					= ObjREVWExcel.cells(comp_case_note_col, 3).Value
					ObjExcel.cells(excel_row, comp_reviewer_col).Value 						= trim(ObjREVWExcel.cells(comp_reviewer_col, 3).Value)
					ObjExcel.cells(excel_row, comp_time_of_case_review_col).Value 			= ObjREVWExcel.cells(comp_time_of_case_review_col, 3).Value
					ObjExcel.cells(excel_row, comp_repair_required_col).Value 				= ObjREVWExcel.cells(comp_repair_required_col, 3).Value
					ObjExcel.cells(excel_row, comp_coaching_col).Value 						= ObjREVWExcel.cells(comp_coaching_col, 3).Value
					ObjExcel.cells(excel_row, comp_fix_summary_col).Value 					= ObjREVWExcel.cells(comp_fix_summary_col, 3).Value
					ObjExcel.cells(excel_row, comp_file_create_date_col).Value 				= this_file_created_date								'saving the date the review was created, which is usually the day the review was completed
					excel_row = excel_row + 1
				End If
				reviewer_found = False
				For horse = 0 to UBOUnd(REVIEW_COMPLETE_ARRAY, 2)
					If UCASE(REVIEW_COMPLETE_ARRAY(reviewer_name_const, horse)) = UCASE(trim(ObjREVWExcel.cells(comp_reviewer_col, 3).Value)) Then
						REVIEW_COMPLETE_ARRAY(reviewer_count_const, horse) = REVIEW_COMPLETE_ARRAY(reviewer_count_const, horse) + 1
						reviewer_found = True
					End If
				Next
				If reviewer_found = False Then
					ReDim Preserve REVIEW_COMPLETE_ARRAY(1, total_reviewers)
					REVIEW_COMPLETE_ARRAY(reviewer_name_const, total_reviewers) = trim(ObjREVWExcel.cells(comp_reviewer_col, 3).Value)
					REVIEW_COMPLETE_ARRAY(reviewer_count_const, total_reviewers) = 1
					total_reviewers = total_reviewers + 1
				End If

				tally = tally + 1
			End If
			ObjREVWExcel.ActiveWorkbook.Close
			ObjREVWExcel.Application.Quit
			ObjREVWExcel.Quit
		End If
	Next

	'save the compilation file and close
	objWorkbook.Save()
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

	Set ObjExcel = Nothing
	Set objWorkbook = Nothing

	On Error Resume Next

	files_failed = ""

	If files_to_move <> "" Then
		files_to_move = left(files_to_move, len(files_to_move)-3)
		file_name_to_move = left(file_name_to_move, len(file_name_to_move)-3)
		files_to_move = split(files_to_move, "~!~")
		file_name_to_move = split(file_name_to_move, "~!~")

		archive_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\V2.1 Cases"
		'Loop through file paths to move
			'move
		For cow = 0 to UBound(files_to_move)
			objFSO.MoveFile files_to_move(cow) , archive_folder & "\" & file_name_to_move(cow) & ".xlsx"    'moving each file to the archive file
			' If Err.Number <> 0 Then MsgBox "Error Number: " & Err.Number
			If Err.Number <> 0 Then files_failed = files_failed & file_name_to_move(cow) & "~!~"
			Err.Clear
		Next
	End If

	If files_to_delete <> "" Then
		files_to_delete = left(files_to_delete, len(files_to_delete)-3)
		file_name_to_delete = left(file_name_to_delete, len(file_name_to_delete)-3)
		files_to_delete = split(files_to_delete, "~!~")
		file_name_to_delete = split(file_name_to_delete, "~!~")

		test_folder_for_delete = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\TEST"
		'Loop through file paths to delete
			'delete
		For sheep = 0 to UBound(files_to_delete)
			objFSO.DeleteFile files_to_delete(sheep)						'If we have determined that we need to delete the file - here we delete it
			' If Err.Number <> 0 Then MsgBox "Error Number: " & Err.Number
			If Err.Number <> 0 Then files_failed = files_failed & file_name_to_delete(cow) & "~!~"
			Err.Clear
		Next
	End If

	On Error Goto 0

	' MsgBox "files_failed - " & files_failed
	If files_failed <> "" Then
		files_failed = left(files_failed, len(files_failed)-3)
		If InStr(files_failed, "~!~") <> 0 Then files_failed_array = split(files_failed, "~!~")
		If InStr(files_failed, "~!~") = 0 Then files_failed_array = array(files_failed)
	End If

	compilation_time = timer - compilation_start_time
	compilation_msg = "COMPILATION DETAILS" & vbCr & "Reviews Found:" & vbCr
	For horse = 0 to UBOUnd(REVIEW_COMPLETE_ARRAY, 2)
		compilation_msg = compilation_msg & "Reviewer: " & REVIEW_COMPLETE_ARRAY(reviewer_name_const, horse) & " - Total Reviews: " & REVIEW_COMPLETE_ARRAY(reviewer_count_const, horse) & vbCr
	Next
	If files_failed <> "" Then
		compilation_msg = compilation_msg & vbCr & "SOME FILE(S) COULD NOT BE MOVED OR DELETED."
		compilation_msg = compilation_msg & vbCr & "This is likely because the file is open by another user."
		compilation_msg = compilation_msg & vbCr & "File Name:"
		For each pig in files_failed_array
			compilation_msg = compilation_msg & vbCr & "  - " & pig
		Next
		compilation_msg = compilation_msg & vbCr & "Move or delete the file manually when available. It HAS been logged." & vbCr
	End If
	compilation_min = int(compilation_time/60)
	compilation_sec = compilation_time MOD 60
	compilation_msg = compilation_msg & vbCr &"Compilation took: " & compilation_min & " minutes " & compilation_sec & " seconds."
End If

'Creating the name for the Excel for the Daily List with date as file name
today_yr = DatePart("yyyy", date)
today_day = Right("00"&DatePart("d", date), 2)
today_mo = Right("00"&DatePart("m", date), 2)
today_review_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Daily Case Lists\" & today_yr & "-" & today_mo & "-" & today_day & ".xlsx"


If add_more_review_files = True Then
	more_review_files_start_time = timer
	If FSO.FileExists(today_review_file) Then
		case_count = 0
		Call excel_open(today_review_file, True, False, ObjExcel, objWorkbook)
		'Finding all of the worksheets available in the file. We will likely open up the main 'Review Report' so the script will default to that one.
		For Each objWorkSheet In objWorkbook.Worksheets
			' If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) &
			current_pop = objWorkSheet.Name
			objExcel.worksheets(current_pop).Activate

			excel_row = 2
			Do
				If trim(ObjExcel.Cells(excel_row, 10).Value) <> "" Then
					ReDim Preserve YESTERDAYS_PENDING_CASES_ARRAY(last_pend_array_const, case_count)
					YESTERDAYS_PENDING_CASES_ARRAY(create_review_file_const, case_count) = True
					YESTERDAYS_PENDING_CASES_ARRAY(worker_number_const, case_count) = ObjExcel.Cells(excel_row, 1).Value
					YESTERDAYS_PENDING_CASES_ARRAY(population_const, case_count) 	= ObjExcel.Cells(excel_row, 2).Value
					YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, case_count) 	= ObjExcel.Cells(excel_row, 3).Value
					YESTERDAYS_PENDING_CASES_ARRAY(case_name_const, case_count) 	= ObjExcel.Cells(excel_row, 4).Value
					YESTERDAYS_PENDING_CASES_ARRAY(appl_date_const, case_count) 	= ObjExcel.Cells(excel_row, 5).Value
					YESTERDAYS_PENDING_CASES_ARRAY(days_pending_const, case_count) 	= ObjExcel.Cells(excel_row, 6).Value
					YESTERDAYS_PENDING_CASES_ARRAY(snap_status_const, case_count) 	= ObjExcel.Cells(excel_row, 7).Value
					YESTERDAYS_PENDING_CASES_ARRAY(cash_status_const, case_count) 	= ObjExcel.Cells(excel_row, 8).Value
					YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, case_count) 	= ObjExcel.Cells(excel_row, 9).Value
					case_count = case_count + 1
				End If
				excel_row = excel_row + 1
				next_case_numb = trim(ObjExcel.Cells(excel_row, 3).Value)
			Loop until next_case_numb = ""
		Next
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
	Else
		Call script_end_procedure("The Daily List has not been created today, the option to create MORE files cannot be run until the main Case Sampling Run is complete.")
	End If
End If

If run_review_selection = True Then

	'Starting the query start time (for the query runtime at the end)
	query_start_time = timer

	'Checking for MAXIS
	Call check_for_MAXIS(False)

	'moving files to the archives
	yesterday = DateAdd("d", -1, date)
	call change_date_to_soonest_working_day(yesterday, "BACK")
	For chick = -1 to -4 step -1
		file_date = DateAdd("d", chick, yesterday)
		file_date_yr = DatePart("yyyy", file_date)
		file_date_day = Right("00"&DatePart("d", file_date), 2)
		file_date_mo = Right("00"&DatePart("m", file_date), 2)
		previous_list_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\REPT-PND2 Lists\" & file_date_yr & "-" & file_date_mo & "-" & file_date_day & ".xlsx"
		previous_list_archive = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\REPT-PND2 Lists\PND2 List Archive\" & file_date_yr & "-" & file_date_mo & "-" & file_date_day & ".xlsx"
		If ObjFSO.FileExists(previous_list_file) Then
			ObjFSO.MoveFile previous_list_file, previous_list_archive
		End If

		previous_review_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Daily Case Lists\" & file_date_yr & "-" & file_date_mo & "-" & file_date_day & ".xlsx"
		previous_review_archive = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Daily Case Lists\Daily List Archive\" & file_date_yr & "-" & file_date_mo & "-" & file_date_day & ".xlsx"
		If ObjFSO.FileExists(previous_review_file) Then
			ObjFSO.MoveFile previous_review_file, previous_review_archive
		End If
	Next


	'CREATE LISTS
	'Pull PND2 for the day to array
		'track the largest days pending
		'select all workers - we can change this potentially to be population specific BUT basket management is an issue
		'check cash and snap
	today_yr = DatePart("yyyy", date)
	today_day = Right("00"&DatePart("d", date), 2)
	today_mo = Right("00"&DatePart("m", date), 2)
	today_list_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\REPT-PND2 Lists\" & today_yr & "-" & today_mo & "-" & today_day & ".xlsx"
	If ObjFSO.FileExists(today_list_file) Then

		Call excel_open(today_list_file, True, False, ObjExcel, objWorkbook)

		excel_row = 2
		today_cases = 0
		Do
			ReDim Preserve TODAYS_PENDING_CASES_ARRAY(last_pend_array_const, today_cases)
			TODAYS_PENDING_CASES_ARRAY(worker_number_const, today_cases) = UCase(ObjExcel.Cells(excel_row, 1).Value)
			TODAYS_PENDING_CASES_ARRAY(case_number_const, today_cases) = trim(ObjExcel.Cells(excel_row, 2).Value)
			TODAYS_PENDING_CASES_ARRAY(case_name_const, today_cases) = ObjExcel.Cells(excel_row, 3).Value
			TODAYS_PENDING_CASES_ARRAY(appl_date_const, today_cases) = ObjExcel.Cells(excel_row, 4).Value
			TODAYS_PENDING_CASES_ARRAY(days_pending_const, today_cases) = abs(ObjExcel.Cells(excel_row, 5).Value)
			TODAYS_PENDING_CASES_ARRAY(snap_status_const, today_cases) = ObjExcel.Cells(excel_row, 6).Value
			TODAYS_PENDING_CASES_ARRAY(cash_status_const, today_cases) = ObjExcel.Cells(excel_row, 7).Value
			TODAYS_PENDING_CASES_ARRAY(cash_prog_const, today_cases) = ObjExcel.Cells(excel_row, 8).Value

			TODAYS_PENDING_CASES_ARRAY(appl_date_const, today_cases) = DateAdd("d", 0, TODAYS_PENDING_CASES_ARRAY(appl_date_const, today_cases))
			TODAYS_PENDING_CASES_ARRAY(pending_today_const, today_cases) = True

			excel_row = excel_row + 1
			today_cases = today_cases + 1
			next_worker_numb = trim(ObjExcel.Cells(excel_row, 1).Value)
		Loop until next_worker_numb = ""
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
	Else
		call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

		case_count = 0
		today_longest_pending_days = 0
		'Reading information from REPT/PND2
		For each worker in worker_array
			back_to_self										'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
			Call navigate_to_MAXIS_screen("REPT", "PND2")       'looking at PND2 to confirm day 30 AND look for MSA cases - which get 60 days
			EMWriteScreen worker, 21, 13
			transmit
			'This code is for bypassing a warning box if the basket has too many cases
			EMWaitReady 0, 0
			row = 1
			col = 1
			EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
			If row <> 0 THEN
				transmit
				send_email_to = "tanya.payne@hennepin.us; ilse.ferris@hennepin.us"
				cc_email_to = ""
				email_subject = worker & " AT PND2 DISPLAY LIMIT"
				email_body = "This is a notice that the basket: " & vbCr & worker & vbCr & "reached the display limit and not all cases have been read during CASE SAMPLING." & vbCr & vbCr & "-- SCRIPT AUTOMATED EMAIL"
				Call create_outlook_email("", send_email_to, cc_email_to, "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
			End If


			'Skips workers with no info
			EMReadScreen has_content_check, 6, 3, 73
			If has_content_check <> "0 Of 0" then
				'Grabbing each case number on screen
				Do
					MAXIS_row = 7
					Do
						EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5	'Reading case number
						EMReadScreen client_name, 22, MAXIS_row, 16		'Reading client name
						EMReadScreen APPL_date, 8, MAXIS_row, 38		'Reading application date
						EMReadScreen days_pending, 4, MAXIS_row, 49		'Reading days pending
						EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
						EMReadScreen cash_prog, 2, MAXIS_row, 56		'Reading cash status
						EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status

						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
						client_name = trim(client_name)
						MAXIS_case_number = trim(MAXIS_case_number)
						If client_name <> "ADDITIONAL APP" Then			'When there is an additional app on this rept, the script actually reads a case number even though one is not visible to the worker on the screen - so we are skipping this ghosting issue because it will ALWAYS find the previous case number.
							If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
							all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
						End If

						If MAXIS_case_number = "" AND client_name = "" Then Exit Do			'Exits do if we reach the end

						'Cleaning up each program's status
						SNAP_status = trim(replace(SNAP_status, "_", ""))
						cash_status = trim(replace(cash_status, "_", ""))

						'If additional application is rec'd then the excel output is the client's name, not ADDITIONAL APP
						If client_name <> "ADDITIONAL APP" then
							EMReadScreen next_client, 22, MAXIS_row + 1, 16
							next_client = trim(next_client)
							If next_client = "ADDITIONAL APP" Then
								client_name = "* " & client_name
								MAXIS_row = MAXIS_row + 1
								If SNAP_status = "" Then EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status
								If cash_status = "" Then
									EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
									EMReadScreen cash_prog, 2, MAXIS_row, 56		'Reading cash status
								End If
								'Cleaning up each program's status
								SNAP_status = trim(replace(SNAP_status, "_", ""))
								cash_status = trim(replace(cash_status, "_", ""))
							End If
						End If

						'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
						If SNAP_status <> "" then add_case_info_to_ARRAY = True
						If cash_status <> "" then add_case_info_to_ARRAY = True

						If Trim(APPL_date) = "" then add_case_info_to_ARRAY = False		'If appl date is blank then we don't want to add it. This is due to a MAXIS error.

						If add_case_info_to_ARRAY = True then
							ReDim Preserve TODAYS_PENDING_CASES_ARRAY(last_pend_array_const, case_count)
							TODAYS_PENDING_CASES_ARRAY(case_number_const, case_count)  		= MAXIS_case_number
							TODAYS_PENDING_CASES_ARRAY(worker_number_const, case_count)  	= UCase(worker)
							TODAYS_PENDING_CASES_ARRAY(case_name_const, case_count)  		= client_name
							TODAYS_PENDING_CASES_ARRAY(appl_date_const, case_count)  		= DateAdd("d", 0, replace(APPL_date, " ", "/"))
							TODAYS_PENDING_CASES_ARRAY(days_pending_const, case_count)  	= abs(days_pending)
							TODAYS_PENDING_CASES_ARRAY(cash_status_const, case_count)  		= cash_status
							TODAYS_PENDING_CASES_ARRAY(cash_prog_const, case_count)  		= cash_prog
							TODAYS_PENDING_CASES_ARRAY(snap_status_const, case_count)  		= SNAP_status
							TODAYS_PENDING_CASES_ARRAY(pending_today_const, case_count)  	= True

							If TODAYS_PENDING_CASES_ARRAY(days_pending_const, case_count) > today_longest_pending_days Then today_longest_pending_days = TODAYS_PENDING_CASES_ARRAY(days_pending_const, case_count)

							case_count = case_count + 1
						End if
						MAXIS_row = MAXIS_row + 1
						add_case_info_to_ARRAY = ""	'Blanking out variable
						MAXIS_case_number = ""			'Blanking out variable
					Loop until MAXIS_row = 19
					PF8
					EMReadScreen last_page_check, 21, 24, 2
				Loop until last_page_check = "THIS IS THE LAST PAGE"
			End if
		next


		'Loop pending cases
			'output to excel
			'No additional worksheets should be created - LIST ONLY by oldest to newest

		'Opening the Excel file
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True
		Set objWorkbook = objExcel.Workbooks.Add()
		objExcel.DisplayAlerts = True

		'Changes name of Excel sheet to "Case information"
		ObjExcel.ActiveSheet.Name = "Case information"

		'Setting the first 4 col as worker, case number, name, and APPL date
		ObjExcel.Cells(1, 1).Value = "WORKER"
		objExcel.Cells(1, 1).Font.Bold = TRUE
		ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
		objExcel.Cells(1, 2).Font.Bold = TRUE
		ObjExcel.Cells(1, 3).Value = "NAME"
		objExcel.Cells(1, 3).Font.Bold = TRUE
		ObjExcel.Cells(1, 4).Value = "APPL DATE"
		objExcel.Cells(1, 4).Font.Bold = TRUE
		ObjExcel.Cells(1, 5).Value = "DAYS PENDING"
		objExcel.Cells(1, 5).Font.Bold = TRUE
		snap_pends_col = 6
		ObjExcel.Cells(1, snap_pends_col).Value = "SNAP?"
		objExcel.Cells(1, snap_pends_col).Font.Bold = TRUE
		cash_pends_col = 7
		ObjExcel.Cells(1, cash_pends_col).Value = "CASH?"
		objExcel.Cells(1, cash_pends_col).Font.Bold = TRUE
		cash_prog_col = 8
		ObjExcel.Cells(1, cash_prog_col).Value = "CASH PROG"
		objExcel.Cells(1, cash_prog_col).Font.Bold = TRUE

		excel_row = 2
		For days_pend = today_longest_pending_days to 1 Step -1
			For dog = 0 to UBound(TODAYS_PENDING_CASES_ARRAY, 2)
				If TODAYS_PENDING_CASES_ARRAY(days_pending_const, dog) = days_pend Then
					ObjExcel.Cells(excel_row, 1).Value = TODAYS_PENDING_CASES_ARRAY(worker_number_const, dog)
					ObjExcel.Cells(excel_row, 2).Value = TODAYS_PENDING_CASES_ARRAY(case_number_const, dog)
					ObjExcel.Cells(excel_row, 3).Value = TODAYS_PENDING_CASES_ARRAY(case_name_const, dog)
					ObjExcel.Cells(excel_row, 4).Value = TODAYS_PENDING_CASES_ARRAY(appl_date_const, dog)
					ObjExcel.Cells(excel_row, 5).Value = TODAYS_PENDING_CASES_ARRAY(days_pending_const, dog)
					ObjExcel.Cells(excel_row, snap_pends_col).Value = TODAYS_PENDING_CASES_ARRAY(snap_status_const, dog)
					ObjExcel.Cells(excel_row, cash_pends_col).Value = TODAYS_PENDING_CASES_ARRAY(cash_status_const, dog)
					ObjExcel.Cells(excel_row, cash_prog_col).Value = TODAYS_PENDING_CASES_ARRAY(cash_prog_const, dog)
					excel_row = excel_row + 1
				End If
			Next
		Next

		'Autofitting columns
		For col_to_autofit = 1 to 8
			ObjExcel.columns(col_to_autofit).AutoFit()
		Next

		'save PND2 Excel with date as file name and close - potential restart process uses this file to fill the array
		objExcel.ActiveWorkbook.SaveAs today_list_file
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
	End If

	' call script_end_procedure("PND2 from today RUN!")

	'open previous days PND2 excel
		'pull all cases into an array
		'Default pending today to FALSE
		'close excel
	yesterday = DateAdd("d", -1, date)
	call change_date_to_soonest_working_day(yesterday, "BACK")
	yestdy_yr = DatePart("yyyy", yesterday)
	yestdy_day = Right("00"&DatePart("d", yesterday), 2)
	yestdy_mo = Right("00"&DatePart("m", yesterday), 2)
	yesterday_list_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\REPT-PND2 Lists\" & yestdy_yr & "-" & yestdy_mo & "-" & yestdy_day & ".xlsx"
	Call excel_open(yesterday_list_file, True, False, ObjExcel, objWorkbook)


	excel_row = 2
	yest_cases = 0
	If days_pending_limit = "" Then days_pending_limit = 0
	days_pending_limit = days_pending_limit-1
	Do
		ReDim Preserve YESTERDAYS_PENDING_CASES_ARRAY(last_pend_array_const, yest_cases)
		YESTERDAYS_PENDING_CASES_ARRAY(worker_number_const, yest_cases) = UCase(ObjExcel.Cells(excel_row, 1).Value)
		YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, yest_cases) = trim(ObjExcel.Cells(excel_row, 2).Value)
		YESTERDAYS_PENDING_CASES_ARRAY(case_name_const, yest_cases) = ObjExcel.Cells(excel_row, 3).Value
		YESTERDAYS_PENDING_CASES_ARRAY(appl_date_const, yest_cases) = ObjExcel.Cells(excel_row, 4).Value
		YESTERDAYS_PENDING_CASES_ARRAY(days_pending_const, yest_cases) = abs(ObjExcel.Cells(excel_row, 5).Value)
		YESTERDAYS_PENDING_CASES_ARRAY(snap_status_const, yest_cases) = ObjExcel.Cells(excel_row, 6).Value
		YESTERDAYS_PENDING_CASES_ARRAY(cash_status_const, yest_cases) = ObjExcel.Cells(excel_row, 7).Value
		YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, yest_cases) = trim(ObjExcel.Cells(excel_row, 8).Value)

		YESTERDAYS_PENDING_CASES_ARRAY(appl_date_const, yest_cases) = DateAdd("d", 0, YESTERDAYS_PENDING_CASES_ARRAY(appl_date_const, yest_cases))
		YESTERDAYS_PENDING_CASES_ARRAY(pending_today_const, yest_cases) = False
		YESTERDAYS_PENDING_CASES_ARRAY(create_review_file_const, yest_cases) = False
		YESTERDAYS_PENDING_CASES_ARRAY(on_daily_list_const, grape) = False

		excel_row = excel_row + 1
		yest_cases = yest_cases + 1
		next_pending_days = trim(ObjExcel.Cells(excel_row, 5).Value)
		If next_pending_days = "" Then Exit Do
		next_pending_days = abs(next_pending_days)
	Loop until next_pending_days < days_pending_limit
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

	Set ObjExcel = Nothing
	Set objWorkbook = Nothing

	'Loop through yesterday's cases
		'Loop through todays cases
			'If found - mark yesterdays' cases pending today as TRUE
		'if not found and days pending of 19 days or over - use functionality from ILSE Test Script - ADMIN - ADD POPULATION to fill population in column
	For duck = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
		case_found = False
		For grape = 0 to UBound(TODAYS_PENDING_CASES_ARRAY, 2)
			If YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, duck) = TODAYS_PENDING_CASES_ARRAY(case_number_const, grape) Then
				YESTERDAYS_PENDING_CASES_ARRAY(pending_today_const, duck) = True
				case_found = True
				Exit For
			End If
		Next
		If case_found = False Then
			population = ""
			If basket_detail.Exists(YESTERDAYS_PENDING_CASES_ARRAY(worker_number_const, duck)) Then
				population = basket_detail.Item(YESTERDAYS_PENDING_CASES_ARRAY(worker_number_const, duck))
			Else
				population = "UNKNOWN"
			End If
			YESTERDAYS_PENDING_CASES_ARRAY(population_const, duck) = population

			If select_counts_by_population_checkbox = unchecked and IsNumeric(total_review_count) Then
				If total_review_count > 0 Then
				End If
			End If

		End If
	Next

	'Creating the Daily List
	first_loop = True
	first_run = True

	'This allows for the script to be run twice and update the existing Daily List
	If ObjFSO.FileExists(today_review_file) Then
		Call excel_open(today_review_file, True, False, ObjExcel, objWorkbook)
		first_loop = False
		first_run = False
	Else
		'open new excel
		'Opening the Excel file
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True
		Set objWorkbook = objExcel.Workbooks.Add()
		objExcel.DisplayAlerts = True
	End If

	'Loop through yesterday's cases - twice - once for Adults and Once for Families - no other populations are considered
		'create worksheet for population
		'Add to list if pending today is FALSE and DAYS pending is 19 days OR OVER AND there cash or snap have a non 'I' status
			'NOTE that the scope is for 20 days pending or more BUT this is yesterdays list so we go from 19 and up
		'add population column (B)
		'add 1 to days pending
		'add date of review column (H) - NO LONGER NEEDED
		'Save with today's date as file name in 'QI\Case Reviews\Daily Case Lists
		'close
	For duck = 0 to UBound(POPULATION_FOR_REVIEWS_ARRAY, 2)
		If POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, duck) = True Then
			excel_row = 1
			sheet_exists = False
			If first_run = False Then								'look for an existing sheet in case it is run a second time in the day
				For Each objWorkSheet In objWorkbook.Worksheets
					If objWorkSheet.Name = POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, duck) Then
						sheet_exists = True
						objExcel.worksheets(objWorkSheet.Name).Activate
						Do
							excel_row = excel_row + 1
							info_here = ObjExcel.Cells(excel_row, 1).Value
						Loop until info_here = ""
						excel_row = excel_row - 1
						Exit For
					End If
				Next
			End If

			If sheet_exists = False Then
				If first_loop = True Then
					'Changes name of Excel sheet to "Families"
					ObjExcel.ActiveSheet.Name = POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, duck)
					first_loop = False
				Else
					'NOW adults list
					ObjExcel.Worksheets.Add().Name = POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, duck)
				End If

				ObjExcel.Cells(1, 1).Value = "WORKER"
				ObjExcel.Cells(1, 2).Value = "Population"
				ObjExcel.Cells(1, 3).Value = "CASE NUMBER"
				ObjExcel.Cells(1, 4).Value = "NAME"
				ObjExcel.Cells(1, 5).Value = "APPL DATE"
				ObjExcel.Cells(1, 6).Value = "DAYS PENDING"
				ObjExcel.Cells(1, 7).Value = "SNAP?"
				ObjExcel.Cells(1, 8).Value = "CASH?"
				ObjExcel.Cells(1, 9).Value = "CASH TYPE"
				ObjExcel.Cells(1, 10).Value = "Date of Review"
			End If

			For grape = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
				review_case = False
				If select_MFIP_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape) = "MF" Then review_case = True
				If select_DWP_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape) = "DW" Then review_case = True
				If select_GA_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape) = "GA" Then review_case = True
				If select_MSA_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape) = "MS" Then review_case = True
				If select_CASH_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape) = "CA" Then review_case = True
				' If select_CASH_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(cash_status_const, grape) <> "" Then review_case = True
				If select_SNAP_checkbox = checked and YESTERDAYS_PENDING_CASES_ARRAY(snap_status_const, grape) <> "" Then review_case = True
				If YESTERDAYS_PENDING_CASES_ARRAY(pending_today_const, grape) = True Then review_case = False
				If YESTERDAYS_PENDING_CASES_ARRAY(population_const, grape) <> POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, duck) Then review_case = False
				If review_case = True Then
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					YESTERDAYS_PENDING_CASES_ARRAY(on_daily_list_const, grape) = True
					POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck) = POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck) + 1
					excel_row = excel_row + 1
					ObjExcel.Cells(excel_row, 1).Value = YESTERDAYS_PENDING_CASES_ARRAY(worker_number_const, grape)
					ObjExcel.Cells(excel_row, 2).Value = YESTERDAYS_PENDING_CASES_ARRAY(population_const, grape)
					ObjExcel.Cells(excel_row, 3).Value = YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, grape)
					ObjExcel.Cells(excel_row, 4).Value = YESTERDAYS_PENDING_CASES_ARRAY(case_name_const, grape)
					ObjExcel.Cells(excel_row, 5).Value = YESTERDAYS_PENDING_CASES_ARRAY(appl_date_const, grape)
					ObjExcel.Cells(excel_row, 6).Value = YESTERDAYS_PENDING_CASES_ARRAY(days_pending_const, grape)
					ObjExcel.Cells(excel_row, 7).Value = YESTERDAYS_PENDING_CASES_ARRAY(snap_status_const, grape)
					ObjExcel.Cells(excel_row, 8).Value = YESTERDAYS_PENDING_CASES_ARRAY(cash_status_const, grape)
					ObjExcel.Cells(excel_row, 9).Value = YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, grape)
				End If
			Next

			'This section is to determine if we should create a review worksheet, This is a random selection based on how many reviews should be completed and how many are possible
			make_random_selection = True																'default to using a random selection
			If IsNumeric(POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck)) = False Then			'if a number was not entered in the beginning dialog, we will not create a random selection
				make_random_selection = False
			Else																						'if the total possible reviews is less than or equal to the max from the dialog there is no random selection needed
				If POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck) =< POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck) Then make_random_selection = False
			End If

			If make_random_selection = True Then														'when we need to make a random selection, we need to determine the correct chance for randomization
				'This creates a chance number that is likely to select close to the right number of cases (if there are 90 possible cases and we only need 30, this will produce a chance number of 3, giving a 1 in 3 chance.)
				chance_number = POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck)/POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck)
				chance_number = FormatNumber(chance_number, 2, -1, 0, -1)
			End If

			'Now we start back at the beginning of the excel to enter when there is a review
			excel_row = 1
			For grape = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
				'only select cases for the population we are reviewing on this loop and were determined to be on the daily list
				If YESTERDAYS_PENDING_CASES_ARRAY(on_daily_list_const, grape) = True  and YESTERDAYS_PENDING_CASES_ARRAY(population_const, grape) = POPULATION_FOR_REVIEWS_ARRAY(pop_basket_title_const, duck) Then
					cases_left = POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck) - (excel_row-1)			'figure out how many cases are sill on the list to evaluate so we  can tell if we need to adjust strategy to meet the max need.
					excel_row = excel_row + 1

					create_review_file = True																		'default to creating a review file
					If make_random_selection = True Then															'if we are randomizing, we will:
						'If the selected review cases plus the cases still to assess is MORE than the max, we will still randomize, this means if selected plus remaining is equal to or less than the max, we leave the review to TRUE and all remaining will be selected
						If cases_left + POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) > POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck) Then
							'if the number of cases selected is equal to or more than the max, we will NOT select the case for review
							If POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) >= POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck) Then create_review_file = False
							'as long as the selected cases is under the max we will call the randomization function, which will change the review selection boolean based on the randomization and the chance number
							If POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) < POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck) Then call random_selection(chance_number, create_review_file)
						End If
					Else																							'if we are NOT randomizing:
						If IsNumeric(POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck)) Then					'if a max slection is indicated we set the selection to false if the selected reviews has met the max
							If POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) >= POPULATION_FOR_REVIEWS_ARRAY(pop_max_count_const, duck) Then create_review_file = False
						End If																						'this means if the max count was blank, we will always default to selecting every file
					End If
					YESTERDAYS_PENDING_CASES_ARRAY(create_review_file_const, grape) = create_review_file			'Add the boolean to the cases array for the next loop when we actually create the files
					If create_review_file = True Then																'next we increment the selected count and add the date to the daily list for the case that is selected
						POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) = POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck) + 1
						ObjExcel.Cells(excel_row, 10).Value = date
					End If
				End if
			Next

			'Autofitting columns
			For col_to_autofit = 1 to 8
				ObjExcel.columns(col_to_autofit).AutoFit()
			Next

			'Creating a table
			POPULATION_FOR_REVIEWS_ARRAY(pop_table_obj, duck) = "A1:J" & excel_row
			ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, POPULATION_FOR_REVIEWS_ARRAY(pop_table_obj, duck), xlYes).Name = POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, duck)
			ObjExcel.ActiveSheet.ListObjects(POPULATION_FOR_REVIEWS_ARRAY(pop_table_name, duck)).TableStyle = POPULATION_FOR_REVIEWS_ARRAY(pop_table_style_const, duck)
		End If
	Next

	objExcel.ActiveWorkbook.SaveAs today_review_file
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

	Set ObjExcel = Nothing
	Set objWorkbook = Nothing
	list_ready_time = time
End If

If add_more_review_files = True or run_review_selection = True Then
	'Create the sampling file for each case
		'naming convention - CASENUMBER PROGS - POPULATION
	review_template_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Template - Application Case Sampling V2.xlsx"

	review_files_created = 0
	For duck = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
		If YESTERDAYS_PENDING_CASES_ARRAY(create_review_file_const, duck) = True Then
			review_file_name = YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, duck) & " "
			If YESTERDAYS_PENDING_CASES_ARRAY(snap_status_const, duck) <> "" Then review_file_name = review_file_name & "SNAP-"
			If YESTERDAYS_PENDING_CASES_ARRAY(cash_status_const, duck) <> "" Then
				cash_prog = "CASH"
				If YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, duck) = "MF" Then cash_prog = "MFIP"
				If YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, duck) = "MS" Then cash_prog = "MSA"
				If YESTERDAYS_PENDING_CASES_ARRAY(cash_prog_const, duck) = "GA" Then cash_prog = "GA"
				review_file_name = review_file_name & cash_prog
			End If
			If Right(review_file_name, 1) = "-" Then review_file_name = left(review_file_name, len(review_file_name)-1)
			review_file_name = review_file_name & " - " & YESTERDAYS_PENDING_CASES_ARRAY(population_const, duck) & ".xlsx"
			' MsgBox "review_file_name - " & review_file_name & vbcr & "EXISTS - " & ObjFSO.FileExists(t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\" & review_file_name)
			If not ObjFSO.FileExists(t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\" & review_file_name) Then
				Call excel_open(review_template_file, False, False, ObjExcel, objWorkbook)
				ObjExcel.Cells(1, 3).Value = YESTERDAYS_PENDING_CASES_ARRAY(case_number_const, duck)

				case_name = replace(YESTERDAYS_PENDING_CASES_ARRAY(case_name_const, duck), "*", "")
				If InStr(case_name, ",") Then
					case_name_array = split(case_name, ",")
					case_name = trim(case_name_array(1)) & " " & trim(case_name_array(0))
				End If
				ObjExcel.Cells(2, 3).Value = case_name

				objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\" & review_file_name
				objExcel.ActiveWorkbook.Close
				objExcel.Application.Quit
				objExcel.Quit

				Set ObjExcel = Nothing
				Set objWorkbook = Nothing
				review_files_created = review_files_created + 1
				' MsgBox "review_files_created - " & review_files_created
			End If
			cash_prog = ""
		End If
	Next
	If run_review_selection = True Then
		review_selection_run_time = timer-query_start_time
		revw_select_run_min = int(review_selection_run_time/60)
		revw_select_run_sec = review_selection_run_time MOD 60

		review_select_msg = "CASE REVIEW SELECTION" & vbCr & "Case selections:"
		For duck = 0 to UBound(POPULATION_FOR_REVIEWS_ARRAY, 2)
			If POPULATION_FOR_REVIEWS_ARRAY(pop_selected_const, duck) = True Then
				review_select_msg = review_select_msg & vbCr & POPULATION_FOR_REVIEWS_ARRAY(pop_name_const, duck) & ": - Total possible Reviews: " & POPULATION_FOR_REVIEWS_ARRAY(pop_list_count_const, duck) & vbCr & " - Review Files Made: " & POPULATION_FOR_REVIEWS_ARRAY(pop_review_count_const, duck)
			End If
		Next
		review_select_msg = review_select_msg & vbCr & "List was ready at " & list_ready_time
		review_select_msg = review_select_msg & vbCr & "(Review file creation started at this time.)"
		review_select_msg = review_select_msg & vbCr & vbCr & "Review creation took: " & revw_select_run_min & " minutes " & revw_select_run_sec & " seconds."
	End If
	If add_more_review_files = True Then
		files_only_run_time = timer-more_review_files_start_time
		files_only_run_min = int(files_only_run_time/60)
		files_only_run_sec = files_only_run_time MOD 60
		review_select_msg = review_select_msg & vbCr & "Review Files Created: " & review_files_created
		review_select_msg = review_select_msg & vbCr & vbCr & "File Only creation took: " & files_only_run_min & " minutes " & files_only_run_sec & " seconds."
	End If
End If

If biweekly_report_run = True Then
	'Setting some variables and file paths
	curr_date_file_format = date
	curr_date_file_format = replace(curr_date_file_format, "/", "-")
	share_point_folder 		= "https://hennepin.sharepoint.com/teams/hs-economic-supports-management/Shared%20Documents/Case%20Sampling/"
	fix_file_path 			= share_point_folder & "Repair%20List%20-%20" & curr_date_file_format & ".xlsx"
	simple_comp_file_path 	= share_point_folder & "Sampling%20List%20-%20" & curr_date_file_format & ".xlsx"

	'These are the template files for making new files
	fix_file_template 			= share_point_folder & "Repair%20List%20TEMPLATE.xlsx"
	simple_comp_file_template	= share_point_folder & "Sampling%20List%20TEMPLATE.xlsx"

	'Open the template files and save them as the specific date file name
	Call excel_open(fix_file_template, True, False, ObjFixExcel, objFixWorkbook)				'Repair List
	ObjFixExcel.ActiveWorkbook.SaveAs fix_file_path

	Call excel_open(simple_comp_file_template, True, False, ObjSimpExcel, objSimpWorkbook)		'Full Review List
	ObjSimpExcel.ActiveWorkbook.SaveAs simple_comp_file_path

	'Open Compilation excel to read the review details
	compilation_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Case Reviews\Support Documents\Data Compilation V2 - SNAP Cash Application Sampling.xlsx"
	Call excel_open(compilation_file_path, True, False, ObjExcel, objWorkbook)
	ObjExcel.WorkSheets("V2.1 Cont.").Activate		'get to the right sheet

	'Setting the start of a loop to read every line on the compilation list.
	'All cases will be added to the simplified review report (ObjSimpExcel)
	'Only cases with a 'Yes' in the column for need repair will be added to the repair report (ObjFixExcel)
	excel_row = 3			'rows start at 3 because there are two header rows
	fix_excel_row = 3
	earliest_date = date	'these dates are to determine the beginning and end date of the review completion dates
	latest_date = #1/1/2025#
	Do
		ObjSimpExcel.cells(excel_row, simple_comp_case_numb_col).Value 						= ObjExcel.cells(excel_row, comp_case_numb_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_appl_date_col).Value 						= ObjExcel.cells(excel_row, comp_appl_date_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_appl_stndrd_col).Value 	= ObjExcel.cells(excel_row, comp_staff_not_meet_appl_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_appl_notes_col).Value 					= ObjExcel.cells(excel_row, comp_appl_notes_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_progs_col).Value 							= ObjExcel.cells(excel_row, comp_progs_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_demo_stndrd_col).Value 	= ObjExcel.cells(excel_row, comp_staff_not_meet_demo_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_prog_hh_comp_notes_col).Value 			= ObjExcel.cells(excel_row, comp_prog_hh_comp_notes_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_intvw_date_col).Value 					= ObjExcel.cells(excel_row, comp_intvw_date_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_intvw_stndrd_col).Value 	= ObjExcel.cells(excel_row, comp_staff_not_meet_intvw_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_intvw_notes_col).Value 					= ObjExcel.cells(excel_row, comp_intvw_notes_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_verif_req_sent_col).Value 				= ObjExcel.cells(excel_row, comp_verif_req_sent_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_verif_stndrd_col).Value 	= ObjExcel.cells(excel_row, comp_staff_not_meet_verif_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_verif_notes_col).Value 					= ObjExcel.cells(excel_row, comp_verif_notes_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_snap_det_exp_col).Value 					= ObjExcel.cells(excel_row, comp_snap_det_exp_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_stat_pact_col).Value 						= ObjExcel.cells(excel_row, comp_stat_pact_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_app_stndrd_col).Value 		= ObjExcel.cells(excel_row, comp_staff_not_meet_app_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_app_notes_col).Value 						= ObjExcel.cells(excel_row, comp_app_notes_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_serve_purpose_col).Value 					= ObjExcel.cells(excel_row, comp_serve_purpose_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_docs_noted_col).Value 					= ObjExcel.cells(excel_row, comp_docs_noted_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_staff_not_meet_note_stndrd_col).Value 	= ObjExcel.cells(excel_row, comp_staff_not_meet_note_stndrd_col).Value
		ObjSimpExcel.cells(excel_row, simple_comp_case_note_col).Value 						= ObjExcel.cells(excel_row, comp_case_note_col).Value

		If trim(ObjExcel.cells(excel_row, comp_repair_required_col).Value) = "Yes" Then		'These are repair cases
			ObjFixExcel.cells(fix_excel_row, fix_case_numb_col).Value 						= ObjExcel.cells(excel_row, comp_case_numb_col).Value
			ObjFixExcel.cells(fix_excel_row, fix_appl_date_col).Value 						= ObjExcel.cells(excel_row, comp_appl_date_col).Value
			ObjFixExcel.cells(fix_excel_row, fix_progs_col).Value 							= ObjExcel.cells(excel_row, comp_progs_col).Value
			ObjFixExcel.cells(fix_excel_row, fix_spec_cash_prog_col).Value 					= ObjExcel.cells(excel_row, comp_spec_cash_prog_col).Value
			ObjFixExcel.cells(fix_excel_row, fix_fix_summary_col).Value 					= ObjExcel.cells(excel_row, comp_fix_summary_col).Value
			fix_excel_row = fix_excel_row + 1
		End If

		'Determining the earliest and latest file dates as the date reviews were completed
		file_date = trim(ObjExcel.cells(excel_row, comp_file_create_date_col).Value)
		If trim(file_date) <> "" Then
			If IsDate(file_date) Then
				file_date = DateAdd("d", 0, file_date)
				If DateDiff("d", file_date, earliest_date) > 0 Then earliest_date = file_date
				If DateDiff("d", latest_date, file_date) > 0 Then latest_date = file_date
			End If
		End If
		excel_row = excel_row + 1

	Loop until trim(ObjExcel.cells(excel_row, comp_case_numb_col).Value) = ""

	'save the report files and close
	objFixWorkbook.Save()
	ObjFixExcel.ActiveWorkbook.Close
	ObjFixExcel.Application.Quit
	ObjFixExcel.Quit

	objSimpWorkbook.Save()
	ObjSimpExcel.ActiveWorkbook.Close
	ObjSimpExcel.Application.Quit
	ObjSimpExcel.Quit

	'This part creates a worksheet in the compilation file for the already captured reviews and blanks out the main compilation sheet for the next biweekly period
	'Find the last existing worksheet on the compilation workbook
	last_worksheet = ""
	For Each objWorkSheet In objWorkbook.Worksheets
		If objWorkSheet.Name <> "V.2 Comp" Then last_worksheet = objWorkSheet.Name
	Next
	'Create a name for the worksheet that was just created
	new_worksheek_name = DatePart("m", earliest_date) & "-" & DatePart("d", earliest_date) & " - " & DatePart("m", latest_date) & "-" & DatePart("d", latest_date)

	'Copy the existing main sheet to the end of the sheets and then rename the sheet to the name created with dates
	objExcel.worksheets("V2.1 Cont.").Copy, objExcel.worksheets(last_worksheet)
	ObjExcel.WorkSheets("V2.1 Cont. (2)").Activate
	ObjExcel.ActiveSheet.Name = new_worksheek_name

	'Go back to the main list
	ObjExcel.WorkSheets("V2.1 Cont.").Activate

	'Email that the reports are ready
	send_email_to = "Jennifer.Frey@hennepin.us"
	cc_email_to = "Tanya.Payne@hennepin.us"
	email_subject = "Case Sampling Biweekly Report"
	email_body = "The Case Sampling reports for the past 2 weeks have been compiled." & "<br>" & "<br>"
	email_body = email_body & "The reports can be found here:" & "<br>"
	email_body = email_body & "&emsp;" & "- " & "<a href=" & chr(34) & fix_file_path & chr(34) & ">" & "Repair List - " & curr_date_file_format & "</a>" & " Case that appear to need to be fixed." & "<br>"
	email_body = email_body & "&emsp;" & "- " & "<a href=" & chr(34) & simple_comp_file_path & chr(34) & ">" & "Sampling List - " & curr_date_file_format & "</a>" & " Full list from the past 2 weeks (" & earliest_date & " - " & latest_date & ")." & "<br>" & "<br>"
	email_body = email_body & "Worklists created by a BlueZone script to support case sampling."

	Call create_outlook_email("", send_email_to, cc_email_to, "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)

	'Now we blank out the reviews on the main list since they are saved on the copy sheet.
	'This keeps this workbook quicker/more responsive
	excel_row = 3
	Do
		ObjExcel.cells(excel_row, comp_case_numb_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_case_name_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_appl_date_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_appl_date_issue_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_prog_date_align_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_appl_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_appl_notes_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_progs_col).Value 						= ""
		ObjExcel.cells(excel_row, comp_spec_cash_prog_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_hh_comp_col).Value 						= ""
		ObjExcel.cells(excel_row, comp_hh_comp_correct_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_demo_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_prog_hh_comp_notes_col).Value 			= ""
		ObjExcel.cells(excel_row, comp_intvw_date_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_intvw_script_used_col).Value 			= ""
		ObjExcel.cells(excel_row, comp_single_intvw_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_mf_orient_complete_col).Value 			= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_intvw_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_intvw_notes_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_verif_req_sent_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_verif_req_blank_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_single_verif_req_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_spec_forms_req_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_unnec_verfs_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_verif_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_verif_notes_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_date_app_col).Value 						= ""
		ObjExcel.cells(excel_row, comp_snap_det_exp_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_same_day_act_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_stat_pact_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_ecf_docs_accept_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_residnt_in_office_col).Value 			= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_app_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_app_notes_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_serve_purpose_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_docs_noted_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_staff_not_meet_note_stndrd_col).Value 	= ""
		ObjExcel.cells(excel_row, comp_case_note_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_reviewer_col).Value 						= ""
		ObjExcel.cells(excel_row, comp_time_of_case_review_col).Value 			= ""
		ObjExcel.cells(excel_row, comp_repair_required_col).Value 				= ""
		ObjExcel.cells(excel_row, comp_coaching_col).Value 						= ""
		ObjExcel.cells(excel_row, comp_fix_summary_col).Value 					= ""
		ObjExcel.cells(excel_row, comp_file_create_date_col).Value 				= ""
		excel_row = excel_row + 1
	Loop until trim(ObjExcel.cells(excel_row, comp_case_numb_col).Value) = ""

	'save and close the compilation file
	objWorkbook.Save()
	ObjExcel.ActiveWorkbook.Close
	ObjExcel.Application.Quit
	ObjExcel.Quit

	'end message output - note that this can be run on an interval other than biweekly the only strict biweekly outputs are in these hard coded strings that do not impact functionality.
	compilation_msg = "BiWeekly reports created:" & vbCr & "Repair List - " & curr_date_file_format & vbCr & "Sampling List - " & curr_date_file_format
End If

done_msg = "Script Run Completed" & vbCr & vbCr
done_msg = done_msg & compilation_msg & vbCr & vbCr
done_msg = done_msg & review_select_msg & vbCr
done_msg = done_msg & vbCr & "All files managed."
call script_end_procedure(done_msg)



'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------3/13/25
'--Tab orders reviewed & confirmed----------------------------------------------3/13/25
'--Mandatory fields all present & Reviewed--------------------------------------3/13/25
'--All variables in dialog match mandatory fields-------------------------------3/13/25
'Review dialog names for content and content fit in dialog----------------------3/13/25
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------3/13/25
'--Create a button to reference instructions------------------------------------N/A							Instructions are in project folder
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------3/13/25
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------N/A
'--BULK - review output of statistics and run time/count (if applicable)--------3/13/25
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------3/13/25
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------3/13/25
'--Incrementors reviewed (if necessary)-----------------------------------------3/13/25
'--Denomination reviewed -------------------------------------------------------3/13/25
'--Script name reviewed---------------------------------------------------------3/13/25
'--BULK - remove 1 incrementor at end of script reviewed------------------------3/13/25

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------3/13/25
'--comment Code-----------------------------------------------------------------3/13/25
'--Update Changelog for release/update------------------------------------------3/13/25
'--Remove testing message boxes-------------------------------------------------3/13/25
'--Remove testing code/unnecessary code-----------------------------------------3/13/25
'--Review/update SharePoint instructions----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------3/13/25
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------3/14/25
'--Update project team/issue contact (if applicable)----------------------------3/16/25
