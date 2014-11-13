'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES

EMConnect ""

'IMIG----------------------------------------------------------------------------------------------------------------------------------
imigration_status = "21"
entry_date = "11/13/2013"
status_date = "11/13/2013"
status_ver = "OT"
status_LPR_adj_from = ""
nationality = "MX"

Function write_panel_to_MAXIS_IMIG(imigration_status, entry_date, status_date, status_ver, status_LPR_adj_from, nationality)
	call create_MAXIS_friendly_date(date, 0, 5, 45)					'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", date), 5, 51
	EMWriteScreen imigration_status, 6, 45							'Writes imig status
	call create_MAXIS_friendly_date(entry_date, 0, 7, 45)			'Enters year as a 2 digit number, so have to modify manually
	EMWriteScreen datepart("yyyy", entry_date), 7, 51
	call create_MAXIS_friendly_date(status_date, 0, 7, 71)			'Enters year as a 2 digit number, so have to modify manually
	EMWriteScreen datepart("yyyy", status_date), 7, 77
	EMWriteScreen status_ver, 8, 45									'Enters status ver
	EMWriteScreen status_LPR_adj_from, 9, 45						'Enters status LPR adj from
	EMWriteScreen nationality, 10, 45								'Enters nationality
	transmit
	transmit
End function

'call write_panel_to_MAXIS_IMIG(imigration_status, entry_date, status_date, status_ver, status_LPR_adj_from, nationality)

'SPON----------------------------------------------------------------------------------------------------------------------------------
SPON_type = "IL"
SPON_ver = "Y"
SPON_name = "Bob Loblaw"
SPON_state = "MN"

Function write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
	EMWriteScreen SPON_type, 6, 38
	EMWriteScreen SPON_ver, 6, 62
	EMWriteScreen SPON_name, 8, 38
	EMWriteScreen SPON_state, 10, 62
	transmit
End function

'call write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)

'DSTT----------------------------------------------------------------------------------------------------------------------------------
DSTT_ongoing_income = "Y"
HH_income_stop_date = "11/13/2014"
income_expected_amt = "300"

Function write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, HH_income_stop_date, income_expected_amt)
	EMWriteScreen DSTT_ongoing_income, 6, 69
	call create_MAXIS_friendly_date(HH_income_stop_date, 0, 9, 69)
	EMWriteScreen income_expected_amt, 12, 71
End function

'call write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, HH_income_stop_date, income_expected_amt)

