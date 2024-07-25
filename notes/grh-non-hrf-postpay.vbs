'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - GRH - NON-HRF-POSTPAY.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 90                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each case
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
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/30/2016", "Case Note title changed to indicate GRH payment.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'FUNCTION edition Block. Need to added this customized navigation FUNCTION=============================================================

'This function checks and compares most active Faci VND address to clients current ADDR. then declares a value to be put into the dialog variant and casenote. Will be called during faci screening
FUNCTION vnd_addr_check
	Call navigate_to_MAXIS_screen ("STAT", "ADDR")
	EMReadScreen addrpnl_address, 22, 6, 43
	addrpnl_address = replace(addrpnl_address, "_","")
	Call navigate_to_MAXIS_screen ("MONY", "VNDS")
	EMWriteScreen faci_vndnumber, 04, 59
	Transmit
	EMreadScreen faci_address, 22, 5, 15
	faci_address = replace(faci_address, "_","")
	If faci_address = addrpnl_address then
		compare_addr_vnds = ", clt's ADDR is the 'SAME.'"
	Else
		compare_addr_vnds = ", clt's ADDR is 'DIFFERENT.'"
	End If
	faci_date_out = "Out Date: " & faci_date_out
	addr_faci_vnds_status = faci_location & faci_date_out & vnd_end_date_footer & compare_addr_vnds
	PF3
	PF3
END FUNCTION

'this function checks pben IAA dates and determines the variant IAA_status
FUNCTION pben_check_IAA_dates
EMReadScreen pben_line_01, 8, 8, 66
	EMReadScreen pben_line_02, 8, 9, 66
	EMReadScreen pben_line_03, 8, 10, 66
	EMReadScreen pben_line_04, 8, 11, 66
	EMReadScreen pben_line_05, 8, 12, 66
	EMReadScreen pben_line_06, 8, 13, 66
	if pben_line_02 = "__ __ __" then pben_line_02 = pben_line_01
	if pben_line_03 = "__ __ __" then pben_line_03 = pben_line_02
	if pben_line_04 = "__ __ __" then pben_line_04 = pben_line_03
	if pben_line_05 = "__ __ __" then pben_line_05 = pben_line_04
	if pben_line_06 = "__ __ __" then pben_line_06 = pben_line_05
	pben_line_01 = replace(pben_line_01, " ","/")
	pben_line_02 = replace(pben_line_02, " ","/")
	pben_line_03 = replace(pben_line_03, " ","/")
	pben_line_04 = replace(pben_line_04, " ","/")
	pben_line_05 = replace(pben_line_05, " ","/")
	pben_line_06 = replace(pben_line_06, " ","/")
	'Determines if all the IAA dates are more than 11 months old to current month. if so, will declare IAA status variant that IAA are expired
	If DateDiff("M", pben_line_01, date()) > 11 and DateDiff("M", pben_line_02, date()) > 11 and DateDiff("M", pben_line_03, date()) > 11 and DateDiff("M", pben_line_04, date()) > 11 and DateDiff("M", pben_line_05, date()) > 11 and DateDiff("M", pben_line_06, date()) > 11 then
		IAA_status = "IAA are expired"
	Else
		IAA_status = "IAA are within 12 months."
	End If
END FUNCTION


'End of Customized FUNCTION BLOCK===================================================================================================


'The Script=========================================================================================================================

EMConnect ""
EMFocus

call check_for_MAXIS(False)	'checking for an active MAXIS session
'Create string of FACI discharge dates to determine most recent discharge date
faci_discharge_dates = ""

'Grabbing case number and putting in the month and year entered from dialog box.
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'Delcares the variable GRH_process_date = footer month/01/year. this is needed to check if FACI outdates for postpay are in the processing footer month/year. If end dates matches processing footer month/year, workers may need to process post pay for that footer month/year.
GRH_process_date = Maxis_footer_month & "/" & "01" & "/" & MAXIS_footer_year
MAXIS_footer_month_confirmation									'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'First Dialog that asks for case number and footer month.
BeginDialog Dialog1, 0, 0, 311, 100, "PostPay Non-HRF"
  EditBox 90, 5, 65, 15, MAXIS_case_number
  EditBox 105, 30, 20, 15, MAXIS_footer_month
  EditBox 135, 30, 20, 15, MAXIS_footer_year
  EditBox 70, 60, 95, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 105, 80, 50, 15
	CancelButton 160, 80, 50, 15
  Text 60, 10, 25, 10, "Case #:"
  Text 130, 35, 5, 10, "/"
  Text 15, 35, 80, 10, "PostPay month (mm/yy):"
  Text 5, 65, 65, 10, "Worker's Signature:"
  Text 175, 15, 120, 10, "This script is for NON-HRF PostPay."
  Text 175, 25, 125, 10, "It will go through the following panels:"
  Text 185, 40, 20, 10, "* FACI"
  Text 185, 50, 30, 10, "* ADDR"
  Text 185, 60, 25, 10, "* JOBS"
  Text 240, 40, 25, 10, "* UNEA"
  Text 240, 50, 25, 10, "* PBEN"
  Text 240, 60, 50, 10, "* VNDS"
  GroupBox 170, 5, 135, 70, "Description:"
EndDialog
'First Dialog. Showing case number, postpay month & year...checking for valid entries of these info.  It'll loop until workers enter the right condition.
Do
	DO
	    err_msg = ""
	    Dialog Dialog1
	    cancel_confirmation
	    Call validate_MAXIS_case_number(err_msg, "*")
		If MAXIS_footer_month = "" OR len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbCr & "You must enter a valid month value of: MM"
	    If MAXIS_footer_year = "" OR len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbCr & "You must enter a valid year value of: YY"
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'navigating to FACI panel. reads if there are FACI panel or not. If none, then the script stop and closes active background excel sheets
CALL navigate_to_MAXIS_screen ("STAT", "FACI")
EMReadScreen faci_pnls, 1, 2, 78			'counts faci pnls
IF faci_pnls = "0" then                     'if none
	script_end_procedure ("Script will end here.  There is no active facility panel created.  Please manually review client status and facility needs.")
End If
'Faci pnls exists will determine any active post pay facilities.
For i = 1 to faci_pnls
		EMWriteScreen "0" & i, 20, 79
		transmit
		EMReadScreen faci_postpay, 1, 13, 71
		IF faci_postpay = "Y" then
			For maxis_row = 14 to 18
				EMReadScreen faci_date_in, 10, maxis_row, 47
				EMReadScreen faci_date_out, 10, maxis_row, 71
				'If finds a faci with a start date but no enddate, then it declares that faci as the most active pnl
				IF faci_date_in <> "__ __ ____" AND faci_date_out = "__ __ ____" then
					faci_postpay_status = true
					EMReadScreen faci_location, 30, 6, 43
					faci_location = replace(faci_location, "_","") & ", "
					EMReadScreen faci_vndnumber, 8, 5, 43
					faci_date_out = "(none) - client is still active"   	'declares no out dates therefore client is still active at current faci
					vnd_addr_check                     				'function to pull mony vnds and compares clt's addr then declares variant to dialog/casenote
					Exit For
				End If
				'If there are no open ended dates, then script formats the date and plots it to hidden spread sheet, later to be recalled and find most recent faci discharged date (the variant will be: Last_Faci_OutDate)
				c_row = maxis_row - 13
				faci_date_out = replace(faci_date_out, " ","/")
				faci_discharge_dates = faci_discharge_dates & faci_date_out & "*" 
			Next
		ElseIf faci_postpay_status <> true then
			addr_faci_vnds_status = "There are no Post Pay facility pnls in this case."     'If no FACI consist of "Y" Post Pay indication, addr_faci_vnds_status variant is declared
		End If
Next

'Remove asterisk from end of list
faci_discharge_dates = left(faci_discharge_dates, len(faci_discharge_dates) - 1)
'Create an array from FACI discharge dates
faci_discharge_dates_array = split(faci_discharge_dates, "*")
'Sort the dates from oldest to newest 
Call sort_dates(faci_discharge_dates_array)
'Identify the newest date
Last_Faci_OutDate = faci_discharge_dates_array(ubound(faci_discharge_dates_array))

'if above faci screening shows that there are none still active, the script will focus on the one with the most recent discharged date.
If faci_date_out <> "(none) - client is still active" Then
	For i = 1 to faci_pnls
		EMWriteScreen "0" & i, 20, 79
		transmit
		EMReadScreen faci_postpay, 1, 13, 71
		IF faci_postpay = "Y" then
			For maxis_row = 14 to 18
				EMReadScreen faci_date_out, 10, maxis_row, 71
				faci_date_out = replace(faci_date_out, " ","/")
				IF faci_date_out = Last_Faci_OutDate then				'once it finds the matching date in that faci pnls. script delcares variant of the current vnd/faci info
					EMReadScreen faci_location, 30, 6, 43
					faci_location = replace(faci_location, "_","") & ", "
					EMReadScreen faci_vndnumber, 8, 5, 43
					If month(Last_Faci_OutDate) = month(GRH_process_date) and year(Last_Faci_OutDate) = year(GRH_process_date) Then    'also checks and declares a variant to see if discharged date is within the same month of GRH_process_date (footer month/year)
						vnd_end_date_footer = ""
					Else
						vnd_end_date_footer = ", can't do post pay for " & GRH_process_date
					End If
					vnd_addr_check		'address/vnd comparison function to declare variant to dialog and casenotes
					Exit For
				End If
			Next
		ElseIf faci_postpay_status <> true then
				addr_faci_vnds_status = "There are no Post Pay facility pnls in this case. "     'If no FACI consist of "Y" Post Pay indication, addr_faci_vnds_status variant is declared
		End If
	Next
End If

'Checks Active Job'
MAXIS_background_check
CALL navigate_to_MAXIS_screen("STAT", "JOBS")
EMReadScreen jobs_pnls, 1, 2, 78  'reads the panel 0 of 0, reading the total panel value..if it's zero or 1+
If jobs_pnls <> "0" then
  For jobs_to_review = 1 to jobs_pnls
    EMWriteScreen "0" & jobs_to_review, 20, 79
	  transmit
    EMReadScreen jobs_hrs_end, 3, 18, 72
			IF jobs_hrs_end <> "___" Then
				jobs_status = "There is an active Job pnl. POSSIBLE HRF process?"
			Else
				jobs_status = "no active jobs"
      End If
		Next
Else
	jobs_status = "no active jobs"
End If

'Checks Active BUSI'
CALL navigate_to_MAXIS_screen("STAT", "BUSI")
EMReadScreen busi_pnls, 1, 2, 78  'reads the panel 0 of 0, reading the total panel value..if it's zero or 1+
If busi_pnls <> "0" then
  For busi_to_review = 1 to busi_pnls
    EMWriteScreen "0" & busi_to_review, 20, 79
	  transmit
    EMReadScreen busi_hrs_end, 3, 13, 74
			IF busi_hrs_end <> "___" Then
				busi_status = "There is an active BUSI pnl. POSSIBLE HRF process?"
			Else
				busi_status = "no active BUSI."
      End If
		Next
Else
	busi_status = "no active BUSI."
End If
'declares variant earnincome_status for dialog/case notes
earnincome_status = jobs_status & ", " & busi_status

'checks UNEA and types of UNEA'
Call MAXIS_case_number_finder(MAXIS_case_number)
Call navigate_to_MAXIS_screen ("STAT", "UNEA")
EMReadScreen unea_pnls, 1, 2, 78		'counts how many active unea pnls
Dim unea_list()					'Dims variable to make an array list of existing UNEA pnls
ReDim unea_list(unea_pnls)
u = 1							'using dummy variables to declare array number. used to combine variables together for the dialog '
r = 0							'using dummy variables to declare array number. used to combine variables together for the dialog '
unea_list(0) = ""
If unea_pnls <> "0" then
  For i = 1 to unea_pnls
    EMWriteScreen "0" & i, 20, 79
	transmit
	EMReadScreen unea_type, 2, 5, 37
	EMReadScreen unea_name, 19, 5, 40
	unea_name = replace(unea_name, " ","")
	If unea_type = "01" or unea_type = "02" or unea_type = "03" then		'filters to see if disability unea types are active: SSI/RSDI/RSDI DISA
		SSA_income = TRUE
		EMReadScreen unea_end, 8, 7, 68
		'If unea pnls are still active with no end date then it builds an array variant to list all active disa unea pnls
		If unea_end = "__ __ __" then
			EMReadScreen unea_amt, 8, 18, 68
			unea_amt = replace(unea_amt, " ","")
      		unea_list(u) = unea_name & " $" & unea_amt & ", " & unea_list(r)			'This array is built/delcares the unea info into variant, then next unea info will be added unto the previous variant each "For" stmt
      		u = u + 1
      		r = r + 1
			IAA_status = "Disability UNEA benefits exists, therefore no updates needed"	'Will declare IAA_status variant since UNEA disa exist
			unea_status = unea_list(r)									'Delcares unea_status the dialog variant and added to the array variant, will be updated with more amts each "For" stmt
		'If unea pnl have an end date, however same as footer month, then income counts for processing post pay for that footer month. repeats with same array logic above
		ElseIF unea_end <> "__ __ __" and month(GRH_process_date) = month(replace(unea_end, " ","/")) and year(GRH_process_date) = year(replace(unea_end, " ","/")) then
			EMReadScreen unea_amt, 8, 18, 68
			unea_amt = replace(unea_amt, " ","")
			unea_list(u) = unea_name & " $" & unea_amt & ", " & unea_list(r)
			u = u + 1
			r = r + 1
			IAA_status = "Disability UNEA benefits exists, however it will end soon on " & replace(unea_end, " ","/")
			unea_status = unea_list(r) & " --> stops this footer month on " & replace(unea_end, " ","/")
		'If unea_pnl have no end date, then unea_status is declared the amt info --> however income ended on (unea_end variant). repeat same array logic above
		ElseIf unea_end <> "__ __ __" then
			EMReadScreen unea_amt, 8, 18, 68
			unea_amt = replace(unea_amt, " ","")
			unea_list(u) = unea_name & " $" & unea_amt & ", " & unea_list(r)
			u = u + 1
			r = r + 1
			IAA_status = "Disability UNEA pnl exists, however it ended on " & replace(unea_end, " ","/")
			unea_status = unea_list(r) & " --> however income ended on " & replace(unea_end, " ","/")
		End If
	Else
		'If there are UNEA but no 'disa' UNEA pnls, then script will look at pben info, finds if IAA are expired or are still within 12 months
		If SSA_income <> TRUE then
			unea_status = "UNEA income exsist, but not disability UNEA."
			call navigate_to_MAXIS_screen("STAT", "PBEN")
			EMReadScreen pben_pnls, 1, 2, 78
			If pben_pnls <> 0 then
				pben_check_IAA_dates
			Else
				'if no IAA pnls exist then variant is declared below
				IAA_status = "There are no PBEN pnls"
				End If
			End If
		End If
	Next
End If

'If no unea pnls, the script checks pben again and declares variants.
If unea_pnls = "0" then
	unea_status = "There are no unea pnls"
	call navigate_to_MAXIS_screen("STAT", "PBEN")
	EMReadScreen pben_pnls, 1, 2, 78
	If pben_pnls <> 0 then
		pben_check_IAA_dates
	Else
		'if there are no IAA pnls and Unea pnls then variants are declared below
		IAA_status = "There are no PBEN pnls. Possible Disability referral"
		unea_status = "No UNEA pnls. Possible Disability referral"
	End If
End If

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'Second Dialog when all info has been grab from case will be called into fields/variants to be reviewed by worker
BeginDialog Dialog1, 0, 0, 456, 275, "GRH NON-HRF CASE NOTE dialog"
  EditBox 80, 5, 365, 15, addr_faci_vnds_status
  EditBox 80, 60, 245, 15, IAA_status
  EditBox 80, 80, 245, 15, earnincome_status
  EditBox 80, 100, 365, 15, unea_status
  EditBox 80, 120, 365, 15, other_notes
  EditBox 80, 140, 365, 15, changes
  EditBox 80, 160, 365, 15, verifs_needed
  EditBox 80, 180, 365, 15, actions_taken
  EditBox 10, 210, 290, 15, Postpay_results
  EditBox 375, 230, 70, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 340, 255, 50, 15
	CancelButton 395, 255, 50, 15
	PushButton 80, 35, 25, 10, "VNDS", VNDS_button
	PushButton 105, 35, 25, 10, "FACI", FACI_button
	PushButton 130, 35, 25, 10, "ADDR", ADDR_button
	PushButton 275, 35, 25, 10, "BUSI", BUSI_button
	PushButton 300, 35, 25, 10, "JOBS", JOBS_button
	PushButton 325, 35, 25, 10, "UNEA", UNEA_button
	PushButton 365, 35, 45, 10, "prev. panel", prev_panel_button
	PushButton 365, 45, 45, 10, "next panel", next_panel_button
	PushButton 340, 75, 25, 10, "MEMB", MEMB_button
	PushButton 365, 75, 25, 10, "MEMI", MEMI_button
	PushButton 390, 75, 25, 10, "REVW", REVW_button
	PushButton 415, 75, 25, 10, "PBEN", PBEN_button
	PushButton 10, 230, 290, 15, "Send case to BGTX", CASE_BGTX
	PushButton 315, 210, 20, 10, "GRH", ELIG_GRH_button
	PushButton 335, 210, 20, 10, "HC", ELIG_HC_button
  Text 35, 65, 40, 10, "IAA Status:"
  Text 30, 125, 40, 10, "Other notes:"
  Text 35, 145, 35, 10, "Changes?:"
  Text 25, 165, 50, 10, "Verifs needed:"
  Text 25, 185, 50, 10, "Actions taken:"
  GroupBox 5, 200, 300, 50, "Post Payment Results"
  Text 310, 235, 60, 10, "Worker Signature:"
  GroupBox 310, 200, 50, 25, "ELIG panels:"
  Text 10, 105, 70, 10, "Active Disa/UNEA?:"
  Text 5, 85, 70, 10, "Earn Income Status:"
  Text 5, 10, 75, 10, "Recent(PostPay)Faci: "
  GroupBox 270, 25, 85, 25, "Income panels"
  GroupBox 75, 25, 85, 25, "Locations"
  GroupBox 360, 25, 85, 35, "STAT-based navigation:"
  GroupBox 335, 65, 110, 25, "other STAT panels:"
EndDialog

'Initiates last dialog: GRH_case_note_dialog
DO
	DO
		DO
			DO
				DO
				 	Dialog Dialog1
					cancel_confirmation
				LOOP UNTIL ButtonPressed <> no_cancel_button
				MAXIS_dialog_navigation
				'Goes to MONY VNDS screen using the most active faci vnd number on case... If EMReadScreen does not read any FACI pnls. MsgBox there are no faci pnls.
				If ButtonPressed = VNDS_button then
					If faci_pnls = "0" then
						MsgBox "There Are No Facility panels"
					Else
						call navigate_to_MAXIS_screen("MONY", "VNDS")
						EMWriteScreen faci_vndnumber, 04, 59
						Transmit
					End If
				End If
				'Button sends case to BGTX. Waits for MAXIS comes back from BG. Then brings Post Pay results into dialog variant
				If ButtonPressed = CASE_BGTX then
					call navigate_to_MAXIS_screen("stat", "memb")
						EMWriteScreen "BGTX", 20, 71  'sending case through background.
						transmit
						MAXIS_background_check
						call navigate_to_MAXIS_screen("elig", "grh")
						EMReadScreen GRPR_check, 4, 3, 47
						If GRPR_check <> "GRPR" then
							MsgBox "The script couldn't find ELIG/GRH. It will now jump to case note."
							Else
							EMWriteScreen "GRSM", 20, 71
						End If
						transmit
					'reads elig/grh info from GRSM for inputting into dialog and case note.
						If GRPR_check = "GRPR" then
							EMReadScreen GRSM_vnd, 9, 10, 31
							GRSM_vnd = replace(GRSM_vnd, " ","")
						End If
						If GRPR_check = "GRPR" then
							EMReadScreen GRSM_payable, 9, 12, 31
							GRSM_payable = replace(GRSM_payable, " ","")
						End If
						If GRPR_check = "GRPR" then
							EMReadScreen GRSM_Obligation, 9, 18, 31
							GRSM_Obligation = replace(GRSM_Obligation, " ","")
						End If
					'Declares variable post pay results for variant and case note
					Postpay_results = "Vendor#: " & GRSM_vnd & ", Payable Amount: $" & GRSM_payable & ", Client Obligation: $" & GRSM_Obligation
				End If
			LOOP UNTIL ButtonPressed = -1 OR ButtonPressed = previous_button
			err_msg = ""
			IF addr_faci_vnds_status = "" THEN err_msg = err_msg & vbCr & "* You must indicate a facility status within the 'Recent(Post Pay)Faci' field."
			IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* Please indicate the actions you have taken."
			IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = "" OR ButtonPressed = previous_button
	LOOP WHILE ButtonPressed = previous_button
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'GRH NON HRF CASE NOTE
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(faci_location & "GRH PAYMENT FOR " & GRH_process_date)      'need to work on how to tell script to stop short of the full listing values indicated, and only stop short of the FACI's name.
call write_variable_in_CASE_NOTE("---------")
call write_bullet_and_variable_in_CASE_NOTE("Most recent FACI/ADDR info", addr_faci_vnds_status)
If PostPay_results <> "" then call write_bullet_and_variable_in_CASE_NOTE("PostPay Results", PostPay_results)
call write_variable_in_CASE_NOTE("---------")
If IAA_status <> "" then call write_bullet_and_variable_in_CASE_NOTE("IAA Status", IAA_status)
If jobs_status <> "" then call write_bullet_and_variable_in_CASE_NOTE("Earn Income Status", earnincome_status)
If unea_status <> "" then call write_bullet_and_variable_in_CASE_NOTE("Active Disa/UNEA?", unea_status)
If other_notes <> "" then call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
If changes <> "" then call write_bullet_and_variable_in_CASE_NOTE("Changes Report", changes)
If verifs_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Verifications needed", verifs_needed)
If actions_taken <> "" then call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE("---------")
call write_variable_in_CASE_NOTE(worker_signature)

'reminding workers to go back to fill in the items that may have left to be fill during the first run.
call script_end_procedure("Success!!! The script will stop here.  Please remember to review, fill-in, postpay code and approved from ELIG results screen if needed."& VbCrLf & VbCrLf &"Thank you!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
