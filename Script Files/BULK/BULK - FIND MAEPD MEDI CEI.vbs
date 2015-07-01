'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - FIND MAEPD MEDI CEI.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

FUNCTION navigate_to_MMIS
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	'The following will select the correct version of MMIS. First it looks for C302, then EK01, then C402.
	row = 1
	col = 1
	EMSearch "C302", row, col
	If row <> 0 then 
		If row <> 1 then 'It has to do this in case the worker only has one option (as many LTC and OSA workers don't have the option to decide between MAXIS and MCRE case access). The MMIS screen will show the text, but it's in the first row in these instances.
			EMWriteScreen "x", row, 4
			transmit
		End if
	Else 'Some staff may only have EK01 (MMIS MCRE). The script will allow workers to use that if applicable.
		row = 1
		col = 1
		EMSearch "EK01", row, col
		If row <> 0 then 
			If row <> 1 then
				EMWriteScreen "x", row, 4
				transmit
			End if
		Else 'Some OSAs have C402 (limited access). This will search for that.
			row = 1
			col = 1
			EMSearch "C402", row, col
			If row <> 0 then 
				If row <> 1 then
					EMWriteScreen "x", row, 4
					transmit
				End if
			Else 'Some OSAs have EKIQ (limited MCRE access). This will search for that.
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 then 
					If row <> 1 then
						EMWriteScreen "x", row, 4
						transmit
					End if
				Else
					script_end_procedure("C402, C302, EKIQ, or EK01 not found. Your access to MMIS may be limited. Contact your script Alpha user if you have questions about using this script.")
				End if
			End if
		End if
	END IF

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit
END FUNCTION

FUNCTION navigate_to_MAXIS(maxis_mode)
	attn
	EMConnect "A"
	IF maxis_mode = "PRODUCTION" THEN
		EMReadScreen prod_running, 7, 6, 15
		IF prod_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect"B"
			attn
			EMReadScreen prod_running, 7, 6, 15
			IF prod_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than 2.")
			END IF
		END IF
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMReadScreen inq_running, 7, 7, 15
		IF inq_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect "B"
			attn
			EMReadScreen inq_running, 7, 7, 15
			IF inq_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than 2.")
			END IF
		END IF
	END IF

	
	EMConnect (x)
	IF maxis_mode = "PRODUCTION" THEN
		EMWriteScreen "1", 2, 15
		transmit
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMWriteScreen "2", 2, 15
		transmit
	END IF		
END FUNCTION

BeginDialog maepd_dlg, 0, 0, 191, 110, "MA-EPD Reimburseables"
  EditBox 85, 10, 65, 15, x_number
  ButtonGroup ButtonPressed
    OkButton 85, 90, 50, 15
    CancelButton 135, 90, 50, 15
  Text 10, 15, 70, 10, "X Number:"
  Text 10, 30, 175, 10, "This script will check REPT/ACTV on this X number."
EndDialog


EMConnect ""

CALL check_for_MAXIS(True)
DO
	err_msg = ""
	Dialog maepd_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF len(x_number) <> 7 THEN err_msg = err_msg & vbCr & "* The X number must be the full 7 digits."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_MAXIS(False)
back_to_SELF

CALL navigate_to_MAXIS_screen("REPT", "ACTV")
EMReadScreen current_rept_actv, 7, 21, 13
IF ucase(current_rept_actv) <> ucase(x_number) THEN CALL write_value_and_transmit(x_number, 21, 13)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True
objExcel.Cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Value = "CLIENT NAME"
objExcel.Cells(1, 3).Value = "NEXT REVW DT"
objExcel.Cells(1, 4).Value = "REIMBURSEMENT ELIG?"

rept_row = 7
excel_row = 2
DO
	EMReadScreen last_page, 21, 24, 2
	DO
		EMReadScreen case_number, 8, rept_row, 12
		case_number = trim(case_number)
		EMReadScreen hc_case, 1, rept_row, 64
		IF case_number <> "" AND hc_case <> " " THEN 
			objExcel.Cells(excel_row, 1).Value = case_number
			EMReadScreen client_name, 21, rept_row, 21
			client_name = trim(client_name)
			EMReadScreen next_revw_dt, 8, rept_row, 42
			next_revw_dt = replace(next_revw_dt, " ", "/")
			objExcel.Cells(excel_row, 2).Value = client_name
			objExcel.Cells(excel_row, 3).Value = next_revw_dt	
			excel_row = excel_row + 1
		END IF
		rept_row = rept_row + 1
	LOOP UNTIL rept_row = 19
	PF8
	rept_row = 7
LOOP UNTIL last_page = "THIS IS THE LAST PAGE"

excel_row = 2
DO
	back_to_SELF
	case_number = objExcel.Cells(excel_row, 1).Value
	CALL find_variable("Environment: ", production_or_inquiry, 10)
	CALL navigate_to_screen("ELIG", "HC")
	hhmm_row = 8
	DO
		EMReadScreen hc_type, 2, hhmm_row, 28
		IF hc_type = "MA" THEN
			EMWriteScreen "X", hhmm_row, 26
			transmit
			EMReadScreen elig_type, 2, 12, 72
			IF elig_type = "DP" THEN
				EMWriteScreen "X", 9, 76
				transmit
				EMReadScreen pct_fpg, 4, 18, 38
				pct_fpg = trim(pct_fpg)
				pct_fpg = pct_fpg * 1
				IF pct_fpg < 201 THEN
					PF3
					PF3
					EMReadScreen hh_memb_num, 2, hhmm_row, 3
					CALL navigate_to_screen("STAT", "MEMB")
					ERRR_screen_check
					EMWriteScreen hh_memb_num, 20, 76
					transmit
					EMReadScreen cl_pmi, 8, 4, 46
					cl_pmi = replace(cl_pmi, " ", "")
					DO
						IF len(cl_pmi) <> 8 THEN cl_pmi = "0" & cl_pmi
					LOOP UNTIL len(cl_pmi) = 8
					navigate_to_MMIS
					DO
						EMReadScreen RKEY, 4, 1, 52
						IF RKEY <> "RKEY" THEN EMWaitReady 0, 0
					LOOP UNTIL RKEY = "RKEY"
					EMWriteScreen "I", 2, 19
					EMWriteScreen cl_pmi, 4, 19
					transmit
					EMWriteScreen "RELG", 1, 8
					transmit
			
					'Reading RELG to determine if the CL is active on MA-EPD		
					EMReadScreen prog01_type, 8, 6, 13
						EMReadScreen elig01_type, 2, 6, 33
						EMReadScreen elig01_end, 8, 7, 36
					EMReadScreen prog02_type, 8, 10, 13
						EMReadScreen elig02_type, 2, 10, 33
						EMReadScreen elig02_end, 8, 11, 36
					EMReadScreen prog03_type, 8, 14, 13
						EMReadScreen elig03_type, 2, 14, 33
						EMReadScreen elig03_end, 8, 15, 36
					EMReadScreen prog04_type, 8, 18, 13
						EMReadScreen elig04_type, 2, 18, 33
						EMReadScreen elig04_end, 8, 19, 36

					IF ((prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR _
						(prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR _
						(prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99/99/99") OR _
						(prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99")) THEN
			
						EMWriteScreen "RMCR", 1, 8
						transmit

						'-----CHECKING FOR ON-GOING MEDICARE PART B-----
						EMReadScreen part_b_begin01, 8, 13, 4
							part_b_begin01 = trim(part_b_begin01)
						EMReadScreen part_b_end01, 8, 13, 15
						EMReadScreen part_b_begin02, 8, 14, 4
							part_b_begin02 = trim(part_b_begin02)
						EMReadScreen part_b_end02, 8, 14, 15
						
						IF (part_b_begin01 <> "" AND part_b_end01 = "99/99/99") THEN		
							EMWriteScreen "RBYB", 1, 8
							transmit
							
							EMReadScreen accrete_date, 8, 5, 66
							EMReadScreen delete_date, 8, 6, 65
							accrete_date = replace(accrete_date, " ", "")

							IF ((accrete_date = "") OR (accrete_date <> "" AND delete_date <> "99/99/99")) THEN
								objExcel.Cells(excel_row, 4).Value = objExcel.Cells(excel_row, 4).Value & ("MEMB " & hh_memb_num & " ELIG FOR REIMBURSEMENT, ")
							END IF
							PF3
						END IF
					ELSE
						PF3
					END IF
					CALL navigate_to_MAXIS(production_or_inquiry)
					hhmm_row = hhmm_row + 1
					CALL navigate_to_screen("ELIG", "HC")
				ELSE
					DO
						EMReadScreen at_hhmm, 4, 3, 51
						IF at_hhmm <> "HHMM" THEN PF3
					LOOP UNTIL at_hhmm = "HHMM"
					hhmm_row = hhmm_row + 1
				END IF
			ELSE
				PF3
				hhmm_row = hhmm_row + 1
			END IF
		ELSE
			hhmm_row = hhmm_row + 1
		END IF
		IF hhmm_row = 20 THEN
			PF8
			EMReadScreen this_is_the_last_page, 21, 24, 2
		END IF
	LOOP UNTIL hc_type = "  " OR this_is_the_last_page = "THIS IS THE LAST PAGE"
	'Deleting the blank results to clean up the spreadsheet
	IF objExcel.Cells(excel_row, 4).Value = "" THEN
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	END IF		
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""

FOR i = 1 to 4
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success!!")

