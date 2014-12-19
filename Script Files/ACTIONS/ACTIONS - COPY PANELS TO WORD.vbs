'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - COPY PANELS TO WORD.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'----------------------------------------------------------------------------------------------------
'ADD TO FUNCTIONS FILE WHEN GITHUB IS WORKING AGAIN
Function copy_screen_to_array(output_array)
	output_array = "" 'resetting array
	Dim screenarray(23)	'24 line array
	row = 1
	For each line in screenarray
		EMReadScreen reading_line, 80, row, 1
		output_array = output_array & reading_line & "UUDDLRLRBA"
		row = row + 1
	Next
	output_array = split(output_array, "UUDDLRLRBA")
End function

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
all_possible_panels = "MEMB MEMI ADDR AREP ALTP ALIA TYPE PROG HCRE PARE SIBL EATS IMIG SPON FACI FCFC FCPL ADME REMO DISA ABPS PREG STRK STWK SCHL WREG EMPS CASH ACCT SECU CARS REST OTHR TRAN STIN STEC PBEN UNEA LUMP RBIC BUSI JOBS TRAC DSTT DCEX WKEX COEX SHEL HEST ACUT PDED PACT FMED ACCI MEDI INSA DIET DISQ SWKR REVW MISC RESI TIME EMMA BILS HCMI BUDG SANC MMSA DFLN MSUR WBSN"


'DIALOG IS TOO LARGE FOR DIALOG EDITOR, CREATED MANUALLY
BeginDialog all_MAXIS_panels_dialog, 0, 0, 371, 190, "All MAXIS panels dialog"
  Checkbox 10, 10, 35, 10, "MEMB", MEMB_check
  Checkbox 60, 10, 35, 10, "MEMI", MEMI_check
  Checkbox 110, 10, 35, 10, "ADDR", ADDR_check
  Checkbox 160, 10, 35, 10, "AREP", AREP_check
  Checkbox 210, 10, 35, 10, "ALTP", ALTP_check
  Checkbox 260, 10, 35, 10, "ALIA", ALIA_check
  Checkbox 10, 25, 35, 10, "TYPE", TYPE_check
  Checkbox 60, 25, 35, 10, "PROG", PROG_check
  Checkbox 110, 25, 35, 10, "HCRE", HCRE_check
  Checkbox 160, 25, 35, 10, "PARE", PARE_check
  Checkbox 210, 25, 35, 10, "SIBL", SIBL_check
  Checkbox 260, 25, 35, 10, "EATS", EATS_check
  Checkbox 10, 40, 35, 10, "IMIG", IMIG_check
  Checkbox 60, 40, 35, 10, "SPON", SPON_check
  Checkbox 110, 40, 35, 10, "FACI", FACI_check
  Checkbox 160, 40, 35, 10, "FCFC", FCFC_check
  Checkbox 210, 40, 35, 10, "FCPL", FCPL_check
  Checkbox 260, 40, 35, 10, "ADME", ADME_check
  Checkbox 10, 55, 35, 10, "REMO", REMO_check
  Checkbox 60, 55, 35, 10, "DISA", DISA_check
  Checkbox 110, 55, 35, 10, "ABPS", ABPS_check
  Checkbox 160, 55, 35, 10, "PREG", PREG_check
  Checkbox 210, 55, 35, 10, "STRK", STRK_check
  Checkbox 260, 55, 35, 10, "STWK", STWK_check
  Checkbox 10, 70, 35, 10, "SCHL", SCHL_check
  Checkbox 60, 70, 35, 10, "WREG", WREG_check
  Checkbox 110, 70, 35, 10, "EMPS", EMPS_check
  Checkbox 160, 70, 35, 10, "CASH", CASH_check
  Checkbox 210, 70, 35, 10, "ACCT", ACCT_check
  Checkbox 260, 70, 35, 10, "SECU", SECU_check	'30
  Checkbox 10, 85, 35, 10, "CARS", CARS_check
  Checkbox 60, 85, 35, 10, "REST", REST_check
  Checkbox 110, 85, 35, 10, "OTHR", OTHR_check
  Checkbox 160, 85, 35, 10, "TRAN", TRAN_check
  Checkbox 210, 85, 35, 10, "STIN", STIN_check
  Checkbox 260, 85, 35, 10, "STEC", STEC_check
  Checkbox 10, 100, 35, 10, "PBEN", PBEN_check
  Checkbox 60, 100, 35, 10, "UNEA", UNEA_check
  Checkbox 110, 100, 35, 10, "LUMP", LUMP_check
  Checkbox 160, 100, 35, 10, "RBIC", RBIC_check
  Checkbox 210, 100, 35, 10, "BUSI", BUSI_check
  Checkbox 260, 100, 35, 10, "JOBS", JOBS_check
  Checkbox 10, 115, 35, 10, "TRAC", TRAC_check
  Checkbox 60, 115, 35, 10, "DSTT", DSTT_check
  Checkbox 110, 115, 35, 10, "DCEX", DCEX_check
  Checkbox 160, 115, 35, 10, "WKEX", WKEX_check
  Checkbox 210, 115, 35, 10, "COEX", COEX_check
  Checkbox 260, 115, 35, 10, "SHEL", SHEL_check
  Checkbox 10, 130, 35, 10, "HEST", HEST_check
  Checkbox 60, 130, 35, 10, "ACUT", ACUT_check	'50
  Checkbox 110, 130, 35, 10, "PDED", PDED_check
  Checkbox 160, 130, 35, 10, "PACT", PACT_check
  Checkbox 210, 130, 35, 10, "FMED", FMED_check
  Checkbox 260, 130, 35, 10, "ACCI", ACCI_check
  Checkbox 10, 145, 35, 10, "MEDI", MEDI_check
  Checkbox 60, 145, 35, 10, "INSA", INSA_check
  Checkbox 110, 145, 35, 10, "DIET", DIET_check
  Checkbox 160, 145, 35, 10, "DISQ", DISQ_check
  Checkbox 210, 145, 35, 10, "SWKR", SWKR_check
  Checkbox 260, 145, 35, 10, "REVW", REVW_check	'60
  Checkbox 10, 160, 35, 10, "MISC", MISC_check
  Checkbox 60, 160, 35, 10, "RESI", RESI_check
  Checkbox 110, 160, 35, 10, "TIME", TIME_check
  Checkbox 160, 160, 35, 10, "EMMA", EMMA_check
  Text 65, 65, 65, 65, "65...spooky!"
  Checkbox 210, 160, 35, 10, "BILS", BILS_check
  Checkbox 260, 160, 35, 10, "HCMI", HCMI_check
  Checkbox 10, 175, 35, 10, "BUDG", BUDG_check
  Checkbox 60, 175, 35, 10, "SANC", SANC_check
  Checkbox 110, 175, 35, 10, "WBSN", WBSN_check
  Checkbox 160, 175, 35, 10, "MMSA", MMSA_check
  Checkbox 210, 175, 35, 10, "DFLN", DFLN_check
  Checkbox 260, 175, 35, 10, "MSUR", MSUR_check
  Checkbox 310, 45, 65, 10, "ALL PANELS", all_panels_check
  ButtonGroup ButtonPressed
    OkButton 310, 5, 50, 15
    CancelButton 310, 25, 50, 15
EndDialog

BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

maxis_check_function

'Finds MAXIS case number
call MAXIS_case_number_finder(case_number)

'Finds MAXIS footer month
row = 1
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Shows case number dialog
Do
	Dialog case_number_and_footer_month_dialog
	If buttonpressed = 0 then stopscript
	If isnumeric(case_number) = False then MsgBox "You must type a valid case number."
Loop until isnumeric(case_number) = True

'Shows the MAXIS panel selection dialog
back_to_SELF

Dialog all_MAXIS_panels_dialog
	If buttonpressed = 0 then stopscript

call navigate_to_screen("STAT", "MEMI")
ERRR_screen_check

call HH_member_custom_dialog(HH_member_array)

'Adding checked objects to the array
IF all_panels_check = checked THEN
	all_panels_selected_array = all_panels_selected_array & all_possible_panels & " "
ELSE

If MEMB_check = checked then all_panels_selected_array = all_panels_selected_array & "MEMB" & " "
If MEMI_check = checked then all_panels_selected_array = all_panels_selected_array & "MEMI" & " "
If ADDR_check = checked then all_panels_selected_array = all_panels_selected_array & "ADDR" & " "
If AREP_check = checked then all_panels_selected_array = all_panels_selected_array & "AREP" & " "
If ALTP_check = checked then all_panels_selected_array = all_panels_selected_array & "ALTP" & " "
If ALIA_check = checked then all_panels_selected_array = all_panels_selected_array & "ALIA" & " "
If TYPE_check = checked then all_panels_selected_array = all_panels_selected_array & "TYPE" & " "
If PROG_check = checked then all_panels_selected_array = all_panels_selected_array & "PROG" & " "
If HCRE_check = checked then all_panels_selected_array = all_panels_selected_array & "HCRE" & " "
If PARE_check = checked then all_panels_selected_array = all_panels_selected_array & "PARE" & " "
If SIBL_check = checked then all_panels_selected_array = all_panels_selected_array & "SIBL" & " "
If EATS_check = checked then all_panels_selected_array = all_panels_selected_array & "EATS" & " "
If IMIG_check = checked then all_panels_selected_array = all_panels_selected_array & "IMIG" & " "
If SPON_check = checked then all_panels_selected_array = all_panels_selected_array & "SPON" & " "
If FACI_check = checked then all_panels_selected_array = all_panels_selected_array & "FACI" & " "
If FCFC_check = checked then all_panels_selected_array = all_panels_selected_array & "FCFC" & " "
If FCPL_check = checked then all_panels_selected_array = all_panels_selected_array & "FCPL" & " "
If ADME_check = checked then all_panels_selected_array = all_panels_selected_array & "ADME" & " "
If REMO_check = checked then all_panels_selected_array = all_panels_selected_array & "REMO" & " "
If DISA_check = checked then all_panels_selected_array = all_panels_selected_array & "DISA" & " "
If ABPS_check = checked then all_panels_selected_array = all_panels_selected_array & "ABPS" & " "
If PREG_check = checked then all_panels_selected_array = all_panels_selected_array & "PREG" & " "
If STRK_check = checked then all_panels_selected_array = all_panels_selected_array & "STRK" & " "
If STWK_check = checked then all_panels_selected_array = all_panels_selected_array & "STWK" & " "
If SCHL_check = checked then all_panels_selected_array = all_panels_selected_array & "SCHL" & " "
If WREG_check = checked then all_panels_selected_array = all_panels_selected_array & "WREG" & " "
If EMPS_check = checked then all_panels_selected_array = all_panels_selected_array & "EMPS" & " "
If CASH_check = checked then all_panels_selected_array = all_panels_selected_array & "CASH" & " "
If ACCT_check = checked then all_panels_selected_array = all_panels_selected_array & "ACCT" & " "
If SECU_check = checked then all_panels_selected_array = all_panels_selected_array & "SECU" & " "
If CARS_check = checked then all_panels_selected_array = all_panels_selected_array & "CARS" & " "
If REST_check = checked then all_panels_selected_array = all_panels_selected_array & "REST" & " "
If OTHR_check = checked then all_panels_selected_array = all_panels_selected_array & "OTHR" & " "
If TRAN_check = checked then all_panels_selected_array = all_panels_selected_array & "TRAN" & " "
If STIN_check = checked then all_panels_selected_array = all_panels_selected_array & "STIN" & " "
If STEC_check = checked then all_panels_selected_array = all_panels_selected_array & "STEC" & " "
If PBEN_check = checked then all_panels_selected_array = all_panels_selected_array & "PBEN" & " "
If UNEA_check = checked then all_panels_selected_array = all_panels_selected_array & "UNEA" & " "
If LUMP_check = checked then all_panels_selected_array = all_panels_selected_array & "LUMP" & " "
If RBIC_check = checked then all_panels_selected_array = all_panels_selected_array & "RBIC" & " "
If BUSI_check = checked then all_panels_selected_array = all_panels_selected_array & "BUSI" & " "
If JOBS_check = checked then all_panels_selected_array = all_panels_selected_array & "JOBS" & " "
If TRAC_check = checked then all_panels_selected_array = all_panels_selected_array & "TRAC" & " "
If DSTT_check = checked then all_panels_selected_array = all_panels_selected_array & "DSTT" & " "
If DCEX_check = checked then all_panels_selected_array = all_panels_selected_array & "DCEX" & " "
If WKEX_check = checked then all_panels_selected_array = all_panels_selected_array & "WKEX" & " "
If COEX_check = checked then all_panels_selected_array = all_panels_selected_array & "COEX" & " "
If SHEL_check = checked then all_panels_selected_array = all_panels_selected_array & "SHEL" & " "
If HEST_check = checked then all_panels_selected_array = all_panels_selected_array & "HEST" & " "
If ACUT_check = checked then all_panels_selected_array = all_panels_selected_array & "ACUT" & " "
If PDED_check = checked then all_panels_selected_array = all_panels_selected_array & "PDED" & " "
If PACT_check = checked then all_panels_selected_array = all_panels_selected_array & "PACT" & " "
If FMED_check = checked then all_panels_selected_array = all_panels_selected_array & "FMED" & " "
If ACCI_check = checked then all_panels_selected_array = all_panels_selected_array & "ACCI" & " "
If MEDI_check = checked then all_panels_selected_array = all_panels_selected_array & "MEDI" & " "
If INSA_check = checked then all_panels_selected_array = all_panels_selected_array & "INSA" & " "
If DIET_check = checked then all_panels_selected_array = all_panels_selected_array & "DIET" & " "
If DISQ_check = checked then all_panels_selected_array = all_panels_selected_array & "DISQ" & " "
If SWKR_check = checked then all_panels_selected_array = all_panels_selected_array & "SWKR" & " "
If REVW_check = checked then all_panels_selected_array = all_panels_selected_array & "REVW" & " "
If MISC_check = checked then all_panels_selected_array = all_panels_selected_array & "MISC" & " "
If RESI_check = checked then all_panels_selected_array = all_panels_selected_array & "RESI" & " "
If TIME_check = checked then all_panels_selected_array = all_panels_selected_array & "TIME" & " "
If EMMA_check = checked then all_panels_selected_array = all_panels_selected_array & "EMMA" & " "
If BILS_check = checked then all_panels_selected_array = all_panels_selected_array & "BILS" & " "
If HCMI_check = checked then all_panels_selected_array = all_panels_selected_array & "HCMI" & " "
If BUDG_check = checked then all_panels_selected_array = all_panels_selected_array & "BUDG" & " "
If SANC_check = checked then all_panels_selected_array = all_panels_selected_array & "SANC" & " "
If MMSA_check = checked then all_panels_selected_array = all_panels_selected_array & "MMSA" & " "
If DFLN_check = checked then all_panels_selected_array = all_panels_selected_array & "DFLN" & " "
If MSUR_check = checked then all_panels_selected_array = all_panels_selected_array & "MSUR" & " "
If WBSN_check = checked then all_panels_selected_array = all_panels_selected_array & "WBSN" & " "		'WBSN needs to be last for whatever reason. The script gets stuck on WBSN and won't read anything after...not sure why

END IF

'Splits the array
all_panels_selected_array = trim(all_panels_selected_array)
all_panels_selected_array = split(all_panels_selected_array, " ")

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 30
objSelection.PageSetup.BottomMargin = 30
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "10"


For each panel_to_scan in all_panels_selected_array
	IF panel_to_scan = "TYPE" OR _
	panel_to_scan = "HEST" OR _
	panel_to_scan = "MISC" OR _
	panel_to_scan = "BUDG" OR _
	panel_to_scan = "PROG" OR _
	panel_to_scan = "DSTT" OR _
	panel_to_scan = "INSA" OR _
	panel_to_scan = "RESI" OR _
	panel_to_scan = "ADDR" OR _
	panel_to_scan = "HCRE" OR _
	panel_to_scan = "ABPS" THEN
		call navigate_to_screen("STAT", panel_to_scan)
				call copy_screen_to_array(screentest)

				'Adds current screen to Word doc
				For each line in screentest
					objSelection.TypeText line & Chr(11)
				Next

				'Determines if the Word doc needs a new page
				If screen_on_page = "" or screen_on_page = 1 then
					screen_on_page = 2
					objSelection.TypeText vbCr & vbCr
				Elseif screen_on_page = 2 then
					screen_on_page = 1
					objSelection.InsertBreak(7)
				End if

	ELSE

		FOR EACH HH_member IN (HH_member_array)
		
			current_panel = ""
			number_of_panels = ""

			IF panel_to_scan = "MEMB" THEN 
				call navigate_to_screen("STAT", "MEMB")
				EMWriteScreen hh_member, 20, 76
				transmit
					call copy_screen_to_array(screentest)

					'Adds current screen to Word doc
					For each line in screentest
						objSelection.TypeText line & Chr(11)
					Next
	
					'Determines if the Word doc needs a new page
					If screen_on_page = "" or screen_on_page = 1 then
						screen_on_page = 2
						objSelection.TypeText vbCr & vbCr
					Elseif screen_on_page = 2 then
						screen_on_page = 1
						objSelection.InsertBreak(7)
					End if

			ELSEIF panel_to_scan = "BILS" THEN
				call navigate_to_screen("STAT", "BILS")
					EMReadScreen total_bils_panel, 1, 3, 78
					IF total_bils_panel = "0" THEN
						call copy_screen_to_array(screentest)
						'Adds current screen to Word doc
						For each line in screentest
							objSelection.TypeText line & Chr(11)
						Next
	
						'Determines if the Word doc needs a new page
						If screen_on_page = "" or screen_on_page = 1 then
							screen_on_page = 2
							objSelection.TypeText vbCr & vbCr
						Elseif screen_on_page = 2 then
							screen_on_page = 1
							objSelection.InsertBreak(7)
						End if
					ELSEIF total_bils_panel = "1" THEN
						call copy_screen_to_array(screentest)
						'Adds current screen to Word doc
						For each line in screentest
							objSelection.TypeText line & Chr(11)
						Next
	
						'Determines if the Word doc needs a new page
						If screen_on_page = "" or screen_on_page = 1 then
							screen_on_page = 2
							objSelection.TypeText vbCr & vbCr
						Elseif screen_on_page = 2 then
							screen_on_page = 1
							objSelection.InsertBreak(7)
						End if
					ELSEIF total_bils_panel <> "0" AND total_bils_panel <> "1" THEN
						DO
							EMReadScreen last_bils_screen, 9, 19, 66
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
	
							'Determines if the Word doc needs a new page
							If screen_on_page = "" or screen_on_page = 1 then
								screen_on_page = 2
								objSelection.TypeText vbCr & vbCr
							Elseif screen_on_page = 2 then
								screen_on_page = 1
								objSelection.InsertBreak(7)
							End if
							PF20
						LOOP until last_bils_screen = "More:   -"	
					END IF

			ELSEIF panel_to_scan = "FMED" THEN
				call navigate_to_screen("STAT", "FMED")
					EMReadScreen more_fmed_screens, 7, 15, 68
					IF more_fmed_screens = "       " THEN
						call copy_screen_to_array(screentest)
						'Adds current screen to Word doc
						For each line in screentest
							objSelection.TypeText line & Chr(11)
						Next
	
						'Determines if the Word doc needs a new page
						If screen_on_page = "" or screen_on_page = 1 then
							screen_on_page = 2
							objSelection.TypeText vbCr & vbCr
						Elseif screen_on_page = 2 then
							screen_on_page = 1
							objSelection.InsertBreak(7)
						End if
					ELSEIF more_fmed_screens = "More: +" THEN
						EMReadScreen more_fmed_screens, 7, 15, 68
						call copy_screen_to_array(screentest)
						'Adds current screen to Word doc
						For each line in screentest
							objSelection.TypeText line & Chr(11)
						Next

						'Determines if the Word doc needs a new page
						If screen_on_page = "" or screen_on_page = 1 then
							screen_on_page = 2
							objSelection.TypeText vbCr & vbCr
						Elseif screen_on_page = 2 then
							screen_on_page = 1
							objSelection.InsertBreak(7)
						End if
						PF20

						EMReadScreen more_fmed_screens, 7, 15, 68
						call copy_screen_to_array(screentest)
						'Adds current screen to Word doc
						For each line in screentest
							objSelection.TypeText line & Chr(11)
						Next

						'Determines if the Word doc needs a new page
						If screen_on_page = "" or screen_on_page = 1 then
							screen_on_page = 2
							objSelection.TypeText vbCr & vbCr
						Elseif screen_on_page = 2 then
							screen_on_page = 1
							objSelection.InsertBreak(7)
						End if
					END IF

			ELSE
				'Goes to the screen for the first HH memb
				call navigate_to_screen("STAT", panel_to_scan)
				EMWriteScreen hh_member, 20, 76
				EMWriteScreen "01", 20, 79
				transmit
				EMReadScreen current_panel, 2, 2, 72
				current_panel = cint(current_panel)
				EMReadScreen number_of_panels, 2, 2, 78
				number_of_panels = cint(number_of_panels)
	
				DO
					EMWriteScreen ("0" & current_panel), 20, 79
					transmit
					call copy_screen_to_array(screentest)

					'Adds current screen to Word doc
					For each line in screentest
						objSelection.TypeText line & Chr(11)
					Next

					'Determines if the Word doc needs a new page
					If screen_on_page = "" or screen_on_page = 1 then
						screen_on_page = 2
						objSelection.TypeText vbCr & vbCr
					Elseif screen_on_page = 2 then
						screen_on_page = 1
						objSelection.InsertBreak(7)
					End if

					current_panel = current_panel + 1
				LOOP UNTIL (left(number_of_panels, 1) = "0") OR (current_panel = (number_of_panels + 1))
			END IF
		NEXT
	END IF
NEXT
