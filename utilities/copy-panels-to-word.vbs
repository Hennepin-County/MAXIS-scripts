'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - COPY PANELS TO WORD.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 43                	'manual run time in seconds
STATS_denomination = "I"       		'I is for each ITEM
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
call changelog_update("08/04/2021", "Updated to work. Previously the script errored with every run", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

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
all_possible_panels = "MEMB MEMI ADDR AREP ALTP ALIA " &_
                      "TYPE PROG HCRE PARE SIBL EATS " &_
					  "IMIG SPON FACI FCFC FCPL ADME " &_
					  "REMO DISA ABPS PREG STRK STWK " &_
					  "SCHL WREG EMPS CASH ACCT SECU " &_
					  "CARS REST OTHR TRAN STIN STEC " &_
					  "PBEN UNEA LUMP RBIC BUSI JOBS " &_
					  "TRAC DSTT DCEX WKEX COEX SHEL " &_
					  "HEST ACUT PDED PACT FMED ACCI " &_
					  "MEDI INSA DIET DISQ SWKR REVW " &_
					  "MISC RESI TIME EMMA BILS HCMI " &_
					  "BUDG SANC MMSA DFLN MSUR SSRT"

const panel_name_const 		= 0
const panel_exists_const 	= 1
const panel_checkbox_const 	= 2
const panel_last_const 		= 3

const MEMB_const = 00
const MEMI_const = 01
const ADDR_const = 02
const AREP_const = 03
const ALTP_const = 04
const ALIA_const = 05
const TYPE_const = 06
const PROG_const = 07
const HCRE_const = 08
const PARE_const = 09
const SIBL_const = 10
const EATS_const = 11
const IMIG_const = 12
const SPON_const = 13
const FACI_const = 14
const FCFC_const = 15
const FCPL_const = 16
const ADME_const = 17
const REMO_const = 18
const DISA_const = 19
const ABPS_const = 20
const PREG_const = 21
const STRK_const = 22
const STWK_const = 23
const SCHL_const = 24
const WREG_const = 25
const EMPS_const = 26
const CASH_const = 27
const ACCT_const = 28
const SECU_const = 29
const CARS_const = 30
const REST_const = 31
const OTHR_const = 32
const TRAN_const = 33
const STIN_const = 34
const STEC_const = 35
const PBEN_const = 36
const UNEA_const = 37
const LUMP_const = 38
const RBIC_const = 39
const BUSI_const = 40
const JOBS_const = 41
const TRAC_const = 42
const DSTT_const = 43
const DCEX_const = 44
const WKEX_const = 45
const COEX_const = 46
const SHEL_const = 47
const HEST_const = 48
const ACUT_const = 49
const PDED_const = 50
const PACT_const = 51
const FMED_const = 52
const ACCI_const = 53
const MEDI_const = 54
const INSA_const = 55
const DIET_const = 56
const DISQ_const = 57
const SWKR_const = 58
const REVW_const = 59
const MISC_const = 60
const RESI_const = 61
const TIME_const = 62
const EMMA_const = 63
const BILS_const = 64
const HCMI_const = 65
const BUDG_const = 66
const SANC_const = 67
const MMSA_const = 68
const DFLN_const = 69
const MSUR_const = 70
const SSRT_const = 71

Dim ALL_THE_PANELS_ARRAY()
ReDim ALL_THE_PANELS_ARRAY(panel_last_const, SSRT_const)



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

Call check_for_MAXIS(False)

'Finds MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)
'Finds MAXIS footer month
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog dialog1, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, MAXIS_footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog

'Shows case number dialog
Do
	Dialog Dialog1
	cancel_without_confirmation
	If isnumeric(MAXIS_case_number) = False then MsgBox "You must type a valid case number."
Loop until isnumeric(MAXIS_case_number) = True

'Shows the MAXIS panel selection dialog
back_to_SELF

Call navigate_to_MAXIS_screen("STAT", "SPAN")

panel_name_array = split(all_possible_panels, " ")

span_row = 7
span_col = 11

for panel_name = 0 to UBound(panel_name_array)
	ALL_THE_PANELS_ARRAY(panel_name_const, panel_name) = panel_name_array(panel_name)
	EMReadScreen panel_counter, 1, span_row, span_col
	ALL_THE_PANELS_ARRAY(panel_exists_const, panel_name) = False
	If panel_counter <> " " Then ALL_THE_PANELS_ARRAY(panel_exists_const, panel_name) = True
	span_col = span_col + 13
	If span_col = 89 then
		span_row = span_row + 1
		span_col = 11
	End If
next

'DIALOG IS TOO LARGE FOR DIALOG EDITOR, CREATED MANUALLY
x_pos = 10
y_pos = 10
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 190, "All MAXIS panels dialog"
  ' Text 10, 10, 300, 10, "MEMB                   MEMI                   ADDR                   AREP                   ALTP                   ALIA "
  For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
  ' If ALL_THE_PANELS_ARRAY(panel_exists_const, panel_counter) = False Then Text x_pos + 10, y_pos, 35, 10, ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter)
	If ALL_THE_PANELS_ARRAY(panel_exists_const, panel_counter) = True Then
		Checkbox x_pos, y_pos, 35, 10, ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter), ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter)
		x_pos = x_pos + 50
		If x_pos = 310 Then
		x_pos = 10
		y_pos = y_pos + 15
		End If
	End If
  Next
  Checkbox 310, 45, 65, 10, "ALL PANELS", all_panels_check
  Checkbox 310, 60, 65, 10, "ALL EXISTING", all_existing_panels_check
  Text 320, 70, 50, 10, "PANELS"
  ButtonGroup ButtonPressed
    OkButton 310, 5, 50, 15
    CancelButton 310, 25, 50, 15
EndDialog

Dialog Dialog1

Cancel_confirmation

call navigate_to_MAXIS_screen("STAT", "MEMI")

call HH_member_custom_dialog(HH_member_array)

'Adding checked objects to the array
IF all_panels_check = checked THEN
	' all_panels_selected_array = all_panels_selected_array & all_possible_panels & " "
	For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
		ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter) = checked
	Next
ELSEIf all_existing_panels_check = checked Then
	For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
		If ALL_THE_PANELS_ARRAY(panel_exists_const, panel_counter) = True Then ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter) = checked
	Next
END IF

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

For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
	If ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter) = checked  Then
		IF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "TYPE" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "HEST" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "MISC" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "BUDG" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "PROG" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "DSTT" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "INSA" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "RESI" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "ADDR" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "HCRE" OR _
		ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "ABPS" THEN
			call navigate_to_MAXIS_screen("STAT", ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter))
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
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter

		ELSE

			FOR EACH HH_member IN (HH_member_array)

				current_panel = ""
				number_of_panels = ""

				IF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "MEMB" THEN
					call navigate_to_MAXIS_screen("STAT", "MEMB")
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
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter

				ELSEIF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "BILS" THEN
					call navigate_to_MAXIS_screen("STAT", "BILS")
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
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
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
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
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
								STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							LOOP until last_bils_screen = "More:   -"
						END IF

				ELSEIF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "FMED" THEN
					call navigate_to_MAXIS_screen("STAT", "FMED")
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
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
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
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
						END IF

				ELSE
					'Goes to the screen for the first HH memb
					call navigate_to_MAXIS_screen("STAT", ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter))
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
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					LOOP UNTIL (left(number_of_panels, 1) = "0") OR (current_panel = (number_of_panels + 1))
				END IF
			NEXT
		END IF
	END IF
	EMWriteScreen "SPAN", 20, 71
	EMWriteScreen "  ", 20, 76
	EMWriteScreen "  ", 20, 79
	transmit
NEXT

STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter
script_end_procedure_with_error_report("Success! Word Document created and opened with PANEL information as selected.")
