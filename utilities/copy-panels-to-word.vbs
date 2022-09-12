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
call changelog_update("09/12/2022", "The script will now output the PIC information from JOBS and UNEA, and BUSI Calculation when selected.", "Casey Love, Hennepin County")	'#735 & #969'
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

function insert_page_break_after_two_panels(screen_on_page)
	'Determines if the Word doc needs a new page
	'screen_on_page - This is a running counter that is updated in this function
	If screen_on_page = "" or screen_on_page = 1 then							'if we are at 1, we need to add some spaces and increment the counter'
		screen_on_page = 2
		objSelection.TypeText vbCr & vbCr
	Elseif screen_on_page = 2 then												'if we are at 2, we need to insert a page breakk and reset the counter
		screen_on_page = 1
		objSelection.InsertBreak(7)
	End if
	STATS_counter = STATS_counter + 1											'also using this to increment the stats counter since we do this with every panel we read.'
end function

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
'These are all the panels as they are laid out in STAT/SPAN
'This layout is to make these easier to review and update as needed.
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

'These are constants for our array of panels.
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

'declaration of the array of all the panels.
'This will tell us which panels exist on the case and which were selected in the dialog.
Dim ALL_THE_PANELS_ARRAY()
ReDim ALL_THE_PANELS_ARRAY(panel_last_const, SSRT_const)

screen_on_page = 1

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

Call check_for_MAXIS(False)			'ensuring we are logged in to MAXIS.

'Finds MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)
'Finds MAXIS footer month
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Inital Dialog for capturing case number'
Dialog1 = ""
BeginDialog dialog1, 0, 0, 166, 65, "Copy Panels to Word Inital Dialog"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 65, 15, MAXIS_case_number
  Text 40, 30, 50, 10, "Footer month:"
  EditBox 95, 25, 20, 15, MAXIS_footer_month
  Text 120, 30, 20, 10, "Year:"
  EditBox 140, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 50, 45, 50, 15
    CancelButton 110, 45, 50, 15
EndDialog

'Shows case number dialog
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If err_msg <> "" Then MsgBox " ************* NOTICE **************" & vbCr & vbCr & "Please Resolve to Continue:" & vbCr & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

Call MAXIS_footer_month_confirmation				'Ensuring we are in SPAN at the right footer month
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SPAN", is_this_priv)
If is_this_priv = True Then Call script_end_procedure("PRIV CASE! This case is privileged and the script cannot continue. Ensure you are allowed access to this case and request access. Once you have access, rerun the script.")

panel_name_array = split(all_possible_panels, " ")		'Creating an array of all of the panel names

'Now we are going to read SPAN to identify which panels exist on this case
'This functionality relies on the panels being in order in the array just created, so that array MUST match SPAN
span_row = 7						'We are reading just the number indicator next to the panel name - this is where the first one is.
span_col = 11
for panel_name = 0 to UBound(panel_name_array)									'look at each panel
	ALL_THE_PANELS_ARRAY(panel_name_const, panel_name) = panel_name_array(panel_name)				'saving the panel name to the LARGE array
	EMReadScreen panel_counter, 1, span_row, span_col												'Reading if SPAN indicates that this panel exists
	ALL_THE_PANELS_ARRAY(panel_exists_const, panel_name) = False									'defaulting the existing to False
	If panel_counter <> " " Then ALL_THE_PANELS_ARRAY(panel_exists_const, panel_name) = True		'If the counter indicator was not blank, then this panel exists. Saving to the LARGE array
	span_col = span_col + 13					'Now we go to the next location of the counter
	If span_col = 89 then
		span_row = span_row + 1
		span_col = 11
	End If
next

'Dialog of the panels to select.
x_pos = 10
y_pos = 10
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 190, "All MAXIS panels dialog"
  Text 10, 11, 300, 10,  "     MEMB               MEMI               ADDR               AREP                ALTP                ALIA "				'These are strings of all of the panel names from SPAN.
  Text 10, 26, 300, 10,  "     TYPE               PROG                HCRE               PARE               SIBL                  EATS "			'They sit behind the checkboxes. We can't put text in for each one ebcause the dialog breaks
  Text 10, 41, 300, 10,  "      IMIG                 SPON               FACI                 FCFC                 FCPL                ADME "
  Text 10, 56, 300, 10,  "      REMO              DISA                 ABPS               PREG               STRK               STWK "
  Text 10, 71, 300, 10,  "      SCHL               WREG              EMPS               CASH                ACCT              SECU "
  Text 10, 86, 300, 10,  "      CARS               REST               OTHR               TRAN               STIN                STEC "
  Text 10, 101, 300, 10, "      PBEN               UNEA               LUMP               RBIC                 BUSI                JOBS "
  Text 10, 116, 300, 10, "      TRAC               DSTT               DCEX                WKEX               COEX              SHEL "
  Text 10, 131, 300, 10, "      HEST               ACUT               PDED                PACT                FMED              ACCI "
  Text 10, 146, 300, 10, "      MEDI                INSA                DIET                 DISQ                 SWKR              REVW "
  Text 10, 161, 300, 10, "      MISC                RESI                TIME                 EMMA               BILS                 HCMI "
  Text 10, 176, 300, 10, "      BUDG              SANC               MMSA               DFLN                MSUR              SSRT "

  'Looping through each othe panels and showing a checkbox for any panel that exists on this case
  For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
  ' If ALL_THE_PANELS_ARRAY(panel_exists_const, panel_counter) = False Then Text x_pos + 10, y_pos, 35, 10, ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter)
	If ALL_THE_PANELS_ARRAY(panel_exists_const, panel_counter) = True Then
		Checkbox x_pos, y_pos, 35, 10, ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter), ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter)
	End If
	x_pos = x_pos + 50			'moving through the spaces the checkboxes go
	If x_pos = 310 Then
		x_pos = 10
		y_pos = y_pos + 15
	End If
  Next
  Checkbox 310, 5, 65, 10, "ALL PANELS", all_panels_check
  Checkbox 310, 20, 65, 10, "ALL EXISTING", all_existing_panels_check
  Text 320, 30, 50, 10, "PANELS"
  Text 310, 40, 60, 20, "For income panels, include PIC?"
  DropListBox 310, 60, 50, 45, "Yes"+chr(9)+"No", include_pics
  ButtonGroup ButtonPressed
    OkButton 315, 150, 50, 15
    CancelButton 315, 170, 50, 15
EndDialog

Do
	Dialog Dialog1
	cancel_without_confirmation

	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False
'There is no looping because there is no mandated selections/information

'Now getting the list of the household members.
call HH_member_custom_dialog(HH_member_array)

'If the 'all panels' checkboxes are used, updating the LARGE array to indicate that the correct panels are checked
IF all_panels_check = checked THEN
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

'Now we go through the LARGE array and grab the panel information if indicates as checked.
For panel_counter = 0 to UBound(ALL_THE_PANELS_ARRAY, 2)
	If ALL_THE_PANELS_ARRAY(panel_checkbox_const, panel_counter) = checked  Then
		'These panels have person association
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

			Call insert_page_break_after_two_panels(screen_on_page)
		ELSE		'the rest of the panels have person association.

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
						Call insert_page_break_after_two_panels(screen_on_page)

				ELSEIF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "BILS" THEN			'BILS works a little different
					call navigate_to_MAXIS_screen("STAT", "BILS")
						EMReadScreen total_bils_panel, 1, 3, 78
						IF total_bils_panel = "0" THEN
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
							Call insert_page_break_after_two_panels(screen_on_page)
						ELSEIF total_bils_panel = "1" THEN
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
							Call insert_page_break_after_two_panels(screen_on_page)
						ELSEIF total_bils_panel <> "0" AND total_bils_panel <> "1" THEN
							DO
								EMReadScreen last_bils_screen, 9, 19, 66
								call copy_screen_to_array(screentest)
								'Adds current screen to Word doc
								For each line in screentest
									objSelection.TypeText line & Chr(11)
								Next
								Call insert_page_break_after_two_panels(screen_on_page)
							LOOP until last_bils_screen = "More:   -"
						END IF

				ELSEIF ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "FMED" THEN		'FMED works a little different'
					call navigate_to_MAXIS_screen("STAT", "FMED")
						EMReadScreen more_fmed_screens, 7, 15, 68
						IF more_fmed_screens = "       " THEN
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
							Call insert_page_break_after_two_panels(screen_on_page)
						ELSEIF more_fmed_screens = "More: +" THEN
							EMReadScreen more_fmed_screens, 7, 15, 68
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
							Call insert_page_break_after_two_panels(screen_on_page)
							PF20

							EMReadScreen more_fmed_screens, 7, 15, 68
							call copy_screen_to_array(screentest)
							'Adds current screen to Word doc
							For each line in screentest
								objSelection.TypeText line & Chr(11)
							Next
							Call insert_page_break_after_two_panels(screen_on_page)
						END IF

				ELSE				'All the other panels
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

						If include_pics = "Yes" Then
							If ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "JOBS" Then
								Call write_value_and_transmit("X", 19, 38)		'SNAP PIC'

								Call insert_page_break_after_two_panels(screen_on_page)

								call copy_screen_to_array(screentest)

								'Adds current screen to Word doc
								For each line in screentest
									objSelection.TypeText line & Chr(11)
								Next
								PF3

								Call write_value_and_transmit("X", 19, 71)		'GRH PIC'

								Call insert_page_break_after_two_panels(screen_on_page)

								call copy_screen_to_array(screentest)

								'Adds current screen to Word doc
								For each line in screentest
									objSelection.TypeText line & Chr(11)
								Next
								PF3
							End If
							If ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "UNEA" Then
								Call write_value_and_transmit("X", 10, 26)		'SNAP PIC'

								Call insert_page_break_after_two_panels(screen_on_page)

								call copy_screen_to_array(screentest)

								'Adds current screen to Word doc
								For each line in screentest
									objSelection.TypeText line & Chr(11)
								Next
								PF3
							End If
							If ALL_THE_PANELS_ARRAY(panel_name_const, panel_counter) = "BUSI" Then
								Call write_value_and_transmit("X", 6, 26)		'Calculation Pop Up'

								Call insert_page_break_after_two_panels(screen_on_page)

								call copy_screen_to_array(screentest)

								'Adds current screen to Word doc
								For each line in screentest
									objSelection.TypeText line & Chr(11)
								Next
								PF3
							End If
						End If
						Call insert_page_break_after_two_panels(screen_on_page)

						current_panel = current_panel + 1
					LOOP UNTIL (left(number_of_panels, 1) = "0") OR (current_panel = (number_of_panels + 1))
				END IF
			NEXT
		END IF
	END IF
NEXT
EMWriteScreen "SPAN", 20, 71		'ending back at span
EMWriteScreen "  ", 20, 76
EMWriteScreen "  ", 20, 79
transmit

STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter
script_end_procedure_with_error_report("Success! Word Document created and opened with PANEL information as selected.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/09/2021
'--Tab orders reviewed & confirmed----------------------------------------------09/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------08/09/2021
'--All variables in dialog match mandatory fields-------------------------------08/09/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------          ----------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------          ----------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------          ----------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/09/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/09/2021
'--PRIV Case handling reviewed -------------------------------------------------08/09/2021
'--Out-of-County handling reviewed----------------------------------------------          ----------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/09/2021
'--BULK - review output of statistics and run time/count (if applicable)--------08/09/2021
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/09/2021
'--Incrementors reviewed (if necessary)-----------------------------------------09/12/2022
'--Denomination reviewed -------------------------------------------------------08/09/2021
'--Script name reviewed---------------------------------------------------------08/09/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/09/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/12/2022
'--comment Code-----------------------------------------------------------------09/12/2022
'--Update Changelog for release/update------------------------------------------09/12/2022
'--Remove testing message boxes-------------------------------------------------09/12/2022
'--Remove testing code/unnecessary code-----------------------------------------09/12/2022
'--Review/update SharePoint instructions----------------------------------------09/12/2022
'--Review Best Practices using BZS page ----------------------------------------08/09/2021
'--Review script information on SharePoint BZ Script List-----------------------08/09/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/09/2021
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/09/2021
'--Complete misc. documentation (if applicable)---------------------------------          ----------------N/A
'--Update project team/issue contact (if applicable)----------------------------          ----------------N/A
