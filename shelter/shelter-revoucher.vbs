'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-REVOUCHER.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
CALL changelog_update("04/12/2022", "Elimination of Self-Pay: Removal of mention from scripts.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2016", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'autofilling the date to the current Date
revoucher_date = date & ""

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 146, 145, "Select a revoucher option"
  EditBox 80, 10, 60, 15, MAXIS_case_number
  DropListBox 80, 30, 60, 10, "Select one..."+chr(9)+"Family"+chr(9)+"Single", revoucher_option
  EditBox 125, 50, 15, 15, goals_accomplished
  EditBox 125, 70, 15, 15, next_goals
  EditBox 5, 100, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 125, 50, 15
    CancelButton 90, 125, 50, 15
  Text 30, 15, 45, 10, "Case number:"
  Text 15, 35, 60, 10, "Revoucher option:"
  Text 20, 55, 105, 10, "How many goals accomplished:"
  Text 35, 75, 90, 10, "Goals for the next voucher:"
  Text 5, 90, 65, 10, "Worker Signature:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
        cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF revoucher_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select a revoucher option."
		If goals_accomplished <> "" AND IsNumeric(goals_accomplished) = False Then err_msg = err_msg & vbNewLine & "* Goals accomplished must be entered as a number, to indicate the number of goals accomplished."
		If next_goals <> "" AND IsNumeric(next_goals) = False Then err_msg = err_msg & vbNewLine & "* Next goals must be entered as a number, to indicate the number of goals set."
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	Call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = False

'-------------------------------------------------------------------------------------------------DIALOG

If revoucher_option = "Family" then
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 336, 95, "Family revoucher"
	  DropListBox 55, 10, 60, 15, "Select one..."+chr(9)+"ACF"+chr(9)+"EA", voucher_type
	  EditBox 195, 10, 55, 15, revoucher_date
	  EditBox 305, 10, 25, 15, num_nights
	  DropListBox 55, 35, 115, 15, "Select one..."+chr(9)+"FMF"+chr(9)+"PSP"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
	  EditBox 225, 35, 25, 15, children
	  EditBox 305, 35, 25, 15, adults
	  EditBox 90, 55, 240, 15, bus_issued
	  EditBox 45, 75, 175, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 225, 75, 50, 15
	    CancelButton 280, 75, 50, 15
	  Text 5, 80, 40, 10, "Comments: "
	  Text 180, 40, 45, 10, "# of Children:"
	  Text 5, 15, 45, 10, "Voucher type:"
	  Text 265, 15, 40, 10, "# of nights:"
	  Text 130, 15, 60, 10, "Date of revoucher:"
	  Text 5, 40, 45, 10, "Shelter name:"
	  Text 5, 60, 85, 10, "Bus tokens/cards issued:"
	  Text 265, 40, 40, 10, "# of Adults:"
	EndDialog

	DO
		DO
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF voucher_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select a voucher type."
			If isDate(revoucher_date) = False then err_msg = err_msg & vbNewLine & "* Please enter the revoucher date."
			If IsNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Please enter the number of nights issued."
			If shelter_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Please choose a shelter name."
			If IsNumeric(children) = False then err_msg = err_msg & vbNewLine & "* Please enter the number of children."
			If IsNumeric(adults) = False then err_msg = err_msg & vbNewLine & "* Please enter the number of adults."
			If bus_issued = "" then err_msg = err_msg & vbNewLine & "* Please enter information about bus cards/tokens issued."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL are_we_passworded_out = False
END IF
'-------------------------------------------------------------------------------------------------DIALOG

If revoucher_option = "Single" then

	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 341, 115, "Single revoucher"
	  DropListBox 55, 10, 60, 15, "Select one..."+chr(9)+"GA/GRH"+chr(9)+"O/C", voucher_type
	  EditBox 195, 10, 55, 15, revoucher_date
	  EditBox 300, 10, 30, 15, num_nights
	  DropListBox 55, 35, 60, 15, "Select one..."+chr(9)+"PSP"+chr(9)+"SA-HL", shelter_droplist
	  EditBox 210, 35, 120, 15, shelter_dates
	  EditBox 90, 55, 240, 15, bus_issued
	  EditBox 90, 75, 240, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 225, 95, 50, 15
	    CancelButton 280, 95, 50, 15
	  Text 45, 80, 40, 10, "Other notes: "
	  Text 125, 40, 85, 10, "Dates shelter issued for:"
	  Text 5, 15, 45, 10, "Voucher type:"
	  Text 130, 15, 60, 10, "Date of revoucher:"
	  Text 5, 40, 45, 10, "Shelter type:"
	  Text 5, 60, 85, 10, "Bus tokens/cards issued:"
	  Text 260, 15, 40, 10, "# of nights:"
	EndDialog

	DO
		DO
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF voucher_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a voucher type."
			If isDate(revoucher_date) = False then err_msg = err_msg & vbNewLine & "* Enter the revoucher date."
			If IsNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Enter the number of nights issued."
			If shelter_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the shelter type."
			If shelter_dates = "" then err_msg = err_msg & vbNewLine & "* Enter the dates of the shelter stay."
			If bus_issued = "" then err_msg = err_msg & vbNewLine & "* Enter information about bus cards/tokens issued."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF

If goals_accomplished = "0" then goals_accomplished = ""
If next_goals = "0" then next_goals = ""

'Dynamic dialog for goals accomplished and next goals----------------------------------------------------------------------------------------------------
If goals_accomplished <> "" then
    Dim goals_accomplished_array()
    ReDim goals_accomplished_array(goals_accomplished - 1)
    goals_number = 1
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BEGINDIALOG Dialog1, 0, 0, 315, (120 + (goals_accomplished * 10)), "Goals accomplished for voucher"   'Creates the dynamic dialog. The height will change based on the number of goals it finds.
      'GroupBox 5, 5, 330, (10 + (i * 20), "Goals accomplished"
      For i = 0 to goals_accomplished - 1
        Text 5, (10 + (i * 20)), 10, 10, goals_number & ":"
        EditBox 20, (10 + (i * 20)), 285, 15, goals_accomplished_array(i)
        goals_number = goals_number + 1
      NEXT
      ButtonGroup buttonpressed
      OkButton 200, (10 + (i * 20)), 50, 15
      CancelButton 255, (10 + (i * 20)), 50, 15
    ENDDIALOG

    dialog Dialog1
    If buttonpressed = 0 then script_end_procedure("You have selected to end this script.")
End if

If next_goals <> "" then
    Dim next_goals_array()
    ReDim next_goals_array(next_goals - 1)
    goals_number = 1

    BEGINDIALOG next_goal_dialog, 0, 0, 315, (120 + (next_goals * 10)), "Goals for the next voucher"   'Creates the dynamic dialog. The height will change based on the number of goals it finds.
      'GroupBox 5, 5, 330, (10 + (i * 20), "Goals accomplished"
      For i = 0 to next_goals - 1
        Text 5, (10 + (i * 20)), 10, 10, goals_number & ":"
        EditBox 20, (10 + (i * 20)), 285, 15, next_goals_array(i)
        goals_number = goals_number + 1
      NEXT
      ButtonGroup buttonpressed
      OkButton 200, (10 + (i * 20)), 50, 15
      CancelButton 255, (10 + (i * 20)), 50, 15
    ENDDIALOG

    dialog next_goal_dialog
    If buttonpressed = 0 then script_end_procedure("You have selected to end this script.")
End if

'Variables for the case note----------------------------------------------------------------------------------------------------
exit_date = dateadd("d", num_nights, revoucher_date)
header_date = revoucher_date & " - " & exit_date

'The case note--------------------------------------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### " & voucher_type & " " & revoucher_option & " voucher " & header_date & " at " & shelter_droplist & " for " & num_nights & " nights ###")
IF revoucher_option = "Family" then Call write_variable_in_CASE_NOTE("* HH comp: " & adults & "A," & children & "C")
Call write_bullet_and_variable_in_CASE_NOTE("Bus tokens/cards issued", bus_issued)
Call write_bullet_and_variable_in_CASE_NOTE("Dates shelter issued for:", shelter_dates)
'Dynamic information for goals and next goals
If goals_accomplished <> "" then
    Call write_variable_in_CASE_NOTE("--Goals Accomplished--")
	goals_number = 1
    FOR i = 0 to goals_accomplished - 1
        call write_bullet_and_variable_in_CASE_NOTE(goals_number, goals_accomplished_array(i))
		goals_number = goals_number + 1
    NEXT
End if

If next_goals <> "" then
    Call write_variable_in_CASE_NOTE("--Next Goals--")
	goals_number = 1
    FOR i = 0 to next_goals	- 1
        call write_bullet_and_variable_in_CASE_NOTE(goals_number,  next_goals_array(i))
		goals_number = goals_number + 1
    NEXT
End if
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

Call script_end_procedure_with_error_report("Revoucher case note entered please follow all next steps to assist the resident.")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/29/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/29/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/29/2022
'--All variables in dialog match mandatory fields-------------------------------04/29/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/29/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------04/29/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/29/2022
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------n/a
'--MAXIS_background_check reviewed (if applicable)------------------------------n/a
'--PRIV Case handling reviewed -------------------------------------------------n/a
'--Out-of-County handling reviewed----------------------------------------------n/a
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/29/2022
'--BULK - review output of statistics and run time/count (if applicable)--------n/a
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------n/a time studies are really hard
'--Incrementors reviewed (if necessary)-----------------------------------------04/29/2022
'--Denomination reviewed -------------------------------------------------------04/29/2022
'--Script name reviewed---------------------------------------------------------04/29/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------n/a

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------04/29/2022
'--comment Code-----------------------------------------------------------------n/a
'--Update Changelog for release/update------------------------------------------04/29/2022
'--Remove testing message boxes-------------------------------------------------04/29/2022
'--Remove testing code/unnecessary code-----------------------------------------04/29/2022
'--Review/update SharePoint instructions----------------------------------------04/29/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/29/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------n/a
'--Complete misc. documentation (if applicable)---------------------------------n/a
'--Update project team/issue contact (if applicable)----------------------------04/29/2022
