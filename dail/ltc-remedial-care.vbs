'Required for statistical purposes===============================================================================
name_of_script = "DAIL - LTC - REMEDIAL CARE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 30          'manual run time in seconds
STATS_denomination = "I"       'I is for item
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
call changelog_update("06/05/2024", "Updated remedial care amount to $278.00 for July 2024.", "Ilse Ferris, Hennepin County")
call changelog_update("12/03/2023", "Updated remedial care amount to $275.00 for January 2024.", "Ilse Ferris, Hennepin County") ''#873
call changelog_update("06/06/2023", "Updated remedial care amount to $271.00 for July 2023.", "Ilse Ferris, Hennepin County") ''#873
call changelog_update("06/17/2022", "Updated remedial care amount to $234.00 for July 2022.", "Ilse Ferris, Hennepin County") ''#873
call changelog_update("01/03/2022", "Updated remedial care amount to $195.00 for January 2022.", "Ilse Ferris, Hennepin County")
call changelog_update("06/10/2021", "Updated remedial care amount to $189.00 for July 2021.", "Ilse Ferris, Hennepin County")
call changelog_update("12/07/2020", "Updated remedial care amount to $177.00 for January 2021.", "Ilse Ferris, Hennepin County")
call changelog_update("06/04/2020", "Updated remedial care amount to $176.00 for July 2020.", "Ilse Ferris, Hennepin County")
call changelog_update("12/11/2019", "Updated back-end funcationality. Added error reporting option.", "Ilse Ferris, Hennepin County")
call changelog_update("12/02/2019", "Updated remedial care amount to $185.00 for January 2020.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2018", "Updated remedial care amount to $196.00 for 2019.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
'EPM Reference for Remedial Care: http://hcopub.dhs.state.mn.us/epm/appendix_f.htm?rhhlterm=remedial%20care&rhsearch=remedial%20care
remedial_care_amt = "278.00"	'Amount that needs to be updated with current remedial care amount.
target_date = "07/01/2024" 'This date is the 1st possible date that a span can be set at for current COLA span updates. This needs to be updated in code at each COLA (Dec for Jan & June for July.)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 256, 65, "LTC Remedial Care BILS Panel Updater"
ButtonGroup ButtonPressed
OkButton 165, 45, 40, 15
CancelButton 210, 45, 40, 15
Text 10, 15, 240, 20, "This script will update the STAT/BILS panel(s) if remedial care (27) entries exist The rate will update to the current deduction standard of $" & remedial_care_amt &"."
GroupBox 5, 5, 245, 35, "About the Script:"
EndDialog

Do
    Dialog dialog1
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call write_value_and_transmit("S", 6, 3)
'PRIV Handling
EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
If priv_check = "PRIVIL" THEN script_end_procedure("This case is privileged. The script will now end.")
EMReadscreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")

Call write_value_and_transmit("BILS", 20, 71)
PF9 'into edit mode

Do
    EMReadScreen page_number, 2, 3, 72
    If page_number = " 1" then exit do
    PF19
Loop until page_number = " 1"

updates_made = 0 'Setting the variable count
bils_row = 6

Do
    EMReadScreen BILS_line, 54, bils_row, 26    'Reading BILS line from 'Ref Nbr' through 'Dpd Ind'
    BILS_line = replace(BILS_line, "$", " ")
    BILS_line = split(BILS_line, "  ") 'splitting elements into an array
        'Array positions
        '0 = Ref Nbr
        '1 = Date
        '2 = Serv (code)
        '3 = Gross ($ amt)
        '4 = Third Payments
        '5 = Ver (code)
    BILS_line(1) = replace(BILS_line(1), " ", "/") 'changing format to be recognized as a date
    If IsDate(BILS_line(1)) = False then exit do

    If datediff("d", target_date, BILS_line(1)) => 0 and BILS_line(2) = 27 and BILS_line(5) <> remedial_care_amt then
        EMWriteScreen remedial_care_amt, bils_row, 48
        EMWriteScreen "C", bils_row, 24
        updates_made = updates_made + 1
        stats_counter = stats_counter + 1
    End if

    bils_row = bils_row + 1
    BILS_line = ""

    If bils_row = 18 then
        PF20
        bils_row = 6
    End If

    EMReadScreen current_page, 1, 3, 73
    EMReadScreen total_pages, 1, 3, 78
Loop until cint(current_page) = cint(total_pages)

PF3
PF3
stats_counter = stats_counter - 1 'get get true count of stats

If updates_made <> 0 then
    script_end_procedure_with_error_report("Success! Updates made: " & updates_made & ".")
elseif updates_made = 0 then
    script_end_procedure_with_error_report("No remedial care entries found to update. You may have already updated this case or need to add new BILS. Use ACTIONS - BILS UPDATER to add new BILS.")
End if

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/17/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/17/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/17/2022------------------N/A
'--All variables in dialog match mandatory fields-------------------------------06/17/2022------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/17/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------06/17/2022------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/17/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/17/2022------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------06/17/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------06/17/2022
'--Out-of-County handling reviewed----------------------------------------------06/17/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/17/2022
'--BULK - review output of statistics and run time/count (if applicable)--------06/17/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/17/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/17/2022
'--Incrementors reviewed (if necessary)-----------------------------------------06/17/2022
'--Denomination reviewed -------------------------------------------------------06/17/2022
'--Script name reviewed---------------------------------------------------------06/17/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------06/17/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/17/2022
'--comment Code-----------------------------------------------------------------06/17/2022
'--Update Changelog for release/update------------------------------------------06/17/2022
'--Remove testing message boxes-------------------------------------------------06/17/2022
'--Remove testing code/unnecessary code-----------------------------------------06/17/2022
'--Review/update SharePoint instructions----------------------------------------06/17/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/17/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/17/2022
'--Complete misc. documentation (if applicable)---------------------------------06/17/2022
'--Update project team/issue contact (if applicable)----------------------------06/17/2022------------------N/A
