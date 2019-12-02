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
call changelog_update("12/02/2019", "Updated remedial care amount to $185.00 for January 2020.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2018", "Updated remedial care amount to $196.00 for 2019.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'<<<GO THROUGH AND REMOVE REDUNDANT FUNCTIONS
EMConnect ""
remedial_care_amt = "185.00"	'Amount that needs to be updated with current remedial care amount.
target_date = "06/30/2020" 'This sets the date range that should be changed, and will need to be updated in code at each COLA.

Do
    BeginDialog Dialog1, 0, 0, 191, 86, "Dialog"
      ButtonGroup ButtonPressed
      OkButton 135, 10, 50, 15
      CancelButton 135, 30, 50, 15
      Text 10, 5, 115, 50, "This script will update your STAT/BILS panel's remedial care (27) entries, to the current deduction rate of $" & remedial_care_amt & "."
      Text 10, 65, 170, 20, "Press OK to start. Remember to case note when you are finished!"
    EndDialog

    Dialog dialog1
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

EMSendKey "s" & "<enter>"
EMWaitReady 0, 0

EMWriteScreen "bils", 20, 71
EMSendKey "<enter>"
EMWaitReady 0, 0

PF9 'into edit mode 

Do
    EMReadScreen page_number, 2, 3, 72
    If page_number = " 1" then exit do
    PF19
Loop until page_number = " 1"

updates_made = 0 'Setting the variable count 
Do
    bils_row = 6
    msgbox bils_row
    EMReadScreen BILS_line, 54, 6, 26    'Reading BILS line from 'Ref Nbr' through 'Dpd Ind'
    BILS_line = replace(BILS_line, "$", " ")
    BILS_line = split(BILS_line, "  ") 'splitting elements into an array
        'Array positions
        '0 = Ref Nbr
        '1 = Date
        '2 = Serv (code)
        '3 = Gross ($ amt)
        '4 = Third Payments 
        '5 = Ver (code)
    BILS_line(1) = replace(BILS_line(1), " ", "/") 'changing format to be recongized as a date 
    If IsDate(BILS_line(1)) = True then
        If datediff("d", target_date, BILS_line(1)) > 0 and BILS_line(2) = 27 and BILS_line(5) <> remedial_care_amt then
            EMWriteScreen remedial_care_amt, 6, 48
            EMWriteScreen "C", 6, 24
            updates_made = updates_made + 1
        End If
    End If
    
    bils_row = bils_row + 1
    BILS_line = ""

    EMReadScreen current_page, 1, 3, 73
    EMReadScreen total_pages, 1, 3, 78
    If cint(current_page) <> cint(total_pages) then
        PF20
        bils_row = 6
    End If
Loop until cint(current_page) = cint(total_pages)

PF3
PF3

If updates_made <> 0 then MsgBox "Success! Updates made: " & updates_made & "."
If updates_made = 0 then MsgBox "No remedial care entries found. You may have already updated this case! Otherwise, this client may be at their renewal, or no remedial care deduction was made. If this appears to be an error, contact the BlueZone Scripts Team."

script_end_procedure("")