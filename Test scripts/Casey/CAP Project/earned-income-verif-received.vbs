'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PAYSTUBS RECEIVED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 473                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("04/23/2018", "Fixed bug in which the lines of the PIC were dupicated in the case note.", "Casey Love, Hennepin County")
CALL changelog_update("12/07/2017", "Removed condition to allow paystubs dated with the current date to be accepted. Updated code to write JOBS verification code in.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/11/2017", "The script has been updated to write to the GRH PIC and to case note that the GRH PIC has been updated.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Find case number and footer month - set a variable with the initially found footer month for the default for every loop
'The footer month may be different for EVERY income source. NEED to add handling to identify if there is a begin date for updating MAXIS (app date that activates the case)
    'A client may apply in april and bring in checks from March but we cannot update MAXIS in March


'DIALOG TO GET CASE NUMBER
'Possibly add worker signature here and take it out of the following dialogs
BeginDialog Dialog1, 0, 0, 191, 105, "Dialog"
  EditBox 110, 5, 70, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 85, 85, 50, 15
    CancelButton 140, 85, 50, 15
  Text 15, 10, 85, 10, "Enter your case number:"
  GroupBox 5, 25, 175, 55, "INSTRUCTIONS - PLEASE READ!!!"
  Text 15, 40, 155, 30, "This script will allow you to update any JOBS/BUSI/RBIC on a case. It can process multiple panels in one run. "
EndDialog

'CREATE ARRAY OF ALL EI panels'
'Put them in a 'FOR-NEXT' to loop through each panel.
'IF all income will be case noted as 1 note then create an ARRAY of all the case note information.


'NAVIGATE TO JOBS for each HH MEMBER and ask if Income information was received for this job.

'This will become dynamic and there will be an array of all the checks listed.
'STILL need some handling for scheduled income with no actual checks or cases where scheduled income is different from actual checks but we get both.
'NEED TO ADD CHECKBOXES FOR PROGRAMS THIS INCOME APPLIES TO - and precheck all the programs that are active on this case'
BeginDialog Dialog1, 0, 0, 656, 135, "Enter ALL Paychecks Received"
  Text 10, 10, 265, 10, "JOBS 01 01 - EMPLOYER"
  Text 10, 30, 60, 10, "JOBS Verif Code:"
  DropListBox 80, 25, 105, 45, "", JOBS_verif_code
  Text 195, 30, 120, 10, "further detail of verification received:"
  EditBox 320, 25, 315, 15, Edit2
  Text 10, 50, 90, 10, "Date verification received:"
  EditBox 105, 45, 50, 15, verif_date
  Text 5, 70, 80, 10, "Pay Date (MM/DD/YY):"
  Text 90, 70, 50, 10, "Gross Amount:"
  Text 145, 70, 25, 10, "Hours:"
  Text 180, 55, 25, 25, "Use in SNAP budget"
  Text 245, 70, 85, 10, "If not used, explain why:"
  Text 360, 55, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
  Text 360, 70, 30, 10, "Amount:"
  Text 415, 70, 30, 10, "Reason:"
  EditBox 10, 85, 65, 15, pay_date
  EditBox 90, 85, 45, 15, gross_amount
  EditBox 145, 85, 25, 15, hours_on_check
  OptionGroup RadioGroup1
    RadioButton 180, 85, 25, 10, "Yes", budget_yes
    RadioButton 210, 85, 25, 10, "No", budget_no
  EditBox 245, 85, 105, 15, reason_not_budgeted
  EditBox 360, 85, 45, 15, not_budgeted_amount
  EditBox 415, 85, 185, 15, amount_not_budgeted_reason
  Text 10, 115, 60, 10, "Worker signature:"
  EditBox 80, 110, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 490, 110, 15, 15, "+", add_another_check
    PushButton 510, 110, 15, 15, "-", take_a_check_away
    OkButton 545, 110, 50, 15
    CancelButton 600, 110, 50, 15
EndDialog

'Script will determine pay frequency and potentially 1st check (if not listed on JOBS)
'Script will determine the initial footer month to change by the pay dates listed.
'Script will create a budget based on the program this income applies to
'Dialog the budget and have the worker confirm - if they decline - pull the check list dialog back up and have them adjust it there.
'Worker must confirm the frequency, first pay, and footer month
'Worker will inicate if future months should be updated - default this to 'yes' as script will update retro and prospective specific to each month
'SNAP PIC, GRH PIC, HC EI EST will be checked to be updated IF any of these programs are open on the case.

'NEED to add handling for future/current changes - start or stop work - get policy on this from SNAP refresher - talk to Melissa.

'NAVIGATE to BUSI for each HH MEMBER and ask if Income Information was received for this Self Employment.
'NAVIGATE to RBIC for each HH MEMBER and ask if Income Information was received for this RBIC
