'Created by Tim DeLong from Stearns County and Ilse Ferris from Hennepin County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - HOUSING GRANT FIATER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 175                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'Function not yet added to the FuncLib----------------------------------------------------------------------------------------------------
FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an initial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date

	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'DIALOG===========================================================================================================================
BeginDialog housing_grant_dialog, 0, 0, 271, 215, "MFIP Housing Grant FIATER"
  EditBox 65, 10, 60, 15, case_number
  EditBox 210, 10, 25, 15, initial_month
  EditBox 240, 10, 25, 15, initial_year
  ButtonGroup ButtonPressed
    OkButton 160, 190, 50, 15
    CancelButton 215, 190, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 145, 15, 60, 10, "Initial month/year:"
  Text 15, 80, 100, 10, "* Caregivers age 60 or older"
  GroupBox 5, 35, 260, 150, "MFIP Housing Grant $50 earned income exemption"
  Text 15, 50, 245, 20, "Only certain people are eligible for the housing grant $50 unearned income exemption. These recipients include:"
  Text 15, 95, 165, 10, "* Caregivers caring for a disabled family member"
  Text 15, 110, 175, 10, "* Caregivers who meet Special Medical Criteria (SMC)"
  Text 15, 125, 245, 20, "* Caregivers who are disabled and do not anticipated being able to work for 20+ hours for more than 30 days"
  Text 15, 150, 100, 10, "* Caregivers who receive SSI"
  Text 15, 165, 180, 10, "* Caregivers who receive Mille Lacs Band Tribal TANF"
EndDialog

'============================================================================================================================
'Connects to MAXIS, grabbing the case case_number
EMConnect ""
Call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

DO
	DO
		err_msg = ""
		dialog housing_grant_dialog
		If buttonpressed = 0 THEN stopscript
		IF len(case_number) > 8 or isnumeric(case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."
		IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial month."
		IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial year."
		IF err_msg <> "" THEN msgbox err_msg & vbCr & "Please resolve to continue."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call MAXIS_footer_month_confirmation		'checking to make sure footer month/year selected by user is the footer month that MAXIS is in, if not it corrects it to the user selected footer month'

'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.
	'We will need to remove the string "/1/" from each element in the array
call date_array_generator(initial_month, initial_year, footer_month_array)

'Create an array of all the counted months
DIM MFIP_months_array()
REDIM MFIP_months_array(ubound(footer_month_array))

'Need to make sure we start in the correct year for maxis'
footer_month = initial_month
footer_year = initial_year














'Navigates to FIAT and selects MFIP.
CALL navigate_to_MAXIS_screen("FIAT", "____")
	EMwritescreen "03", 4, 34
	EMwritescreen "x", 9, 22
	transmit

'Selects View Case Budget.
	EMwritescreen "x", 18, 4
	transmit
'Selects the Subsidy/Tribal pop-up then the Housing Subsidy sub-pop-up
	EMwritescreen "x", 17, 5
	transmit
	EMwritescreen "x", 8, 13
	transmit
'Changes the prospective column to $0
	EMwritescreen "0       ", 8, 51
	transmit
	transmit
	transmit

'script ends where the worker can see if the housing grant is showing as eligible and pops up a msg box to do so.
script_end_procedure ("Verify that the results showing are what were expected." & vbNewline & vbNewline &_
	"If results are correct, PF3 twice to exit FIAT then retain results." & vbNewline & vbNewline &_
	"Run the script for any other months needed and approve.")
