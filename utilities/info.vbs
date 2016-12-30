'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - INFO.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Variables to declare---------------------------------------------------------------------
'The amount of active scriptwriters needs to go in these arrays manually. Subtract one because 0 counts!
Dim scriptwriter_array(17)
Dim email_button_array(17)

'This creates a scriptwriter class. This class can be used to easily recall info about each scriptwriter. It's cleaner than a straight array.
class scriptwriter
	public name
	public agency
	public role
	public formerrole
	public email
end class

'Using this through adding the remaining scriptwriters' info
scriptwriter_counter = 0

'Setting each scriptwriter in alphabetical order by last name, with DHS staff at the top and county staff following

'Veronica Cary, DHS
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			= 	"Veronica Cary"
scriptwriter_array(scriptwriter_counter).agency			= 	"DHS"
scriptwriter_array(scriptwriter_counter).role			= 	"PRISM Project Manager"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"MAXIS Project Manager"
scriptwriter_array(scriptwriter_counter).email			= 	"Veronica.Cary@state.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Charles Potter, DHS
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Charles Potter"
scriptwriter_array(scriptwriter_counter).agency			= 	"DHS"
scriptwriter_array(scriptwriter_counter).role			= 	"MAXIS Project Manager"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"Finanacial Worker/Mentor"
scriptwriter_array(scriptwriter_counter).email			= 	"Charles.D.Potter@state.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'David Courtright, St. Louis
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"David Courtright"
scriptwriter_array(scriptwriter_counter).agency			= 	"St. Louis County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"CourtrightD@StLouisCountyMN.gov"
scriptwriter_counter = scriptwriter_counter + 1

'Tim Delong, Stearns
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Tim Delong"
scriptwriter_array(scriptwriter_counter).agency			= 	"Stearns County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"Timothy.Delong@co.stearns.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Megan Dietz, Anoka
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Megan Dietz"
scriptwriter_array(scriptwriter_counter).agency			= 	"Anoka County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"megan.dietz@co.anoka.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Travis Farleigh, Carlton
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Travis Farleigh"
scriptwriter_array(scriptwriter_counter).agency			= 	"Carlton County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"Travis.Farleigh@co.carlton.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Ilse Ferris, Hennepin
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Ilse Ferris"
scriptwriter_array(scriptwriter_counter).agency			= 	"Hennepin County"
scriptwriter_array(scriptwriter_counter).role			= 	"Case Reviewer"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).email			= 	"Ilse.Ferris@hennepin.us"
scriptwriter_counter = scriptwriter_counter + 1

'Robert Fewins-Kalb, Anoka
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Robert Fewins-Kalb"
scriptwriter_array(scriptwriter_counter).agency			= 	"Anoka County"
scriptwriter_array(scriptwriter_counter).role			= 	"Business Analyst"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).email			= 	"Robert.Kalb@co.anoka.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Melissa Fox, Stearns
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Melissa Fox"
scriptwriter_array(scriptwriter_counter).agency			= 	"Stearns County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"MELISSA.FOX@CO.STEARNS.MN.US"
scriptwriter_counter = scriptwriter_counter + 1

'Kelly Hiestand, Wright
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Kelly Hiestand"
scriptwriter_array(scriptwriter_counter).agency			= 	"Wright County"
scriptwriter_array(scriptwriter_counter).role			= 	"Case Aide"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"kelly.hiestand@co.wright.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Devonne Kent, Wright
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Devonne Kent"
scriptwriter_array(scriptwriter_counter).agency			= 	"Wright County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"devonne.kent@co.wright.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Kelly Kobbervig, Ramsey
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Kaley Kobbervig"
scriptwriter_array(scriptwriter_counter).agency			= 	"Ramsey County"
scriptwriter_array(scriptwriter_counter).role			= 	"Project Manager"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).email			= 	"Kaley.Kobbervig@CO.RAMSEY.MN.US"
scriptwriter_counter = scriptwriter_counter + 1

'Laura Larson, Olmsted
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Laura Larson"
scriptwriter_array(scriptwriter_counter).agency			= 	"Olmsted County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"larson.laura@co.olmsted.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Kenny Lee, Ramsey
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Kenny Lee"
scriptwriter_array(scriptwriter_counter).agency			= 	"Ramsey County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"kenneth.a.lee@CO.RAMSEY.MN.US"
scriptwriter_counter = scriptwriter_counter + 1

'Casey Love, Ramsey
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Casey Love"
scriptwriter_array(scriptwriter_counter).agency			= 	"Ramsey County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"casey.love@co.ramsey.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Lucas Shanley, St. Louis
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Lucas Shanley"
scriptwriter_array(scriptwriter_counter).agency			= 	"St. Louis County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"ShanleyL@StLouisCountyMN.gov"
scriptwriter_counter = scriptwriter_counter + 1

'Gay Sikkink, Stearns
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Gay Sikkink"
scriptwriter_array(scriptwriter_counter).agency			= 	"Stearns County"
scriptwriter_array(scriptwriter_counter).role			= 	"Statistician"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"gay.sikkink@co.stearns.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Roy Walz, Stearns
set scriptwriter_array(scriptwriter_counter) = new scriptwriter
scriptwriter_array(scriptwriter_counter).name			=	"Roy Walz"
scriptwriter_array(scriptwriter_counter).agency			= 	"Stearns County"
scriptwriter_array(scriptwriter_counter).role			= 	"Financial Worker"
scriptwriter_array(scriptwriter_counter).formerrole		= 	"-"
scriptwriter_array(scriptwriter_counter).email			= 	"roy.walz@co.stearns.mn.us"
scriptwriter_counter = scriptwriter_counter + 1

'Here's the actual dialog---------------------------------------------
'Text layout: X, Y, size X, size Y
BeginDialog Dialog1, 0, 0, 375, 450, "DHS BlueZone Scripts Info dialog"
  ButtonGroup ButtonPressed
    OkButton 320, 430, 50, 15

	'header
	Text 5, 10, 370, 10, "============================= BLUEZONE SCRIPTS INFORMATION ============================="

	'We want to see right away if we're using the master branch
	If use_master_branch = true then
		Text 290, 30, 80, 10, "Branch used: MASTER"
	Else
		Text 290, 30, 80, 10, "Branch used: RELEASE"
	End if

	'This stuff is all pulled from global variables
  	Text 5, 30, 200, 10, "Scripts most recent install date: " & scripts_updated_date
	Text 5, 40, 365, 10, "Worker county code: " & worker_county_code
	Text 5, 50, 365, 10, "Agency code from installer: " & code_from_installer
	Text 5, 60, 365, 10, "User logged in: " & windows_user_ID
	Text 5, 70, 365, 10, "EDMS choice: " & EDMS_choice
	Text 5, 80, 365, 10, "BNDX variance threshold: $" & county_bndx_variance_threshold
	Text 5, 90, 365, 10, "Emergency ''percent rule'' amount: " & emer_percent_rule_amt & "%"
	Text 5, 100, 365, 10, "Number of days-worth-of-income to be verified for emergency: " & emer_number_of_income_days
	Text 5, 110, 365, 10, "CLS x1 number: " & CLS_x1_number

	'The users who select a worker is either set to True (for everyone in the agency), or set to False and manually entered into global variables. This reads off who's covered by that.
	If all_users_select_a_worker = True then
		Text 5, 120, 365, 30, "Nav scripts users set to select a worker mode: ALL"
	Else
		For each user in users_using_select_a_user
			user_string = user_string & user & ", "
		Next
		'This handy trick lops off the last character without tons of complex code, and adds a handy "none! if it's no users"
		user_string = user_string & "none!"
		user_string = replace(user_string, ", none!", "")
		Text 5, 120, 365, 30, "Nav scripts users set to select a worker mode: " & user_string
	End if

	'This should only be tripped by scriptwriters who are running scripts locally
	If run_locally = true then
		Text 5, 150, 300, 10, "==========================================================================="
		Text 5, 160, 300, 10, ">>>>>>>>>>>SCRIPTS ARE RUNNING LOCALLY, NOT THROUGH GITHUB<<<<<<<<<<<"
		Text 5, 170, 300, 10, "     PROCEED WITH CAUTION, KNOWING YOUR SCRIPTS ARE RUNNING LOCALLY"
		Text 5, 180, 300, 10, "==========================================================================="
	Else
		Text 5, 150, 300, 10, "Your BlueZone Scripts are being loaded from the repository at GitHub.com."
		Text 5, 160, 300, 10, "         Please note: no client data is EVER passed through GitHub.com. Github"
		Text 5, 170, 300, 10, "         is used a storage medium for the latest scripts, and was approved for our"
		Text 5, 180, 300, 10, "         use by state IT in 2014."
	End if

	'Here's some logic to create a list of scriptwriters based on the above info--------------
	'First some headers
	Text 5, 200, 370, 10, 	"========================= LIST OF SCRIPTWRITERS AS OF 08/01/2016 ========================="
  	Text 5, 210, 70, 10, "---NAME---"
  	Text 75, 210, 40, 10, "---AGENCY---"
  	Text 155, 210, 90, 10, "---CURRENT ROLE---"
  	Text 245, 210, 90, 10, "---FORMER ROLE---"
  	Text 335, 210, 35, 10, "---EMAIL---"

	'This loop takes info from above and turns it into coordinates on the dialog
	For i = 0 to ubound(scriptwriter_array)
  		y_pos = (i * 10) + 220
  		Text 5, y_pos, 70, 10, scriptwriter_array(i).name
  		Text 75, y_pos, 80, 10, scriptwriter_array(i).agency
  		Text 155, y_pos, 90, 10, scriptwriter_array(i).role
  		Text 245, y_pos, 90, 10, scriptwriter_array(i).formerrole
		If scriptwriter_array(i).email <> "" THEN PushButton 335, y_pos, 35, 10, "email", email_button_array(i)
	Next
EndDialog

'Shows the dialog
Dialog

'If the ButtonPressed wasn't OK or cancel, it ended because one of the email buttons was hit. This uses "mailto:" and a shell object to load a blank email addressed to the scriptwriter
If ButtonPressed <> OK and ButtonPressed <> Cancel then
	For i = 0 to ubound(email_button_array)
		If ButtonPressed = email_button_array(i) then CreateObject("WScript.Shell").Run("mailto:" & scriptwriter_array(i).email)
	Next
End if

'ends the script
script_end_procedure("")
