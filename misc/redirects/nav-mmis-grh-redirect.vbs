'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------
'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the
'	installer, but doing so should not be undertaken lightly.

'CONFIG FOR HOW SCRIPTS WORK===================

'Default directory: used by the script to determine if we're scriptwriters or not (scriptwriters use a default directory traditionally).
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.

default_directory = "S:\D040\CLS\Comm Living Support\GRH\SSR reduction script process\"

'Run locally: if this is set to "True", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts.
run_locally = false

'========================================================================================================================================

'COUNTY NAME AND INFO==========================

'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts).
worker_county_code = "x191"

'========================================================================================================================================

'COLLECTING STATISTICS=========================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = true

'This is a variable used to determine if the agency is using a SQL database or not. Set to true if you're using SQL. Otherwise, set to false.
using_SQL_database = true

'This is the file path for the statistics Access database.
stats_database_path = "hssqlpw017;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = true

'If set to true, the case number will be collected and input into the database
collect_MAXIS_case_number = true

'BRANCH CONFIGURATION=====================

'This is a variable which sets the scripts to use the master branch (common with scriptwriters)
use_master_branch = False

'This is a setting to determine if changes to scripts will be displayed in messageboxes in real time to end users
changelog_enabled = true

'This is a setting for the script team to identify if they are aware of a git hub outage
git_hub_issue_known = false


'========================================================================================================================================
'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'Some screens require the two digit county code, and this determines what that code is. It only does it for single-county agencies
'(ie, DHS and other multicounty agencies follow different logic, which will be fleshed out in the individual scripts affected)
If worker_county_code <> "MULTICOUNTY" then two_digit_county_code = right(worker_county_code, 2)

'This is the URL of our script repository, and should only change if the agency is a scriptwriting agency. Scriptwriters can elect to use the master branch, allowing them to test new tools, etc.
IF use_master_branch = TRUE THEN		'scriptwriters typically use the master branch
	script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
ELSE							'Everyone else (who isn't a scriptwriter) typically uses the release branch
	script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
END IF

'----------------------------------------------------------------------------------------------------LOADING SCRIPT - REDIRECT FILE 
script_url = script_repository & "/nav/nav-mmis-grh.vbs"
IF run_locally = False THEN
    SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_URL
    req.open "GET", script_URL, FALSE									'Attempts to open the script_URL
    req.send													'Sends request
    IF req.Status = 200 THEN									'200 means great success
    	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
    	Execute req.responseText								'Executes the script code
    ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
    	If git_hub_issue_known = TRUE Then
            MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
            vbCr &_
            "The BlueZone Script Team is aware of an issue on GitHub and are monitoring the progress of the fix." & vbCr &_
            vbCr &_
            "There is no support for NAV scripts at this time. Some essential scripts have been saved locally for access during outages. Press any of the 'MAXIS Script Category' buttons and if the outage is still in effect, the special outage menu will appear to access these exxential scripts."
        End If
    	StopScript
    END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_url)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF