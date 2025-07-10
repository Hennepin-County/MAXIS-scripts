'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------
'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the
'	installer, but doing so should not be undertaken lightly.

'CONFIG FOR HOW SCRIPTS WORK===================

'Default directory: used by the script to determine if we're scriptwriters or not (scriptwriters use a default directory traditionally).
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.
default_directory = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\"

'Run locally: if this is set to "True", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts.
run_locally = true

'========================================================================================================================================

'COUNTY NAME AND INFO==========================

'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts).
worker_county_code = "X127"

'This merely exists to help the installer determine which dropdown box to default. It is not used by any scripts.
code_from_installer = "27 - Hennepin County"

'This is an "updated date" variable, which is updated dynamically by the intaller.
scripts_updated_date = "01/01/2099"

'ALL-COUNTY SCRIPT CONFIGURATION===============

'This is used by scripts which tell the worker where to find a doc to send to a client (ie "Send form using Compass Pilot")
EDMS_choice = "ECF"

'This is used to allow some agencies to decline to case note intake/rein dates on denied progs and closed progs. We're hoping to convince these agencies to case note this info, so that we can drop this field.
case_noting_intake_dates = False

'This moves "verifs needed" to be at the top of the CAF case note template, instead of the bottom.
move_verifs_needed = False

'This threshold is what the BNDX scrubber will use to determine what is considered "within the realm" of the currently budgeted income.
county_bndx_variance_threshold = "0"

'These two variables determine the percent rule for EA/EGA, as well as the number of days income is evaluated.
emer_percent_rule_amt = "30"
emer_number_of_income_days = "30"

'This is the X1/PW number to send closed cases to in the INAC scrubber.
CLS_x1_number = ""

'This is a setting for the TYMA TIKLer script. When set to "true", TYMA TIKLer will TIKL all TYMA months simultaneously, as opposed to only the first month. Defaullt is "false".
TYMA_TIKL_all_at_once = false

'This is a setting to determine if changes to scripts will be displayed in messageboxes in real time to end users
changelog_enabled = true

'NAVIGATION SCRIPT CONFIGURATION================

'If all users use "select a worker" nav scripts, this will be True. (Example: case banking county)
all_users_select_a_worker = True

'If the above is False, we need a list of workers who do use the "select a worker" nav scripts.
users_using_select_a_user = array()

'DATABASE SETTINGS AND DETAILS=========================

'specific settings for the database schema and tables to use.
db_provider = "SQLOLEDB.1"
db_data_source = "hssqlpw139"
db_catalog = "BlueZone_Statistics"
db_security = "SSPI"
db_translate = "False"

'string to use for database calls in scripts.
db_full_string = "Provider = " & db_provider & ";Data Source= " & db_data_source & ";Initial Catalog= " & db_catalog & "; Integrated Security=" & db_security & ";Auto Translate=" & db_translate & ";"

'COLLECTING STATISTICS=========================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = true

'This is a variable used to determine if the agency is using a SQL database or not. Set to true if you're using SQL. Otherwise, set to false.
using_SQL_database = true

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
stats_database_path = db_data_source & ";Initial Catalog= " & db_catalog & "; Integrated Security=" & db_security & ";Auto Translate=" & db_translate & ";"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = true

'If set to true, the case number will be collected and input into the database
collect_MAXIS_case_number = true


'ABAWD BANKED MONTHS TRACKING CONFIG================

'This determines whether-or-not banked months tracking happens at all
banked_months_db_tracking = false

'Add the path to the database file using banked_month_database_path, replacing this path with wherever you have the file installed
banked_month_database_path = ""


'BRANCH CONFIGURATION=====================

'This is a variable which sets the scripts to use the master branch (common with scriptwriters)
use_master_branch = False

'TRAINING CASE SCENARIO SETTINGS==========

'This is a variable which decides the default location of training case scenario Excel sheets
' training_case_creator_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Instructions & Training Materials\Training materials\Training cases\TRAINING CASE SCENARIOS.xlsx"
training_case_creator_excel_file_path = "T:\Eligibility Support\Scripts\Script Files\TRAINING CASE SCENARIOS.xlsx"

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

'If run locally is set to "True", the scripts will totally bypass GitHub and run locally.
IF run_locally = TRUE THEN script_repository = "C:\MAXIS-Scripts\"

'LOADING LIST OF TESTERS====================================================================================================
tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
'END LIST OF TESTERS BLOCK==================================================================================================
