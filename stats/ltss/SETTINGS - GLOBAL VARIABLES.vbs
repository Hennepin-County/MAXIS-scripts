'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------

'COUNTY NAME AND INFO==========================

'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts). | Example x186 or x106
'worker_county_code = "x118"

' Enter your county number followed by County Name | Example 86 - Wright County
'code_from_installer = "18 - Crow Wing County"


'ALL-COUNTY SCRIPT CONFIGURATION===============

'This is used by scripts which tell the worker where to find a doc to send to a client (ie "Send form using Compass Pilot") | Sample "OnBase" "DHS eDocs" or "Compass Pilot"
'EDMS_choice = "Compass Forms"

'This is used to allow some agencies to decline to case note intake/rein dates on denied progs and closed progs. We're hoping to convince these agencies to case note this info, so that we can drop this field.
'case_noting_intake_dates = True

'This moves "verifs needed" to be at the top of the CAF case note template, instead of the bottom.
'move_verifs_needed = False

'This threshold is what the BNDX scrubber will use to determine what is considered "within the realm" of the currently budgeted income.
'county_bndx_variance_threshold = "0"

'These two variables determine the percent rule for EA/EGA, as well as the number of days income is evaluated.
'emer_percent_rule_amt = "30"
'emer_number_of_income_days = "30"

'This is the X1/PW number to send closed cases to in the INAC scrubber.
'CLS_x1_number = ""

'This is a TRUE/FALSE that will tell the INAC scrubber to hold onto MAGI cases that closed for no/incomplete renewals for 4 months or not.
'MAGI_cases_closed_four_month_TIKL_no_XFER = FALSE

'This is a setting for the TYMA TIKLer script. When set to "true", TYMA TIKLer will TIKL all TYMA months simultaneously, as opposed to only the first month. Defaullt is "false".
'TYMA_TIKL_all_at_once = false

'NAVIGATION SCRIPT CONFIGURATION================

'If all users use "select a worker" nav scripts, this will be True. (Example: case banking county)
'all_users_select_a_worker = True

'If the above is False, we need a list of workers who do use the "select a worker" nav scripts.
'users_using_select_a_user = array("USERS_USING_SELECT_A_USER=ARRAY()")




'COLLECTING STATISTICS========================= OPTIONAL SETTINGS ======================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = False

'This is a variable used to determine if the agency is using a SQL database or not. Set to true if you're using SQL. Otherwise, set to false.
using_SQL_database = False

'This is the file path for the statistics Access database.
stats_database_path = script_repository & "Databases for script usage\usage statistics.accdb"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = false


'ABAWD BANKED MONTHS TRACKING CONFIG================ OPTIONAL SETTINGS ======================

'This determines whether-or-not banked months tracking happens at all
'banked_months_db_tracking = true

'Add the path to the database file using banked_month_database_path, replacing this path with wherever you have the file installed
'banked_month_database_path = "C:\Users\Public\Documents\BlueZone\Scripts\State of Minnesota\MAXIS Scripts\Databases for script usage\banked month tracking.accdb"


'TRAINING CASE SCENARIO SETTINGS========== OPTIONAL SETTINGS ======================

'This is a variable which decides the default location of training case scenario Excel sheets
'training_case_creator_excel_file_path = script_repository & "SETTINGS - TRAINING CASE SCENARIOS.xlsx"













' ############################  DO NOT EDIT BELOW THIS LINE UNLESS YOU KNOW WHAT YOU ARE DOING ###################################

'BRANCH CONFIGURATION=====================

'This is a variable which sets the scripts to use the master branch (common with scriptwriters)
use_master_branch = false


'CONFIG FOR HOW SCRIPTS WORK===================
'This is a variable to set the contact information for any problems encountered with scripts
contact_admin = "tss.hd.processing@state.mn.us"

'Default directory: used by the script to determine if we're scriptwriters or not (scriptwriters use a default directory traditionally).
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.
default_directory = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\MnChoices"

'Run locally: if this is set to "True", the scripts will run locally and bypass the internet entirely. This is great for debugging or developing scripts.
run_locally = true

'========================================================================================================================================

'This is a variable used to determine the version and release active and the date activated
version_release = "2.00"

'This is an "updated date" variable, which is updated when the version release changes.
scripts_updated_date = "06/26/2020"


'NAVIGATION SCRIPT CONFIGURATION================

'If all users use "select a worker" nav scripts, this will be True. (Example: case banking county)
'all_users_select_a_worker = False

'If the above is False, we need a list of workers who do use the "select a worker" nav scripts.
'users_using_select_a_user = array()


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

