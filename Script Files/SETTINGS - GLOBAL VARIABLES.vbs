'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------
'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the 
'	installer, but doing so should not be undertaken lightly.

'CONFIG FOR HOW SCRIPTS WORK===================

'Default directory: used by the script to determine if we're scriptwriters or not (scriptwriters use a default directory traditionally).
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.
default_directory = "C:\DHS-MAXIS-Scripts\Script Files\"

'Run locally: if this is set to "True", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts.
run_locally = False

'========================================================================================================================================

'COUNTY NAME AND INFO==========================

'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts).
worker_county_code = "x102"

'This is used for MEMO scripts, such as appointment letter
county_name = "Anoka County"

'This merely exists to help the installer determine which dropdown box to default. It is not used by any scripts.
code_from_installer = "02 - Anoka County"

'Creates a double array of county offices, first by office (using the ~), then by address line (using the |). Dynamically added with the installer.
county_office_array = split("2100 3rd Ave Suite 400|Anoka, MN 55303~1201 89th Ave NE Suite 400|Blaine, MN 55434~3980 Central Ave NE|Columbia Heights, MN 55421~4175 Lovell RD NE|Lexington, MN 55014", "~")



'ALL-COUNTY SCRIPT CONFIGURATION===============

'This is used by scripts which tell the worker where to find a doc to send to a client (ie "Send form using Compass Pilot")
EDMS_choice = "Compass Pilot"

'This is used to allow some agencies to decline to case note intake/rein dates on denied progs and closed progs. We're hoping to convince these agencies to case note this info, so that we can drop this field.
case_noting_intake_dates = True

'This moves "verifs needed" to be at the top of the CAF case note template, instead of the bottom.
move_verifs_needed = False

'This threshold is what the BNDX scrubber will use to determine what is considered "within the realm" of the currently budgeted income.
county_bndx_variance_threshold = "1"

'These two variables determine the percent rule for EA/EGA, as well as the number of days income is evaluated.
emer_percent_rule_amt = "30"
emer_number_of_income_days = "30"

'This is the X1/PW number to send closed cases to in the INAC scrubber.
CLS_x1_number = "X102CLS"



'NAVIGATION SCRIPT CONFIGURATION================

'If all users use "select a worker" nav scripts, this will be True. (Example: case banking county)
all_users_select_a_worker = False

'If the above is False, we need a list of workers who do use the "select a worker" nav scripts.
users_using_select_a_user = array("VKC", "VKCARY", "PWVKC45")


'COLLECTING STATISTICS=========================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = False

'This is the file path for the statistics Access database.
stats_database_path = "C:\DHS-MAXIS-Scripts\Statistics\usage statistics.accdb"



'BETA AGENCY CONFIGURATION=====================

'This is a variable which signifies the agency is beta (affects script URL)
beta_agency = True

'========================================================================================================================================
'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------

'Making a list of offices to be used in various scripts
For each office in county_office_array
	new_office_array = split(office, "|")									'Assigned earlier in the FUNCTIONS FILE script. Splits into an array, containing each line of the address.
	comma_location_in_address_line_02 = instr(new_office_array(1), ",")				'Finds the location of the first comma in the second line of the address (because everything before this is the city)
	city_for_array = left(new_office_array(1), comma_location_in_address_line_02 - 1)		'Pops this city into a variable
	county_office_list = county_office_list & chr(9) & city_for_array					'Adds the city to the variable called "county_office_list", which also contains a new line, so that it works correctly in dialogs.
Next


is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'Some screens require the two digit county code, and this determines what that code is. It only does it for single-county agencies 
'(ie, DHS and other multicounty agencies follow different logic, which will be fleshed out in the individual scripts affected)
If worker_county_code <> "MULTICOUNTY" then two_digit_county_code = right(worker_county_code, 2)

'The following code looks in C:\USERS\''windows_user_ID''\My Documents for a text file called workersig.txt.
'If the file exists, it pulls the contents (generated by ACTIONS - UPDATE WORKER SIGNATURE.vbs) and populates worker_signature automatically.
Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = objNet.UserName

Dim oTxtFile 
With (CreateObject("Scripting.FileSystemObject"))
	If .FileExists("C:\users\" & windows_user_ID & "\my documents\workersig.txt") Then
		Set get_worker_sig = CreateObject("Scripting.FileSystemObject")
		Set worker_sig_command = get_worker_sig.OpenTextFile("C:\users\" & windows_user_ID & "\my documents\workersig.txt")
		worker_sig = worker_sig_command.ReadAll
		IF worker_sig <> "" THEN worker_signature = worker_sig	
		worker_sig_command.Close
	END IF
END WITH

'This is the URL of our script repository, and should only change if the agency is beta or standard, or if there's a scriptwriter in the group.
IF default_directory <> "C:\DHS-MAXIS-Scripts\Script Files\" THEN	'For folks who did NOT install from GitHub, which uses this path...
	IF beta_agency = TRUE THEN		'Beta agencies use the beta branch
		script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/BETA/Script Files/"
	ELSE							'Everyone else (who isn't a scriptwriter) uses the release branch
		script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/RELEASE/Script Files/"
	END IF
ELSE	'Scriptwriters use the master branch
	script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files/"
END IF

'If run locally is set to "True", the scripts will totally bypass GitHub and run locally. 
IF run_locally = TRUE THEN script_repository = "C:\DHS-MAXIS-Scripts\Script Files"
