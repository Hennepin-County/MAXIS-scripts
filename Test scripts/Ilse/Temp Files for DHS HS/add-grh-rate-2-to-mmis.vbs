 'This is used by almost every script which calls a specific agency worker number (like the REPT/ACTV nav and list gen scripts).
worker_county_code = "x191"

'This is an "updated date" variable, which is updated dynamically by the intaller.
scripts_updated_date = "01/01/2099"

'This is a setting to determine if changes to scripts will be displayed in messageboxes in real time to end users
changelog_enabled = true

'COLLECTING STATISTICS=========================

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = true

'This is a variable used to determine if the agency is using a SQL database or not. Set to true if you're using SQL. Otherwise, set to false.
using_SQL_database = true

'This is the file path for the statistics Access database.
stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'If the "enhanced database" is used (with new features added in January 2016), this variable should be set to true
STATS_enhanced_db = true

'If set to true, the case number will be collected and input into the database
collect_MAXIS_case_number = true
'Required for statistical purposes===============================================================================
name_of_script = "ACTIONS - ADD GRH RATE 2 TO MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 900                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
'END OF stats block==============================================================================================

'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
Dim checked, unchecked, cancel, OK, blank, t_drive, STATS_counter, STATS_manualtime, STATS_denomination, script_run_lowdown, testing_run, MAXIS_case_number		'Declares this for Option Explicit users

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs
blank = ""
t_drive = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team"

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

'Global function to actually RUN'
If IsArray(tester_array) = True Then Call confirm_tester_information

'Time arrays which can be used to fill an editbox with the convert_array_to_droplist_items function
time_array_15_min = array("7:00 AM", "7:15 AM", "7:30 AM", "7:45 AM", "8:00 AM", "8:15 AM", "8:30 AM", "8:45 AM", "9:00 AM", "9:15 AM", "9:30 AM", "9:45 AM", "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "1:00 PM", "1:15 PM", "1:30 PM", "1:45 PM", "2:00 PM", "2:15 PM", "2:30 PM", "2:45 PM", "3:00 PM", "3:15 PM", "3:30 PM", "3:45 PM", "4:00 PM", "4:15 PM", "4:30 PM", "4:45 PM", "5:00 PM", "5:15 PM", "5:30 PM", "5:45 PM", "6:00 PM")
time_array_30_min = array("7:00 AM", "7:30 AM", "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM")

'Array of all the upcoming holidays
HOLIDAYS_ARRAY = Array(#11/10/23#, #11/23/23#, #11/24/23#, #12/25/23#, #1/1/24#, #1/15/24#, #2/19/24#, #5/27/24#, #6/19/24#, #7/4/24#, #9/2/24#, #11/11/24#, #11/28/24#, #11/29/24#, #12/25/24#)

'Determines CM and CM+1 month and year using the two rightmost chars of both the month and year. Adds a "0" to all months, which will only pull over if it's a single-digit-month
Dim CM_mo, CM_yr, CM_plus_1_mo, CM_plus_1_yr, CM_plus_2_mo, CM_plus_2_yr, CM_plus_3_mo, CM_plus_3_yr, CM_minus_1_mo, CM_minus_1_yr, CM_minus_2_mo, CM_minus_2_yr, CM_minus_3_mo, CM_minus_3_yr
'var equals...  the right part of...    the specific part...    of either today or next month... just the right 2 chars!
CM_mo =         right("0" &             DatePart("m",           date                             ), 2)
CM_yr =         right(                  DatePart("yyyy",        date                             ), 2)

CM_plus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", 1, date)            ), 2)
CM_plus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", 1, date)            ), 2)

CM_plus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", 2, date)            ), 2)
CM_plus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", 2, date)            ), 2)

CM_plus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", 3, date)            ), 2)
CM_plus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", 3, date)            ), 2)

CM_minus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)

CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)

CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)

If worker_county_code   = "" then worker_county_code = "MULTICOUNTY"
IF PRISM_script <> true then county_name = ""		'VKC NOTE 08/12/2016: ADDED IF...THEN CONDITION BECAUSE PRISM IS STILL USING THIS VARIABLE IN ALL SCRIPTS.vbs. IT WILL BE REMOVED AND THIS CAN BE RESTORED.

If ButtonPressed <> "" then ButtonPressed = ""		'Defines ButtonPressed if not previously defined, allowing scripts the benefit of not having to declare ButtonPressed all the time

function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ access_type:indicates if we need to read or write to the ADDR panel, only options are 'READ' or 'WRITE' - case does not matter - the function wil ucase it
'~~~~~ notes_on_address: string - information read from the panel and used typically in an editbox
'~~~~~ resi_line_one: string - information from the first line of the residence street address
'~~~~~ resi_line_two: string - information from the second line of the residence street address
'~~~~~ resi_street_full: string - information from the first and second line of the residence street address
'~~~~~ resi_city: string - information from the city line of the residence address
'~~~~~ resi_state: string - information from the state line of the residence address - this is formatted as MN - Minnesota
'~~~~~ resi_zip: string - information of the zip code from the residence address
'~~~~~ resi_county: string - information of the county of the residence address - this is formatted as 27 Hennepin
'~~~~~ addr_verif: string - information of the verification listed on ADDR - formattes as SF - Shlter Form
'~~~~~ addr_homeless: string of 'Yes' or 'No' from the homeless yes/no code on ADDR
'~~~~~ addr_reservation: string of 'Yes' or 'No' from the reservation yes/no code on ADDR
'~~~~~ addr_living_sit: string - information of the living situation - formatted as 01 - Own Home, or as Blank
'~~~~~ reservation_name: string - information of the reservation name - formatted as LL - Leech Lake
'~~~~~ mail_line_one: string - information from first line of street address of the mailing address
'~~~~~ mail_line_two: string - information from the second line of street address of the mailing address
'~~~~~ mail_street_full: string - information of both the first and second line of the street address of the mailing address
'~~~~~ mail_city: string - information of the city listed on the mailing address
'~~~~~ mail_state: string - information of the state listed on the mailing address - this is formatted as MN - Minnesota
'~~~~~ mail_zip: string - information of the zip code listed on the mailing address
'~~~~~ addr_eff_date: string - the date listed as the effective date on the top of the panel. Formatted as MM/DD/YY - script may read as a date
'~~~~~ addr_future_date: string - the date listed at the top if there is a future date of a change in a future month. formatted as MM/DD/YY - script may read as a date
'~~~~~ phone_one: string - information listed on the first phone line - formatted as xxx-xxx-xxxx
'~~~~~ phone_two: string - information listed on the second phone line - formatted as xxx-xxx-xxxx
'~~~~~ phone_three: string - information listed on the third phone line - formatted as xxx-xxx-xxxx
'~~~~~ type_one: string - information listed for the type of the first phone  - formatted as C - Cell
'~~~~~ type_two: string - information listed for the type of the second phone  - formatted as C - Cell
'~~~~~ type_three: string - information listed for the type of the third phone  - formatted as C - Cell
'~~~~~ text_yn_one: string - information listed as the authorization on the first phone for texting - formatted as Y or N
'~~~~~ text_yn_two: string - information listed as the authorization on the second phone for texting - formatted as Y or N
'~~~~~ text_yn_three: string - information listed as the authorization on the third phone for texting - formatted as Y or N
'~~~~~ addr_email: string - information listed on the email line. any trailing '_' will be trimmed off but NOT any in the middle
'~~~~~ verif_received: string - StRING - used as entry on WRITE only
'~~~~~ original_information: string that saves the original detail known on the panel - this is used to see if there is an update to any information
'~~~~~ update_attempted: boolean - output after 'WRITE' to indicate that an update of panel information was attempted.
'===== Keywords: MAXIS, update, ADDR
	access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")							'going to ADDR
		EMReadScreen curr_addr_footer_month, 2, 20, 55							'reading the footer month and year on the panel because there is footer months pecific differences
		EMReadScreen curr_addr_footer_year, 2, 20, 58
		the_footer_month_date = curr_addr_footer_month &  "/1/" & curr_addr_footer_year		'making a date out of the footer month
		the_footer_month_date = DateAdd("d", 0, the_footer_month_date)
		new_version = True														'defaulting to using the new verion of the panel with text authorization and email
		If DateDiff("d", the_footer_month_date, #10/1/2021#) > 0 Then new_version = False	'if the footer months is BEFORE 10/1/2021, then we need to read the old version

        EMReadScreen line_one, 22, 6, 43										'reading the information from the top half of the ADDR panel
        EMReadScreen line_two, 22, 7, 43
        EMReadScreen city_line, 15, 8, 43
        EMReadScreen state_line, 2, 8, 66
        EMReadScreen zip_line, 7, 9, 43
        EMReadScreen county_line, 2, 9, 66
        EMReadScreen verif_line, 2, 9, 74
        EMReadScreen homeless_line, 1, 10, 43
        EMReadScreen reservation_line, 1, 10, 74
        EMReadScreen living_sit_line, 2, 11, 43
		EMReadScreen reservation_name, 2, 11, 74

        resi_line_one = replace(line_one, "_", "")								'formatting the residence address information to remove the '_'
        resi_line_two = replace(line_two, "_", "")
		resi_street_full = trim(resi_line_one & " " & resi_line_two)
        resi_city = replace(city_line, "_", "")
        resi_zip = replace(zip_line, "_", "")

        If county_line = "01" Then addr_county = "01 Aitkin"					'Adding the county name to the county string
        If county_line = "02" Then addr_county = "02 Anoka"
        If county_line = "03" Then addr_county = "03 Becker"
        If county_line = "04" Then addr_county = "04 Beltrami"
        If county_line = "05" Then addr_county = "05 Benton"
        If county_line = "06" Then addr_county = "06 Big Stone"
        If county_line = "07" Then addr_county = "07 Blue Earth"
        If county_line = "08" Then addr_county = "08 Brown"
        If county_line = "09" Then addr_county = "09 Carlton"
        If county_line = "10" Then addr_county = "10 Carver"
        If county_line = "11" Then addr_county = "11 Cass"
        If county_line = "12" Then addr_county = "12 Chippewa"
        If county_line = "13" Then addr_county = "13 Chisago"
        If county_line = "14" Then addr_county = "14 Clay"
        If county_line = "15" Then addr_county = "15 Clearwater"
        If county_line = "16" Then addr_county = "16 Cook"
        If county_line = "17" Then addr_county = "17 Cottonwood"
        If county_line = "18" Then addr_county = "18 Crow Wing"
        If county_line = "19" Then addr_county = "19 Dakota"
        If county_line = "20" Then addr_county = "20 Dodge"
        If county_line = "21" Then addr_county = "21 Douglas"
        If county_line = "22" Then addr_county = "22 Faribault"
        If county_line = "23" Then addr_county = "23 Fillmore"
        If county_line = "24" Then addr_county = "24 Freeborn"
        If county_line = "25" Then addr_county = "25 Goodhue"
        If county_line = "26" Then addr_county = "26 Grant"
        If county_line = "27" Then addr_county = "27 Hennepin"
        If county_line = "28" Then addr_county = "28 Houston"
        If county_line = "29" Then addr_county = "29 Hubbard"
        If county_line = "30" Then addr_county = "30 Isanti"
        If county_line = "31" Then addr_county = "31 Itasca"
        If county_line = "32" Then addr_county = "32 Jackson"
        If county_line = "33" Then addr_county = "33 Kanabec"
        If county_line = "34" Then addr_county = "34 Kandiyohi"
        If county_line = "35" Then addr_county = "35 Kittson"
        If county_line = "36" Then addr_county = "36 Koochiching"
        If county_line = "37" Then addr_county = "37 Lac Qui Parle"
        If county_line = "38" Then addr_county = "38 Lake"
        If county_line = "39" Then addr_county = "39 Lake Of Woods"
        If county_line = "40" Then addr_county = "40 Le Sueur"
        If county_line = "41" Then addr_county = "41 Lincoln"
        If county_line = "42" Then addr_county = "42 Lyon"
        If county_line = "43" Then addr_county = "43 Mcleod"
        If county_line = "44" Then addr_county = "44 Mahnomen"
        If county_line = "45" Then addr_county = "45 Marshall"
        If county_line = "46" Then addr_county = "46 Martin"
        If county_line = "47" Then addr_county = "47 Meeker"
        If county_line = "48" Then addr_county = "48 Mille Lacs"
        If county_line = "49" Then addr_county = "49 Morrison"
        If county_line = "50" Then addr_county = "50 Mower"
        If county_line = "51" Then addr_county = "51 Murray"
        If county_line = "52" Then addr_county = "52 Nicollet"
        If county_line = "53" Then addr_county = "53 Nobles"
        If county_line = "54" Then addr_county = "54 Norman"
        If county_line = "55" Then addr_county = "55 Olmsted"
        If county_line = "56" Then addr_county = "56 Otter Tail"
        If county_line = "57" Then addr_county = "57 Pennington"
        If county_line = "58" Then addr_county = "58 Pine"
        If county_line = "59" Then addr_county = "59 Pipestone"
        If county_line = "60" Then addr_county = "60 Polk"
        If county_line = "61" Then addr_county = "61 Pope"
        If county_line = "62" Then addr_county = "62 Ramsey"
        If county_line = "63" Then addr_county = "63 Red Lake"
        If county_line = "64" Then addr_county = "64 Redwood"
        If county_line = "65" Then addr_county = "65 Renville"
        If county_line = "66" Then addr_county = "66 Rice"
        If county_line = "67" Then addr_county = "67 Rock"
        If county_line = "68" Then addr_county = "68 Roseau"
        If county_line = "69" Then addr_county = "69 St. Louis"
        If county_line = "70" Then addr_county = "70 Scott"
        If county_line = "71" Then addr_county = "71 Sherburne"
        If county_line = "72" Then addr_county = "72 Sibley"
        If county_line = "73" Then addr_county = "73 Stearns"
        If county_line = "74" Then addr_county = "74 Steele"
        If county_line = "75" Then addr_county = "75 Stevens"
        If county_line = "76" Then addr_county = "76 Swift"
        If county_line = "77" Then addr_county = "77 Todd"
        If county_line = "78" Then addr_county = "78 Traverse"
        If county_line = "79" Then addr_county = "79 Wabasha"
        If county_line = "80" Then addr_county = "80 Wadena"
        If county_line = "81" Then addr_county = "81 Waseca"
        If county_line = "82" Then addr_county = "82 Washington"
        If county_line = "83" Then addr_county = "83 Watonwan"
        If county_line = "84" Then addr_county = "84 Wilkin"
        If county_line = "85" Then addr_county = "85 Winona"
        If county_line = "86" Then addr_county = "86 Wright"
        If county_line = "87" Then addr_county = "87 Yellow Medicine"
        If county_line = "89" Then addr_county = "89 Out-of-State"
        resi_county = addr_county

		Call get_state_name_from_state_code(state_line, resi_state, TRUE)		'This function makes the state code to be the state name written out - including the code

        If homeless_line = "Y" Then addr_homeless = "Yes"						'formatting the Y and N to 'Yes' or 'No'
        If homeless_line = "N" Then addr_homeless = "No"
        If reservation_line = "Y" Then addr_reservation = "Yes"
        If reservation_line = "N" Then addr_reservation = "No"

		If reservation_name = "__" Then reservation_name = ""					'Filling in the detail of the reservation name
		If reservation_name = "BD" Then reservation_name = "BD - Bois Forte - Deer Creek"
		If reservation_name = "BN" Then reservation_name = "BN - Bois Forte - Nett Lake"
		If reservation_name = "BV" Then reservation_name = "BV - Bois Forte - Vermillion Lk"
		If reservation_name = "FL" Then reservation_name = "FL - Fond du Lac"
		If reservation_name = "GP" Then reservation_name = "GP - Grand Portage"
		If reservation_name = "LL" Then reservation_name = "LL - Leach Lake"
		If reservation_name = "LS" Then reservation_name = "LS - Lower Sioux"
		If reservation_name = "ML" Then reservation_name = "ML - Mille Lacs"
		If reservation_name = "PL" Then reservation_name = "PL - Prairie Island Community"
		If reservation_name = "RL" Then reservation_name = "RL - Red Lake"
		If reservation_name = "SM" Then reservation_name = "SM - Shakopee Mdewakanton"
		If reservation_name = "US" Then reservation_name = "US - Upper Sioux"
		If reservation_name = "WE" Then reservation_name = "WE - White Earth"

        If verif_line = "SF" Then addr_verif = "SF - Shelter Form"				'filling in the detail of the verification listed
        If verif_line = "CO" Then addr_verif = "CO - Coltrl Stmt"
		If verif_line = "LE" Then addr_verif = "LE - Lease/Rent Doc"
        If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
        If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
        If verif_line = "CD" Then addr_verif = "CD - Contrct for Deed"
        If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
        If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
        If verif_line = "OT" Then addr_verif = "OT - Other Document"
        If verif_line = "NO" Then addr_verif = "NO - No Verif"
        If verif_line = "?_" Then addr_verif = "? - Delayed"
        If verif_line = "__" Then addr_verif = "Blank"


        If living_sit_line = "__" Then living_situation = "Blank"				'Adding detail to the living situation code
        If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roommate"
        If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
        If living_sit_line = "03" Then living_situation = "03 -  servc prvdr- foster/group home"
        If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
        If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
        If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
        If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
        If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
        If living_sit_line = "09" Then living_situation = "09 - Declined"
        If living_sit_line = "10" Then living_situation = "10 - Unknown"
        addr_living_sit = living_situation

        EMReadScreen addr_eff_date, 8, 4, 43									'reading the dates on the panel
        EMReadScreen addr_future_date, 8, 4, 66

		If new_version = False Then												'reading the bottom half of the panel based on if we are in the new or old version
	        EMReadScreen mail_line_one, 22, 13, 43
	        EMReadScreen mail_line_two, 22, 14, 43
	        EMReadScreen mail_city_line, 15, 15, 43
	        EMReadScreen mail_state_line, 2, 16, 43
	        EMReadScreen mail_zip_line, 7, 16, 52

			EMReadScreen phone_one, 14, 17, 45
			EMReadScreen phone_two, 14, 18, 45
			EMReadScreen phone_three, 14, 19, 45

			EMReadScreen type_one, 1, 17, 67
			EMReadScreen type_two, 1, 18, 67
			EMReadScreen type_three, 1, 19, 67
		End If

		If new_version = True Then
			EMReadScreen mail_line_one, 22, 12, 49
			EMReadScreen mail_line_two, 22, 13, 49
			EMReadScreen mail_city_line, 15, 14, 49
			EMReadScreen mail_state_line, 2, 15, 49
			EMReadScreen mail_zip_line, 7, 15, 58

			EMReadScreen phone_one, 12, 16, 33
			EMReadScreen phone_two, 12, 17, 33
			EMReadScreen phone_three, 12, 18, 33

			EMReadScreen type_one, 1, 16, 61
			EMReadScreen type_two, 1, 17, 61
			EMReadScreen type_three, 1, 18, 61

			EMReadScreen text_yn_one, 1, 16, 76
			EMReadScreen text_yn_two, 1, 17, 76
			EMReadScreen text_yn_three, 1, 18, 76


			EMReadScreen addr_email, 50, 19, 31
		End If

        addr_eff_date = replace(addr_eff_date, " ", "/")						'formatting the information from the second half
        addr_future_date = trim(addr_future_date)
        addr_future_date = replace(addr_future_date, " ", "/")
        mail_line_one = replace(mail_line_one, "_", "")
        mail_line_two = replace(mail_line_two, "_", "")
		mail_street_full = trim(mail_line_one & " " & mail_line_two)
        mail_city = replace(mail_city_line, "_", "")
        mail_state = replace(mail_state_line, "_", "")
        mail_zip = replace(mail_zip_line, "_", "")

		If text_yn_one = "_" Then text_yn_one = ""								'Changing blanks to nulls
		If text_yn_two = "_" Then text_yn_two = ""
		If text_yn_three = "_" Then text_yn_three = ""
		'here we remove the '_' from the front and back of the string - we cannot use replace here because '_' may be a part of the email address.
		Do
			If right(addr_email, 1) = "_" Then addr_email = left(addr_email, len(addr_email) - 1)
			If left(addr_email, 1) = "_" Then addr_email = right(addr_email, len(addr_email) - 1)
		Loop until right(addr_email, 1) <> "_" AND  left(addr_email, 1) <> "_"

        notes_on_address = "Address effective on " & addr_eff_date & "."
        ' If mail_line_one <> "" Then
        '     If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        '     If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        ' End If
        If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."
		notes_on_address = notes_on_address & "; "

        phone_one = replace(phone_one, " ", "-")									'formatting phone numbers
        If phone_one = "___-___-____" Then phone_one = ""							'ALERT - phone numbers from panels before 10/21 will be formatted weird

        phone_two = replace(phone_two, " ", "-")
        If phone_two = "___-___-____" Then phone_two = ""

		phone_three = replace(phone_three, " ", "-")
        If phone_three = "___-___-____" Then phone_three = ""

        If type_one = "H" Then type_one = "H - Home"
        If type_one = "W" Then type_one = "W - Work"
        If type_one = "C" Then type_one = "C - Cell"
        If type_one = "M" Then type_one = "M - Message"
        If type_one = "T" Then type_one = "T - TTY/TDD"
        If type_one = "_" Then type_one = ""

        If type_two = "H" Then type_two = "H - Home"
        If type_two = "W" Then type_two = "W - Work"
        If type_two = "C" Then type_two = "C - Cell"
        If type_two = "M" Then type_two = "M - Message"
        If type_two = "T" Then type_two = "T - TTY/TDD"
        If type_two = "_" Then type_two = ""

        If type_three = "H" Then type_three = "H - Home"
        If type_three = "W" Then type_three = "W - Work"
        If type_three = "C" Then type_three = "C - Cell"
        If type_three = "M" Then type_three = "M - Message"
        If type_three = "T" Then type_three = "T - TTY/TDD"
        If type_three = "_" Then type_three = ""

		'here we save the information we gathered to start with so that we can compare it and know if it changed
		original_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							   addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							   mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&_
							   text_yn_one&"|"&text_yn_two&"|"&text_yn_three&"|"&addr_email&"|"&addr_verif
		original_information = UCASE(original_information)
    End If

    If access_type = "WRITE" Then
		' verif_received 'add functionality to change how this is updated based on if we have verif or not.

		update_attempted = False
		resi_line_one = trim(resi_line_one)
		resi_line_two = trim(resi_line_two)
		resi_street_full = trim(resi_street_full)
		resi_city = trim(resi_city)
		resi_state = trim(resi_state)
		resi_zip = trim(resi_zip)
		resi_county = trim(resi_county)
		addr_verif = trim(addr_verif)
		addr_homeless = trim(addr_homeless)
		addr_reservation = trim(addr_reservation)
		addr_living_sit = trim(addr_living_sit)
		mail_line_one = trim(mail_line_one)
		mail_line_two = trim(mail_line_two)
		mail_street_full = trim(mail_street_full)
		mail_city = trim(mail_city)
		mail_state = trim(mail_state)
		mail_zip = trim(mail_zip)
		addr_eff_date = trim(addr_eff_date)
		addr_future_date = trim(addr_future_date)
		phone_one = trim(phone_one)
		phone_two = trim(phone_two)
		phone_three = trim(phone_three)
		type_one = trim(type_one)
		type_two = trim(type_two)
		type_three = trim(type_three)
		verif_received = trim(verif_received)

		current_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							  addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							  mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&_
							  text_yn_one&"|"&text_yn_two&"|"&text_yn_three&"|"&addr_email&"|"&addr_verif


		current_information = UCase(current_information)
		' MsgBox "THIS" & vbCR & "ORIGINAL" & vbCr & original_information & vbCr & vbCr & "CURRENT" & vbCr & current_information
		If current_information <> original_information Then						'If the information in the beginning and the information inthe end do not match - we need to update
			update_attempted = True

			Call navigate_to_MAXIS_screen("STAT", "ADDR")							'going to ADDR
			EMReadScreen curr_addr_footer_month, 2, 20, 55							'reading the footer month and year on the panel because there is footer months pecific differences
			EMReadScreen curr_addr_footer_year, 2, 20, 58
			the_footer_month_date = curr_addr_footer_month &  "/1/" & curr_addr_footer_year		'making a date out of the footer month
			the_footer_month_date = DateAdd("d", 0, the_footer_month_date)
			new_version = True														'defaulting to using the new verion of the panel with text authorization and email
			If DateDiff("d", the_footer_month_date, #10/1/2021#) > 0 Then new_version = False	'if the footer months is BEFORE 10/1/2021, then we need to read the old version

	        PF9																	'Put it in edit mode

			Call clear_line_of_text(6, 43) 	'residence addr 1
			Call clear_line_of_text(7, 43) 	'residence addr 2
			Call clear_line_of_text(8, 43) 	'residence city
			Call clear_line_of_text(8, 66) 	'residence state
			Call clear_line_of_text(9, 43) 	'residence zip

			If new_version = True then
				Call clear_line_of_text(12, 49) 	'mail addr 1
				Call clear_line_of_text(13, 49) 	'mail addr 2
				Call clear_line_of_text(14, 49) 	'mail city
				Call clear_line_of_text(15, 49) 	'mail state
				Call clear_line_of_text(15, 58) 	'mail zip

				Call clear_line_of_text(16, 39)		'phone information'
				Call clear_line_of_text(16, 45)
				Call clear_line_of_text(16, 49)
				Call clear_line_of_text(16, 61)
				Call clear_line_of_text(16, 76)

				Call clear_line_of_text(17, 39)
				Call clear_line_of_text(17, 45)
				Call clear_line_of_text(17, 49)
				Call clear_line_of_text(17, 61)
				Call clear_line_of_text(17, 76)

				Call clear_line_of_text(18, 39)
				Call clear_line_of_text(18, 45)
				Call clear_line_of_text(18, 49)
				Call clear_line_of_text(18, 61)
				Call clear_line_of_text(18, 76)

				Call clear_line_of_text(19, 31)
			End If
			If new_version = False then
				Call clear_line_of_text(13, 43) 	'mail addr 1
				Call clear_line_of_text(14, 43) 	'mail addr 2
				Call clear_line_of_text(15, 43) 	'mail city
				Call clear_line_of_text(16, 43) 	'mail state
				Call clear_line_of_text(16, 52) 	'mail zip

				Call clear_line_of_text(17, 45)		'phone information'
				Call clear_line_of_text(17, 15)
				Call clear_line_of_text(17, 55)
				Call clear_line_of_text(17, 67)

				Call clear_line_of_text(18, 45)
				Call clear_line_of_text(18, 15)
				Call clear_line_of_text(18, 55)
				Call clear_line_of_text(18, 67)

				Call clear_line_of_text(19, 45)
				Call clear_line_of_text(19, 15)
				Call clear_line_of_text(19, 55)
				Call clear_line_of_text(19, 67)
			End If

			'Now we write all the information
	        Call create_mainframe_friendly_date(addr_eff_date, 4, 43, "YY")

			resi_street_full = trim(trim(resi_line_one) & " " & trim(resi_line_two))
			If len(resi_line_one) > 22 or len(resi_line_two) > 22 Then
				'This functionality will write the address word by word so that if it needs to word wrap to the second line, it can move the words to the next line
	            resi_words = split(resi_street_full, " ")
	            write_resi_line_one = ""
	            write_resi_line_two = ""
	            For each word in resi_words
	                If write_resi_line_one = "" Then
	                    write_resi_line_one = word
	                ElseIf len(write_resi_line_one & " " & word) =< 22 Then
	                    write_resi_line_one = write_resi_line_one & " " & word
	                Else
	                    If write_resi_line_two = "" Then
	                        write_resi_line_two = word
	                    Else
	                        write_resi_line_two = write_resi_line_two & " " & word
	                    End If
	                End If
	            Next
	        Else
	            write_resi_line_one = resi_line_one
				write_resi_line_two = resi_line_two
	        End If
	        EMWriteScreen write_resi_line_one, 6, 43
	        EMWriteScreen write_resi_line_two, 7, 43
	        EMWriteScreen resi_city, 8, 43
	        EMWriteScreen left(resi_county, 2), 9, 66
	        EMWriteScreen left(resi_state, 2), 8, 66
	        EMWriteScreen resi_zip, 9, 43
			If addr_living_sit <> "Blank" AND addr_living_sit <> "Select" AND len(addr_living_sit) >=2 Then EMWriteScreen left(addr_living_sit, 2), 11, 43
			EMWriteScreen left(addr_homeless, 1), 10, 43
			EMWriteScreen left(addr_reservation, 1), 10, 74
			IF addr_reservation = "No" THEN Call clear_line_of_text(11, 74)

	        EMWriteScreen left(addr_verif, 2), 9, 74

			mail_street_full = trim(trim(mail_line_one) & " " & trim(mail_line_two))
			If len(mail_line_one) > 22 or len(mail_line_two) > 22 Then
				'This functionality will write the address word by word so that if it needs to word wrap to the second line, it can move the words to the next line
	            mail_words = split(mail_street_full, " ")
	            write_mail_line_one = ""
	            write_mail_line_two = ""
	            For each word in mail_words
	                If write_mail_line_one = "" Then
	                    write_mail_line_one = word
	                ElseIf len(write_mail_line_one & " " & word) =< 22 Then
	                    write_mail_line_one = write_mail_line_one & " " & word
	                Else
	                    If write_mail_line_two = "" Then
	                        write_mail_line_two = word
	                    Else
	                        write_mail_line_two = write_mail_line_two & " " & word
	                    End If
	                End If
	            Next
	        Else
	            write_mail_line_one = mail_line_one
				write_mail_line_two = mail_line_two
	        End If


			If new_version = False then
				EMWriteScreen write_mail_line_one, 13, 43
				EMWriteScreen write_mail_line_two, 14, 43
				EMWriteScreen mail_city, 15, 43
				If write_mail_line_one <> "" Then EMWriteScreen left(mail_state, 2), 16, 43
				EMWriteScreen mail_zip, 16, 52

		        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
		        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
		        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

		        EMWriteScreen phone_one_left, 17, 45
		        EMWriteScreen phone_one_mid, 17, 51
		        EMWriteScreen phone_one_right, 17, 55
		        If type_one <> "Select One..." Then EMWriteScreen type_one, 17, 67

		        EMWriteScreen phone_two_left, 18, 45
		        EMWriteScreen phone_two_mid, 18, 51
		        EMWriteScreen phone_two_right, 18, 55
		        If type_two <> "Select One..." Then EMWriteScreen type_two, 18, 67

		        EMWriteScreen phone_three_left, 19, 45
		        EMWriteScreen phone_three_mid, 19, 51
		        EMWriteScreen phone_three_right, 19, 55
		        If type_three <> "Select One..." Then EMWriteScreen type_three, 19, 67
			End If

			If new_version = True then
				EMWriteScreen write_mail_line_one, 12, 49
		        EMWriteScreen write_mail_line_two, 13, 49
		        EMWriteScreen mail_city, 14, 49
		        If write_mail_line_one <> "" Then EMWriteScreen left(mail_state, 2), 15, 49
		        EMWriteScreen mail_zip, 15, 58

		        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
		        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
		        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

		        EMWriteScreen phone_one_left, 16, 39
		        EMWriteScreen phone_one_mid, 16, 45
		        EMWriteScreen phone_one_right, 16, 49
		        If type_one <> "Select One..." Then EMWriteScreen type_one, 16, 61
				If phone_one <> "" Then EMWriteScreen text_yn_one, 16, 76

		        EMWriteScreen phone_two_left, 17, 39
		        EMWriteScreen phone_two_mid, 17, 45
		        EMWriteScreen phone_two_right, 17, 49
		        If type_two <> "Select One..." Then EMWriteScreen type_two, 17, 61
				If phone_two <> "" Then EMWriteScreen text_yn_two, 17, 76

		        EMWriteScreen phone_three_left, 18, 39
		        EMWriteScreen phone_three_mid, 18, 45
		        EMWriteScreen phone_three_right, 18, 49
		        If type_three <> "Select One..." Then EMWriteScreen type_three, 18, 61
				If phone_three <> "" Then EMWriteScreen text_yn_three, 18, 76

				EMSetCursor 19, 31
				EMSendKey "<eraseEOF>"
				EMWriteScreen addr_email, 19, 31
			End If

	        save_attempt = 1
	        Do
	            transmit
				' MsgBox "Pause - " & save_attempt
	            EMReadScreen resi_standard_note, 33, 24, 2
	            If resi_standard_note = "RESIDENCE ADDRESS IS STANDARDIZED" Then transmit
	            EMReadScreen mail_standard_note, 31, 24, 2
	            If mail_standard_note = "MAILING ADDRESS IS STANDARDIZED" Then transmit

				EMReadScreen warn_msg, 60, 24, 2
				warn_msg = trim(warn_msg)
				If warn_msg = "ENTER A VALID COMMAND OR PF-KEY" Then Exit Do
	            row = 0
	            col = 0
	            EMSearch "Warning:", row, col

	            If row <> 0 Then
	                Do
	                    EMReadScreen warning_note, 55, row, col
	                    warning_note = trim(warning_note)
	                    warning_message = warning_message & "; " & warning_note
	                Loop until warning_note = ""
	            End If

	            save_attempt = save_attempt + 1
	        Loop until save_attempt = 20
		End IF
    End If
end function

function cancel_confirmation()
'--- This function asks if the user if they want to cancel. If you say yes, the script will end. If no, the dialog will appear for the user again.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
	End if
end function

function cancel_without_confirmation()
'--- This function ends a script after a user presses cancel. There is no confirmation message box but the end message for statistical information that cancel was pressed.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
        script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
        'Left the If...End If in the tier in case we want more stats or error handling, or if we need specialty processing for workflows
    End if
end function

function check_for_MAXIS(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				check_for_MAXIS_msg = "*** The script cannot identify if you are currently logged into MAXIS. ***"
				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & vbCr &"   This may be for a few different reasons, including:"
				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & "      - You are passworded out."
				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & "      - You are in a different system (PRISM, MMIS, etc.)"
				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & "      - The area of MAXIS has a different header."
				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & "        (The top of your screen should say MAXIS.)"

				check_for_MAXIS_msg = check_for_MAXIS_msg & vbCr & vbCr & "The script has stopped, please check your MAXIS screen and try again."
				script_end_procedure(check_for_MAXIS_msg)
			Else
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 241, 145, "Password Dialog"
				ButtonGroup ButtonPressed
					OkButton 130, 125, 50, 15
					CancelButton 185, 125, 50, 15
				Text 5, 10, 245, 10, "*** The script cannot identify if you are currently logged into MAXIS. ***"
				Text 20, 25, 175, 10, "This may be for a few different reasons, including:"
				Text 35, 40, 90, 10, "- You are passworded out."
				Text 35, 50, 170, 10, "- You are in a different system (PRISM, MMIS, etc.)"
				Text 35, 60, 170, 10, "- The area of MAXIS has a different header."
				Text 40, 70, 170, 10, "(The top of your screen should say MAXIS.)"
				Text 5, 90, 180, 20, "Password back in or navigate to a main area of MAXIS and press 'OK'' to continue."
				Text 5, 110, 125, 10, "Or press 'Cancel'' to stop the script."
				EndDialog
                Do
                    Do
                        dialog Dialog1
                        cancel_confirmation
                    Loop until ButtonPressed = -1
                    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
                Loop until are_we_passworded_out = false					'loops until user passwords back in
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

function check_for_password(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Transmit 'transmitting to see if the password screen appears
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function clear_line_of_text(row, start_column)
'--- This function clears out a single line of text
'~~~~~ row: coordinate of row to clear
'~~~~~ start_column: coordinate of column to start clearing
'===== Keywords: MAXIS, PRISM, production, clear
  EMSetCursor row, start_column
  EMSendKey "<EraseEof>"
  EMWaitReady 0, 0
end function

function get_county_code()
'--- This function determines county_name from worker_county_code, and asks for it if it's blank
'===== Keywords: MAXIS, MMIS, PRISM, county
	If left(code_from_installer, 2) = "PT" then 'special handling for Pine Tech
		worker_county_code = "PWVTS"
	Else
		If worker_county_code = "MULTICOUNTY" or worker_county_code = "" then 		'If the user works for many counties (i.e. SWHHS) or isn't assigned (i.e. a scriptwriter) it asks.
			Do
				two_digit_county_code_variable = inputbox("Select the county to proxy as. Ex: ''01''")
				If two_digit_county_code_variable = "" then stopscript
				If len(two_digit_county_code_variable) <> 2 or isnumeric(two_digit_county_code_variable) = False then MsgBox "Your county proxy code should be two digits and numeric."
			Loop until len(two_digit_county_code_variable) = 2 and isnumeric(two_digit_county_code_variable) = True
			worker_county_code = "X1" & two_digit_county_code_variable
			If two_digit_county_code_variable = "91" then worker_county_code = "PW"	'For DHS folks without proxy
		End If
	End if

    'Determining county name
    if worker_county_code = "x101" then
        county_name = "Aitkin County"
    elseif worker_county_code = "x102" then
        county_name = "Anoka County"
    elseif worker_county_code = "x103" then
        county_name = "Becker County"
    elseif worker_county_code = "x104" then
        county_name = "Beltrami County"
    elseif worker_county_code = "x105" then
        county_name = "Benton County"
    elseif worker_county_code = "x106" then
        county_name = "Big Stone County"
    elseif worker_county_code = "x107" then
        county_name = "Blue Earth County"
    elseif worker_county_code = "x108" then
        county_name = "Brown County"
    elseif worker_county_code = "x109" then
        county_name = "Carlton County"
    elseif worker_county_code = "x110" then
        county_name = "Carver County"
    elseif worker_county_code = "x111" then
        county_name = "Cass County"
    elseif worker_county_code = "x112" then
        county_name = "Chippewa County"
    elseif worker_county_code = "x113" then
        county_name = "Chisago County"
    elseif worker_county_code = "x114" then
        county_name = "Clay County"
    elseif worker_county_code = "x115" then
        county_name = "Clearwater County"
    elseif worker_county_code = "x116" then
        county_name = "Cook County"
    elseif worker_county_code = "x117" then
        county_name = "Cottonwood County"
    elseif worker_county_code = "x118" then
        county_name = "Crow Wing County"
    elseif worker_county_code = "x119" then
        county_name = "Dakota County"
    elseif worker_county_code = "x120" then
        county_name = "Dodge County"
    elseif worker_county_code = "x121" then
        county_name = "Douglas County"
    elseif worker_county_code = "x122" then
        county_name = "Faribault County"
    elseif worker_county_code = "x123" then
        county_name = "Fillmore County"
    elseif worker_county_code = "x124" then
        county_name = "Freeborn County"
    elseif worker_county_code = "x125" then
        county_name = "Goodhue County"
    elseif worker_county_code = "x126" then
        county_name = "Grant County"
    elseif worker_county_code = "x127" then
        county_name = "Hennepin County"
    elseif worker_county_code = "x128" then
        county_name = "Houston County"
    elseif worker_county_code = "x129" then
        county_name = "Hubbard County"
    elseif worker_county_code = "x130" then
        county_name = "Isanti County"
    elseif worker_county_code = "x131" then
        county_name = "Itasca County"
    elseif worker_county_code = "x132" then
        county_name = "Jackson County"
    elseif worker_county_code = "x133" then
        county_name = "Kanabec County"
    elseif worker_county_code = "x134" then
        county_name = "Kandiyohi County"
    elseif worker_county_code = "x135" then
        county_name = "Kittson County"
    elseif worker_county_code = "x136" then
        county_name = "Koochiching County"
    elseif worker_county_code = "x137" then
        county_name = "Lac Qui Parle County"
    elseif worker_county_code = "x138" then
        county_name = "Lake County"
    elseif worker_county_code = "x139" then
        county_name = "Lake of the Woods County"
    elseif worker_county_code = "x140" then
        county_name = "LeSueur County"
    elseif worker_county_code = "x141" then
        county_name = "Lincoln County"
    elseif worker_county_code = "x142" then
        county_name = "Lyon County"
    elseif worker_county_code = "x143" then
        county_name = "Mcleod County"
    elseif worker_county_code = "x144" then
        county_name = "Mahnomen County"
    elseif worker_county_code = "x145" then
        county_name = "Marshall County"
    elseif worker_county_code = "x146" then
        county_name = "Martin County"
    elseif worker_county_code = "x147" then
        county_name = "Meeker County"
    elseif worker_county_code = "x148" then
        county_name = "Mille Lacs County"
    elseif worker_county_code = "x149" then
        county_name = "Morrison County"
    elseif worker_county_code = "x150" then
        county_name = "Mower County"
    elseif worker_county_code = "x151" then
        county_name = "Murray County"
    elseif worker_county_code = "x152" then
        county_name = "Nicollet County"
    elseif worker_county_code = "x153" then
        county_name = "Nobles County"
    elseif worker_county_code = "x154" then
        county_name = "Norman County"
    elseif worker_county_code = "x155" then
        county_name = "Olmsted County"
    elseif worker_county_code = "x156" then
        county_name = "Otter Tail County"
    elseif worker_county_code = "x157" then
        county_name = "Pennington County"
    elseif worker_county_code = "x158" then
        county_name = "Pine County"
    elseif worker_county_code = "x159" then
        county_name = "Pipestone County"
    elseif worker_county_code = "x160" then
        county_name = "Polk County"
    elseif worker_county_code = "x161" then
        county_name = "Pope County"
    elseif worker_county_code = "x162" then
        county_name = "Ramsey County"
    elseif worker_county_code = "x163" then
        county_name = "Red Lake County"
    elseif worker_county_code = "x164" then
        county_name = "Redwood County"
    elseif worker_county_code = "x165" then
        county_name = "Renville County"
    elseif worker_county_code = "x166" then
        county_name = "Rice County"
    elseif worker_county_code = "x167" then
        county_name = "Rock County"
    elseif worker_county_code = "x168" then
        county_name = "Roseau County"
    elseif worker_county_code = "x169" then
        county_name = "St. Louis County"
    elseif worker_county_code = "x170" then
        county_name = "Scott County"
    elseif worker_county_code = "x171" then
        county_name = "Sherburne County"
    elseif worker_county_code = "x172" then
        county_name = "Sibley County"
    elseif worker_county_code = "x173" then
        county_name = "Stearns County"
    elseif worker_county_code = "x174" then
        county_name = "Steele County"
    elseif worker_county_code = "x175" then
        county_name = "Stevens County"
    elseif worker_county_code = "x176" then
        county_name = "Swift County"
    elseif worker_county_code = "x177" then
        county_name = "Todd County"
    elseif worker_county_code = "x178" then
        county_name = "Traverse County"
    elseif worker_county_code = "x179" then
        county_name = "Wabasha County"
    elseif worker_county_code = "x180" then
        county_name = "Wadena County"
    elseif worker_county_code = "x181" then
        county_name = "Waseca County"
    elseif worker_county_code = "x182" then
        county_name = "Washington County"
    elseif worker_county_code = "x183" then
        county_name = "Watonwan County"
    elseif worker_county_code = "x184" then
        county_name = "Wilkin County"
    elseif worker_county_code = "x185" then
        county_name = "Winona County"
    elseif worker_county_code = "x186" then
        county_name = "Wright County"
    elseif worker_county_code = "x187" then
        county_name = "Yellow Medicine County"
    elseif worker_county_code = "x188" then
        county_name = "Mille Lacs Band"
    elseif worker_county_code = "x192" then
        county_name = "White Earth Nation"
    elseif worker_county_code = "PWVTS" then
    	county_name = "Pine Tech"
    end if
end function

function MAXIS_case_number_finder(variable_for_MAXIS_case_number)
'--- This function finds the MAXIS case number if listed on a MAXIS screen
'~~~~~ variable_for_MAXIS_case_number: this should be <code>MAXIS_case_number</code>
'===== Keywords: MAXIS, case number
	EMReadScreen variable_for_SELF_check, 4, 2, 50
	IF variable_for_SELF_check = "SELF" then
		EMReadScreen variable_for_MAXIS_case_number, 8, 18, 43
		variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
		variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
	ELSE
		row = 1
		col = 1
		EMSearch "Case Nbr:", row, col
		If row <> 0 then
			EMReadScreen variable_for_MAXIS_case_number, 8, row, col + 10
			variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
			variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
		END IF
	END IF

end function

function MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'--- This function finds the MAXIS footer month/year for MAXIS cases in SELF, MAXIS panels or MEMO screens.
'~~~~~ MAXIS_footer_month: needs to be <code>MAXIS_footer_month</code>
'~~~~~ MAXIS_footer_year: needs to be <code>MAXIS_footer_year</code>
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50
    EMReadScreen MEMO_check, 4, 2, 47
    EMReadScreen casenote_check, 4, 2, 45
    Call find_variable("Function: ", MAXIS_function, 4)

	IF SELF_check = "SELF" THEN
		EMReadScreen MAXIS_footer_month, 2, 20, 43
		EMReadScreen MAXIS_footer_year, 2, 20, 46
	ELSEIF MEMO_check = "MEMO" or MEMO_check = "WCOM" Then
		EMReadScreen MAXIS_footer_month, 2, 19, 54
		EMReadScreen MAXIS_footer_year, 2, 19, 57
    ELSEIF casenote_check = "NOTE" then
    	EMReadScreen MAXIS_footer_month, 2, 20, 54
        EMReadScreen MAXIS_footer_year, 2, 20, 57
	ELSEIF MAXIS_function = "STAT" then
		EMReadScreen MAXIS_footer_month, 2, 20, 55
        EMReadScreen MAXIS_footer_year, 2, 20, 58
    ELSEIF MAXIS_function = "REPT" then
    		EMReadScreen MAXIS_footer_month, 2, 20, 54
            EMReadScreen MAXIS_footer_year, 2, 20, 57
    Else
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
    End if
end function

function MAXIS_footer_month_confirmation()
'--- This function is for checking and changing the footer month to the MAXIS_footer_month & MAXIS_footer_year selected by the user in the inital dialog if necessary
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50			'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN
		EMReadScreen panel_footer_month, 2, 20, 43
		EMReadScreen panel_footer_year, 2, 20, 46
	ELSE
		Call find_variable("Month: ", MAXIS_footer, 5)	'finding footer month and year if not on the SELF screen
		panel_footer_month = left(MAXIS_footer, 2)
		panel_footer_year = right(MAXIS_footer, 2)
		If row <> 0 then
  			panel_footer_month = panel_footer_month		'Establishing variables
			panel_footer_year =panel_footer_year
		END IF
	END IF
	panel_date = panel_footer_month & panel_footer_year		'creating new variable combining month and year for the date listed on the MAXIS panel
	dialog_date = MAXIS_footer_month & MAXIS_footer_year	'creating new variable combining the MAXIS_footer_month & MAXIS_footer_year to measure against the panel date
	IF panel_date <> dialog_date then 						'if dates are not equal
		back_to_SELF
		EMWriteScreen MAXIS_footer_month, 20, 43			'goes back to self and enters the date that the user selcted'
		EMWriteScreen MAXIS_footer_year, 20, 46
	END IF
end function

Function MMIS_panel_confirmation(panel_name, col)
'--- This function confirms that a user in on the correct MMIS panel.
'~~~~~ panel_name: name of the panel you are confirming as a string in ""
'~~~~~ col: The column which to start reading the panel name. For instance, this is usually 51 or 52 in MMIS.
'===== Keywords: MMIS, navigate, confirm
	Do
		EMReadScreen panel_check, 4, 1, col
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

function navigate_to_MAXIS(maxis_mode)
'--- This function is to be used when navigating back to MAXIS from another function in BlueZone (MMIS, PRISM, INFOPAC, etc.)
'~~~~~ maxis_mode: This parameter needs to be "maxis_mode"
'===== Keywords: MAXIS, navigate
    attn
    Do
        EMReadScreen MAI_check, 3, 1, 33
        If MAI_check <> "MAI" then EMWaitReady 1, 1
    Loop until MAI_check = "MAI"

    EMReadScreen prod_check, 7, 6, 15
    IF prod_check = "RUNNING" THEN
        Call write_value_and_transmit("1", 2, 15)
    ELSE
        EMConnect"A"
        attn
        EMReadScreen prod_check, 7, 6, 15
        IF prod_check = "RUNNING" THEN
            Call write_value_and_transmit("1", 2, 15)
        ELSE
            EMConnect"B"
            attn
            EMReadScreen prod_check, 7, 6, 15
            IF prod_check = "RUNNING" THEN
                Call write_value_and_transmit("1", 2, 15)
            Else
                script_end_procedure("You do not appear to have Production mode running. This script will now stop. Please make sure you have production and MMIS open in the same session, and re-run the script.")
            END IF
        END IF
    END IF
end function

function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
'--- This function is to be used to navigate to a specific MAXIS screen
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'===== Keywords: MAXIS, navigate
	function_to_go_to = UCase(function_to_go_to)
	command_to_go_to = UCase(command_to_go_to)
	EMSendKey "<enter>"
	EMWaitReady 0, 0
	EMReadScreen MAXIS_check, 5, 1, 39
	If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
		already_at_the_correct_screen = False
		review_footer_month = False
		at_correct_footer_month = False

		EMReadScreen locked_panel, 23, 2, 30
		IF locked_panel = "Program History Display" then
			PF3 'Checks to see if on Program History panel - which does not allow the Command line to be updated
		END IF
		row = 1
		col = 1
		EMSearch "Function: ", row, col
		If row <> 0 then
			EMReadScreen MAXIS_function, 4, row, col + 10
			EMReadScreen STAT_note_check, 4, 2, 45
			row = 1
			col = 1
			EMSearch "Case Nbr: ", row, col
			EMReadScreen current_case_number, 8, row, col + 10
			current_case_number = replace(current_case_number, "_", "")
			current_case_number = trim(current_case_number)
			EMReadScreen all_of_function_row, 80, row, 1
			footer_month_location = InStr(all_of_function_row, "Month")
			footer_month_location = footer_month_location + 7
			If function_to_go_to = "STAT" Then review_footer_month = True
			If function_to_go_to = "ELIG" Then review_footer_month = True
			If MAXIS_footer_month = "" OR MAXIS_footer_year = "" Then review_footer_month = False
			EMReadScreen current_footer_month, 2, row, footer_month_location
			EMReadScreen current_footer_year, 2, row, footer_month_location+3
			If review_footer_month = True Then
				If current_footer_month = MAXIS_footer_month AND current_footer_year = MAXIS_footer_year Then at_correct_footer_month = True
			Else
				at_correct_footer_month = True
			End If

			EMReadScreen all_of_row_two, 80, 2, 1
			If current_case_number = MAXIS_case_number and MAXIS_function = function_to_go_to AND InStr(all_of_row_two, command_to_go_to) <> 0 AND at_correct_footer_month = True Then already_at_the_correct_screen = True
		End if

		If already_at_the_correct_screen = False Then
			If current_case_number = MAXIS_case_number and MAXIS_function = function_to_go_to and STAT_note_check <> "NOTE" and at_correct_footer_month = True then
				row = 1
				col = 1
				EMSearch "Command: ", row, col
				EMWriteScreen command_to_go_to, row, col + 9
				EMSendKey "<enter>"
				EMWaitReady 0, 0
			Else
				Do
					EMSendKey "<PF3>"
					EMWaitReady 0, 0
					EMReadScreen SELF_check, 4, 2, 50
				Loop until SELF_check = "SELF"
				EMWriteScreen function_to_go_to, 16, 43
				EMWriteScreen "________", 18, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen MAXIS_footer_month, 20, 43
				EMWriteScreen MAXIS_footer_year, 20, 46
				EMWriteScreen command_to_go_to, 21, 70
				EMSendKey "<enter>"
				EMWaitReady 0, 0
				EMReadScreen abended_check, 7, 9, 27
				If abended_check = "abended" then
					EMSendKey "<enter>"
					EMWaitReady 0, 0
				End if
				EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
				If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
			End if
		End If
	End if
end function

function navigate_to_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
            BeginDialog Password_dialog, 0, 0, 156, 55, "Password Dialog"
            ButtonGroup ButtonPressed
            OkButton 45, 35, 50, 15
            CancelButton 100, 35, 50, 15
            Text 5, 5, 150, 25, "You have passworded out. Please enter your password, then press OK to continue. Press CANCEL to stop the script. "
            EndDialog
            Do
                Do
                    dialog Password_dialog
                    cancel_confirmation
                Loop until ButtonPressed = -1
			    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	 		Loop until are_we_passworded_out = false					'loops until user passwords back in
		End if
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

            menu_to_enter = "PRIOR AUTHORIZATION   "

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

            menu_to_enter = "RECIPIENT FILE APPLICATION"

		End Select

        'Now it finds the recipient file application feature and selects it.
        row = 1
        col = 1
        EMSearch menu_to_enter, row, col
        EMWriteScreen "X", row, col - 3
        transmit

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "X", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "X", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		End Select
        EMWaitReady 0, 0
	End If
end function


function script_end_procedure(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
	script_run_end_time = time
	script_run_end_date = date
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then MsgBox closing_message '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Defaulting script success to successful
        SCRIPT_success = -1

		'Determines if the value of the MAXIS case number - BULK scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" then MAXIS_CASE_NUMBER = ""

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		IF using_SQL_database = TRUE then
    		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
		ELSE
			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
		END IF

        'Adds some data for users of the old database, but adds lots more data for users of the new.
        If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
    		'Opening usage_log and adding a record
    		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
    		"VALUES ('" & user_ID & "', '" & script_run_end_date & "', '" & script_run_end_time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & script_run_end_date & "', '" & script_run_end_time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & script_run_end_date & "', '" & script_run_end_time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function

function write_bullet_and_variable_in_CASE_NOTE(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CASE notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CASE note
'===== Keywords: MAXIS, bullet, CASE note
	If trim(variable) <> "" then
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
			If character_test <> "" or noting_row >= 18 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
				End if
			End if
		Loop until character_test = ""

		'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
		If len(bullet) >= 14 then
			indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
		Else
			indent_length = len(bullet) + 4 'It's four more for the reason explained above.
		End if

		'Writes the bullet
		EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

		'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + (len(bullet) + 4)

		'Splits the contents of the variable into an array of words
        variable = trim(variable)
        If right(variable, 1) = ";" Then variable = left(variable, len(variable) - 1)
        If left(variable, 1) = ";" Then variable = right(variable, len(variable) - 1)
		variable_array = split(variable, " ")

		For each word in variable_array
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

                'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
                EMReadScreen end_of_case_note_check, 1, 24, 2
                If end_of_case_note_check = "A" then
                    EMSendKey "<PF3>"												'PF3s
                    EMWaitReady 0, 0
                    EMSendKey "<PF9>"												'PF9s (opens new note)
                    EMWaitReady 0, 0
                    EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
                    EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
                    noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
                Else
                    noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                End If
			End if

			'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
			If noting_col = 3 then
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
			If right(word, 1) = ";" then
				noting_row = noting_row + 1
				noting_col = 3
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function


function write_value_and_transmit(input_value, row, col)
'--- This function writes a specific value and transmits.
'~~~~~ input_value: information to be entered
'~~~~~ row: row to write the input_value
'~~~~~ col: column to write the input_value
'===== Keywords: MAXIS, PRISM, case note, three columns, format
	EMWriteScreen input_value, row, col
	transmit
end function

function write_variable_in_CASE_NOTE(variable)
'--- This function writes a variable in CASE note
'~~~~~ variable: information to be entered into CASE note from script/edit box
'===== Keywords: MAXIS, CASE note
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
            If character_test <> "" or noting_row >= 18 then

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
                Else
                    noting_row = noting_row + 1
				End if
			End if
		Loop until character_test = ""

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		For each word in variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End if
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(true)
get_county_code
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 216, 170, "Add GRH Rate 2 to MMIS"
  EditBox 110, 10, 50, 15, MAXIS_case_number
  EditBox 110, 30, 20, 15, MAXIS_footer_month
  EditBox 140, 30, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 75, 55, 40, 15
    CancelButton 120, 55, 40, 15
  Text 10, 130, 195, 25, "Before you use the script, you must have approved GRH results that reflect the SSR information in the SSR pop-up on ELIG/GRFB for the selected footer month/year."
  Text 55, 15, 50, 10, "Case Number:"
  Text 10, 95, 195, 25, "This script is to be used when a new service agreement needs to be added into MMIS. If you need to update an agreement, please do that manually."
  Text 45, 35, 60, 10, "Initial month/year:"
  GroupBox 5, 80, 205, 85, "Add GRH Rate 2 to MMIS script:"
EndDialog

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "Enter a valid case number."		'mandatory field
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

EMReadScreen grh_status, 4, 9, 74		'Ensuring that the case is active on GRH. If not, case will not be updated in MMIS.
If grh_status <> "ACTV" then
    If trim(grh_status) = "" then grh_status = "Inactive"
	script_end_procedure("GRH case status is " & grh_status & ". The script will now end.")
End if

EMReadscreen current_county, 4, 21, 21
If current_county <> UCase(worker_county_code) then script_end_procedure("Out-of-county case. Cannot update. The script will now end.")

Call HCRE_panel_bypass			'Function to bypass a jenky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

'----------------------------------------------------------------------------------------------------UNEA panel
SSA_disa = ""
Call navigate_to_MAXIS_screen("STAT", "UNEA")
Call write_value_and_transmit("01", 20, 76)
Call write_value_and_transmit("01", 20, 79)
EMReadScreen total_panels, 1, 2, 78
If total_panels = "0" then
    SSA_disa = false
Else
    Do
        EmReadscreen UNEA_type, 2, 5, 37
        If UNEA_type = "01" or UNEA_type = "02" or UNEA_type = "03" then
            SSA_disa = True
            exit do
        ELSE
            SSA_disa = False
            transmit
        End if
        EmReadscreen error_check, 5, 24, 2
    Loop until error_check = "ENTER"
End if

'----------------------------------------------------------------------------------------------------DISA panel'
Call navigate_to_MAXIS_screen("STAT", "DISA")
Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.
EMReadScreen waiver_type, 1, 14, 59
If waiver_type <> "_" then script_end_procedure("Client is active on a waiver. Should not be Rate 2. Please review waiver information in MMIS, and update MAXIS if applicable. The script will now end.")

If SSA_disa = True then
    EmReadscreen cert_start_date, 10, 7, 47
    EmReadscreen cert_end_date, 10, 7, 69
    If (SSA_disa = True and cert_start_date = "__ __ ____") then
        script_end_procedure("Client is certified disabled through SSA. Both SSA disability dates and PSN dates need to be listed on STAT/DISA. The script will now end.")
    Else
        DISA_start = replace(cert_start_date, " ", "/")
        If DISA_start = "__/__/____" then DISA_start = ""

        DISA_end = replace(cert_end_date, " ", "/")
        If DISA_end = "__/__/____" then DISA_end = ""
    End if
Else
    EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
	EMReadScreen disa_end_date, 10, 6, 69
	IF disa_start_date <> "__/__/____" then disa_start = Replace(disa_start_date," ","/")		'cleans up DISA dates
	If disa_end_date <> "__ __ ____" then disa_end = Replace(disa_end_date," ","/")

    DISA_start = replace(disa_start_date, " ", "/")
    If DISA_start = "__/__/____" then DISA_start = ""

    DISA_end = replace(disa_end_date, " ", "/")
    If DISA_end = "__/__/____" then DISA_end = ""
End if

If cdate(DISA_start) <= cdate("02/01/2018") then DISA_start = "02/01/2018"

'logic to ensure that the disa end date extends through the end of the month if necessary.
If disa_end <> "" then
    next_month = DateAdd("M", 1, DISA_end)
    next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
    DISA_end = dateadd("d", -1, next_month)
End if

'----------------------------------------------------------------------------------------------------BUSI and JOBS panels
CSR_required = ""   'Value that will be established as True or False based on if someone is working or not.

Call navigate_to_MAXIS_screen("STAT", "JOBS")
EmWriteScreen "01", 20, 76
Call write_value_and_transmit("01", 20, 79)
EMReadScreen total_panels, 1, 2, 78
If total_panels = "0" then
    CSR_required = FALSE
Else
    Do
        EmReadscreen JOBS_end_date, 8, 9, 49
        If JOBS_end_date = "__ __ __" then
            CSR_required = True
            exit do
        ELSE
            CSR_required = FALSE
            transmit
        End if
        EmReadscreen error_check, 5, 24, 2
    Loop until error_check = "ENTER"
End if

If CSR_requied <> True then
    Call navigate_to_MAXIS_screen("STAT", "BUSI")
    EmWriteScreen "01", 20, 76
    Call write_value_and_transmit("01", 20, 79)
    EMReadScreen total_panels, 1, 2, 78
    If total_panels = "0" then
        CSR_required = FALSE
    Else
        Do
            EmReadscreen BUSI_end_date, 8, 5, 72
            If BUSI_end_date = "__ __ __" then
                CSR_required = True
                exit do
            ELSE
                CSR_required = FALSE
                transmit
            End if
            EmReadscreen error_check, 5, 24, 2
        Loop until error_check = "ENTER"
    End if
End if
'----------------------------------------------------------------------------------------------------REVW panel
Call navigate_to_MAXIS_screen("STAT", "REVW")
Call write_value_and_transmit("X", 5, 35)

If CSR_required = True then
    EmReadscreen SR_month, 2, 9, 26
    EmReadscreen SR_year, 2, 9, 32
    If SR_month = "__" then script_end_procedure("A CSR is required for this case due to earned income. Please update the case, and run the script again if needed. The script will now end.")
End if

PF3 'back to stat/revw screen
EmReadscreen next_revw_month, 2, 9, 37
EmReadscreen next_revw_day, 2, 9, 40
EmReadscreen next_revw_year, 2, 9, 43

next_revw_date = next_revw_month & "/" & next_revw_day & "/" & next_revw_year
revw_end = dateadd("d", -1, next_revw_date)
IF CSR_required = true then
    revw_start = dateadd("M", - 6, next_revw_date)
else
    revw_start = dateadd("M", - 12, next_revw_date)
End if

If cdate(revw_start) <= cdate("02/01/2018") then revw_start = "02/01/2018"

'----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the FACI dates match.
 Call navigate_to_MAXIS_screen ("STAT", "SSRT")
 Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.
 Call write_value_and_transmit ("01", 20, 79)    'Ensuring we're on the 1st panel

EMReadScreen SSRT_total_check, 1, 2, 78
If SSRT_total_check = "0" then
	script_end_procedure("SSRT panel needs to be created. The script will now end.")
Elseif SSRT_total_check = "1" then
    SSRT_found = True
Else
    Do
        confirm_SSRT = msgbox("Is this the facility/vendor you'd like to create an agreement for? Press NO to check next facility. Press YES to continue.", vbYesNoCancel + vbQuestion, "More than one SSRT panel exists.")
	    If confirm_SSRT = vbCancel then script_end_procedure("You have pressed Cancel. The script will now end.")
        If confirm_SSRT = vbYes then
            SSRT_found = True
            exit do
        End if
        If confirm_SSRT = vbNo then
            SSRT_found = False
            transmit
            EmReadscreen last_panel, 5, 24, 2
        End if
    Loop until last_panel = "ENTER"
    If SSRT_found = False then script_end_procedure("All facility/vendors have reviewed without being selected. The script will now end.")
End if

'Trying to find a suggested date based on the SSRT panel
EMReadScreen SSRT_vendor_number, 8, 5, 43		'Enters vendor number
EmReadscreen SSRT_vendor_name, 30, 6, 43
SSRT_vendor_name = replace(SSRT_vendor_name, "_", "")
EMReadScreen NPI_number, 10, 7, 43

If trim(NPI_number) = "" then script_end_procedure("No NPI number on SSRT panel. Agreement cannot be loaded into MMIS. Please report this NPI number to DHS. The script will now end.")
If instr(SSRT_vendor_name, "ANDREW RESIDENCE") then script_end_procedure("Andrew Residence facilities do not get loaded into MMIS. The script will now end.")

current_faci = false
row = 14
Do
    EMReadScreen ssrt_out_date, 10, row, 71
    EMReadScreen ssrt_in_date, 10, row, 47
    If ssrt_out_date = "__ __ ____" then
        If ssrt_in_date = "__ __ ____" then
            current_faci = False
            row = row - 1
        else
            current_faci = True
            Exit do
        End if
    Else
        current_faci = true
        exit do
    End if
    If row = 9 then
        transmit
        row = 14
    End if
Loop until row = 9

SSRT_start = replace(ssrt_in_date, " ", "/")
If SSRT_start = "__/__/____" then SSRT_start = ""
If cdate(SSRT_start) <= cdate("02/01/2018") then SSRT_start = "02/01/2018"

SSRT_end = replace(ssrt_out_date, " ", "/")
If SSRT_end = "__/__/____" then SSRT_end = ""

'----------------------------------------------------------------------------------------------------MEMB and ADDR panels
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadScreen client_PMI, 8, 4, 46
client_PMI = trim(client_PMI)
client_PMI = right("00000000" & client_pmi, 8)

EMReadScreen client_DOB, 10, 8, 42
client_DOB = replace(client_DOB, " ", "")

Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
If mail_line_one = "" then
	addr_line_01 = resi_line_one
	addr_line_02 = resi_line_two
	city_line = resi_city
	state_line = resi_state
	zip_line = resi_zip
Else
	addr_line_01 = mail_line_one
	addr_line_02 = mail_line_two
	city_line = mail_city
	state_line = mail_state
	zip_line = mail_zip
End if

'----------------------------------------------------------------------------------------------------FACI panel
Call navigate_to_MAXIS_screen("STAT", "FACI")
Call write_value_and_transmit ("01", 20, 76) 	'For member 01 - All GRH cases should be for member 01.
Call write_value_and_transmit ("01", 20, 79)    'Ensuring we're on the 1st panel

'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
EMReadScreen FACI_total_check, 1, 2, 78
If FACI_total_check = "0" then script_end_procedure("Case does not have a FACI panel. The script will now end.")
'Matching the FACI panel vendor number to the SSRT panel vendor number
faci_found = False
Do
    EMReadScreen faci_vendor_number, 8, 5, 43		'Enters vendor number
    If faci_vendor_number = SSRT_vendor_number then
        faci_found = true
        'Gathering approval county information
        EMReadScreen approval_county, 2, 12, 71
        approval_county = "0" & approval_county
        exit do	'when the correct
    else
        faci_found = False
    End if
    transmit
    EMReadScreen last_panel, 5, 24, 2
Loop until last_panel = "ENTER"	'This means that there are no other faci panels

If (faci_found = true and approval_county = "__") then script_end_procedure("Please fill in the 'Approval Cty' field on the FACI panel.")
If faci_found = False then script_end_procedure("FACI panel could not be found for the SSRT panel vendor. The script will now end.")

'----------------------------------------------------------------------------------------------------VNDS/VND2
Call Navigate_to_MAXIS_screen("MONY", "VNDS")
Call write_value_and_transmit(SSRT_vendor_number, 4, 59)
Call write_value_and_transmit("VND2", 20, 70)
EMReadScreen VND2_check, 4, 2, 54
If VND2_check <> "VND2" then script_end_procedure("Unable to find MONY/VND2 panel. The script will now end.")
EMReadScreen service_rate, 8, 16, 68		'Reading the service rate to input into MMIS
If IsNumeric(service_rate) = False then EMReadScreen service_rate, 8, 15, 72        'Handling for vendors with Rate 3 information
service_rate = replace(service_rate, ".", "")	'removing the period for input into MMIS
service_rate = trim(service_rate)

'----------------------------------------------------------------------------------------------------ELIG/GRH
'Trimming the vendor number of the preceding 0's since ELIG/GRH doesn't show the 0's.
If left(SSRT_vendor_number, 1) = "0" then
    Do
        SSRT_vendor_number = right(SSRT_vendor_number, len(SSRT_vendor_number) - 1)
    Loop until left(SSRT_vendor_number, 1) <> "0"
End if

Call Navigate_to_MAXIS_screen("ELIG", "GRH ")
EMReadScreen no_grh, 10, 24, 2		'NO GRH version means no conversion to MMIS will take place
If no_grh = "NO VERSION" then script_end_procedure("There are no GRH eligibility results. Please review. The script will now end.")

Call write_value_and_transmit("99", 20, 79)
'This brings up the FS versions of eligibility results to search for approved versions
status_row = 7
Do
	EMReadScreen app_status, 8, status_row, 50
	If trim(app_status) = "" then script_end_procedure("There are no GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". The script will now end.")
	If app_status = "UNAPPROV" Then status_row = status_row + 1
	IF app_status = "APPROVED" then
		EMReadScreen vers_number, 1, status_row, 23
		Call write_value_and_transmit(vers_number, 18, 54)
		exit do
 	End if
Loop until app_status = "APPROVED" or trim(app_status) = ""

If app_status <> "APPROVED" then script_end_procedure("There are no approved GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". The script will now end.")

'----------------------------------------------------------------------------------------------------ELIG/GRFB
Call write_value_and_transmit("GRFB", 20, 71)
Call write_value_and_transmit("x", 11, 3)
'Ensuring a rate 2 is found. If none or more than one are found, MMIS will not be updated.
row = 15
Do
    EMReadScreen rate_two_check, 8, row, 8
    rate_two_check = Trim(rate_two_check)
    If rate_two_check = SSRT_vendor_number then
        exit do
    else
        row = row + 1
    End if
Loop until row = 20

If rate_two_check = "" then script_end_procedure("GRH eligibility doesn't reflect Rate 2 vendor information, or the SSRT vendor number did not match ELIG/GRFB vendor number. The script will now end.")
PF3' out of ELIG/GRFB

'----------------------------------------------------------------------------------------------------Main selection dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 170, "Select the SSR start and end dates for "  & SSRT_vendor_name
  CheckBox 60, 25, 65, 10, DISA_start, disa_start_checkbox
  CheckBox 150, 25, 65, 10, DISA_end, disa_end_checkbox
  CheckBox 60, 45, 65, 10, revw_start, revw_start_checkbox
  CheckBox 150, 45, 65, 10, revw_end, revw_end_checkbox
  CheckBox 60, 65, 65, 10, SSRT_start, SSRT_start_checkbox
  CheckBox 150, 65, 65, 10, SSRT_end, SSRT_end_checkbox
  EditBox 60, 85, 55, 15, custom_start
  EditBox 150, 85, 55, 15, custom_end
  EditBox 85, 110, 190, 15, custom_dates_explained
  EditBox 85, 130, 190, 15, other_notes
  EditBox 85, 150, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 150, 40, 15
    CancelButton 235, 150, 40, 15
    PushButton 15, 25, 30, 10, "DISA", DISA_Button
    PushButton 15, 45, 30, 10, "REVW", REVW_button
    PushButton 15, 65, 30, 10, "SSRT", SSRT_Button
  Text 5, 115, 75, 10, "Explain custom dates:"
  Text 5, 90, 45, 10, "Custom date:"
  Text 5, 135, 75, 10, "Other SSR/GRH notes:"
  GroupBox 145, 10, 85, 95, "Select the SSR end date"
  Text 20, 155, 60, 10, "Worker signature:"
  GroupBox 50, 10, 85, 95, "Select the SSR start date"
  GroupBox 235, 10, 40, 75, "Navigation"
  ButtonGroup ButtonPressed
    PushButton 240, 40, 30, 10, "JOBS", JOBS_button
    PushButton 240, 25, 30, 10, "FACI", FACI_button
    PushButton 240, 55, 30, 10, "MAXIS", MAXIS_button
    PushButton 240, 70, 30, 10, "MMIS", MMIS_button
EndDialog

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
    DO
        DO
            dialog Dialog1				'main dialog
            cancel_confirmation
            'Navigation button handling
            MAXIS_dialog_navigation
            If ButtonPressed = MAXIS_button then Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
            If ButtonPressed = MMIS_button then
                Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
                Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen
                Call write_value_and_transmit(client_PMI, 10, 36)
            End if

            start_date = "" 'revaluing the variables for the start_date date
            end_date = ""   'revaluing the variables for the end date
            custom_date = ""
            total_units = ""
            If disa_start_checkbox = checked then start_date = start_date & disa_start
            If revw_start_checkbox = checked then start_date = start_date & revw_start
            If SSRT_start_checkbox = checked then start_date = start_date & SSRT_start
            If trim(custom_start) <> "" then
                start_date = start_date & custom_start
                custom_date = true
            End if

            If Disa_end_checkbox = checked then end_date = end_date & DISA_end
            If revw_end_checkbox = checked then end_date = end_date & revw_end
            If SSRT_end_checkbox = checked then end_date = end_date & SSRT_end
            If trim(custom_end) <> "" then
                end_date = end_date & custom_end
                custom_date = true
            End if
        Loop until ButtonPressed = -1

        err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
        If trim(start_date) = "" or IsDate(start_date) = false THEN err_msg = err_msg & vbCr & "Select/enter one valid start date."		'mandatory field
        IF trim(end_date) = "" or IsDate(end_date) = false THEN err_msg = err_msg & vbCr & "Select/enter one valid end date."		'mandatory field
        'If total_units > 365 THEN err_msg = err_msg & vbCr & "You cannot enter an agreement for more than 365 days. Select a new start and/or end dates."   ' Cannot be over 365 days.
        If (custom_date = True and trim(custom_dates_explained) = "") THEN err_msg = err_msg & vbCr & "Explain the reason for selecting custom dates."		'mandatory field
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "Enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'------------------------------------------------------------------------------------------------------calcuationas and conversion for MMIS
total_units = datediff("d", start_date, end_date) + 1   'Determining the total units to enter into MMIS.
MAXIS_agree_period = start_date & "-" & end_date

start_mo =  right("0" &  DatePart("m",    start_date), 2)
start_day = right("0" &  DatePart("d",    start_date), 2)
start_yr =  right(       DatePart("yyyy", start_date), 2)

output_start_date = start_mo & start_day & start_yr

end_mo =  right("0" &  DatePart("m",    end_date), 2)
end_day = right("0" &  DatePart("d",    end_date), 2)
end_yr =  right(       DatePart("yyyy", end_date), 2)

output_end_date = end_mo & end_day & end_yr

'----------------------------------------------------------------------------------------------------MMIS portion of the script
Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen

EmWriteScreen client_PMI, 10, 36
Call write_value_and_transmit("C", 3, 22)	'Checking to make sure that more than one agreement is not listed by trying to change (C) the information for the PMI selected.
EMReadScreen active_agreement, 12, 24, 2

If active_agreement = "NO DOCUMENTS" then
    duplicate_agreement = False 'no agreements exists in MMIS
else
	EMReadScreen AGMT_status, 31, 3, 19
    AGMT_status = trim(AGMT_status)
    If AGMT_status = "START DT:        END DT:" then
        row = 6
        Do
            EMReadScreen agreement_status, 1, row, 60
            EMReadScreen ASEL_start_date, 6, row, 63
            EmReadscreen ASEL_end_date, 6, row, 70
            ASEL_period = ASEL_start_date & "-" & ASEL_end_date
            output_period = output_start_date & "-" & output_end_date
            If agreement_status = "A" then
                If ASEL_period = output_period then
                    duplicate_agreement = True
                    script_end_procedure("An approved agreement already exists for the time frame selected. Please review the case. The script will now end.")
                Else
                    duplicate_agreement = False
                    row = row + 1
                End if
            ElseIf agreement_status = "D" then
                duplicate_agreement = False
                row = row + 1
            Else
                duplicate_agreement = False
            End if
        Loop until trim(agreement_status) = ""
    Else
        EMReadScreen agreement_status, 8, 3, 19
        If agreement_status = "APPROVED" then
            EMReadScreen ASA1_start_date, 6, row, 63
            EmReadscreen ASA1_end_date, 6, row, 70
            ASA1_period = ASA1_start_date & "-" & ASA1_end_date
            output_period = output_start_date & "-" & output_end_date

            If ASA1_period = output_period then
                duplicate_agreement = True
                script_end_procedure("An approved agreement already exists for the time frame selected. Please review the case. The script will now end.")
            Else
                duplicate_agreement = false
            End if
        Else
            duplicate_agreement = False
        End if
    End if
    PF6 'back to AKEY screen
End if

If duplicate_agreement = true then script_end_procedure("It appears an approved agreement already exists. Please review the case. The script will now end.")

If duplicate_agreement = False then
    Call clear_line_of_text(10, 36) 	'clears out the PMI number. Cannot add new agreement with PMI listed on AKEY.
    EmWriteScreen "A", 3, 22					'Selects the action code (A)
    EmWriteScreen "T", 3, 71					'Selecs the service agreement option (T)
    Call write_value_and_transmit("2", 7, 77)	'Enters the agreement type and transmits

    '----------------------------------------------------------------------------------------------------ASA1 screen
    Call MMIS_panel_confirmation("ASA1", 51)         'ensuring we are on the right MMIS screen

    EmWriteScreen output_start_date, 4, 64				'Start date
    EmWriteScreen output_end_date, 4, 71				'End date
    EmWriteScreen client_PMI, 8, 64						'Enters the client's PMI
    EmWriteScreen client_DOB, 9, 19						'Enters the client's DOB
    EmWriteScreen approval_county, 11, 19				'Enters 3 digit CO of SVC
    EmWriteScreen approval_county, 11, 39				'Enters 3 digit CO of RES
    Call write_value_and_transmit(approval_county, 11, 64)	'Enters 3 digit CO of FIN RESP and transmits

    Call MMIS_panel_confirmation("ASA2", 51)         'ensuring we are on the right MMIS screen
    transmit 	'no action required on ASA2
    '----------------------------------------------------------------------------------------------------ASA3 screen
    Call MMIS_panel_confirmation("ASA3", 51)         'ensuring we are on the right MMIS screen
    EMWriteScreen "H0043", 7, 36
    EMWriteScreen "U5", 7, 44
    EmWriteScreen output_start_date, 8, 60
    EmWriteScreen output_end_date, 8, 67
    EMWriteScreen service_rate, 9, 20			'Enters service rate from VND2
    EMWriteScreen total_units, 9, 60

    Call write_value_and_transmit(NPI_number, 10, 20)	'Enters the NPI number then transmits
    Emreadscreen NPI_issue, 26, 24, 1
    If NPI_issue = "CORRECT HIGHLIGHTED FIELDS" then
    	Update_MMIS = False
    	script_end_procedure("Issue with NPI# in MMIS. Please review case/report issue to the Quality Improvement Team. The script will now end.")
    	Call clear_line_of_text(10, 20) 	'clears out the NPI number so that the rest of the information can be saved.
    	PF3
    else
        '----------------------------------------------------------------------------------------------------PPOP screen handling
        EMReadScreen PPOP_check, 4, 1, 52
        If PPOP_check = "PPOP" then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 180, 90, "PPOP screen - Choose Facility"
                ButtonGroup ButtonPressed
                OkButton 65, 70, 50, 15
                CancelButton 120, 70, 50, 15
                Text 5, 5, 170, 35, "Please select the correct facility name/address from the list in PPOP by putting a 'X' next to the name. DO NOT TRANSMIT. Press OK when ready. Press CANCEL to stop the script."
                Text 5, 45, 175, 20, "* Provider types for GRH must be '18/H COMM PRV' and the status must be '1 ACTIVE.'"
            EndDialog
            Do
                dialog Dialog1
                cancel_confirmation
            Loop until ButtonPressed = -1
			EMReadScreen PPOP_check, 4, 1, 52
            If PPOP_check = "PPOP" then transmit     'to exit PPOP
            If PPOP_check = "SA3 " then transmit    'to navigate to ACF1 - this is the partial screen check for ASA3
            transmit ' to next available screen (does not need to be updated)
            Call write_value_and_transmit("ACF1", 1, 51)
        End if

        '----------------------------------------------------------------------------------------------------ACF1 screen
        Call MMIS_panel_confirmation("ACF1", 51)         'ensuring we are on the right MMIS screen
        EmWriteScreen addr_line_01, 5, 8	'enters the clients address
        EmWriteScreen addr_line_02, 5, 37
        EmWriteScreen city_line, 6, 8
        EmWriteScreen state_line, 6, 34
        EmWriteScreen zip_line, 6, 42
        Call write_value_and_transmit("ASA1", 1, 8)		'direct navigating to ASA1

        '----------------------------------------------------------------------------------------------------ASA1 screen
        Call MMIS_panel_confirmation("ASA1", 51)         'ensuring we are on the right MMIS screen
         PF9 								'triggering stat edits
        EmreadScreen error_codes, 79, 20, 2	'checking for stat edits
        If trim(error_codes) <> "00 140  4          01 140  4" then
        	script_end_procedure("MMIS stat edits exist. Edit codes are: " & error_codes & vbcr & "PF3 to save what's been updated in MMIS, and follow up on the error codes. The script will now end.")
        else
        	EMWriteScreen "A", 3, 17						'Updating the AMT type/STAT to A for approved
        	Call write_value_and_transmit("ASA3", 1, 8)		'direct navigating to ASA3
        	Call MMIS_panel_confirmation("ASA3", 51)         'ensuring we are on the right MMIS screen
        	EMWriteScreen "A", 12, 19						'Updating the STAT CD/DATE to A for approved
        	Update_MMIS = true
            PF3 '	to save changes

            Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen
            EMReadScreen authorization_number, 13, 9, 36
            authorization_number = trim(authorization_number)
            EMReadscreen approval_message, 16, 24, 2
        End if
    End if
End if

'----------------------------------------------------------------------------------------------------Back to MAXIS & CASE/NOTE
If Update_MMIS = True then
    Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
    Call check_for_MAXIS(False)

    If disa_start_checkbox = checked then start_date_source = ", PSN start date."
    If revw_start_checkbox = checked then start_date_source = ", start of certification period."
    If SSRT_start_checkbox = checked then start_date_source = ", SSRT start date."

    If Disa_end_checkbox = checked then end_date_source = ", PSN end date."
    If revw_end_checkbox = checked then end_date_source = ", end of certification period."
    If SSRT_end_checkbox = checked then end_date_source = ", SSRT end date."

    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9
    Call write_variable_in_CASE_NOTE("GRH Rate 2 SSR added to MMIS for " & SSRT_vendor_name)
    Call write_bullet_and_variable_in_CASE_NOTE("NPI #", npi_number)
    Call write_bullet_and_variable_in_CASE_NOTE("MMIS authorization number", authorization_number)
    Call write_variable_in_CASE_NOTE("* SSR start date: " & start_date & start_date_source)   'Hard coded for now
    Call write_variable_in_CASE_NOTE("* SSR end date: " & end_date & end_date_source)
    Call write_bullet_and_variable_in_CASE_NOTE("Explanation of custom date", custom_dates_explained)
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
    PF3
End if

script_end_procedure("Success! Your case has been updated in MMIS and case noted in MAXIS.")
