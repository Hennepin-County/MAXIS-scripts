'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script "library" contains functions and variables that the other BlueZone scripts use very commonly. The other BlueZone scripts contain a few lines of code that run
'this script and get the functions. This saves time in writing and copy/pasting the same functions in many different places. Only add functions to this script if they've
'been tested in other scripts first. This document is actively used by live scripts, so it needs to be functionally complete at all times.
'
'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============

'CHANGELOG BLOCK ===========================================================================================================
actual_script_name = name_of_script
name_of_script = "Functions Library"
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/25/2019", "We want to hear from YOU! Please respond to our Survey, link and details can be found in Hot Topics.", "Casey Love, Hennepin County")
call changelog_update("06/21/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
name_of_script = actual_script_name
'END CHANGELOG BLOCK =======================================================================================================

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
Dim checked, unchecked, cancel, OK, blank		'Declares this for Option Explicit users

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs
blank = ""

Dim STATS_counter, STATS_manualtime, STATS_denomination, script_run_lowdown

'Time arrays which can be used to fill an editbox with the convert_array_to_droplist_items function
time_array_15_min = array("7:00 AM", "7:15 AM", "7:30 AM", "7:45 AM", "8:00 AM", "8:15 AM", "8:30 AM", "8:45 AM", "9:00 AM", "9:15 AM", "9:30 AM", "9:45 AM", "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "1:00 PM", "1:15 PM", "1:30 PM", "1:45 PM", "2:00 PM", "2:15 PM", "2:30 PM", "2:45 PM", "3:00 PM", "3:15 PM", "3:30 PM", "3:45 PM", "4:00 PM", "4:15 PM", "4:30 PM", "4:45 PM", "5:00 PM", "5:15 PM", "5:30 PM", "5:45 PM", "6:00 PM")
time_array_30_min = array("7:00 AM", "7:30 AM", "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM")

'Array of all the upcoming holidays
HOLIDAYS_ARRAY = Array(#1/1/19#, #1/21/19#, #2/18/19#, #5/27/19#, #7/4/19#, #9/2/19#, #11/11/19#, #11/28/19#, #11/29/19#, #12/24/19#, #12/15/19#, #1/1/20#)

'Determines CM and CM+1 month and year using the two rightmost chars of both the month and year. Adds a "0" to all months, which will only pull over if it's a single-digit-month
Dim CM_mo, CM_yr, CM_plus_1_mo, CM_plus_1_yr, CM_plus_2_mo, CM_plus_2_yr
'var equals...  the right part of...    the specific part...    of either today or next month... just the right 2 chars!
CM_mo =         right("0" &             DatePart("m",           date                             ), 2)
CM_yr =         right(                  DatePart("yyyy",        date                             ), 2)

CM_plus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", 1, date)            ), 2)
CM_plus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", 1, date)            ), 2)

CM_plus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", 2, date)            ), 2)
CM_plus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", 2, date)            ), 2)

If worker_county_code   = "" then worker_county_code = "MULTICOUNTY"
IF PRISM_script <> true then county_name = ""		'VKC NOTE 08/12/2016: ADDED IF...THEN CONDITION BECAUSE PRISM IS STILL USING THIS VARIABLE IN ALL SCRIPTS.vbs. IT WILL BE REMOVED AND THIS CAN BE RESTORED.

If ButtonPressed <> "" then ButtonPressed = ""		'Defines ButtonPressed if not previously defined, allowing scripts the benefit of not having to declare ButtonPressed all the time

'preloading boolean variables for tabbing dialogs
pass_one = False
pass_two = False
pass_three = False
pass_four = False
pass_five = False
pass_six = False
pass_seven = False
pass_eight = False
pass_nine = False
pass_ten = False

show_one = True
show_two = True
show_three = True
show_four = True
show_five = True
show_six = True
show_seven = True
show_eight = True
show_nine = True
show_ten = True

tab_button = False

'Preloading worker_signature, as a constant to be used in scripts---------------------------------------------------------------------------------------------------------

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

'Now it determines the signature
With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
	If .FileExists(user_myDocs_folder & "workersig.txt") Then												'If the workersig.txt file exists...
		Set get_worker_sig = CreateObject("Scripting.FileSystemObject")										'Create another FSO
		Set worker_sig_command = get_worker_sig.OpenTextFile(user_myDocs_folder & "workersig.txt")			'Open the text file
		worker_sig = worker_sig_command.ReadAll																'Read the text file
		IF worker_sig <> "" THEN worker_signature = worker_sig												'If it isn't blank then the worker_signature should be populated with the contents of the file
		worker_sig_command.Close																			'Closes the file
	END IF
END WITH

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName

'=========================================================================================================================================================================== FUNCTIONS RELATED TO GLOBAL CONSTANTS
FUNCTION income_test_SNAP_categorically_elig(household_size, income_limit) '165% FPG
	'See Combined Manual 0019.06
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "17") OR (MAXIS_footer_year = "18") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1670										'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 2264										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 2858
		IF household_size = 4 THEN income_limit = 3452
		IF household_size = 5 THEN income_limit = 4046
		IF household_size = 6 THEN income_limit = 4640
		IF household_size = 7 THEN income_limit = 5234
		IF household_size = 8 THEN income_limit = 5828
		IF household_size > 8 THEN income_limit = 5828 + (594 * (household_size- 8))
	ELSE
        IF household_size = 1 THEN income_limit = 1634
        IF household_size = 2 THEN income_limit = 2203
        IF household_size = 3 THEN income_limit = 2772
        IF household_size = 4 THEN income_limit = 3342
        IF household_size = 5 THEN income_limit = 3911
        IF household_size = 6 THEN income_limit = 4480
        IF household_size = 7 THEN income_limit = 5051
        IF household_size = 8 THEN income_limit = 5623
        IF household_size > 8 THEN income_limit = 5623 + (572 * (household_size- 8))
	END IF

	valid_through_date = #10/01/2019#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

FUNCTION income_test_SNAP_gross(household_size, income_limit) '130% FPG
	'See Combined Manual 0019.06
	'Also used for sponsor income
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "17") OR (MAXIS_footer_year = "18") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1316										'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 1784										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 2252
		IF household_size = 4 THEN income_limit = 2720
		IF household_size = 5 THEN income_limit = 3188
		IF household_size = 6 THEN income_limit = 3656
		IF household_size = 7 THEN income_limit = 4124
		IF household_size = 8 THEN income_limit = 4592
		IF household_size > 8 THEN income_limit = 4592 + (468 * (household_size- 8))
	ELSE
        IF household_size = 1 THEN income_limit = 1307
        IF household_size = 2 THEN income_limit = 1760
        IF household_size = 3 THEN income_limit = 2213
        IF household_size = 4 THEN income_limit = 2665
        IF household_size = 5 THEN income_limit = 3118
        IF household_size = 6 THEN income_limit = 3571
        IF household_size = 7 THEN income_limit = 4024
        IF household_size = 8 THEN income_limit = 4477
        IF household_size > 8 THEN income_limit = 4477 + (453 * (sponsor_HH_size - 8))
	END IF

	valid_through_date = #10/01/2019#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

FUNCTION income_test_SNAP_net(household_size, income_limit)
	'See Combined Manual 0020.12 - Net income standard 100% FPG
	'When using this function, you can pass (ubound(hh_array) + 1) for household_size
	IF ((MAXIS_footer_month * 1) >= "10" AND (MAXIS_footer_year * 1) >= "16") OR (MAXIS_footer_year = "17") THEN  'This will allow the function to be used during the transition period when both income limits can be used.
		IF household_size = 1 THEN income_limit = 1012										'Going forward you should only have to change the years and this should hold.
		IF household_size = 2 THEN income_limit = 1372										'Multipled the footer months by 1 to insure they become numeric
		IF household_size = 3 THEN income_limit = 1732
		IF household_size = 4 THEN income_limit = 2092
		IF household_size = 5 THEN income_limit = 2452
		IF household_size = 6 THEN income_limit = 2812
		IF household_size = 7 THEN income_limit = 3172
		IF household_size = 8 THEN income_limit = 3532
		IF household_size > 8 THEN income_limit = 3532 + (360 * (household_size- 8))
	ELSE
        IF household_size = 1 THEN income_limit = 1005
        IF household_size = 2 THEN income_limit = 1354
        IF household_size = 3 THEN income_limit = 1702
        IF household_size = 4 THEN income_limit = 2050
        IF household_size = 5 THEN income_limit = 2399
        IF household_size = 6 THEN income_limit = 2747
        IF household_size = 7 THEN income_limit = 3095
        IF household_size = 8 THEN income_limit = 3444
        IF household_size > 8 THEN income_limit = 3444 + (349 * (household_size- 8))
	END IF

	valid_through_date = #10/01/2019#
	IF DateDiff("D", date, valid_through_date) <= 0 THEN
		out_of_date_warning = MsgBox ("This script appears to be using out of date income limits. Please contact a scripts administrator to have this updated." & vbNewLine & vbNewLine & "Press OK to continue the script. Press CANCEL to stop the script.", vbOKCancel + vbCritical + vbSystemModal, "NOTICE!!!")
		IF out_of_date_warning = vbCancel THEN script_end_procedure("")
	END IF
END FUNCTION

FUNCTION ten_day_cutoff_check(MAXIS_footer_month, MAXIS_footer_year, ten_day_cutoff)
	'All 10-day cutoff dates are provided in POLI/TEMP TE19.132
	IF MAXIS_footer_month = "01" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #01/18/2019#
	ELSEIF MAXIS_footer_month = "02" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #02/15/2019#
	ELSEIF MAXIS_footer_month = "03" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #03/21/2019#
	ELSEIF MAXIS_footer_month = "04" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #04/18/2019#
	ELSEIF MAXIS_footer_month = "05" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #05/21/2019#
	ELSEIF MAXIS_footer_month = "06" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #06/20/2019#
	ELSEIF MAXIS_footer_month = "07" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #07/19/2019#
	ELSEIF MAXIS_footer_month = "08" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #08/21/2019#
	ELSEIF MAXIS_footer_month = "09" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #09/19/2019#
	ELSEIF MAXIS_footer_month = "10" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #10/21/2019#
	ELSEIF MAXIS_footer_month = "11" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #11/19/2019#
	ELSEIF MAXIS_footer_month = "12" AND MAXIS_footer_year = "19" THEN
		ten_day_cutoff = #12/19/2019#
    ELSEIF MAXIS_footer_month = "12" AND MAXIS_footer_year = "18" THEN
    	ten_day_cutoff = #12/20/2018#                                      'last month of current year
	ELSE
		MsgBox "You have entered a date (" & MAXIS_footer_month & "/" & MAXIS_footer_year & ") not supported by this function. Please contact a scripts administrator to determine if the script requires updating.", vbInformation + vbSystemModal, "NOTICE"
	END IF
END FUNCTION

'=========================================================================================================================================================================== CLASSES USED BY SCRIPTS
'A class for each script item
class script

	public script_name             	'The familiar name of the script
	public file_name               	'The actual file name
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
    public SIR_instructions_URL    	'The instructions URL in SIR
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run
	public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property

end class
'=========================================================================================================================================================================== END OF CLASSES

























'BELOW ARE THE ACTUAL FUNCTIONS--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function ABAWD_FSET_exemption_finder()
'--- This function screens for ABAWD/FSET exemptions for SNAP.
'===== Keywords: MAXIS, ABAWD, FSET, exemption, SNAP
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    '>>>>>Checking for privileged<<<<<
    row = 1
    col = 1
    EMSearch "PRIVILEGED", row, col
    IF row <> 0 THEN script_end_procedure("This case appears to be privileged. The script cannot access it.")

    DO
    	CALL HH_member_custom_dialog(HH_member_array)
    	IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
    LOOP UNTIL uBound(HH_member_array) <> -1

    'Building a placeholder array for EATS group comparison
    placeholder_HH_array = ""
    person_count = 0
    FOR EACH person IN HH_member_array
    	placeholder_HH_array = placeholder_HH_array & person & ","
    NEXT

    CALL check_for_MAXIS(False)

    closing_message = ""

    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen cl_age, 2, 8, 76
    		IF cl_age = "  " THEN cl_age = 0
    		cl_age = cl_age * 1
    		IF cl_age < 18 OR cl_age >= 50 THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have exemption. Age = " & cl_age & "."
    	END IF
    NEXT

    CALL navigate_to_MAXIS_screen("STAT", "DISA")
    FOR EACH person IN HH_member_array
    	disa_status = false
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_DISA, 1, 2, 78
    		IF num_of_DISA <> "0" THEN
    			EMReadScreen disa_end_dt, 10, 6, 69
    			disa_end_dt = replace(disa_end_dt, " ", "/")
    			EMReadScreen cert_end_dt, 10, 7, 69
    			cert_end_dt = replace(cert_end_dt, " ", "/")
    			IF IsDate(disa_end_dt) = True THEN
    				IF DateDiff("D", date, disa_end_dt) > 0 THEN
    					closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA end date = " & disa_end_dt & "."
    					disa_status = True
    				END IF
    			ELSE
    				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    					closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA has no end date."
    					disa_status = True
    				END IF
    			END IF
    			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
    				IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
    			ELSE
    				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
    					EMReadScreen cert_begin_dt, 8, 7, 47
    					IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have disability exemption. DISA certification has no end date."
    				END IF
    			END IF
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>>> EATS GROUP
    FOR EACH person IN HH_member_array
    	CALL navigate_to_MAXIS_screen("STAT", "EATS")
    	eats_group_members = ""
    	memb_found = True
    	EMReadScreen all_eat_together, 1, 4, 72
    	IF all_eat_together = "_" THEN
    		eats_group_members = "01" & ","
    	ELSEIF all_eat_together = "Y" THEN
    		eats_row = 5
    		DO
    			EMReadScreen eats_person, 2, eats_row, 3
    			eats_person = replace(eats_person, " ", "")
    			IF eats_person <> "" THEN
    				eats_group_members = eats_group_members & eats_person & ","
    				eats_row = eats_row + 1
    			END IF
    		LOOP UNTIL eats_person = ""
    	ELSEIF all_eat_together = "N" THEN
    		eats_row = 13
    		DO
    			EMReadScreen eats_group, 38, eats_row, 39
    			find_memb01 = InStr(eats_group, person)
    			IF find_memb01 = 0 THEN
    				eats_row = eats_row + 1
    				IF eats_row = 18 THEN
    					memb_found = False
    					EXIT DO
    				END IF
    			END IF
    		LOOP UNTIL find_memb01 <> 0
    		eats_col = 39
    		DO
    			EMReadScreen eats_group, 2, eats_row, eats_col
    			IF eats_group <> "__" THEN
    				eats_group_members = eats_group_members & eats_group & ","
    				eats_col = eats_col + 4
    			END IF
    		LOOP UNTIL eats_group = "__"
    	END IF

    	IF memb_found = True THEN
    		IF placeholder_HH_array <> eats_group_members THEN script_end_procedure("You are asking the script to verify ABAWD and SNAP E&T exemptions for a household that does not match the EATS group. The script cannot support this request. It will now end." & vbCr & vbCr & "Please re-run the script selecting only the individuals in the EATS group.")
    		eats_group_members = trim(eats_group_members)
    		eats_group_members = split(eats_group_members, ",")

    		IF all_eat_together <> "_" THEN
    			CALL write_value_and_transmit("MEMB", 20, 71)
    			FOR EACH eats_pers IN eats_group_members
    				IF eats_pers <> "" AND person <> eats_pers THEN
    					CALL write_value_and_transmit(eats_pers, 20, 76)
    					EMReadScreen cl_age, 2, 8, 76
    					IF cl_age = "  " THEN cl_age = 0
    						cl_age = cl_age * 1
    						IF cl_age =< 17 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": May have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
    						END IF
    				END IF
    			NEXT
    		END IF

    		CALL write_value_and_transmit("DISA", 20, 71)
    		FOR EACH disa_pers IN eats_group_members
    			disa_status = false
    			IF disa_pers <> "" AND disa_pers <> person THEN
    				CALL write_value_and_transmit(disa_pers, 20, 76)
    				EMReadScreen num_of_DISA, 1, 2, 78
    				IF num_of_DISA <> "0" THEN
    					EMReadScreen disa_end_dt, 10, 6, 69
    					disa_end_dt = replace(disa_end_dt, " ", "/")
    					EMReadScreen cert_end_dt, 10, 7, 69
    					cert_end_dt = replace(cert_end_dt, " ", "/")
    					IF IsDate(disa_end_dt) = True THEN
    						IF DateDiff("D", date, disa_end_dt) > 0 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": MAY have an exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = TRUE
    						END IF
    					ELSEIF IsDate(disa_end_dt) = False THEN
    						IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    							closing_message = closing_message & vbCr & "* M" & person & " : MAY have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = true
    						END IF
    					END IF
    					IF IsDate(cert_end_dt) = True AND disa_status = False THEN
    						IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
    					ELSE
    						IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN
    							EMReadScreen cert_begin_dt, 8, 7, 47
    							IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
    						END IF
    					END IF
    				END IF
    			END IF
    		NEXT
    	END IF
    NEXT

    '>>>>>>>>>>>>>>EARNED INCOME
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		prosp_inc = 0
    		prosp_hrs = 0
    		prospective_hours = 0

    		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
    		EMWritescreen person, 20, 76
    		EMWritescreen "01", 20, 79				'ensures that we start at 1st job
    		transmit
    		EMReadScreen num_of_JOBS, 1, 2, 78
    		IF num_of_JOBS <> "0" THEN
    			DO
    			 	EMReadScreen jobs_end_dt, 8, 9, 49
    				EMReadScreen cont_end_dt, 8, 9, 73
    				IF jobs_end_dt = "__ __ __" THEN
    					CALL write_value_and_transmit("X", 19, 38)
    					EMReadScreen prosp_monthly, 8, 18, 56
    					prosp_monthly = trim(prosp_monthly)
    					IF prosp_monthly = "" THEN prosp_monthly = 0
    					prosp_inc = prosp_inc + prosp_monthly
    					EMReadScreen prosp_hrs, 8, 16, 50
    					IF prosp_hrs = "        " THEN prosp_hrs = 0
    					prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
    					EMReadScreen pay_freq, 1, 5, 64
    					IF pay_freq = "1" THEN
    						prosp_hrs = prosp_hrs
    					ELSEIF pay_freq = "2" THEN
    						prosp_hrs = (2 * prosp_hrs)
    					ELSEIF pay_freq = "3" THEN
    						prosp_hrs = (2.15 * prosp_hrs)
    					ELSEIF pay_freq = "4" THEN
    						prosp_hrs = (4.3 * prosp_hrs)
    					END IF
    					prospective_hours = prospective_hours + prosp_hrs
    				ELSE
    					jobs_end_dt = replace(jobs_end_dt, " ", "/")
    					IF DateDiff("D", date, jobs_end_dt) > 0 THEN
    						'Going into the PIC for a job with an end date in the future
    						CALL write_value_and_transmit("X", 19, 38)
    						EMReadScreen prosp_monthly, 8, 18, 56
    						prosp_monthly = trim(prosp_monthly)
    						IF prosp_monthly = "" THEN prosp_monthly = 0
    						prosp_inc = prosp_inc + prosp_monthly
    						EMReadScreen prosp_hrs, 8, 16, 50
    						IF prosp_hrs = "        " THEN prosp_hrs = 0
    						prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
    						EMReadScreen pay_freq, 1, 5, 64
    						IF pay_freq = "1" THEN
    							prosp_hrs = prosp_hrs
    						ELSEIF pay_freq = "2" THEN
    							prosp_hrs = (2 * prosp_hrs)
    						ELSEIF pay_freq = "3" THEN
    							prosp_hrs = (2.15 * prosp_hrs)
    						ELSEIF pay_freq = "4" THEN
    							prosp_hrs = (4.3 * prosp_hrs)
    						END IF
    						'added seperate incremental variable to account for multiple jobs
    						prospective_hours = prospective_hours + prosp_hrs
    					END IF
    				END IF
    				transmit		'to exit PIC
    				EMReadScreen JOBS_panel_current, 1, 2, 73
    				'looping until all the jobs panels are calculated
    				If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
    			Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
    		END IF

    		EMWriteScreen "BUSI", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_BUSI, 1, 2, 78
    		IF num_of_BUSI <> "0" THEN
    			DO
    				EMReadScreen busi_end_dt, 8, 5, 72
    				busi_end_dt = replace(busi_end_dt, " ", "/")
    				IF IsDate(busi_end_dt) = True THEN
    					IF DateDiff("D", date, busi_end_dt) > 0 THEN
    						EMReadScreen busi_inc, 8, 10, 69
    						busi_inc = trim(busi_inc)
    						EMReadScreen busi_hrs, 3, 13, 74
    						busi_hrs = trim(busi_hrs)
    						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
    						prosp_inc = prosp_inc + busi_inc
    						prosp_hrs = prosp_hrs + busi_hrs
    						prospective_hours = prospective_hours + busi_hrs
    					END IF
    				ELSE
    					IF busi_end_dt = "__/__/__" THEN
    						EMReadScreen busi_inc, 8, 10, 69
    						busi_inc = trim(busi_inc)
    						EMReadScreen busi_hrs, 3, 13, 74
    						busi_hrs = trim(busi_hrs)
    						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
    						prosp_inc = prosp_inc + busi_inc
    						prosp_hrs = prosp_hrs + busi_hrs
    						prospective_hours = prospective_hours + busi_hrs
    					END IF
    				END IF
    				transmit
    				EMReadScreen enter_a_valid, 13, 24, 2
    			LOOP UNTIL enter_a_valid = "ENTER A VALID"
    		END IF

    		EMWriteScreen "RBIC", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_RBIC, 1, 2, 78
    		IF num_of_RBIC <> "0" THEN closing_message = closing_message & vbCr & "* M" & person & ": Has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
    		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working 30 hours/wk (regardless of wage level) or earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
    		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption and SNAP E&T exemptions."
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>>>UNEA
    CALL navigate_to_MAXIS_screen("STAT", "UNEA")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_UNEA, 1, 2, 78
    		IF num_of_UNEA <> "0" THEN
    			DO
    				EMReadScreen unea_type, 2, 5, 37
    				EMReadScreen unea_end_dt, 8, 7, 68
    				unea_end_dt = replace(unea_end_dt, " ", "/")
    				IF IsDate(unea_end_dt) = True THEN
    					IF DateDiff("D", date, unea_end_dt) > 0 THEN
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
    					END IF
    				ELSE
    					IF unea_end_dt = "__/__/__" THEN
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
    					END IF
    				END IF
    				transmit
    				EMReadScreen enter_a_valid, 13, 24, 2
    			LOOP UNTIL enter_a_valid = "ENTER A VALID"
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>PBEN
    CALL navigate_to_MAXIS_screen("STAT", "PBEN")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		EMWriteScreen "PBEN", 20, 71
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_PBEN, 1, 2, 78
    		IF num_of_PBEN <> "0" THEN
    			pben_row = 8
    			DO
    			    IF pben_type = "12" THEN		'UI pending'
    					EMReadScreen pben_disp, 1, pben_row, 77
    					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
    						closing_message = closing_message & vbCr & "* M" & person & ": Appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption."
    						EXIT DO
    					END IF
    				ELSE
    					pben_row = pben_row + 1
    				END IF
    			LOOP UNTIL pben_row = 14
    		END IF
    	END IF
    NEXT

    '>>>>>>>>>>PREG
    CALL navigate_to_MAXIS_screen("STAT", "PREG")
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_PREG, 1, 2, 78
            EMReadScreen preg_due_dt, 8, 10, 53
            preg_due_dt = replace(preg_due_dt, " ", "/")
    		EMReadScreen preg_end_dt, 8, 12, 53

    		IF num_of_PREG <> "0" THen
                If preg_due_dt <> "__/__/__" Then
                    If DateDiff("d", date, preg_due_dt) > 0 AND preg_end_dt = "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active pregnancy. Please review for ABAWD exemption."
                    If DateDiff("d", date, preg_due_dt) < 0 Then closing_message = closing_message & vbCr & "* M" & person & ": Appears to have an overdue pregnancy, person may meet a minor child exemption. Contact client."
                End If
            End If
        END IF
    NEXT

    '>>>>>>>>>>PROG
    CALL navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash1_status, 4, 6, 74
    EMReadScreen cash2_status, 4, 7, 74
    IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN closing_message = closing_message & vbCr & "* Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."

    '>>>>>>>>>>ADDR
    CALL navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_code, 1, 10, 43
    EmReadscreen addr_line_01, 16, 6, 43

    IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN closing_message = closing_message & vbCr & "* Client is claiming homelessness. If client has barriers to employment, they could meet the 'Unfit for Employment' exemption. Exemption began 05/2018."

    '>>>>>>>>>SCHL/STIN/STEC
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
            CALL navigate_to_MAXIS_screen("STAT", "SCHL")
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_SCHL, 1, 2, 78
    		IF num_of_SCHL = "1" THEN
    			EMReadScreen school_status, 1, 6, 40
    			IF school_status <> "N" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
    		ELSE
    			EMWriteScreen "STIN", 20, 71
    			CALL write_value_and_transmit(person, 20, 76)
    			EMReadScreen num_of_STIN, 1, 2, 78
    			IF num_of_STIN = "1" THEN
    				STIN_row = 8
    				DO
    					EMReadScreen cov_thru, 5, STIN_row, 67
    					IF cov_thru <> "__ __" THEN
    						cov_thru = replace(cov_thru, " ", "/01/")
    						cov_thru = DateAdd("M", 1, cov_thru)
    						cov_thru = DateAdd("D", -1, cov_thru)
    						IF DateDiff("D", date, cov_thru) > 0 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
    							EXIT DO
    						ELSE
    							STIN_row = STIN_row + 1
    							IF STIN_row = 18 THEN
    								PF20
    								STIN_row = 8
    								EMReadScreen last_page, 21, 24, 2
    								IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
    							END IF
    						END IF
    					ELSE
    						EXIT DO
    					END IF
    				LOOP
    			ELSE
    				EMWriteScreen "STEC", 20, 71
    				CALL write_value_and_transmit(person, 20, 76)
    				EMReadScreen num_of_STEC, 1, 2, 78
    				IF num_of_STEC = "1" THEN
    					STEC_row = 8
    					DO
    						EMReadScreen stec_thru, 5, STEC_row, 48
    						IF stec_thru <> "__ __" THEN
    							stec_thru = replace(stec_thru, " ", "/01/")
    							stec_thru = DateAdd("M", 1, stec_thru)
    							stec_thru = DateAdd("D", -1, stec_thru)
    							IF DateDiff("D", date, stec_thru) > 0 THEN
    								closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
    								EXIT DO
    							ELSE
    								STEC_row = STEC_row + 1
    								IF STEC_row = 17 THEN
    									PF20
    									STEC_row = 8
    									EMReadScreen last_page, 21, 24, 2
    									IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
    								END IF
    							END IF
    						ELSE
    							EXIT DO
    						END IF
    					LOOP
    				END IF
    			END IF
    		END IF
    	END IF
    	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
    NEXT

    household_persons = ""
    pers_count = 0

    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		IF pers_count = uBound(HH_member_array) THEN
    			IF pers_count = 0 THEN
    				household_persons = household_persons & person
    			ELSE
    				household_persons = household_persons & "and " & person
    			END IF
    		ELSE
    			household_persons = household_persons & person & ", "
    			pers_count = pers_count + 1
    		END IF
    	END IF
    NEXT

    IF closing_message = "" THEN
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "It appears there are NO missed exemptions for ABAWD or SNAP E&T in MAXIS for this case. The script has checked ADDR, EATS, MEMB, DISA, JOBS, BUSI, RBIC, UNEA, PREG, PROG, PBEN, SCHL, STIN, and STEC for member(s) " & household_persons & "." & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    ELSE
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "The script has checked for ABAWD and SNAP E&T exemptions coded in MAXIS for member(s) " & household_persons & "." & vbCr & closing_message & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    END IF

    'Displaying the results...now with added MsgBox bling.
    'vbSystemModal will keep the results in the foreground.
    MsgBox closing_message, vbInformation + vbSystemModal, "ABAWD/FSET Exemption Check -- Results"

    STATS_counter = STATS_counter - 1		'Removing one instance from the STATS Counter as it started with one at the beginning
End Function


function add_ACCI_to_variable(ACCI_variable)
'--- This function adds STAT/ACCI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ ACCI_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, ACCI
  EMReadScreen ACCI_date, 8, 6, 73
  ACCI_date = replace(ACCI_date, " ", "/")
  If datediff("yyyy", ACCI_date, now) < 5 then
    EMReadScreen ACCI_type, 2, 6, 47
    If ACCI_type = "01" then ACCI_type = "Auto"
    If ACCI_type = "02" then ACCI_type = "Workers Comp"
    If ACCI_type = "03" then ACCI_type = "Homeowners"
    If ACCI_type = "04" then ACCI_type = "No Fault"
    If ACCI_type = "05" then ACCI_type = "Other Tort"
    If ACCI_type = "06" then ACCI_type = "Product Liab"
    If ACCI_type = "07" then ACCI_type = "Med Malprac"
    If ACCI_type = "08" then ACCI_type = "Legal Malprac"
    If ACCI_type = "09" then ACCI_type = "Diving Tort"
    If ACCI_type = "10" then ACCI_type = "Motorcycle"
    If ACCI_type = "11" then ACCI_type = "MTC or Other Bus Tort"
    If ACCI_type = "12" then ACCI_type = "Pedestrian"
    If ACCI_type = "13" then ACCI_type = "Other"
    ACCI_variable = ACCI_variable & ACCI_type & " on " & ACCI_date & ".; "
  End if
end function

function add_ACCT_to_variable(ACCT_variable)
'--- This function adds STAT/ACCT data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ ACCT_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, ACCT
  EMReadScreen ACCT_amt, 8, 10, 46
  ACCT_amt = trim(ACCT_amt)
  ACCT_amt = "$" & ACCT_amt
  EMReadScreen ACCT_type, 2, 6, 44
  EMReadScreen ACCT_location, 20, 8, 44
  ACCT_location = replace(ACCT_location, "_", "")
  ACCT_location = split(ACCT_location)
  For each ACCT_part in ACCT_location
    If ACCT_part <> "" then
      first_letter = ucase(left(ACCT_part, 1))
      other_letters = LCase(right(ACCT_part, len(ACCT_part) -1))
      If len(ACCT_part) > 3 then
        new_ACCT_location = new_ACCT_location & first_letter & other_letters & " "
      Else
        new_ACCT_location = new_ACCT_location & ACCT_part & " "
      End if
    End if
  Next
  EMReadScreen ACCT_ver, 1, 10, 63
  If ACCT_ver = "N" then
    ACCT_ver = ", no proof provided"
  Else
    ACCT_ver = ""
  End if
  ACCT_variable = ACCT_variable & ACCT_type & " at " & new_ACCT_location & "(" & ACCT_amt & ")" & ACCT_ver & ".; "
  new_ACCT_location = ""
end function

function add_BUSI_to_variable(variable_name_for_BUSI)
'--- This function adds STAT/BUSI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ BUSI_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, BUSI

	'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
	EMReadScreen BUSI_footer_month, 5, 20, 55
	BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")

	'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
	If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		EMWriteScreen "x", 7, 26
		EMSendKey "<enter>"
		EMWaitReady 0, 0
		If cash_check = 1 then
			EMReadScreen BUSI_ver, 1, 9, 73
		ElseIf HC_check = 1 then
			EMReadScreen BUSI_ver, 1, 12, 73
			If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
		ElseIf SNAP_check = 1 then
			EMReadScreen BUSI_ver, 1, 11, 73
		End if
		EMSendKey "<PF3>"
		EMWaitReady 0, 0
		If SNAP_check = 1 then
			EMReadScreen BUSI_amt, 8, 11, 68
			BUSI_amt = trim(BUSI_amt)
		ElseIf cash_check = 1 then
			EMReadScreen BUSI_amt, 8, 9, 54
			BUSI_amt = trim(BUSI_amt)
		ElseIf HC_check = 1 then
			EMWriteScreen "x", 17, 29
			EMSendKey "<enter>"
			EMWaitReady 0, 0
			EMReadScreen BUSI_amt, 8, 15, 54
			If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
			BUSI_amt = trim(BUSI_amt)
			EMSendKey "<PF3>"
			EMWaitReady 0, 0
		End if
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI"
		EMReadScreen BUSI_income_end_date, 8, 5, 71
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
		If IsDate(BUSI_income_end_date) = True then
			variable_name_for_BUSI = variable_name_for_BUSI & " (ended " & BUSI_income_end_date & ")"
		Else
			If BUSI_amt <> "" then variable_name_for_BUSI = variable_name_for_BUSI & ", ($" & BUSI_amt & "/monthly)"
		End if
		If BUSI_ver = "N" or BUSI_ver = "?" then
			variable_name_for_BUSI = variable_name_for_BUSI & ", no proof provided.; "
		Else
			variable_name_for_BUSI = variable_name_for_BUSI & ".; "
		End if
	Else		'------------This was updated 01/07/2015.
		'Checks the current footer month. If this is the future, it will know later on to read the HC pop-up
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		If datediff("d", date, BUSI_footer_month) > 0 then
			pull_future_HC = TRUE
		Else
			pull_future_HC = FALSE
		End if

		'Converting BUSI type code to a human-readable string
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"

		'Reading and converting BUSI Self employment method into human-readable
		EMReadScreen BUSI_method, 2, 16, 53
		IF BUSI_method = "01" THEN BUSI_method = "50% Gross Income"
		IF BUSI_method = "02" THEN BUSI_method = "Tax Forms"

		'Going to the Gross Income Calculation pop-up
		EMWriteScreen "x", 6, 26
		transmit

		'Getting the verification codes for each type. Only does income, expenses are not included at this time.
		EMReadScreen BUSI_cash_ver, 1, 9, 73
		EMReadScreen BUSI_IVE_ver, 1, 10, 73
		EMReadScreen BUSI_SNAP_ver, 1, 11, 73
		EMReadScreen BUSI_HCA_ver, 1, 12, 73
		EMReadScreen BUSI_HCB_ver, 1, 13, 73

		'Converts each ver type to human readable
		If BUSI_cash_ver = "1" then BUSI_cash_ver = "tax returns provided"
		If BUSI_cash_ver = "2" then BUSI_cash_ver = "receipts provided"
		If BUSI_cash_ver = "3" then BUSI_cash_ver = "client ledger provided"
		If BUSI_cash_ver = "6" then BUSI_cash_ver = "other doc provided"
		If BUSI_cash_ver = "N" then BUSI_cash_ver = "no proof provided"
		If BUSI_cash_ver = "?" then BUSI_cash_ver = "no proof provided"
		If BUSI_IVE_ver = "1" then BUSI_IVE_ver = "tax returns provided"
		If BUSI_IVE_ver = "2" then BUSI_IVE_ver = "receipts provided"
		If BUSI_IVE_ver = "3" then BUSI_IVE_ver = "client ledger provided"
		If BUSI_IVE_ver = "6" then BUSI_IVE_ver = "other doc provided"
		If BUSI_IVE_ver = "N" then BUSI_IVE_ver = "no proof provided"
		If BUSI_IVE_ver = "?" then BUSI_IVE_ver = "no proof provided"
		If BUSI_SNAP_ver = "1" then BUSI_SNAP_ver = "tax returns provided"
		If BUSI_SNAP_ver = "2" then BUSI_SNAP_ver = "receipts provided"
		If BUSI_SNAP_ver = "3" then BUSI_SNAP_ver = "client ledger provided"
		If BUSI_SNAP_ver = "6" then BUSI_SNAP_ver = "other doc provided"
		If BUSI_SNAP_ver = "N" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_SNAP_ver = "?" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_HCA_ver = "1" then BUSI_HCA_ver = "tax returns provided"
		If BUSI_HCA_ver = "2" then BUSI_HCA_ver = "receipts provided"
		If BUSI_HCA_ver = "3" then BUSI_HCA_ver = "client ledger provided"
		If BUSI_HCA_ver = "6" then BUSI_HCA_ver = "other doc provided"
		If BUSI_HCA_ver = "N" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCA_ver = "?" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCB_ver = "1" then BUSI_HCB_ver = "tax returns provided"
		If BUSI_HCB_ver = "2" then BUSI_HCB_ver = "receipts provided"
		If BUSI_HCB_ver = "3" then BUSI_HCB_ver = "client ledger provided"
		If BUSI_HCB_ver = "6" then BUSI_HCB_ver = "other doc provided"
		If BUSI_HCB_ver = "N" then BUSI_HCB_ver = "no proof provided"
		If BUSI_HCB_ver = "?" then BUSI_HCB_ver = "no proof provided"

		'Back to the main screen
		PF3

		'Reading each income amount, trimming them to clean out unneeded spaces.
		EMReadScreen BUSI_cash_retro_amt, 8, 8, 55
		BUSI_cash_retro_amt = trim(BUSI_cash_retro_amt)
		EMReadScreen BUSI_cash_pro_amt, 8, 8, 69
		BUSI_cash_pro_amt = trim(BUSI_cash_pro_amt)
		EMReadScreen BUSI_IVE_amt, 8, 9, 69
		BUSI_IVE_amt = trim(BUSI_IVE_amt)
		EMReadScreen BUSI_SNAP_retro_amt, 8, 10, 55
		BUSI_SNAP_retro_amt = trim(BUSI_SNAP_retro_amt)
		EMReadScreen BUSI_SNAP_pro_amt, 8, 10, 69
		BUSI_SNAP_pro_amt = trim(BUSI_SNAP_pro_amt)

		'Pulls prospective amounts for HC, either from prosp side or from HC inc est.
		If pull_future_HC = False then
			EMReadScreen BUSI_HCA_amt, 8, 11, 69
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 12, 69
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
		Else
			EMWriteScreen "x", 17, 27
			transmit
			EMReadScreen BUSI_HCA_amt, 8, 15, 54
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 16, 54
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
			PF3
		End if

		'Reads end date logic (in case it ended), converts to an actual date
		EMReadScreen BUSI_income_end_date, 8, 5, 72
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")

		'Entering the variable details based on above
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI:; "
		If IsDate(BUSI_income_end_date) = True then	variable_name_for_BUSI = variable_name_for_BUSI & "- Income ended " & BUSI_income_end_date & ".; "
		If BUSI_cash_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH retro: $" & BUSI_cash_retro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_cash_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH pro: $" & BUSI_cash_pro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_IVE_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- IV-E: $" & BUSI_IVE_amt & " budgeted, " & BUSI_IVE_ver & "; "
		If BUSI_SNAP_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP retro: $" & BUSI_SNAP_retro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_SNAP_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP pro: $" & BUSI_SNAP_pro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_HCA_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method A: $" & BUSI_HCA_amt & " budgeted, " & BUSI_HCA_ver & "; "
		If BUSI_HCB_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method B: $" & BUSI_HCB_amt & " budgeted, " & BUSI_HCB_ver & "; "
		'Checks to see if pre 01/15 or post 02/15 then decides what to put in case note based on what was found/needed on the self employment method.
		If IsDate(BUSI_income_end_date) = false then
			IF BUSI_method <> "__" or BUSI_method = "" THEN
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: " & BUSI_method & "; "
			Else
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: None; "
			END IF
		End if
	End if
end function

function add_CARS_to_variable(CARS_variable)
'--- This function adds STAT/CARS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ CARS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, CARS
  EMReadScreen CARS_year, 4, 8, 31
  EMReadScreen CARS_make, 15, 8, 43
  CARS_make = replace(CARS_make, "_", "")
  EMReadScreen CARS_model, 15, 8, 66
  CARS_model = replace(CARS_model, "_", "")
  CARS_type = CARS_year & " " & CARS_make & " " & CARS_model
  CARS_type = split(CARS_type)
  For each CARS_part in CARS_type
    If len(CARS_part) > 1 then
      first_letter = ucase(left(CARS_part, 1))
      other_letters = LCase(right(CARS_part, len(CARS_part) -1))
      new_CARS_type = new_CARS_type & first_letter & other_letters & " "
    End if
  Next
  EMReadScreen CARS_amt, 8, 9, 45
  CARS_amt = trim(CARS_amt)
  CARS_amt = "$" & CARS_amt
  CARS_variable = CARS_variable & trim(new_CARS_type) & ", (" & CARS_amt & "); "
  new_CARS_type = ""
end function

function add_JOBS_to_variable(variable_name_for_JOBS)
'--- This function adds STAT/JOBS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ JOBS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, JOBS
  EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
  JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
  EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
  JOBS_type = replace(JOBS_type, "_", ""	)
  JOBS_type = trim(JOBS_type)
  JOBS_type = split(JOBS_type)
  For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
    If JOBS_part <> "" then
      first_letter = ucase(left(JOBS_part, 1))
      other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
      new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
    End if
  Next
  EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
  jobs_hourly_wage = replace(jobs_hourly_wage, "_", "")   'trimming any underscores
' Navigates to the FS PIC
    EMWriteScreen "x", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
		EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
		jobs_SNAP_prospective_amt = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
    EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
    transmit
'Navigats to GRH PIC
	EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
	IF GRH_PIC_check = "GRH" THEN
		EMWriteScreen "x", 19, 71
		transmit
		EMReadScreen GRH_JOBS_amt, 8, 16, 69
		GRH_JOBS_amt = trim(GRH_JOBS_amt)
		EMReadScreen GRH_pay_frequency, 1, 3, 63
		EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
		GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
		PF3
	END IF
'  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    retro_JOBS_amt = trim(retro_JOBS_amt)
'  Reads the information on the prospective side of JOBS
	EMReadScreen prospective_JOBS_amt, 8, 17, 67
	prospective_JOBS_amt = trim(prospective_JOBS_amt)
'  Reads the information about health care off of HC Income Estimator
    EMReadScreen pay_frequency, 1, 18, 35
	EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
	IF HC_income_est_check = "Est" Then 'this is the old position
		EMWriteScreen "x", 19, 54
	ELSE								'this is the new position
		EMWriteScreen "x", 19, 48
	END IF
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit

  EMReadScreen JOBS_ver, 1, 6, 38
  EMReadScreen JOBS_income_end_date, 8, 9, 49
	'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
  If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
  If IsDate(JOBS_income_end_date) = True then
    variable_name_for_JOBS = variable_name_for_JOBS & new_JOBS_type & "(ended " & JOBS_income_end_date & "); "
  Else
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
	If GRH_pay_frequency = "1" then GRH_pay_frequency = "monthly"
    If GRH_pay_frequency = "2" then GRH_pay_frequency = "semimonthly"
    If GRH_pay_frequency = "3" then GRH_pay_frequency = "biweekly"
    If GRH_pay_frequency = "4" then GRH_pay_frequency = "weekly"
    variable_name_for_JOBS = variable_name_for_JOBS & "EI from " & trim(new_JOBS_type) & ", " & JOBS_month  & " amts:; "
    If SNAP_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- SNAP PIC: $" & SNAP_JOBS_amt & "/" & snap_pay_frequency & ", SNAP PIC Prospective: $" & jobs_SNAP_prospective_amt & ", calculated " & date_of_pic_calc & "; "
    If GRH_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- GRH PIC: $" & GRH_JOBS_amt & "/" & GRH_pay_frequency & ", calculated " & GRH_date_of_pic_calc & "; "
	If retro_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- Retrospective: $" & retro_JOBS_amt & " total; "
    IF prospective_JOBS_amt <> "" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Prospective: $" & prospective_JOBS_amt & " total; "
    IF isnumeric(jobs_hourly_wage) THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Hourly Wage: $" & jobs_hourly_wage & "; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF MAXIS_footer_month = current_month_for_hc_est THEN
	IF HC_JOBS_amt <> "________" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- HC Inc Est: $" & HC_JOBS_amt & "/" & pay_frequency & "; "
    END IF
	If JOBS_ver = "N" or JOBS_ver = "?" then variable_name_for_JOBS = variable_name_for_JOBS & "- No proof provided for this panel; "
  End if
end function

function add_OTHR_to_variable(OTHR_variable)
'--- This function adds STAT/OTHR data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ OTHR_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, OTHR
  EMReadScreen OTHR_type, 16, 6, 43
  OTHR_type = trim(OTHR_type)
  EMReadScreen OTHR_amt, 10, 8, 40
  OTHR_amt = trim(OTHR_amt)
  OTHR_amt = "$" & OTHR_amt
  OTHR_variable = OTHR_variable & trim(OTHR_type) & ", (" & OTHR_amt & ").; "
  new_OTHR_type = ""
end function

function add_RBIC_to_variable(variable_name_for_RBIC)
'--- This function adds STAT/RBIC data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ RBIC_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, RBIC
	EMReadScreen RBIC_month, 5, 20, 55
	RBIC_month = replace(RBIC_month, " ", "/")
	EMReadScreen RBIC_type, 14, 5, 48
	RBIC_type = trim(RBIC_type)
	EMReadScreen RBIC01_pro_amt, 8, 10, 62
	RBIC01_pro_amt = trim(RBIC01_pro_amt)
	EMReadScreen RBIC02_pro_amt, 8, 11, 62
	RBIC02_pro_amt = trim(RBIC02_pro_amt)
	EMReadScreen RBIC03_pro_amt, 8, 12, 62
	RBIC03_pro_amt = trim(RBIC03_pro_amt)
	EMReadScreen RBIC01_retro_amt, 8, 10, 47
	IF RBIC01_retro_amt <> "________" THEN RBIC01_retro_amt = trim(RBIC01_retro_amt)
	EMReadScreen RBIC02_retro_amt, 8, 11, 47
	IF RBIC02_retro_amt <> "________" THEN RBIC02_retro_amt = trim(RBIC02_retro_amt)
	EMReadScreen RBIC03_retro_amt, 8, 12, 47
	IF RBIC03_retro_amt <> "________" THEN RBIC03_retro_amt = trim(RBIC03_retro_amt)
	EMReadScreen RBIC_group_01, 17, 10, 25
		RBIC_group_01 = replace(RBIC_group_01, " __", "")
		RBIC_group_01 = replace(RBIC_group_01, " ", ", ")
	EMReadScreen RBIC_group_02, 17, 11, 25
		RBIC_group_02 = replace(RBIC_group_02, " __", "")
		RBIC_group_02 = replace(RBIC_group_02, " ", ", ")
	EMReadScreen RBIC_group_03, 17, 12, 25
		RBIC_group_03 = replace(RBIC_group_03, " __", "")
		RBIC_group_03 = replace(RBIC_group_03, " ", ", ")

	EMReadScreen RBIC_01_verif, 1, 10, 76
	IF RBIC_01_verif = "N" THEN
		RBIC01_pro_amt = RBIC01_pro_amt & ", not verified"
		RBIC01_retro_amt = RBIC01_retro_amt & ", not verified"
	END IF

	EMReadScreen RBIC_02_verif, 1, 11, 76
	IF RBIC_02_verif = "N" THEN
		RBIC02_pro_amt = RBIC02_pro_amt & ", not verified"
		RBIC02_retro_amt = RBIC02_retro_amt & ", not verified"
	END IF

	EMReadScreen RBIC_03_verif, 1, 12, 76
	IF RBIC_03_verif = "N" THEN
		RBIC03_pro_amt = RBIC03_pro_amt & ", not verified"
		RBIC03_retro_amt = RBIC03_retro_amt & ", not verified"
	END IF

	RBIC_expense_row = 15
	DO
		EMReadScreen RBIC_expense_type, 13, RBIC_expense_row, 28
		RBIC_expense_type = trim(RBIC_expense_type)
		EMReadScreen RBIC_expense_amt, 8, RBIC_expense_row, 62
		RBIC_expense_amt = trim(RBIC_expense_amt)
		EMReadScreen RBIC_expense_verif, 1, RBIC_expense_row, 76
		IF RBIC_expense_type <> "" THEN
			total_RBIC_expenses = total_RBIC_expenses & "- " & RBIC_expense_type & ", $" & RBIC_expense_amt
			IF RBIC_expense_verif <> "N" THEN
				total_RBIC_expenses = total_RBIC_expenses & "; "
			ELSE
				total_RBIC_expenses = total_RBIC_expenses & ", not verified; "
			END IF
			RBIC_expense_row = RBIC_expense_row + 1
			IF RBIC_expense_row = 19 THEN
				PF20
				EMReadScreen RBIC_last_page, 21, 24, 2
				RBIC_expense_row = 15
			END IF
		END IF
	LOOP UNTIL RBIC_expense_type = "" OR RBIC_last_page = "THIS IS THE LAST PAGE"
	EMReadScreen RBIC_ver, 1, 10, 76
	If RBIC_ver = "N" then RBIC_ver = ", no proof provided"
	EMReadScreen RBIC_end_date, 8, 6, 68
	RBIC_end_date = replace(RBIC_end_date, " ", "/")
	If isdate(RBIC_end_date) = True then
		variable_name_for_RBIC = variable_name_for_RBIC & trim(RBIC_type) & " RBIC, ended " & RBIC_end_date & RBIC_ver & "; "
	Else
		IF left(RBIC01_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Prospective, ($" & RBIC01_pro_amt & "); "
		IF left(RBIC01_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Retrospective, ($" & RBIC01_retro_amt & "); "
		IF left(RBIC02_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Prospective, ($" & RBIC02_pro_amt & "); "
		IF left(RBIC02_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Retrospective, ($" & RBIC02_retro_amt & "); "
		IF left(RBIC03_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Prospective, ($" & RBIC03_pro_amt & "); "
		IF left(RBIC03_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Retrospective, ($" & RBIC03_retro_amt & "); "
		IF total_RBIC_expenses <> "" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC Expenses:; " & total_RBIC_expenses
	End if
end function

function add_REST_to_variable(REST_variable)
'--- This function adds STAT/REST data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ REST_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, REST
  EMReadScreen REST_type, 16, 6, 41
  REST_type = trim(REST_type)
  EMReadScreen REST_amt, 10, 8, 41
  REST_amt = trim(REST_amt)
  REST_amt = "$" & REST_amt
  REST_variable = REST_variable & trim(REST_type) & ", (" & REST_amt & ").; "
  new_REST_type = ""
end function

function add_SECU_to_variable(SECU_variable)
'--- This function adds STAT/SECU data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ SECU_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, SECU
  EMReadScreen SECU_amt, 8, 10, 52
  SECU_amt = trim(SECU_amt)
  SECU_amt = "$" & SECU_amt
  EMReadScreen SECU_type, 2, 6, 50
  EMReadScreen SECU_location, 20, 8, 50
  SECU_location = replace(SECU_location, "_", "")
  SECU_location = split(SECU_location)
  For each SECU_part in SECU_location
    If SECU_part <> "" then
      first_letter = ucase(left(SECU_part, 1))
      other_letters = LCase(right(SECU_part, len(SECU_part) -1))
      If len(a) > 3 then
        new_SECU_location = new_SECU_location & b & c & " "
      Else
        new_SECU_location = new_SECU_location & a & " "
      End if
    End if
  Next
  EMReadScreen SECU_ver, 1, 11, 50
  If SECU_ver = "1" then SECU_ver = "agency form provided"
  If SECU_ver = "2" then SECU_ver = "source doc provided"
  If SECU_ver = "3" then SECU_ver = "verified via phone"
  If SECU_ver = "5" then SECU_ver = "other doc verified"
  If SECU_ver = "N" then SECU_ver = "no proof provided"
  SECU_variable = SECU_variable & SECU_type & " at " & new_SECU_location & " (" & SECU_amt & "), " & SECU_ver & ".; "
  new_SECU_location = ""
end function

function add_UNEA_to_variable(variable_name_for_UNEA)
'--- This function adds STAT/UNEA data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ UNEA_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, UNEA
  EMReadScreen UNEA_month, 5, 20, 55
  UNEA_month = replace(UNEA_month, " ", "/")
  EMReadScreen UNEA_type, 16, 5, 40
  If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
  If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
  If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
  UNEA_type = trim(UNEA_type)
  EMReadScreen UNEA_ver, 1, 5, 65
  EMReadScreen UNEA_income_end_date, 8, 7, 68
  If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
  If IsDate(UNEA_income_end_date) = True then
    variable_name_for_UNEA = variable_name_for_UNEA & UNEA_type & " (ended " & UNEA_income_end_date & "); "
  Else
    EMReadScreen UNEA_amt, 8, 18, 68
    UNEA_amt = trim(UNEA_amt)
      EMWriteScreen "x", 10, 26
      transmit
      EMReadScreen SNAP_UNEA_amt, 8, 17, 56
      SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
      EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
      transmit
      EMReadScreen retro_UNEA_amt, 8, 18, 39
      retro_UNEA_amt = trim(retro_UNEA_amt)
	EMReadScreen prosp_UNEA_amt, 8, 18, 68
	prosp_UNEA_amt = trim(prosp_UNEA_amt)
      EMWriteScreen "x", 6, 56
      transmit
      EMReadScreen HC_UNEA_amt, 8, 9, 65
      HC_UNEA_amt = trim(HC_UNEA_amt)
      EMReadScreen pay_frequency, 1, 10, 63
      transmit
      If HC_UNEA_amt = "________" then
        EMReadScreen HC_UNEA_amt, 8, 18, 68
        HC_UNEA_amt = trim(HC_UNEA_amt)
        pay_frequency = "mo budgeted prospectively"
    End If
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
    variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amts:; "
    If SNAP_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
    If retro_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
    If prosp_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF MAXIS_footer_month = current_month_for_hc_est THEN
    	If HC_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- HC Inc Est: $" & HC_UNEA_amt & "/" & pay_frequency & "; "
    END IF
    If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "
  End if
end function

function assess_button_pressed()
'--- This fuction will review the button pressed on a dialog with tabs to go to the correct dialog if the do loop structure is used.
'===== Keywords: DIALOGS, NAVIGATES
    If ButtonPressed = dlg_one_button Then
        pass_one = false
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = true
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_two_button Then
        pass_one = true
        pass_two = False
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = true
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_three_button Then
        pass_one = true
        pass_two = true
        pass_three = false
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = true
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_four_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = false
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = true
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_five_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = false
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = true
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_six_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = false
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = true
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_seven_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = false
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = true
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_eight_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = false
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = true
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_nine_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = false
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = true
        show_ten = true

        tab_button = True
    End If
    If ButtonPressed = dlg_ten_button Then
        pass_one = true
        pass_two = true
        pass_three = true
        pass_four = true
        pass_five = true
        pass_six = true
        pass_seven = true
        pass_eight = true
        pass_nine = true
        pass_ten = false

        show_one = false
        show_two = false
        show_three = false
        show_four = false
        show_five = false
        show_six = false
        show_seven = false
        show_eight = false
        show_nine = false
        show_ten = true

        tab_button = True
    End If
end function


function assign_county_address_variables(address_line_01, address_line_02)
'--- This function will assign an address to a variable selected from the interview_location variable in the Appt Letter script.
'~~~~~ address_line_01: 1st line of address (street address) from new_office_array
'~~~~~ address_line_02: 2nd line of address (city/state/zip) from new_office_array
'===== Keywords: MAXIS, APPT LETTER, ADDRESS
	For each office in county_office_array				'Splits the county_office_array, which is set by the config program and declared earlier in this file
		If instr(office, interview_location) <> 0 then		'If the name of the office is found in the "interview_location" variable, which is contained in the MEMO - appt letter script.
			new_office_array = split(office, "|")		'Split the office into its own array
			address_line_01 = new_office_array(0)		'Line 1 of the address is the first part of this array
			address_line_02 = new_office_array(1)		'Line 2 of the address is the second part of this array
		End if
	Next
end function

function attn()
 '--- This function sends or hits the ESC (escape) key.
  '===== Keywords: MAXIS, MMIS, PRISM, ESC
  EMSendKey "<attn>"
  EMWaitReady -1, 0
end function

function autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
 '--- This function autofills information for all HH members idenified from the HH_member_array from a selected MAXIS panel into an edit box in a dialog.
 '~~~~~ HH_member_array: array of HH members from function HH_member_custom_dialog(HH_member_array). User selects which HH members are added to array.
 '~~~~~ read_panel_from: first four characters because we use separate handling for HCRE-retro. This is something that should be fixed someday!!!!!!!!!
 '~~~~~ variable_written_to: the variable used by the editbox you wish to autofill.
 '===== Keywords: MAXIS, autofill, HH_member_array
  call navigate_to_MAXIS_screen("stat", left(panel_read_from, 4))

  'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
  EMReadScreen panel_total_check, 6, 2, 73
  IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

  If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then
      Do
        'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
        EMReadScreen support_coop_code, 1, 4, 73
        If support_coop_code = "N" then
          EMReadScreen caregiver_ref_nbr, 2, 4, 47
          If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated.
        End if
        'Then it gets info on the ABPS themself.
        EMReadScreen ABPS_current, 45, 10, 30
        If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
        ABPS_current = replace(ABPS_current, "  First:", ",")
        ABPS_current = replace(ABPS_current, "_", "")
        ABPS_current = split(ABPS_current)
        For each ABPS_part in ABPS_current
          first_letter = ucase(left(ABPS_part, 1))
          other_letters = LCase(right(ABPS_part, len(ABPS_part) -1))
          If len(ABPS_part) > 1 then
            new_ABPS_current = new_ABPS_current & first_letter & other_letters & " "
          Else
            new_ABPS_current = new_ABPS_current & ABPS_part & " "
          End if
        Next
        ABPS_row = 15 'Setting variable for do...loop
        Do
          Do 'Using a do...loop to determine which MEMB numbers are with this parent
            EMReadScreen child_ref_nbr, 2, ABPS_row, 35
            If child_ref_nbr <> "__" then
              amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
              children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
            End if
            ABPS_row = ABPS_row + 1
          Loop until ABPS_row > 17		'End of the row
          EMReadScreen more_check, 7, 19, 66
          If more_check = "More: +" then
            EMSendKey "<PF20>"
            EMWaitReady 0, 0
            ABPS_row = 15
          End if
        Loop until more_check <> "More: +"
        'Cleaning up the "children_for_ABPS" variable to be more readable
        children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
        children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around
        if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
        if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
        variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
        'Resetting variables for the do...loop in case this function runs again
        new_ABPS_current = ""
        amt_of_children_for_ABPS = 0
        children_for_ABPS = ""
        'Checking to see if it needs to run again, if it does it transmits or else the loop stops
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
      'Combining the two variables (support coop and the variable written to)
      variable_written_to = support_coop & variable_written_to
    End if
  Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCI_total, 1, 2, 78
      If ACCI_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCI_to_variable(variable_written_to)
          EMReadScreen ACCI_panel_current, 1, 2, 73
          If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
        Loop until cint(ACCI_panel_current) = cint(ACCI_total)
      End if
    Next
  Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCT_total, 2, 2, 78
	  ACCT_total = trim(ACCT_total)   'deleting space if one digit.
      If ACCT_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCT_to_variable(variable_written_to)
          EMReadScreen ACCT_panel_current, 2, 2, 72
          ACCT_panel_current = trim(ACCT_panel_current)
          If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
        Loop until cint(ACCT_panel_current) = cint(ACCT_total)
      End if
    Next
  ElseIf panel_read_from = "ACUT" Then '----------------------------------------------------------------------------------------------------ACUT
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        transmit
        EMReadScreen ACUT_total, 1, 2, 78
        If ACUT_total <> 0 then
            EMReadScreen share_yn, 1, 6, 42
            EMReadScreen retro_heat_verif, 1, 10, 35
            EMReadScreen retro_heat_amount, 8, 10, 41
            EMReadScreen retro_air_verif, 1, 11, 35
            EMReadScreen retro_air_amount, 8, 11, 41
            EMReadScreen retro_elec_verif, 1, 12, 35
            EMReadScreen retro_elec_amount, 8, 12, 41
            EMReadScreen retro_fuel_verif, 1, 13, 35
            EMReadScreen retro_fuel_amount, 8, 13, 41
            EMReadScreen retro_garbage_verif, 1, 14, 35
            EMReadScreen retro_garbage_amount, 8, 14, 41
            EMReadScreen retro_water_verif, 1, 15, 35
            EMReadScreen retro_water_amount, 8, 15, 41
            EMReadScreen retro_sewer_verif, 1, 16, 35
            EMReadScreen retro_sewer_amount, 8, 16, 41
            EMReadScreen retro_other_verif, 1, 17, 35
            EMReadScreen retro_other_amount, 8, 17, 41

            EMReadScreen prosp_heat_verif, 1, 10, 55
            EMReadScreen prosp_heat_amount, 8, 10, 61
            EMReadScreen prosp_air_verif, 1, 11, 55
            EMReadScreen prosp_air_amount, 8, 11, 61
            EMReadScreen prosp_elec_verif, 1, 12, 55
            EMReadScreen prosp_elec_amount, 8, 12, 61
            EMReadScreen prosp_fuel_verif, 1, 13, 55
            EMReadScreen prosp_fuel_amount, 8, 13, 61
            EMReadScreen prosp_garbage_verif, 1, 14, 55
            EMReadScreen prosp_garbage_amount, 8, 14, 61
            EMReadScreen prosp_water_verif, 1, 15, 55
            EMReadScreen prosp_water_amount, 8, 15, 61
            EMReadScreen prosp_sewer_verif, 1, 16, 55
            EMReadScreen prosp_sewer_amount, 8, 16, 61
            EMReadScreen prosp_other_verif, 1, 17, 55
            EMReadScreen prosp_other_amount, 8, 17, 61

            EMReadScreen dwp_phone_yn, 1, 18, 55

            variable_written_to = "Actutal Utilitiy Expense for M" & HH_member
            If share_yn = "Y" Then variable_written_to = variable_written_to & " - this expense is shared."
            If retro_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (retro) $" & trim(retro_heat_amount) & " - Verif: " & retro_heat_verif & "."
            If prosp_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (prosp) $" & trim(prosp_heat_amount) & " - Verif: " & prosp_heat_verif & "."
            If retro_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (retro) $" & trim(retro_air_amount) & " - Verif: " & retro_air_verif & "."
            If prosp_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (prosp) $" & trim(prosp_air_amount) & " - Verif: " & prosp_air_verif & "."
            If retro_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (retro) $" & trim(retro_elec_amount) & " - Verif: " & retro_elec_verif & "."
            If prosp_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (prosp) $" & trim(prosp_elec_amount) & " - Verif: " & prosp_elec_verif & "."
            If retro_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (retro) $" & trim(retro_fuel_amount) & " - Verif: " & retro_fuel_verif & "."
            If prosp_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (prosp) $" & trim(prosp_fuel_amount) & " - Verif: " & prosp_fuel_verif & "."
            If retro_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (retro) $" & trim(retro_garbage_amount) & " - Verif: " & retro_garbage_verif & "."
            If prosp_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (prosp) $" & trim(prosp_garbage_amount) & " - Verif: " & prosp_garbage_verif & "."
            If retro_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (retro) $" & trim(retro_water_amount) & " - Verif: " & retro_water_verif & "."
            If prosp_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (prosp) $" & trim(prosp_water_amount) & " - Verif: " & prosp_water_verif & "."
            If retro_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (retro) $" & trim(retro_sewer_amount) & " - Verif: " & retro_sewer_verif & "."
            If prosp_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (prosp) $" & trim(prosp_sewer_amount) & " - Verif: " & prosp_sewer_verif & "."
            If retro_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (retro) $" & trim(retro_other_amount) & " - Verif: " & retro_other_verif & "."
            If prosp_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (prosp) $" & trim(prosp_other_amount) & " - Verif: " & prosp_other_verif & "."
            If dwp_phone_yn = "Y" Then variable_written_to = variable_written_to & " Standard DWP Phone allowance of $35."

        End If
    Next
  Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
    EMReadScreen addr_line_01, 22, 6, 43
    EMReadScreen addr_line_02, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 12, 9, 43
    variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
    variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
  Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
    EMReadScreen AREP_name, 37, 4, 32
    AREP_name = replace(AREP_name, "_", "")
    AREP_name = split(AREP_name)
    For each word in AREP_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
    EMReadScreen BILS_amt, 1, 2, 78
    If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
  Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen BUSI_total, 1, 2, 78
      If BUSI_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_BUSI_to_variable(variable_written_to)
          EMReadScreen BUSI_panel_current, 1, 2, 73
          If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
        Loop until cint(BUSI_panel_current) = cint(BUSI_total)
      End if
    Next
  Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen CARS_total, 2, 2, 78
	  CARS_total = trim(CARS_total)
      If CARS_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_CARS_to_variable(variable_written_to)
          EMReadScreen CARS_panel_current, 2, 2, 72
		  CARS_panel_current = trim(CARS_panel_current)
          If cint(CARS_panel_current) < cint(CARS_total) then transmit
        Loop until cint(CARS_panel_current) = cint(CARS_total)
      End if
    Next
  Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen cash_amt, 8, 8, 39
      cash_amt = trim(cash_amt)
      If cash_amt <> "________" then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
      End if
    Next
  Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen support_amt, 8, 10, 63
      support_amt = trim(support_amt)
      If support_amt <> "________" then
        EMReadScreen support_ver, 1, 10, 36
        If support_ver = "?" or support_ver = "N" then
          support_ver = ", no proof provided"
        Else
          support_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
      End if
      EMReadScreen alimony_amt, 8, 11, 63
      alimony_amt = trim(alimony_amt)
      If alimony_amt <> "________" then
        EMReadScreen alimony_ver, 1, 11, 36
        If alimony_ver = "?" or alimony_ver = "N" then
          alimony_ver = ", no proof provided"
        Else
          alimony_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
      End if
      EMReadScreen tax_dep_amt, 8, 12, 63
      tax_dep_amt = trim(tax_dep_amt)
      If tax_dep_amt <> "________" then
        EMReadScreen tax_dep_ver, 1, 12, 36
        If tax_dep_ver = "?" or tax_dep_ver = "N" then
          tax_dep_ver = ", no proof provided"
        Else
          tax_dep_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
      End if
      EMReadScreen other_COEX_amt, 8, 13, 63
      other_COEX_amt = trim(other_COEX_amt)
      If other_COEX_amt <> "________" then
        EMReadScreen other_COEX_ver, 1, 13, 36
        If other_COEX_ver = "?" or other_COEX_ver = "N" then
          other_COEX_ver = ", no proof provided"
        Else
          other_COEX_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
      End if
    Next
  Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadScreen DCEX_total, 1, 2, 78
      If DCEX_total <> 0 then
		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		Do
			DCEX_row = 11
			Do
				EMReadScreen expense_amt, 8, DCEX_row, 63
				expense_amt = trim(expense_amt)
				If expense_amt <> "________" then
					EMReadScreen child_ref_nbr, 2, DCEX_row, 29
					EMReadScreen expense_ver, 1, DCEX_row, 41
					If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
						expense_ver = ", no proof provided"
					Else
						expense_ver = ""
					End if
					variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
				End if
				DCEX_row = DCEX_row + 1
			Loop until DCEX_row = 17
			EMReadScreen DCEX_panel_current, 1, 2, 73
			If cint(DCEX_panel_current) < cint(DCEX_total) then transmit
		Loop until cint(DCEX_panel_current) = cint(DCEX_total)
	  End if
    Next
  Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DIET_row = 8 'Setting this variable for the next do...loop
      EMReadScreen DIET_total, 1, 2, 78
      If DIET_total <> 0 then
        DIET = DIET & "Member " & HH_member & "- "
        Do
          EMReadScreen diet_type, 2, DIET_row, 40
          EMReadScreen diet_proof, 1, DIET_row, 51
          If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then
            diet_proof = ", no proof provided"
          Else
            diet_proof = ""
          End if
          If diet_type = "01" then diet_type = "High Protein"
          If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
          If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
          If diet_type = "04" then diet_type = "Lo Cholesterol"
          If diet_type = "05" then diet_type = "High Residue"
          If diet_type = "06" then diet_type = "Preg/Lactation"
          If diet_type = "07" then diet_type = "Gluten Free"
          If diet_type = "08" then diet_type = "Lactose Free"
          If diet_type = "09" then diet_type = "Anti-Dumping"
          If diet_type = "10" then diet_type = "Hypoglycemic"
          If diet_type = "11" then diet_type = "Ketogenic"
          If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
          DIET_row = DIET_row + 1
        Loop until DIET_row = 19
      End if
    Next
  Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
    For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadscreen DISA_total, 1, 2, 78
	  IF DISA_total <> 0 THEN
		'Reads and formats CASH/GRH disa status
		EMReadScreen CASH_DISA_status, 2, 11, 59
		EMReadScreen CASH_DISA_verif, 1, 11, 69
		IF CASH_DISA_status = "01" or CASH_DISA_status = "02" or CASH_DISA_status = "03" OR CASH_DISA_status = "04" THEN CASH_DISA_status = "RSDI/SSI certified"
		IF CASH_DISA_status = "06" THEN CASH_DISA_status = "SMRT/SSA pends"
		IF CASH_DISA_status = "08" THEN CASH_DISA_status = "Certified Blind"
		IF CASH_DISA_status = "09" THEN CASH_DISA_status = "Ill/Incap"
		IF CASH_DISA_status = "10" THEN CASH_DISA_status = "Certified disabled"
		IF CASH_DISA_verif = "?" OR CASH_DISA_verif = "N" THEN
			CASH_DISA_verif = ", no proof provided"
		ELSE
			CASH_DISA_verif = ""
		END IF

		'Reads and formats SNAP disa status
		EmreadScreen SNAP_DISA_status, 2, 12, 59
		EMReadScreen SNAP_DISA_verif, 1, 12, 69
		IF SNAP_DISA_status = "01" or SNAP_DISA_status = "02" or SNAP_DISA_status = "03" OR SNAP_DISA_status = "04" THEN SNAP_DISA_status = "RSDI/SSI certified"
		IF SNAP_DISA_status = "08" THEN SNAP_DISA_status = "Certified Blind"
		IF SNAP_DISA_status = "09" THEN SNAP_DISA_status = "Ill/Incap"
		IF SNAP_DISA_status = "10" THEN SNAP_DISA_status = "Certified disabled"
		IF SNAP_DISA_status = "11" THEN SNAP_DISA_status = "VA determined PD disa"
		IF SNAP_DISA_status = "12" THEN SNAP_DISA_status = "VA (other accept disa)"
		IF SNAP_DISA_status = "13" THEN SNAP_DISA_status = "Cert RR Ret Disa & on MEDI"
		IF SNAP_DISA_status = "14" THEN SNAP_DISA_status = "Other Govt Perm Disa Ret Bnft"
		IF SNAP_DISA_status = "15" THEN SNAP_DISA_status = "Disability from MINE list"
		IF SNAP_DISA_status = "16" THEN SNAP_DISA_status = "Unable to p&p own meal"
		IF SNAP_DISA_verif = "?" OR SNAP_DISA_verif = "N" THEN
			SNAP_DISA_verif = ", no proof provided"
		ELSE
			SNAP_DISA_verif = ""
		END IF

		'Reads and formats HC disa status/verif
		EMReadScreen HC_DISA_status, 2, 13, 59
		EMReadScreen HC_DISA_verif, 1, 13, 69
		If HC_DISA_status = "01" or HC_DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
		If HC_DISA_status = "06" then HC_DISA_status = "SMRT/SSA pends"
		If HC_DISA_status = "08" then HC_DISA_status = "Certified blind"
		If HC_DISA_status = "10" then HC_DISA_status = "Certified disabled"
		If HC_DISA_status = "11" then HC_DISA_status = "Spec cat- disa child"
		If HC_DISA_status = "20" then HC_DISA_status = "TEFRA- disabled"
		If HC_DISA_status = "21" then HC_DISA_status = "TEFRA- blind"
		If HC_DISA_status = "22" then HC_DISA_status = "MA-EPD"
		If HC_DISA_status = "23" then HC_DISA_status = "MA/waiver"
		If HC_DISA_status = "24" then HC_DISA_status = "SSA/SMRT appeal pends"
		If HC_DISA_status = "26" then HC_DISA_status = "SSA/SMRT disa deny"
		IF HC_DISA_verif = "?" OR HC_DISA_verif = "N" THEN
			HC_DISA_verif = ", no proof provided"
		ELSE
			HC_DISA_verif = ""
		END IF
		'cleaning to make variable to write
		IF CASH_DISA_status = "__" THEN
			CASH_DISA_status = ""
		ELSE
			IF CASH_DISA_status = SNAP_DISA_status THEN
				SNAP_DISA_status = "__"
				CASH_DISA_status = "CASH/SNAP: " & CASH_DISA_status & " "
			ELSE
				CASH_DISA_status = "CASH: " & CASH_DISA_status & " "
			END IF
		END IF
		IF SNAP_DISA_status = "__" THEN
			SNAP_DISA_status = ""
		ELSE
			SNAP_DISA_status = "SNAP: " & SNAP_DISA_status & " "
		END IF
		IF HC_DISA_status = "__" THEN
			HC_DISA_status = ""
		ELSE
			HC_DISA_status = "HC: " & HC_DISA_status & " "
		END IF
		'Adding verif code info if N or ?
		IF CASH_DISA_verif <> "" THEN CASH_DISA_status = CASH_DISA_status & CASH_DISA_verif & " "
		IF SNAP_DISA_verif <> "" THEN SNAP_DISA_status = SNAP_DISA_status & SNAP_DISA_verif & " "
		IF HC_DISA_verif <> "" THEN HC_DISA_status = HC_DISA_status & HC_DISA_verif & " "
		'Creating final variable
		IF CASH_DISA_status <> "" THEN FINAL_DISA_status = CASH_DISA_status
		IF SNAP_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & SNAP_DISA_status
		IF HC_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & HC_DISA_status

		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		variable_written_to = variable_written_to & FINAL_DISA_status & "; "
	  END IF
    Next
  Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
    row = 14
    Do
      EMReadScreen reference_numbers_current_row, 40, row, 39
      reference_numbers = reference_numbers + reference_numbers_current_row
      row = row + 1
    Loop until row = 18
    reference_numbers = replace(reference_numbers, "  ", " ")
    reference_numbers = split(reference_numbers)
    For each member in reference_numbers
      If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
    Next
    EATS_info = trim(EATS_info)
    if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
    If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
 Elseif panel_read_from = "EMPS" then '----------------------------------------------------------------------------------------------------EMPS
    	For each HH_member in HH_member_array
  		'blanking out variables for the next HH member
  		EMPS_info = ""
  		ES_exemptions = ""
  		ES_info = ""
  		EMWriteScreen HH_member, 20, 76
  		EMWriteScreen "01", 20, 79
  		transmit
  		EMReadScreen EMPS_total, 1, 2, 78
  		If EMPS_total <> 0 then
  			'orientation info (EMPS_info variable)-------------------------------------------------------------------------
  			EMReadScreen EMPS_orientation_date, 8, 5, 39
  			IF EMPS_orientation_date = "__ __ __" then
  				EMPS_orientation_date = "none"
  			ElseIf EMPS_orientation_date <> "__ __ __" then
  				EMPS_orientation_date = replace(EMPS_orientation_date, " ", "/")
  				EMPS_info = EMPS_info & " Fin orient: " & EMPS_orientation_date & ","
  			END IF
  	  		EMReadScreen EMPS_orientation_attended, 1, 5, 65
  			IF EMPS_orientation_attended <> "_" then EMPS_info = EMPS_info & " Attended orient: " & EMPS_orientation_attended & ","
  			'Good cause (EMPS_info variable)
  			EMReadScreen EMPS_good_cause, 2, 5, 79
  			IF EMPS_good_cause <> "__" then
  				If EMPS_good_cause = "01" then EMPS_good_cause = "01-No Good Cause"
  				If EMPS_good_cause = "02" then EMPS_good_cause = "02-No Child Care"
  				If EMPS_good_cause = "03" then EMPS_good_cause = "03-Ill or Injured"
  				If EMPS_good_cause = "04" then EMPS_good_cause = "04-Care Ill/Incap. Family Member"
  				If EMPS_good_cause = "05" then EMPS_good_cause = "05-Lack of Transportation"
  				If EMPS_good_cause = "06" then EMPS_good_cause = "06-Emergency"
  				If EMPS_good_cause = "07" then EMPS_good_cause = "07-Judicial Proceedings"
  				If EMPS_good_cause = "08" then EMPS_good_cause = "08-Conflicts with Work/School"
  				If EMPS_good_cause = "09" then EMPS_good_cause = "09-Other Impediments"
  				If EMPS_good_cause = "10" then EMPS_good_cause = "10-Special Medical Criteria "
  				If EMPS_good_cause = "20" then EMPS_good_cause = "20-Exempt--Only/1st Caregiver Employed 35+ Hours"
  				If EMPS_good_cause = "21" then EMPS_good_cause = "21-Exempt--2nd Caregiver Employed 20+ Hours"
  				If EMPS_good_cause = "22" then EMPS_good_cause = "22-Exempt--Preg/Parenting Caregiver < Age 20"
  				If EMPS_good_cause = "23" then EMPS_good_cause = "23-Exempt--Special Medical Criteria"
  				IF EMPS_good_cause <> "__" then EMPS_info = EMPS_info & " Good cause: " & EMPS_good_cause & ","
  			END IF

  			'sanction dates (EMPS_info variable)
  			EMReadScreen EMPS_sanc_begin, 8, 6, 39
  			If EMPS_sanc_begin <> "__ 01 __" then
  				EMPS_sanc_begin = replace(EMPS_sanc_begin, "_", "/")
  				sanction_date = sanction_date & EMPS_sanc_begin
  			END IF
  			EMReadScreen EMPS_sanc_end, 8, 6, 65
  			If EMPS_sanc_end <> "__ 01 __" then
  				EMPS_sanc_end = replace(EMPS_sanc_end, "_", "/")
  				sanction_date = sanction_date & "-" & EMPS_sanc_end
  			END IF
  			IF sanction_date <> "" then EMPS_info = EMPS_info & " Sanction dates: " & sanction_date & ","
  			'cleaning up ES_info variable
  			If right(EMPS_info, 1) = "," then EMPS_info = left(EMPS_info, len(EMPS_info) - 1)
  			IF trim(EMPS_info) <> "" then EMPS_info = EMPS_info & "."

  			'other sanction dates (ES_exemptions variable)--------------------------------------------------------------------------------
  			'special medical criteria
			EMReadScreen EMPS_memb_at_home, 1, 8, 76
  			IF EMPS_memb_at_home <> "N" then
				If EMPS_memb_at_home = "1" then EMPS_memb_at_home = "Home-Health/Waiver service"
				IF EMPS_memb_at_home = "2" then EMPS_memb_at_home = "Child w/ severe emotional dist"
				IF EMPS_memb_at_home = "3" then EMPS_memb_at_home = "Adult/Serious Persistent MI"
				ES_exemptions = ES_exemptions & " Special med criteria: " & EMPS_memb_at_home & ","
  			END IF

			EMReadScreen EMPS_care_family, 1, 9, 76
  			IF EMPS_care_family = "Y" then ES_exemptions = ES_exemptions & " Care of ill/incap memb: " & EMPS_care_family & ","
  			EMReadScreen EMPS_crisis, 1, 10, 76
  			IF EMPS_crisis = "Y" then ES_exemptions = ES_exemptions & " Family crisis: " & EMPS_crisis & ","

			'hard to employ
			EMReadScreen EMPS_hard_employ, 2, 11, 76
  			IF EMPS_hard_employ <> "NO" then
				IF EMPS_hard_employ = "IQ" then EMPS_hard_employ = "IQ tested at < 80"
				IF EMPS_hard_employ = "LD" then EMPS_hard_employ = "Learning Disabled"
				IF EMPS_hard_employ = "MI" then EMPS_hard_employ = "Mentally ill"
				IF EMPS_hard_employ = "DD" then EMPS_hard_employ = "Dev Disabled"
				IF EMPS_hard_employ = "UN" then EMPS_hard_employ = "Unemployable"
				ES_exemptions = ES_exemptions & " Hard to employ: " & EMPS_hard_employ & ","
  			END IF

  			'EMPS under 1 coding and dates used(ES_exemptions variable)
  			EMReadScreen EMPS_under1, 1, 12, 76
  			IF EMPS_under1 = "Y" then
  				ES_exemptions = ES_exemptions & " FT child under 1: " & EMPS_under1 & ","
  				EMWriteScreen "x", 12, 39
  				transmit
  				MAXIS_row = 7
  				MAXIS_col = 22
  				DO
  					EMReadScreen exemption_date, 9, MAXIS_row, MAXIS_col
  					If trim(exemption_date) = "" then exit do
  					If exemption_date <> "__ / ____" then
  						child_under1_dates = child_under1_dates & exemption_date & ", "
  						MAXIS_col = MAXIS_col + 11
  						If MAXIS_col = 66 then
  							MAXIS_row = MAXIS_row + 1
  							MAXIS_col = 22
  						END IF
  					END IF
  				LOOP until exemption_date = "__ / ____" or (MAXIS_row = 9 and MAXIS_col = 66)
  				PF3
  				'cleaning up excess comma at the end of child_under1_dates variable
  				If right(child_under1_dates,  2) = ", " then child_under1_dates = left(child_under1_dates, len(child_under1_dates) - 2)
  				If trim(child_under1_dates) = "" then child_under1_dates = " N/A"
  				ES_exemptions = ES_exemptions & " Child under 1 exeption dates: " & child_under1_dates & ","
  			END IF

  			'cleaning up ES_exemptions variable
  			If right(ES_exemptions, 1) = "," then ES_exemptions = left(ES_exemptions, len(ES_exemptions) - 1)
  			IF trim(ES_exemptions) <> "" then ES_exemptions = ES_exemptions & "."

  			'Reading ES Information (for ES_info variable)
  			EMReadScreen ES_status, 40, 15, 40
  			ES_status = trim(ES_status)
  			IF ES_status <> "" then ES_info = ES_info & " ES status: " & ES_status & ","
  			EMReadScreen ES_referral_date, 8, 16, 40
  			If ES_referral_date <> "__ __ __" then
  				ES_referral_date = replace(ES_referral_date, " ", "/")
  				ES_info = ES_info & " ES referral date: " & ES_referral_date & ","
  			END IF

  			EMReadScreen DWP_plan_date, 8, 17, 40
  			IF DWP_plan_date <> "__ __ __" then
  				DWP_plan_date = replace(DWP_plan_date, "_", "/")
  				ES_info = ES_info & " DWP plan date: " & DWP_plan_date & ","
  			END IF

			EMReadScreen minor_ES_option, 2, 16, 76
			If minor_ES_option <> "__" then
				IF minor_ES_option = "SC" then minor_ES_option = "Secondary Education"
				IF minor_ES_option = "EM" then minor_ES_option = "Employment"
				ES_info = ES_info & " 18/19 yr old ES option: " & minor_ES_option & ","
			END if

			'cleaning up ES_info variable
  			If right(ES_info, 1) = "," then ES_info = left(ES_info, len(ES_info) - 1)

  			variable_written_to = variable_written_to & "Member " & HH_member & "- "
  			variable_written_to = variable_written_to & EMPS_info & ES_exemptions & ES_info & "; "
  		END IF
  	next
  Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen FACI_total, 1, 2, 78
      If FACI_total <> 0 then
        row = 14
        Do
          EMReadScreen date_in_check, 4, row, 53
		  EMReadScreen date_in_month_day, 5, row, 47
          EMReadScreen date_out_check, 4, row, 77
		  date_in_month_day = replace(date_in_month_day, " ", "/") & "/"
          If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
          If row > 18 then
            EMReadScreen FACI_page, 1, 2, 73
            If FACI_page = FACI_total then
              FACI_status = "Not in facility"
            Else
              transmit
              row = 14
            End if
          End if
        Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
        EMReadScreen client_FACI, 30, 6, 43
        client_FACI = replace(client_FACI, "_", "")
        FACI_array = split(client_FACI)
        For each a in FACI_array
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            new_FACI = new_FACI & b & c & " "
          End if
        Next
        client_FACI = new_FACI
        If FACI_status = "Not in facility" then
          client_FACI = ""
        Else
          variable_written_to = variable_written_to & "Member " & HH_member & "- "
          variable_written_to = variable_written_to & client_FACI & " Date in: " & date_in_month_day & date_in_check & "; "
        End if
      End if
    Next
  Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
	For each HH_member in HH_member_array
	  EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	  If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      fmed_row = 9 'Setting this variable for the next do...loop
      EMReadScreen fmed_total, 1, 2, 78
      If fmed_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
		  use_expense = False					'<--- Used to determine if an FMED expense that has an end date is going to be counted.
          EMReadScreen fmed_type, 2, fmed_row, 25
          EMReadScreen fmed_proof, 2, fmed_row, 32
          EMReadScreen fmed_amt, 8, fmed_row, 70
		  EMReadScreen fmed_end_date, 5, fmed_row, 60		'reading end date to see if this one even gets added.
		  IF fmed_end_date <> "__ __" THEN
			fmed_end_date = replace(fmed_end_date, " ", "/01/")
			fmed_end_date = dateadd("M", 1, fmed_end_date)
			fmed_end_date = dateadd("D", -1, fmed_end_date)
			IF datediff("D", date, fmed_end_date) > 0 THEN use_expense = True		'<--- If the end date of the FMED expense is the current month or a future month, the expense is going to be counted.
		  END IF
		  If fmed_end_date = "__ __" OR use_expense = TRUE then					'Skips entries with an end date or end dates in the past.
            If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then
              fmed_proof = ", no proof provided"
            Else
              fmed_proof = ""
            End if
            If fmed_amt = "________" then
              fmed_amt = ""
            Else
              fmed_amt = " ($" & trim(fmed_amt) & ")"
            End if
            If fmed_type = "01" then fmed_type = "Nursing Home"
            If fmed_type = "02" then fmed_type = "Hosp/Clinic"
            If fmed_type = "03" then fmed_type = "Physicians"
            If fmed_type = "04" then fmed_type = "Prescriptions"
            If fmed_type = "05" then fmed_type = "Ins Premiums"
            If fmed_type = "06" then fmed_type = "Dental"
            If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
            If fmed_type = "08" then fmed_type = "Vision Care"
            If fmed_type = "09" then fmed_type = "Medicare Prem"
            If fmed_type = "10" then fmed_type = "Mo. Spdwn Amt/Waiver Obl"
            If fmed_type = "11" then fmed_type = "Home Care"
            If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
            If fmed_type = "15" then fmed_type = "Medi Part D premium"
            If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
			IF fmed_end_date <> "__ __" THEN					'<--- If there is a counted FMED expense with a future end date, the script will modify the way that end date is displayed.
				fmed_end_date = datepart("M", fmed_end_date) & "/" & right(datepart("YYYY", fmed_end_date), 2)		'<--- Begins pulling apart fmed_end_date to format it to human speak.
				IF left(fmed_end_date, 1) <> "0" THEN fmed_end_date = "0" & fmed_end_date
				variable_written_to = left(variable_written_to, len(variable_written_to) - 2) & ", counted through " & fmed_end_date & "; "			'<--- Putting variable_written_to back together with FMED expense end date information.
			END IF
          End if
          fmed_row = fmed_row + 1
          If fmed_row = 15 then
            PF20
            fmed_row = 9
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
          End if
        Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
      End if
    Next
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
    EMReadScreen variable_written_to, 8, 10, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
    If isdate(variable_written_to) = False then variable_written_to = ""
  Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
    EMReadScreen variable_written_to, 5, 10, 64
    If isdate(variable_written_to) = True then
      variable_written_to = replace(variable_written_to, " ", "/01/")
      If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
        variable_written_to = variable_written_to
      Else
        variable_written_to = ""
      End if
    End if
  Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
    EMReadScreen HEST_total, 1, 2, 78
    If HEST_total <> 0 then
      EMReadScreen heat_air_check, 6, 13, 75
      If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
      EMReadScreen electric_check, 6, 14, 75
      If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
      EMReadScreen phone_check, 6, 15, 75
      If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
    End if
  Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen IMIG_total, 1, 2, 78
      If IMIG_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen IMIG_type, 30, 6, 48
        variable_written_to = variable_written_to & trim(IMIG_type) & "; "
      End if
    Next
  Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
    EMReadScreen INSA_amt, 1, 2, 78
    If INSA_amt <> 0 then
      'Runs once per INSA screen
		For i = 1 to INSA_amt step 1
			insurance_name = ""
			'Goes to the correct screen
			EMWriteScreen "0" & i, 20, 79
			transmit
			'Gather Insurance Name
			EMReadScreen INSA_name, 38, 10, 38
			INSA_name = replace(INSA_name, "_", "")
			INSA_name = split(INSA_name)
			For each word in INSA_name
				If trim(word) <> "" then
						first_letter_of_word = ucase(left(word, 1))
						rest_of_word = LCase(right(word, len(word) -1))
						If len(word) > 4 then
							insurance_name = insurance_name & first_letter_of_word & rest_of_word & " "
						Else
							insurance_name = insurance_name & word & " "
						End if
				End if
			Next
			'Create a list of members covered by this insurance
			INSA_row = 15 : INSA_col = 30
			insured_count = 0
			member_list = ""
			Do
				EMReadScreen insured_member, 2, INSA_row, INSA_col
				If insured_member <> "__" then
					if member_list = "" then member_list = insured_member
					if member_list <> "" then member_list = member_list & ", " & insured_member
					INSA_col = INSA_col + 4
					If INSA_col = 70 then
						INSA_col = 30 : INSA_row = 16
					End If
				End If
			loop until insured_member = "__"
			'Retain "variable_written_to" as is while also adding members covered by the insurance policy
			'Example - "Members: 01, 03, 07 are covered by Blue Cross Blue Shield; "
			variable_written_to = variable_written_to & "Members: " & member_list & " are covered by " & trim(insurance_name) & "; "
		Next
		'This will loop and add the above statement for all insurance policies listed
	End if
  Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_JOBS_to_variable(variable_written_to)
          EMReadScreen JOBS_panel_current, 1, 2, 73
          If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
        Loop until cint(JOBS_panel_current) = cint(JOBS_total)
      End if
    Next
  Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen MEDI_amt, 1, 2, 78
      If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
    Next
  Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen rel_to_applicant, 2, 10, 42
      EMReadScreen client_age, 3, 8, 76
      If client_age = "   " then client_age = 0
      If cint(client_age) >= 21 or rel_to_applicant = "02" then
        number_of_adults = number_of_adults + 1
      Else
        number_of_children = number_of_children + 1
      End if
    Next
    If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
    If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
    If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
  Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen citizen, 1, 11, 49
      If citizen = "Y" then citizen = "US citizen"
      If citizen = "N" then citizen = "non-citizen"
      EMReadScreen citizenship_ver, 2, 11, 78
      EMReadScreen SSA_MA_citizenship_ver, 1, 12, 49
      If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
      If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
      If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
      variable_written_to = variable_written_to & "Member " & HH_member & "- "
      variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
    Next
  ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
    EMReadScreen variable_written_to, 8, 6, 39
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen OTHR_total, 2, 2, 78
	  OTHR_total = trim(OTHR_total)
      If OTHR_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_OTHR_to_variable(variable_written_to)
          EMReadScreen OTHR_panel_current, 2, 2, 72
		  OTHR_panel_current = trim(OTHR_panel_current)
          If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
        Loop until cint(OTHR_panel_current) = cint(OTHR_total)
      End if
    Next
  Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen panel_amt, 1, 2, 78
      If panel_amt <> "0" then
        PBEN = PBEN & "Member " & HH_member & "- "
        row = 8
        Do
          EMReadScreen PBEN_type, 12, row, 28
          EMReadScreen PBEN_disp, 1, row, 77
          If PBEN_disp = "A" then PBEN_disp = " appealing"
          If PBEN_disp = "D" then PBEN_disp = " denied"
          If PBEN_disp = "E" then PBEN_disp = " eligible"
          If PBEN_disp = "P" then PBEN_disp = " pends"
          If PBEN_disp = "N" then PBEN_disp = " not applied yet"
          If PBEN_disp = "R" then PBEN_disp = " refused"
          If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
          row = row + 1
        Loop until row = 14
      End if
    Next
    If PBEN <> "" then variable_written_to = variable_written_to & PBEN
  Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen PREG_total, 1, 2, 78
      If PREG_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen PREG_due_date, 8, 10, 53
        If PREG_due_date = "__ __ __" then
          PREG_due_date = "unknown"
        Else
          PREG_due_date = replace(PREG_due_date, " ", "/")
        End if
        variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
      End if
    Next
  Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
    row = 6
    Do
      EMReadScreen appl_prog_date, 8, row, 33
      If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
      row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    variable_written_to = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
      if CDate(appl_prog_date_array(i)) > variable_written_to then
        variable_written_to = CDate(appl_prog_date_array(i))
      End if
    next
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen RBIC_total, 1, 2, 78
      If RBIC_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_RBIC_to_variable(variable_written_to)
          EMReadScreen RBIC_panel_current, 1, 2, 73
          If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
        Loop until cint(RBIC_panel_current) = cint(RBIC_total)
      End if
    Next
  Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen REST_total, 2, 2, 78
	  REST_total = trim(REST_total)
      If REST_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_REST_to_variable(variable_written_to)
          EMReadScreen REST_panel_current, 2, 2, 72
		  REST_panel_current = trim(REST_panel_current)
          If cint(REST_panel_current) < cint(REST_total) then transmit
        Loop until cint(REST_panel_current) = cint(REST_total)
      End if
    Next
  Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
    EMReadScreen variable_written_to, 8, 13, 37
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
	For each HH_member in HH_member_array
			EMWriteScreen HH_member, 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen school_type, 2, 7, 40							'Reading the school type code and converting it into words
			If school_type = "01" then school_type = "elementary school"
			If school_type = "11" then school_type = "middle school"
			If school_type = "02" then school_type = "high school"
			If school_type = "03" then school_type = "GED"
			If school_type = "07" then school_type = "IEP"
			If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
			If school_type = "12" then school_type = "adult basic education"
			If school_type = "13" then school_type = "English as a 2nd language"
			If school_type = "06" or school_type = "__" or school_type = "?_" then  'if the school type is blank, child not in school, or postponed default type to blank.
				school_type = ""
			Else
				EMReadScreen SCHL_ver, 2, 6, 63
				If SCHL_ver = "?_" or SCHL_ver = "NO" then								'If the verification field is postponed or NO it defaults to no proof provided
					school_proof_type = ", no proof provided"
				Else
					school_proof_type = ""
				End if
				EMReadScreen FS_eligibility_status_SCHL, 2, 16, 63				'Reading the FS eligibility status and converting it to words
				IF FS_eligibility_status_SCHL = "01" THEN FS_eligibility_status_SCHL = ", FS Elig Status: < 18 or 50+"
				IF FS_eligibility_status_SCHL = "02" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Disabled"
				IF FS_eligibility_status_SCHL = "03" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Attenting Higher Ed or Attending < 1/2"
				IF FS_eligibility_status_SCHL = "04" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Employed 20 Hours/Wk"
				IF FS_eligibility_status_SCHL = "05" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Fed/State Work Study Program"
				IF FS_eligibility_status_SCHL = "06" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant Under 6"
				IF FS_eligibility_status_SCHL = "07" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant 6-11, daycare not available"
				IF FS_eligibility_status_SCHL = "09" THEN FS_eligibility_status_SCHL = ", FS Elig Status: WIA, TAA, TRA, or FSET placement"
				IF FS_eligibility_status_SCHL = "10" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Full Time Single Parent with Child under 12"
				IF FS_eligibility_status_SCHL = "99" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Eligible"
				IF FS_eligibility_status_SCHL = "__" or FS_eligibility_status_SCHL = "?_" THEN FS_eligibility_status_SCHL = ""
				'formatting the output variable for the function
				variable_written_to = variable_written_to & "Member " & HH_member & "- "
				variable_written_to = variable_written_to & school_type & school_proof_type & FS_eligibility_status_SCHL & "; "
			End if
		Next
  Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SECU_total, 2, 2, 78
	  SECU_total = trim(SECU_total)
      If SECU_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_SECU_to_variable(variable_written_to)
          EMReadScreen SECU_panel_current, 2, 2, 72
		  SECU_panel_current = trim(SECU_panel_current)
          If cint(SECU_panel_current) < cint(SECU_total) then transmit
        Loop until cint(SECU_panel_current) = cint(SECU_total)
      End if
    Next
  Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SHEL_total, 1, 2, 78
      If SHEL_total <> 0 then
        member_number_designation = "Member " & HH_member & "- "
        row = 11
        Do
          EMReadScreen SHEL_amount, 8, row, 56
          If SHEL_amount <> "________" then
            EMReadScreen SHEL_type, 9, row, 24
            EMReadScreen SHEL_proof_check, 2, row, 67
            If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then
              SHEL_proof = ", no proof provided"
            Else
              SHEL_proof = ""
            End if
            SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ". ;"
          End if
          row = row + 1
        Loop until row = 19
        variable_written_to = variable_written_to & member_number_designation & SHEL_expense
      End if
      SHEL_expense = ""
    Next
 Elseif panel_read_from = "SWKR" then '---------------------------------------------------------------------------------------------------SWKR
    EMReadScreen SWKR_name, 35, 6, 32
    SWKR_name = replace(SWKR_name, "_", "")
    SWKR_name = split(SWKR_name)
    For each word in SWKR_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen STWK_total, 1, 2, 78
      If STWK_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen STWK_verification, 1, 7, 63
        If STWK_verification = "N" then
          STWK_verification = ", no proof provided"
        Else
          STWK_verification = ""
        End if
        EMReadScreen STWK_employer, 30, 6, 46
        STWK_employer = replace(STWK_employer, "_", "")
        STWK_employer = split(STWK_employer)
        For each STWK_part in STWK_employer
          If STWK_part <> "" then
            first_letter = ucase(left(STWK_part, 1))
            other_letters = LCase(right(STWK_part, len(STWK_part) -1))
            If len(STWK_part) > 3 then
              new_STWK_employer = new_STWK_employer & first_letter & other_letters & " "
            Else
              new_STWK_employer = new_STWK_employer & STWK_part & " "
            End if
          End if
        Next
        EMReadScreen STWK_income_stop_date, 8, 8, 46
        If STWK_income_stop_date = "__ __ __" then
          STWK_income_stop_date = "at unknown date"
        Else
          STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
        End if
      EMReadScreen voluntary_quit, 1, 10, 46
	vol_quit_info = ", Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
      End if
      new_STWK_employer = "" 'clearing variable to prevent duplicates
    Next
  Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen UNEA_total, 1, 2, 78
      If UNEA_total <> 0 then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_UNEA_to_variable(variable_written_to)
          EMReadScreen UNEA_panel_current, 1, 2, 73
          If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
        Loop until cint(UNEA_panel_current) = cint(UNEA_total)
      End if
    Next
  Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
    EMReadScreen wreg_total, 1, 2, 78
    IF wreg_total <> "0" THEN
	EmWriteScreen "x", 13, 57
	transmit
	 bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
	  bene_yr_row = 10
       abawd_counted_months = 0
       second_abawd_period = 0
 	 month_count = 0
 	   DO
	   		'establishing variables for specific ABAWD counted month dates
	 		If bene_mo_col = "19" then counted_date_month = "01"
	 		If bene_mo_col = "23" then counted_date_month = "02"
	 		If bene_mo_col = "27" then counted_date_month = "03"
	 		If bene_mo_col = "31" then counted_date_month = "04"
	 		If bene_mo_col = "35" then counted_date_month = "05"
	 		If bene_mo_col = "39" then counted_date_month = "06"
	 		If bene_mo_col = "43" then counted_date_month = "07"
	 		If bene_mo_col = "47" then counted_date_month = "08"
	 		If bene_mo_col = "51" then counted_date_month = "09"
	 		If bene_mo_col = "55" then counted_date_month = "10"
	 		If bene_mo_col = "59" then counted_date_month = "11"
	 		If bene_mo_col = "63" then counted_date_month = "12"
	 		'reading to see if a month is counted month or not
  		  	EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
			'counting and checking for counted ABAWD months
			IF is_counted_month = "X" or is_counted_month = "M" THEN
				EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
				abawd_counted_months_string = counted_date_month & "/" & counted_date_year
				abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
				abawd_counted_months = abawd_counted_months + 1				'adding counted months
			END IF

			'declaring & splitting the abawd months array
			If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
			abawd_months_array = Split(abawd_info_list, ",")

			'counting and checking for second set of ABAWD months
			IF is_counted_month = "Y" or is_counted_month = "N" THEN
				EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
				second_abawd_period = second_abawd_period + 1				'adding counted months
				second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
				second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
			END IF

			'declaring & splitting the second set of abawd months array
			If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
			second_months_array = Split(second_set_info_list,",")

			bene_mo_col = bene_mo_col - 4
    		IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
   	     		bene_mo_col = 63
   	   	    END IF
    		month_count = month_count + 1
  	   LOOP until month_count = 36
  		PF3

	EmreadScreen read_WREG_status, 2, 8, 50
	If read_WREG_status = "03" THEN  WREG_status = "WREG = incap"
	If read_WREG_status = "04" THEN  WREG_status = "WREG = resp for incap HH memb"
	If read_WREG_status = "05" THEN  WREG_status = "WREG = age 60+"
	If read_WREG_status = "06" THEN  WREG_status = "WREG = < age 16"
	If read_WREG_status = "07" THEN  WREG_status = "WREG = age 16-17, live w/prnt/crgvr"
	If read_WREG_status = "08" THEN  WREG_status = "WREG = resp for child < 6 yrs old"
	If read_WREG_status = "09" THEN  WREG_status = "WREG = empl 30 hrs/wk or equiv"
	If read_WREG_status = "10" THEN  WREG_status = "WREG = match grant part"
	If read_WREG_status = "11" THEN  WREG_status = "WREG = rec/app for unemp ins"
	If read_WREG_status = "12" THEN  WREG_status = "WREG = in schl, train prog or higher ed"
	If read_WREG_status = "13" THEN  WREG_status = "WREG = in CD prog"
	If read_WREG_status = "14" THEN  WREG_status = "WREG = rec MFIP"
	If read_WREG_status = "20" THEN  WREG_status = "WREG = pend/rec DWP or WB"
	If read_WREG_status = "22" THEN  WREG_status = "WREG = app for SSI"
	If read_WREG_status = "15" THEN  WREG_status = "WREG = age 16-17 not live w/ prnt/crgvr"
	If read_WREG_status = "16" THEN  WREG_status = "WREG = 50-59 yrs old"
	If read_WREG_status = "21" THEN  WREG_status = "WREG = resp for child < 18"
	If read_WREG_status = "17" THEN  WREG_status = "WREG = rec RCA or GA"
	If read_WREG_status = "18" THEN  WREG_status = "WREG = provide home schl"
	If read_WREG_status = "30" THEN  WREG_status = "WREG = mand FSET part"
	If read_WREG_status = "02" THEN  WREG_status = "WREG = non-coop w/ FSET"
	If read_WREG_status = "33" THEN  WREG_status = "WREG = non-coop w/ referral"
	If read_WREG_status = "__" THEN  WREG_status = "WREG = blank"

	EmreadScreen read_abawd_status, 2, 13, 50
	If read_abawd_status = "01" THEN  abawd_status = "ABAWD = work reg exempt."
    	If read_abawd_status = "02" THEN  abawd_status = "ABAWD = < age 18."
	If read_abawd_status = "03" THEN  abawd_status = "ABAWD = age 50+."
	If read_abawd_status = "04" THEN  abawd_status = "ABAWD = crgvr of minor child."
	If read_abawd_status = "05" THEN  abawd_status = "ABAWD = pregnant."
	If read_abawd_status = "06" THEN  abawd_status = "ABAWD = emp ave 20 hrs/wk."
	If read_abawd_status = "07" THEN  abawd_status = "ABAWD = work exp participant."
	If read_abawd_status = "08" THEN  abawd_status = "ABAWD = othr E & T service."
	If read_abawd_status = "09" THEN  abawd_status = "ABAWD = reside in waiver area."
	IF read_abawd_status = "10" AND abawd_counted_months = "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo."
	Elseif read_abawd_status = "10" AND second_abawd_period = "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & "."
	Elseif read_abawd_status = "10" AND second_abawd_period <> "0" THEN
		abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & ". Counted second set months: " & second_set_info_list & "."
	END IF
	If read_abawd_status = "11" THEN  abawd_status = "ABAWD = Using second set of ABAWD months. Counted second set months: " & second_set_info_list & "."
	If read_abawd_status = "12" THEN  abawd_status = "ABAWD = RCA or GA recip."
	If read_abawd_status = "13" THEN  abawd_status = "ABAWD = ABAWD banked months."
	If read_abawd_status = "__" THEN  abawd_status = "ABAWD = blank"

	variable_written_to = variable_written_to & "Member " & HH_member & "- " & WREG_status & ", " & abawd_status & "; "
     END IF
    Next
  End if
  variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
  if right(variable_written_to, 1) = ";" then variable_written_to = left(variable_written_to, len(variable_written_to) - 1)
end function

function back_to_SELF()
'--- This function will return back to the 'SELF' menu or the MAXIS home menu
'===== Keywords: MAXIS, SELF, navigate
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
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

function change_client_name_to_FML(client_name)
'--- This function changes the format of a participant name. client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".
'~~~~~ client_name: variable used within the script for name to be converted
'===== Keywords: PRISM, MAXIS, name, change

	client_name = trim(client_name)
	length = len(client_name)

	'Adds handling for names that have no spaces or 1 space
	If Instr(client_name, ", ") then
		position = InStr(client_name, ", ")
		first_name = Right(client_name, length-position - 1)
	elseif Instr(client_name, ",") then
		position = InStr(client_name, ",")                           '
		first_name = Right(client_name, length-position)
	END if
	last_name = Left(client_name, position - 1)

	'final formating of the client name
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
end function

function change_date_to_soonest_working_day(date_to_change)
'--- This function will change a date that is on a weekend or Hennepin County holiday to the next working date before the date provided, the date will remain the same if it is not a holiday or weekend.
'~~~~~ date_to_change: variable in the form of a date - this will change once the function is called
'===== Keywords: MAXIS, date, change
    Do
        is_holiday = FALSE
        For each holiday in HOLIDAYS_ARRAY
            If holiday = date_to_change Then
                is_holiday = TRUE
                date_to_change = DateAdd("d", -1, date_to_change)
            End If
        Next
        If WeekdayName(WeekDay(date_to_change)) = "Saturday" Then date_to_change = DateAdd("d", -1, date_to_change)
        If WeekdayName(WeekDay(date_to_change)) = "Sunday" Then date_to_change = DateAdd("d", -2, date_to_change)
    Loop until is_holiday = FALSE
end function

function changelog_display()
'--- This function determines if the user has been informed of a change to a script, and if not will display a mesage box with the script's change log information
'===== Keywords: MAXIS, PRISM, change, info, information
	If changelog_enabled = "" Then changelog_enabled = true
	If changelog_enabled <> false Then
		'Needs to determine MyDocs directory before proceeding.
		Set wshshell = CreateObject("WScript.Shell")
		user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

		'Now determines name of file
		local_changelog_path = user_myDocs_folder & "scripts-local-changelog-entries.txt"

		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8

		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")

		'Before doing comparisons, it needs to see what the most recent item added to the list was.
		last_item_added_to_changelog = split(changelog(0), " | ")

		With objFSO

			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream

			'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

			If .FileExists(local_changelog_path) = False then
				'Setting the object to open the text file for appending the new data
				Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)

				'Write the contents of the text file
				objTextStream.WriteLine date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Close the object so it can be opened again shortly
				objTextStream.Close

				'Since the file was new, we can simply exit the function
				exit function
			End if

			'Setting the object to open the text file for reading the data already in the file
			Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

			'Reading the entire text file into a string
			every_line_in_text_file = objTextStream.ReadAll

			'Splitting the text file contents into an array which will be sorted
			local_changelog_array = split(every_line_in_text_file, vbNewLine)

			'Looks to see if the script has been used before!
			'for each local_changelog_item in local_changelog_array
			for i = 0 to ubound(local_changelog_array)
				If local_changelog_array(i) <> "" then 'some are likely blank
					'splits the local_changelog_array(i) into an array: 0 -> date, 1 -> name_of_script, 2 -> text_of_change
					local_changelog_item_array = split(local_changelog_array(i), " | ")

					'Looking to see if the script is in fact in the local changelog list. If it is, we will then check the text against the listed changes to see what needs to be displayed.
					if local_changelog_item_array(1) = name_of_script then
						script_in_local_changelog = true
						if local_changelog_item_array(2) <> last_item_added_to_changelog(1) then
							display_changelog = true
							local_changelog_text_of_change = trim(local_changelog_item_array(2))
							line_in_local_changelog_array_to_delete = i
						Else
							display_changelog = false
						End if
					End if
				End if
			next

			'Close the file
			objTextStream.Close

			'If the script is not in the local changelog, it needs to be added. If this is the case, it shouldn't display the changelog at all, because it'll be the first time the script was run.
			If script_in_local_changelog <> true then

				'Setting the object to open the text file for appending the new data
				Set objTextStream = .OpenTextFile(local_changelog_path, ForAppending, true)

				'Write the contents of the text file
				objTextStream.WriteLine date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Close the file and clean up
				objTextStream.Close

				'Setting this to false. We don't want to display the changelog if the script has never been added to the local list of changelog events
				display_changelog = false

			End if

			'So, if the script IS in the local changelog, and needs to be displayed, it takes special handling to ensure that's done.
			If display_changelog = true then

				'Splitting the changelog into different variables for making things prettier
				For each changelog_entry in changelog
					date_of_change = left(changelog_entry, instr(changelog_entry, " | ") - 1)
					scriptwriter_of_change = trim(right(changelog_entry, len(changelog_entry) - instrrev(changelog_entry, "|") ))
					text_of_change = replace(replace(replace(changelog_entry, scriptwriter_of_change, ""), date_of_change, ""), " | ", "")

					'If the text_of_change is the same as that stored in the local changelog, that means the user is up-to-date to this point, and the script should exit without displaying any more updates. Otherwise, add it to the contents.
					if trim(text_of_change) = trim(local_changelog_text_of_change) then
					 	exit for
					else
                        If name_of_script = "Functions Library" Then
                            changelog_msgbox = changelog_msgbox & "-----" & cdate(date_of_change) & "-----" & vbNewLine & text_of_change & vbNewLine & vbNewLine & "Thank you!" & vbNewLine & "The BlueZone Script Team" & vbNewLine & vbNewLine
                        Else
                            changelog_msgbox = changelog_msgbox & "-----" & cdate(date_of_change) & "-----" & vbNewLine & text_of_change & vbNewLine & "Completed by " & scriptwriter_of_change & vbNewLine & vbNewLine
                        End If
					end if

				Next

				If changelog_msgbox <> "" then
                    If name_of_script = "Functions Library" Then
                        message_of_change = MsgBox("Script Announcement: " & vbNewLine & vbNewLine & changelog_msgbox, vbSystemModal, "BZST Communication")
                        'MsgBox "Script Announcement: " & vbNewLine & vbNewLine & changelog_msgbox
                    Else
                        message_of_change = MsgBox("Recent changes in this script: " & vbNewLine & vbNewLine & changelog_msgbox, vbSystemModal, "BZST Changes to Script")
                        'MsgBox "Recent changes in this script: " & vbNewLine & vbNewLine & changelog_msgbox
                    End If
				End if

				'Now we need to determine what the most recent change is, in order to add this to our text file
				string_to_enter_into_local_changelog = date & " | " & name_of_script & " | " & last_item_added_to_changelog(1)

				'Lastly, if it displayed a changelog, it should go through and update the record to remove the old entry and replace it with this one.
				Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)						'Opening the file one last time
				for i = 0 to ubound(local_changelog_array)
					If i = line_in_local_changelog_array_to_delete then local_changelog_array(i) = string_to_enter_into_local_changelog
					if local_changelog_array(i) <> "" then objTextStream.WriteLine local_changelog_array(i)
				next

			end if

			'Close the file
			objTextStream.Close
		End with
	End If

end function

function changelog_update(date_of_change, text_of_change, scriptwriter_of_change)
'--- This function adds the change to the scripts to the user change log to be displayed in function changelog_display()
'~~~~~ date_of_change: date the change was made/committed to the script file. Surround date in ""
'~~~~~ text_of_change: information about the change to the script that users statewide will see. Please be clear about your updates. You can write several sentences. Surround text in "".
'~~~~~ scriptwriter_of_change: scriptwriter name and county seperated by a comma. Surround name and county name with "".
'===== Keywords: MAXIS, PRISM, change, info, information
	If changelog_enabled = "" Then changelog_enabled = true
	If changelog_enabled <> false Then
		ReDim Preserve changelog(UBound(changelog) + 1)
		changelog(ubound(changelog)) = date_of_change & " | " & text_of_change & " | " & scriptwriter_of_change
	End If
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
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
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
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

function check_for_MMIS(end_script)
'--- This function checks to ensure the user is in a MMIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MMIS screen.
'===== Keywords: MMIS, production, script_end_procedure
	Do
		transmit
		row = 1
		col = 1
		EMSearch "MMIS", row, col
		IF row <> 1 then
			If end_script = True then
				script_end_procedure("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until row = 1
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

function check_for_password_without_transmit(are_we_passworded_out)
'--- This function checks to make sure a user is not passworded out. If they are, it allows the user to password back in. NEEDS TO BE ADDED INTO dialog DO...lOOPS
'~~~~~ are_we_passworded_out: When adding to dialog enter "Call check_for_password(are_we_passworded_out)", then Loop until are_we_passworded_out = false. Parameter will remain true if the user still needs to input password.
'===== Keywords: MAXIS, PRISM, password
	Emreadscreen password_check, 8, 2, 33 'checking for the word password which will indicate you are passworded out
	If password_check = "PASSWORD" then 'If the word password is found then it will tell the worker and set the parameter to be true, otherwise it will be set to false.
		Msgbox "Are you passworded out? Press OK and the dialog will reappear. Once it does, you can enter your password."
		are_we_passworded_out = true
	Else
		are_we_passworded_out = false
	End If
end function

function check_for_PRISM(end_script)
'--- This function checks to ensure the user is in a PRISM panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a PRISM screen.
'===== Keywords: PRISM, production, script_end_procedure
	EMReadScreen PRISM_check, 5, 1, 36
	if end_script = True then
		If PRISM_check <> "PRISM" then script_end_procedure("You do not appear to be in PRISM. You may be passworded out. Please check your PRISM screen and try again.")
	else
		If PRISM_check <> "PRISM" then MsgBox "You do not appear to be in PRISM. You may be passworded out. Please enter your password before pressing OK."
	end if
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

function convert_array_to_droplist_items(array_to_convert, output_droplist_box)
'--- This function converts an array into a droplist to be used within dialog
'~~~~~ array_to_convert: name of the array
'~~~~~ output_droplist_box: name of droplist variant/variable
'===== Keywords: MAXIS, PRISM, production, array, droplist
	For each item in array_to_convert
		If output_droplist_box = "" then
			output_droplist_box = item
		Else
			output_droplist_box = output_droplist_box & chr(9) & item
		End if
	Next
end function

function convert_date_into_MAXIS_footer_month(date_to_convert, MAXIS_footer_month, MAXIS_footer_year)
'--- This function converts a date (MM/DD/YY or MM/DD/YYYY format) into a separate footer month and footer year variables.
'~~~~~ date_to_convert: variable name of date you want to convert
'~~~~~ MAXIS_footer_month: footer month to convert the date into
'~~~~~ MAXIS_footer_month: footer year to convert the date into
'===== Keywords: MAXIS, production, array, droplist, convert
	MAXIS_footer_month = DatePart("m", date_to_convert)										'Uses DatePart function to copy the month from date_to_convert into the MAXIS_footer_month variable.
	IF Len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month		'Uses Len function to determine if the MAXIS_footer_month is a single digit month. If so, it adds a 0, which MAXIS needs.
	MAXIS_footer_year = DatePart("yyyy", date_to_convert)									'Uses DatePart function to copy the year from date_to_convert into the MAXIS_footer_year variable.
	MAXIS_footer_year = Right(MAXIS_footer_year, 2)											'Uses Right function to reduce the MAXIS_footer_year variable to it's right 2 characters (allowing for a 2 digit footer year).
end function

function convert_digit_to_excel_column(col_in_excel)
'--- This function converts a numeric digit to an Excel column, up to 104 digits (columns).
'~~~~~ col_in_excel: must be a numeric, cannot exceed 104. Do not put in "".
'===== Keywords: MAXIS, PRISM, convert, Excel
	'Create string with the alphabet

	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	'Assigning a letter, based on that column. Uses "mid" function to determine it. If number > 26, it handles by adding a letter (per Excel).
	convert_digit_to_excel_column = Mid(alphabet, col_in_excel, 1)
	If col_in_excel >= 27 and col_in_excel < 53 then convert_digit_to_excel_column = "A" & Mid(alphabet, col_in_excel - 26, 1)
	If col_in_excel >= 53 and col_in_excel < 79 then convert_digit_to_excel_column = "B" & Mid(alphabet, col_in_excel - 52, 1)
	If col_in_excel >= 79 and col_in_excel < 105 then convert_digit_to_excel_column = "C" & Mid(alphabet, col_in_excel - 78, 1)

	'Closes script if the number gets too high (very rare circumstance, just errorproofing)
	If col_in_excel >= 105 then script_end_procedure("This script is only able to assign excel columns to 104 distinct digits. You've exceeded this number, and this script cannot continue.")
end function


function create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
'--- This function is used to grab all active X numbers according to the supervisor X number(s) inputted
'~~~~~ array_name: name of array that will contain all the supervisor's staff x numbers
'~~~~~ supervisor_array: list of supervisor's x numbers seperated by comma
'===== Keywords: MAXIS, array, supervisor, worker number, create
	'Create string with the alphabet
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")

	'Sorting by supervisor
	PF5
	PF5

	'Reseting array_name
	array_name = ""

	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)

			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
end function

function create_array_of_all_active_x_numbers_in_county(array_name, county_code)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	call navigate_to_MAXIS_screen("rept", "user")

	'Hitting PF5 to force sorting, which allows directly selecting a county
	PF5

	'Inserting county
	EMWriteScreen county_code, 21, 6
	transmit

	'Declaring the MAXIS row
	MAXIS_row = 7

	'Blanking out array_name in case this has been used already in the script
	array_name = ""

	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
			array_name = trim(array_name & " " & worker_ID)				'writing to variable
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then
			PF8			'getting to next screen
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank
	array_name = split(array_name)
end function

function create_mainframe_friendly_date(date_variable, screen_row, screen_col, year_type)
'--- This function creates a mainframe friendly date. This can be used for both year formats and input spacing.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'~~~~~ year_type: formatting to export date year as "YY" or "YYYY"
'===== Keywords: MAXIS, PRISM, MMIS, date, create
	var_month = datepart("m", date_variable)
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month & "/", screen_row, screen_col
	var_day = datepart("d", date_variable)
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day & "/", screen_row, screen_col + 3
	If year_type = "YY" then
		var_year = right(datepart("yyyy", date_variable), 2)
	ElseIf year_type = "YYYY" then
		var_year = datepart("yyyy", date_variable)
	Else
		MsgBox "Year type entered incorrectly. Fourth parameter of function create_mainframe_friendly_date should read ""YYYY"" or ""YY"". The script will now stop."
		StopScript
	END IF
	EMWriteScreen var_year, screen_row, screen_col + 6
end function

function create_MAXIS_friendly_date(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM DD YY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	If len(var_month) = 1 then var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	If len(var_day) = 1 then var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
end function

function create_MAXIS_friendly_date_three_spaces_between(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM  DD  YY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))		'determines the date based on the variable length: month
	If len(var_month) = 1 then var_month = "0" & var_month				'adds a '0' in front of a single digit month
	EMWriteScreen var_month, screen_row, screen_col					'writes in var_month at coordinates set in function line
	var_day = datepart("d", dateadd("d", variable_length, date_variable)) 		'determines the date based on the variable length: day
	If len(var_day) = 1 then var_day = "0" & var_day 				'adds a '0' in front of a single digit day
	EMWriteScreen var_day, screen_row, screen_col + 5 				'writes in var_day at coordinates set in function line, and starts 5 columns into date field in MAXIS
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable)) 	'determines the date based on the variable length: year
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 10 			'writes in var_year at coordinates set in function line , and starts 5 columns into date field in MAXIS
end function

function create_MAXIS_friendly_date_with_YYYY(date_variable, variable_length, screen_row, screen_col)
'--- This function creates a MM DD YYYY date entry into BlueZone.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ variable_length:the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'===== Keywords: MAXIS, date, create
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen var_year, screen_row, screen_col + 6
end function

function create_MAXIS_friendly_phone_number(phone_number_variable, screen_row, screen_col)
'--- This function creates a MAXIS friendly phone number
'~~~~~ phone_number_variable: the name of the variable to output
'~~~~~ screen_row: row to start writing phone number
'~~~~~ screen_col: column to start writing phone number
'===== Keywords: MAXIS, date, create
	WITH (new RegExp)                                                            	'Uses RegExp to bring in special string functions to remove the unneeded strings
                .Global = True                                                   	'I don't know what this means but David made it work so we're going with it
                .Pattern = "\D"                                                	 	'Again, no clue. Just do it.
                phone_number_variable = .Replace(phone_number_variable, "")    	 	'This replaces the non-digits of the phone number with nothing. That leaves us with a bunch of numbers
	END WITH
	EMWriteScreen left(phone_number_variable, 3), screen_row, screen_col 		'writes in left 3 digits of the phone number in variable
	EMWriteScreen mid(phone_number_variable, 4, 3), screen_row, screen_col + 6	'writes in middle 3 digits of the phone number in variable
	EMWriteScreen right(phone_number_variable, 4), screen_row, screen_col + 12	'writes in right 4 digits of the phone number in variable
end function

FUNCTION create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
'--- This function creates a an outlook appointment
'~~~~~ (appt_date): date of the appointment
'~~~~~ (appt_start_time): start time of the appointment - format example: "08:00 AM"
'~~~~~ (appt_end_time): end time of the appointment - format example: "08:00 AM"
'~~~~~ (appt_subject): subject of the email in quotations or a variable
'~~~~~ (appt_body): body of the email in quotations or a variable
'~~~~~ (appt_location): name of location in quotations or a variable
'~~~~~ (appt_reminder): reminder for appointment. Set to TRUE or FALSE
'~~~~~ (reminder_in_minutes): enter the number of minutes prior to the appointment to set the reminder. Set as 0 if at the time of the appoint. Set to "" if appt_reminder is set to FALSE
'~~~~~ (appt_category): can be left "" or assgin to the set the name of the category in quotations
'===== Keywords: MAXIS, PRISM, create, outlook, appointment

	'Assigning needed numbers as variables for readability
	olAppointmentItem = 1
	olRecursDaily = 0

	'Creating an Outlook object item
	Set objOutlook = CreateObject("Outlook.Application")
	Set objAppointment = objOutlook.CreateItem(olAppointmentItem)

	'Assigning individual appointment options
	objAppointment.Start = appt_date & " " & appt_start_time		'Start date and time are carried over from parameters
	objAppointment.End = appt_date & " " & appt_end_time			'End date and time are carried over from parameters
	objAppointment.AllDayEvent = False 								'Defaulting to false for this. Perhaps someday this can be true. Who knows.
	objAppointment.Subject = appt_subject							'Defining the subject from parameters
	objAppointment.Body = appt_body									'Defining the body from parameters
	objAppointment.Location = appt_location							'Defining the location from parameters
	If appt_reminder = FALSE then									'If the reminder parameter is false, it skips the reminder, otherwise it sets it to match the number here.
		objAppointment.ReminderSet = False
	Else
		objAppointment.ReminderSet = True
		objAppointment.ReminderMinutesBeforeStart = reminder_in_minutes
	End if
	objAppointment.Categories = appt_category						'Defines a category
	objAppointment.Save												'Saves the appointment
END FUNCTION

Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'--- This function creates a an outlook appointment
'~~~~~ (email_recip): email address for recipeint - seperated by semicolon
'~~~~~ (email_recip_CC): email address for recipeints to cc - seperated by semicolon
'~~~~~ (email_subject): subject of email in quotations or a variable
'~~~~~ (email_body): body of email in quotations or a variable
'~~~~~ (email_attachment): set as "" if no email or file location
'~~~~~ (send_email): set as TRUE or FALSE
'===== Keywords: MAXIS, PRISM, create, outlook, email

	'Setting up the Outlook application
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    objMail.Display                                 'To display message

    'Adds the information to the email
    objMail.to = email_recip                        'email recipient
    objMail.cc = email_recip_CC                     'cc recipient
    objMail.Subject = email_subject                 'email subject
    objMail.Body = email_body                       'email body
    If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
    'Sends email
    If send_email = true then objMail.Send	                   'Sends the email
    Set objMail =   Nothing
    Set objOutlook = Nothing
End Function

function create_panel_if_nonexistent()
'--- This function creates a panel if a panel does not exist. This is currently only used within the FuncLib itself.
'~~~~~ (): keep this parameter empty
'===== Keywords: MAXIS, FuncLib only, create
	EMWriteScreen reference_number , 20, 76
	transmit
	EMReadScreen case_panel_check, 44, 24, 2
	If case_panel_check = "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen quantity_of_screens, 1, 2, 78
		If quantity_of_screens <> "0" then
			PF9
		ElseIf quantity_of_screens = "0" then
			EMWriteScreen "__", 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	ElseIf case_panel_check <> "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen error_scan, 80, 24, 1
		error_scan = trim(error_scan)
		EMReadScreen quantity_of_screens, 1, 2, 78
		If error_scan = "" and quantity_of_screens <> "0" then
			PF9
		ElseIf error_scan <> "" and quantity_of_screens <> "0" then
			'FIX ERROR HERE
			msgbox("Error: " & error_scan)
		ElseIf error_scan <> "" and quantity_of_screens = "0" then
			EMWriteScreen reference_number, 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	End If
end function

function date_array_generator(initial_month, initial_year, date_array)
'--- This function creates a series of dates (Example: for each footer month/year through current month plus 1)
'~~~~~ initial_month: first footer month
'~~~~~ initial_year: first footer year
'~~~~~ date_array: the name of the array that holds the dates/number of months to create dates for
'===== Keywords: MAXIS, create, date, array
	'defines an intial date from the initial_month and initial_year parameters
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
end function

function date_converter_PALC_PAPL(date_variable)
'--- This function creates a creates a date in MM/DD/YY format
'~~~~~ date_variable: name of variable that holds the date info
'===== Keywords: PRISM, date convert, PALC, PAPL
	date_year = left (date_variable, 2)
	date_day = right (date_variable, 2)
	date_month = right (left (date_variable, 4), 2)

	date_variable = date_month & "/" & date_day & "/" & date_year
end function

function dynamic_calendar_dialog(selected_dates_array, month_to_use, text_prompt, one_date_only, disable_weekends, disable_month_change, start_date, end_date)
'--- This function creates a dynamic calendar that users can select dates from to be used in scheduleing. This is used in BULK - REVS SCRUBBER.
'~~~~~ selected_dates_array:the output array it will contain dates in MM/DD/YY format
'~~~~~ month_to_use: this can be MM/YY or MM/DD/YY format as long as it is considered a date it will work.
'~~~~~ one_date_only: this is a True/false parameter which will restrict the function to only allow one date to be selected if set to TRUE
'~~~~~ disable_weekends: this is a True/false parameter which will restrict the selection of weekends if set to TRUE
'~~~~~ disable_month_change: this is a True/false parameter which will restrict the selection of different months if set to TRUE
'~~~~~ start_date & end_date: this will provide a range of dates which cannot be selected. These are to be entered as numbers. For example start_date = 3 and end_date = 14 the days 3 through 14 will be unavailable to select
'===== Keywords: MAXIS, PRISM, create, date, calendar, dialog
	'dimming array to display the dates
	DIM display_dates_array
	DO
		full_date_to_display = ""					'resetting variables to make sure loops work properly.
		selected_dates_array = ""
		'Determining the number of days in the calendar month.
		display_month = DatePart("M", month_to_use) & "/01/" & DatePart("YYYY", month_to_use)			'Converts whatever the month_to_use variable is to a MM/01/YYYY format
		num_of_days = DatePart("D", (DateAdd("D", -1, (DateAdd("M", 1, display_month)))))								'Determines the number of days in a month by using DatePart to get the day of the last day of the month, and just using the day variable gives us a total

		'Redeclares the available dates array to be sized appropriately (with the right amount of dates) and another dimension for whether-or-not it was selected
		Redim display_dates_array(num_of_days, 0)


		'Actually displays the dialog
		BeginDialog dialog1, 0, 0, 280, 190, "Select Date(s)"
			Text 5, 10, 265, 50, text_prompt
			'This next part`creates a line showing the month displayed"
			Text 120, 70, 55, 10, (MonthName(DatePart("M", display_month)) & " " & DatePart("YYYY", display_month))

			'Defining the vertical position starting point for the for...next which displays dates in the dialog
			vertical_position = 85
			'This for...next displays dates in the dialog, and has checkboxes for available dates (defined in-code as dates before the 8th)
			for day_to_display = 1 to num_of_days																						'From first day of month to last day of month...
				full_date_to_display = (DatePart("M", display_month) & "/" & day_to_display & "/" & DatePart("YYYY", display_month))		'Determines the full date to display in the dialog. It needs the full date to determine the day-of-week (we obviously don't want weekends)
				horizontal_position = 15 + (40 * (WeekDay(full_date_to_display) - 1))													'horizontal position of this is the weekday numeric value (1-7) * 40, minus 1, and plus 15 pixels
				IF WeekDay(full_date_to_display) = vbSunday AND day_to_display <> 1 THEN vertical_position = vertical_position + 15		'If the day of the week isn't Sunday and the day isn't first of the month, kick the vertical position up another 15 pixels

				'This blocks out anything that's an unavailable date, currently defined as any date before the 8th. Other dates display as a checkbox.
				IF day_to_display <= end_date AND day_to_display >= start_date THEN
					Text horizontal_position, vertical_position, 30, 10, " X " & day_to_display
					display_dates_array(day_to_display, 0) = unchecked 'unchecking so selections cannot be made the range between start_date and end_date
				ELSE
					IF (disable_weekends = TRUE AND WeekDay(full_date_to_display) = vbSunday) OR (disable_weekends = TRUE AND WeekDay(full_date_to_display) = vbSaturday) THEN		'If the weekends are disabled this will change them to text rather than checkboxes
						Text horizontal_position, vertical_position, 30, 10, " X " & day_to_display
					ELSE
						CheckBox horizontal_position, vertical_position, 35, 10, day_to_display, display_dates_array(day_to_display, 0)
					END IF
				END IF
			NEXT
			ButtonGroup ButtonPressed
			OkButton 175, 170, 50, 15
			CancelButton 225, 170, 50, 15
			IF disable_month_change = FALSE THEN
				PushButton 85, 65, 20, 15, "<", prev_month_button
				PushButton 180, 65, 20, 15, ">", next_month_button
			END IF
		EndDialog

		IF one_date_only = TRUE THEN										'if only one date is allowed to be selected the script will act one way. Else it will allow for an large array of dates from a month to be build.
			DO
				selected_dates_array = ""									' declaring array at start of do loop.
				Dialog
				cancel_confirmation
				IF ButtonPressed = prev_month_button THEN month_to_use = dateadd("M", -1, month_to_use)				'changing the month_to_use based on previous or next month
				IF ButtonPressed = next_month_button THEN month_to_use = dateadd("M", 1, month_to_use)				'this will allow us to get to a new month when the dialog is rebuild.
				FOR i = 0 to num_of_days																			'checking each checkbox in the array to see what dates were selected.
					IF display_dates_array(i, 0) = 1 THEN 															'if the date has been checked
						IF len(DatePart("M", month_to_use)) = 1 THEN												'adding a leading 0 to the month if needed
							output_month = "0" & DatePart("M", month_to_use)
						ELSE
							output_month =  DatePart("M", month_to_use)
						END IF
						IF len(i) = 1 THEN 																			'building the output array with dates in MM/DD/YY format
							selected_dates_array = selected_dates_array & output_month & "/0" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
						ELSE
							selected_dates_array = selected_dates_array & output_month & "/" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
						END IF
					END IF
				NEXT
				selected_dates_array = selected_dates_array & "end"						'this will allow us to delete the extra entry in the array
				selected_dates_array = replace(selected_dates_array, ";end", "")
				selected_dates_array = Split(selected_dates_array, ";")					'splitting array
			IF Ubound(selected_dates_array) <> 0 AND (buttonpressed <> prev_month_button or buttonpressed <> next_month_button) THEN msgbox "Please select just one date."
			LOOP until Ubound(selected_dates_array) = 0
		ELSE
			Dialog
			cancel_confirmation
			IF ButtonPressed = prev_month_button THEN month_to_use = dateadd("M", -1, month_to_use)					'changing the month_to_use based on previous or next month
			IF ButtonPressed = next_month_button THEN month_to_use = dateadd("M", 1, month_to_use)					'this will allow us to get to a new month when the dialog is rebuild.
			FOR i = 0 to num_of_days																				'checking each checkbox in the array to see what dates were selected.
				IF display_dates_array(i, 0) = 1 THEN 																'if the date has been checked
					IF len(DatePart("M", month_to_use)) = 1 THEN 													'adding a leading 0 to the month if needed
						output_month = "0" & DatePart("M", month_to_use)
					ELSE
						output_month =  DatePart("M", month_to_use)
					END IF
					IF len(i) = 1 THEN 																				'building the output array with dates in MM/DD/YY format addding leading 0 to DD if needed.
						selected_dates_array = selected_dates_array & output_month & "/0" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
					ELSE
						selected_dates_array = selected_dates_array & output_month & "/" & i & "/" & right(DatePart("YYYY", month_to_use), 2) & ";"
					END IF
				END IF
			NEXT
			selected_dates_array = selected_dates_array & "end"							'this will allow us to delete the extra entry in the array
			selected_dates_array = replace(selected_dates_array, ";end", "")
			selected_dates_array = Split(selected_dates_array, ";")						'splitting array
		END IF
	LOOP until buttonpressed = -1								'looping until someone hits the ok button, this makes the previous and next buttons work.
end function

function end_excel_and_script()
'--- This function might not be needed anymore.
'===== Keywords: likely depreciated
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
end function

function enter_PRISM_case_number(case_number_variable, row, col)
'--- This function enters a PRISM case number.
'~~~~~ case_number_variable: always use <code>PRISM_case_number</code>
'~~~~~ row: row to write case number
'~~~~~ col: column to write case number
'===== Keywords: PRISM, case number
	EMSetCursor row, col
	EMSendKey replace(case_number_variable, "-", "")                                                                                                                                       'Entering the specific case indicated
	EMSendKey "<enter>"
	EMWaitReady 0, 0
end function

function excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = alerts_status
end function

function find_variable(opening_string, variable_name, length_of_variable)
'--- This function finds a string on a page in BlueZone
'~~~~~ opening_string: string to search for
'~~~~~ variable_name: variable name of the string
'~~~~~ length_of_variable: length of the string
'===== Keywords: MAXIS, MMIS, PRISM, find
  row = 1
  col = 1
  EMSearch opening_string, row, col
  If row <> 0 then EMReadScreen variable_name, length_of_variable, row, col + len(opening_string)
end function

function file_selection_system_dialog(file_selected, file_extension_restriction)
'--- This function allows a user to select a file to be opened in a script
'~~~~~ file_selected: variable for the name of the file
'~~~~~ file_extension_restriction: restricts all other file type besides allowed file type. Example: ".csv" only allows a CSV file to be accessed.
'===== Keywords: MAXIS, MMIS, PRISM, file
	'Creates a Windows Script Host object
	Set wShell=CreateObject("WScript.Shell")

	'This loops until the right file extension is selected. If it isn't specified (= ""), it'll always exit here.
	Do
		'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
		Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE ><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

		'Creates the file_selected variable from the exit
		file_selected = oExec.StdOut.ReadLine

		'If no file is selected the script will stop
		If file_selected = "" then stopscript

		'If the rightmost characters of the file selected don't match what was in the file_extension_restriction argument, it'll tell the user. Otherwise the loop (and function) ends.
		If right(file_selected, len(file_extension_restriction)) <> file_extension_restriction then MsgBox "You've entered an incorrect file type. The allowable file type is: " & file_extension_restriction & "."
	Loop until right(file_selected, len(file_extension_restriction)) = file_extension_restriction
end function

function find_MAXIS_worker_number(x_number)
'--- This function finds a MAXIS worker's X number
'~~~~~ x_number: worker number variable
'===== Keywords: MAXIS, worker number
	EMReadScreen SELF_check, 4, 2, 50		'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN				'if on the self screen then x # is read from coordinates
		EMReadScreen x_number, 7, 22, 8
	ELSE
		Call find_variable("PW: ", x_number, 7)	'if not, then the PW: variable is searched to find the worker #
		If isnumeric(MAXIS_worker_number) = true then 	 'making sure that the worker # is a number
			MAXIS_worker_number = x_number				'delcares the MAXIS_worker_number to be the x_number
		End if
	END if
end function

'This function fixes the case for a phrase. For example, "ROBERT P. ROBERTSON" becomes "Robert P. Robertson".
'	It capitalizes the first letter of each word.
function fix_case(phrase_to_split, smallest_length_to_skip)										'Ex: fix_case(client_name, 3), where 3 means skip words that are 3 characters or shorter
	phrase_to_split = split(phrase_to_split)													'splits phrase into an array
	For each word in phrase_to_split															'processes each word independently
		If word <> "" then																		'Skip blanks
			first_character = ucase(left(word, 1))												'grabbing the first character of the string, making uppercase and adding to variable
			remaining_characters = LCase(right(word, len(word) -1))								'grabbing the remaining characters of the string, making lowercase and adding to variable
			If len(word) > smallest_length_to_skip then											'skip any strings shorter than the smallest_length_to_skip variable
				output_phrase = output_phrase & first_character & remaining_characters & " "	'output_phrase is the output of the function, this combines the first_character and remaining_characters
			Else
				output_phrase = output_phrase & word & " "										'just pops the whole word in if it's shorter than the smallest_length_to_skip variable
			End if
		End if
	Next
	phrase_to_split = output_phrase																'making the phrase_to_split equal to the output, so that it can be used by the rest of the script.
end function

function fix_case_for_name(name_variable)
'--- This function takes in a client's name and outputs the name (accounting for hyphenated surnames) with Ucase first character & and lcase the rest. This is like fix_case but this function is a bit more specific for names
'~~~~~ name_variable: should be client_name for function to work
'===== Keywords: MAXIS, MMIS, PRISM, name, case
	name_variable = split(name_variable, " ")
	FOR EACH client_name IN name_variable
		IF client_name <> "" THEN
			IF InStr(client_name, "-") = 0 THEN
				client_name = UCASE(left(client_name, 1)) & LCASE(right(client_name, len(client_name) - 1))
				output_variable = output_variable & " " & client_name
			ELSE				'When the client has a hyphenated surname
				hyphen_location = InStr(client_name, "-")
				first_part = left(client_name, hyphen_location - 1)
				first_part = UCASE(left(first_part, 1)) & LCASE(right(first_part, len(first_part) - 1))
				second_part = right(client_name, len(client_name) - hyphen_location)
				second_part = UCASE(left(second_part, 1)) & LCASE(right(second_part, len(second_part) - 1))
				output_variable = output_variable & " " & first_part & "-" & second_part
			END IF
		END IF
	NEXT
	name_variable = output_variable
end function

function fix_read_data(search_string)
'--- This function fixes data that we are reading from PRISM that includes underscores. The function searches the variable and removes underscores. Then, the fix case function is called to format the string to the correct case & the data is trimmed to remove any excess spaces.
'~~~~~ search_string: the string for the variable to be searched
'===== Keywords: MAXIS, MMIS, PRISM, name, data, fix
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
end function

Function generate_client_list(list_for_dropdown, initial_text)
'--- This function creates a variable formatted for a DropListBox or ComboBox in a dialog to have all the clients on a case as an option.
'~~~~~ list_for_dropdown: the variable to put in the dialog for the list
'~~~~~ initial_text: the words to have in the top position of the list
'===== Keywords: MAXIS, DIALOG, CLIENTS
	memb_row = 5
    list_for_dropdown = initial_text

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next
End Function

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
			worker_county_code = "x1" & two_digit_county_code_variable
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

function get_to_MMIS_session_begin()
'--- This function brings a MMIS user all the way out of MMIS by PF6'ing until the session is terminated.
'===== Keywords: MMIS, PF6
  Do
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
end function

function HH_member_custom_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadscreen last_name, 25, 6, 30
		EMReadscreen first_name, 12, 6, 63
		EMReadscreen mid_initial, 1, 6, 79
		last_name = trim(replace(last_name, "_", "")) & " "
		first_name = trim(replace(first_name, "_", "")) & " "
		mid_initial = replace(mid_initial, "_", "")
		client_string = ref_nbr & last_name & first_name & mid_initial
		client_array = client_array & client_string & "|"
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_array = TRIM(client_array)
	test_array = split(client_array, "|")
	total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

	DIM all_client_array()
	ReDim all_clients_array(total_clients, 1)

	FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
		Interim_array = split(client_array, "|")
		all_clients_array(x, 0) = Interim_array(x)
		all_clients_array(x, 1) = 1
	NEXT

	BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		Text 10, 5, 105, 10, "Household members to look at:"
		FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
			IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		NEXT
		ButtonGroup ButtonPressed
		OkButton 185, 10, 50, 15
		CancelButton 185, 30, 50, 15
	ENDDIALOG
													'runs the dialog that has been dynamically created. Streamlined with new functions.
	Dialog HH_memb_dialog
	If buttonpressed = 0 then stopscript
	check_for_maxis(True)

	HH_member_array = ""

	FOR i = 0 to total_clients
		IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
			IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
				'msgbox all_clients_
				HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
			END IF
		END IF
	NEXT

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
	HH_member_array = SPLIT(HH_member_array, " ")
end function

function is_date_holiday_or_weekend(date_to_review, boolean_variable)
'--- This function reviews a date to determine if it falls on a weekend or Hennepin County holiday
'~~~~~ date_to_review: this should be in the form of a date
'~~~~~ boolean_variable: this returns TRUE if the date is a weekend or holiday, FALSE if it is not
'==== Keywords: MAXIS, dates, boolean
    non_working_day = FALSE
    day_of_week = WeekdayName(WeekDay(date_to_review))
    If day_of_week = "Saturday" OR day_of_week = "Sunday" Then non_working_day = TRUE
    For each holiday in HOLIDAYS_ARRAY
        If holiday = date_to_review Then non_working_day = TRUE
    Next
    boolean_variable = non_working_day
end function

function log_usage_stats_without_closing()
'--- This function allows logging usage stats but then running another script without closing, i.e. DAIL scrubber
'===== Keywords: MAXIS, MMIS, PRISM, statistics
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\DHS-MAXIS-Scripts\Statistics\usage statistics.accdb"

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & "" & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
end function

function MAXIS_background_check()
'--- This function checks to see if a user is in background
'===== Keywords: MAXIS, background
	Do
		call navigate_to_MAXIS_screen("STAT", "SUMM")
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			PF3
			Pause 2
		End if
	Loop until SELF_check <> "SELF"
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

function MAXIS_dialog_navigation()
'--- This function navigates to various panels in MAXIS. You need to name your buttons using the button names in the function.
'===== Keywords: MAXIS, dialog, navigation
	'This part works with the prev/next buttons on several of our dialogs. You need to name your buttons prev_panel_button, next_panel_button, prev_memb_button, and next_memb_button in order to use them.
	EMReadScreen STAT_check, 4, 20, 21
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = prev_memb_button then
			HH_memb_row = HH_memb_row - 1
			EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
			If isnumeric(prev_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen prev_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		ELSEIF ButtonPressed = next_memb_button then
			HH_memb_row = HH_memb_row + 1
			EMReadScreen next_HH_memb, 2, HH_memb_row, 3
			If isnumeric(next_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen next_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		End if
	End if

	'This part takes care of remaining navigation buttons, designed to go to a single panel.
	If ButtonPressed = ABPS_button then call navigate_to_MAXIS_screen("stat", "ABPS")
	If ButtonPressed = ACCI_button then call navigate_to_MAXIS_screen("stat", "ACCI")
	If ButtonPressed = ACCT_button then call navigate_to_MAXIS_screen("stat", "ACCT")
	If ButtonPressed = ADDR_button then call navigate_to_MAXIS_screen("stat", "ADDR")
	If ButtonPressed = ADME_button then call navigate_to_MAXIS_screen("stat", "ADME")
	If ButtonPressed = ALTP_button then call navigate_to_MAXIS_screen("stat", "ALTP")
	If ButtonPressed = AREP_button then call navigate_to_MAXIS_screen("stat", "AREP")
	If ButtonPressed = BILS_button then call navigate_to_MAXIS_screen("stat", "BILS")
	If ButtonPressed = BUDG_button then call navigate_to_MAXIS_screen("stat", "BUDG")
	If ButtonPressed = BUSI_button then call navigate_to_MAXIS_screen("stat", "BUSI")
	If ButtonPressed = CARS_button then call navigate_to_MAXIS_screen("stat", "CARS")
	If ButtonPressed = CASH_button then call navigate_to_MAXIS_screen("stat", "CASH")
	If ButtonPressed = COEX_button then call navigate_to_MAXIS_screen("stat", "COEX")
	If ButtonPressed = DCEX_button then call navigate_to_MAXIS_screen("stat", "DCEX")
	If ButtonPressed = DFLN_button then call navigate_to_MAXIS_screen("stat", "DFLN")
	If ButtonPressed = DIET_button then call navigate_to_MAXIS_screen("stat", "DIET")
	If ButtonPressed = DISA_button then call navigate_to_MAXIS_screen("stat", "DISA")
	If ButtonPressed = DISQ_button then call navigate_to_MAXIS_screen("stat", "DISQ")
	If ButtonPressed = EATS_button then call navigate_to_MAXIS_screen("stat", "EATS")
	If ButtonPressed = ELIG_DWP_button then call navigate_to_MAXIS_screen("elig", "DWP_")
	If ButtonPressed = ELIG_FS_button then call navigate_to_MAXIS_screen("elig", "FS__")
	If ButtonPressed = ELIG_GA_button then call navigate_to_MAXIS_screen("elig", "GA__")
	If ButtonPressed = ELIG_HC_button then call navigate_to_MAXIS_screen("elig", "HC__")
	If ButtonPressed = ELIG_MFIP_button then call navigate_to_MAXIS_screen("elig", "MFIP")
	If ButtonPressed = ELIG_MSA_button then call navigate_to_MAXIS_screen("elig", "MSA_")
	If ButtonPressed = ELIG_WB_button then call navigate_to_MAXIS_screen("elig", "WB__")
	If ButtonPressed = ELIG_GRH_button then call navigate_to_MAXIS_screen("elig", "GRH_")
	If ButtonPressed = EMMA_button then call navigate_to_MAXIS_screen("stat", "EMMA")
	If ButtonPressed = EMPS_button then call navigate_to_MAXIS_screen("stat", "EMPS")
	If ButtonPressed = FACI_button then call navigate_to_MAXIS_screen("stat", "FACI")
	If ButtonPressed = FMED_button then call navigate_to_MAXIS_screen("stat", "FMED")
	If ButtonPressed = HCMI_button then call navigate_to_MAXIS_screen("stat", "HCMI")
	If ButtonPressed = HCRE_button then call navigate_to_MAXIS_screen("stat", "HCRE")
	If ButtonPressed = HEST_button then call navigate_to_MAXIS_screen("stat", "HEST")
	If ButtonPressed = IMIG_button then call navigate_to_MAXIS_screen("stat", "IMIG")
	If ButtonPressed = INSA_button then call navigate_to_MAXIS_screen("stat", "INSA")
	If ButtonPressed = JOBS_button then call navigate_to_MAXIS_screen("stat", "JOBS")
	If ButtonPressed = MEDI_button then call navigate_to_MAXIS_screen("stat", "MEDI")
	If ButtonPressed = MEMB_button then call navigate_to_MAXIS_screen("stat", "MEMB")
	If ButtonPressed = MEMI_button then call navigate_to_MAXIS_screen("stat", "MEMI")
	If ButtonPressed = MMSA_button then call navigate_to_MAXIS_screen("stat", "MMSA")
	If ButtonPressed = MONT_button then call navigate_to_MAXIS_screen("stat", "MONT")
	If ButtonPressed = OTHR_button then call navigate_to_MAXIS_screen("stat", "OTHR")
	If ButtonPressed = PACT_button then call navigate_to_MAXIS_screen("stat", "PACT")
	If ButtonPressed = PARE_button then call navigate_to_MAXIS_screen("stat", "PARE")
	If ButtonPressed = PBEN_button then call navigate_to_MAXIS_screen("stat", "PBEN")
	If ButtonPressed = PDED_button then call navigate_to_MAXIS_screen("stat", "PDED")
	If ButtonPressed = PREG_button then call navigate_to_MAXIS_screen("stat", "PREG")
	If ButtonPressed = PROG_button then call navigate_to_MAXIS_screen("stat", "PROG")
	If ButtonPressed = RBIC_button then call navigate_to_MAXIS_screen("stat", "RBIC")
	If ButtonPressed = REMO_button then call navigate_to_MAXIS_screen("stat", "REMO")
	If ButtonPressed = REST_button then call navigate_to_MAXIS_screen("stat", "REST")
	If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("stat", "REVW")
	If ButtonPressed = SANC_button then call navigate_to_MAXIS_screen("stat", "SANC")
	If ButtonPressed = SCHL_button then call navigate_to_MAXIS_screen("stat", "SCHL")
	If ButtonPressed = SECU_button then call navigate_to_MAXIS_screen("stat", "SECU")
	If ButtonPressed = SHEL_button then call navigate_to_MAXIS_screen("stat", "SHEL")
	If ButtonPressed = SIBL_button then call navigate_to_MAXIS_screen("stat", "SIBL")
	If ButtonPressed = SPON_button then call navigate_to_MAXIS_screen("stat", "SPON")
    If ButtonPressed = SSRT_button then call navigate_to_MAXIS_screen("stat", "SSRT")
	If ButtonPressed = STEC_button then call navigate_to_MAXIS_screen("stat", "STEC")
	If ButtonPressed = STIN_button then call navigate_to_MAXIS_screen("stat", "STIN")
	If ButtonPressed = STWK_button then call navigate_to_MAXIS_screen("stat", "STWK")
	If ButtonPressed = SWKR_button then call navigate_to_MAXIS_screen("stat", "SWKR")
	If ButtonPressed = TIME_button then call navigate_to_MAXIS_screen("stat", "TIME")
	If ButtonPressed = TRAN_button then call navigate_to_MAXIS_screen("stat", "TRAN")
	If ButtonPressed = TYPE_button then call navigate_to_MAXIS_screen("stat", "TYPE")
	If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("stat", "UNEA")
    If ButtonPressed = WKEX_button then call navigate_to_MAXIS_screen("stat", "WKEX")
	If ButtonPressed = WREG_button then call navigate_to_MAXIS_screen("stat", "WREG")
end function

function MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'--- This function finds the MAXIS footer month/year for MAXIS cases in SELF, MAXIS panels or MEMO screens.
'~~~~~ MAXIS_footer_month: needs to be <code>MAXIS_footer_month</code>
'~~~~~ MAXIS_footer_year: needs to be <code>MAXIS_footer_year</code>
'===== Keywords: MAXIS, footer, month, year
	EMReadScreen SELF_check, 4, 2, 50
	IF SELF_check = "SELF" THEN
		EMReadScreen MAXIS_footer_month, 2, 20, 43
		EMReadScreen MAXIS_footer_year, 2, 20, 46
	ELSE
		EMReadScreen MEMO_check, 4, 2, 47
		IF MEMO_check = "MEMO" Then
			EMReadScreen MAXIS_footer_month, 2, 19, 54
			EMReadScreen MAXIS_footer_year, 2, 49, 57
		ELSE
			Call find_variable("Month: ", MAXIS_footer, 5)
			MAXIS_footer_month = left(MAXIS_footer, 2)
			MAXIS_footer_year = right(MAXIS_footer, 2)
		END IF
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

function MMIS_RKEY_finder()
'--- This function finds the 'RKEY' screen in MMIS
'===== Keywords: MMIS, find, panel
  Do	  						'Now we use a Do Loop to get to the start screen for MMIS.
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
  'Now we get back into MMIS. We have to skip past the intro screens.
  EMWriteScreen "mw00", 1, 2
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  'This section may not work for all OSAs, since some only have EK01. This will find EK01 and enter it.
  MMIS_row = 1
  MMIS_col = 1
  EMSearch "EK01", MMIS_row, MMIS_col
  If MMIS_row <> 0 then
    EMWriteScreen "x", MMIS_row, 4
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
  'This section starts from EK01. OSAs may need to skip the previous section.
  EMWriteScreen "x", 10, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function month_change(interval, starting_month, starting_year, result_month, result_year)
'--- This function may be deleted soon. Waiting for feedback from scriptwriters.
'~~~~~ interval: numeric amount of intervals
'~~~~~ starting_month: month to start
'~~~~~ starting_year: year to start
'~~~~~ result_month: This should be 'result_month'...maybe
'~~~~~ result_year: This should be 'result_year'...maybe
'===== Keywords: MAXIS, month, year, change
	result_month = abs(starting_month)
	result_year = abs(starting_year)
	valid_month = FALSE
	IF result_month = 1 OR result_month = 2 OR result_month = 3 OR result_month = 4 OR result_month = 5 OR result_month = 6 OR result_month = 7 OR result_month = 8 OR result_month = 9 OR result_month = 10 OR result_month = 11 OR result_month = 12 Then valid_month = TRUE
	If valid_month = FALSE Then
		Month_Input_Error_Msg = MsgBox("The month to start from is not a number between 1 and 12, these are the only valid entries for this function. Your data will have the wrong month." & vbnewline & "The month input was: " & result_month & vbnewline & vbnewline & "Do you wish to continue?", vbYesNo + vbSystemModal, "Input Error")
		If Month_Input_Error_Msg = VBNo Then script_end_procedure("")
	End If
	Do
		If left(interval, 1) = "-" Then
			result_month = result_month - 1
			If result_month = 0 then
				result_month = 12
				result_year = result_year - 1
			End If
			interval = interval + 1
		Else
			result_month = result_month + 1
			If result_month = 13 then
				result_month = 1
				result_year = result_year + 1
			End if
			interval = interval - 1
		End If
	Loop until interval = 0
	result_month = right("00" & result_month, 2)
	result_year = right(result_year, 2)
end function

function navigate_to_MAXIS(maxis_mode)
'--- This function is to be used when navigating back to MAXIS from another function in BlueZone (MMIS, PRISM, INFOPAC, etc.)
'~~~~~ maxis_mode: This parameter needs to be "maxis_mode"
'===== Keywords: MAXIS, navigate
    EMWaitReady 0, 0
    attn
    EMWaitReady 0, 0
	EMConnect "A"
    EMWaitReady 0, 0

	IF maxis_mode = "PRODUCTION" THEN
		EMReadScreen prod_running, 7, 6, 15
		IF prod_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect"B"
            EMWaitReady 0, 0
			attn
            EMWaitReady 0, 0
			EMReadScreen prod_running, 7, 6, 15
			IF prod_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than S2.")
			END IF
		END IF
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMReadScreen inq_running, 7, 7, 15
		IF inq_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect "B"
            EMWaitReady 0, 0
			attn
            EMWaitReady 0, 0
			EMReadScreen inq_running, 7, 7, 15
			IF inq_running = "RUNNING" THEN
				x = "B"
			ELSE
				script_end_procedure("Please do not run this script in a session larger than 2.")
			END IF
		END IF
	END IF

    EMWaitReady 0, 0
	EMConnect (x)
    EMWaitReady 0, 0
	IF maxis_mode = "PRODUCTION" THEN
		EMWriteScreen "1", 2, 15
		transmit
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMWriteScreen "2", 2, 15
		transmit
	END IF
end function

function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
'--- This function is to be used to navigate to a specific MAXIS screen
'~~~~~ function_to_go_to: needs to be MAXIS function like "STAT" or "REPT"
'~~~~~ command_to_go_to: needs to be MAXIS function like "WREG" or "ACTV"
'===== Keywords: MAXIS, navigate
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
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
    End if
    If current_case_number = MAXIS_case_number and MAXIS_function = ucase(function_to_go_to) and STAT_note_check <> "NOTE" then
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
  End if
end function

function navigate_to_MMIS()
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
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

	'The following will select the correct version of MMIS. First it looks for C302, then EK01, then C402.
	row = 1
	col = 1
	EMSearch ("C3" & right(worker_county_code, 2)), row, col
	If row <> 0 then
		If row <> 1 then 'It has to do this in case the worker only has one option (as many LTC and OSA workers don't have the option to decide between MAXIS and MCRE case access). The MMIS screen will show the text, but it's in the first row in these instances.
			EMWriteScreen "x", row, 4
			transmit
		End if
	Else 'Some staff may only have EK01 (MMIS MCRE). The script will allow workers to use that if applicable.
		row = 1
		col = 1
		EMSearch "EK01", row, col
		If row <> 0 then
			If row <> 1 then
				EMWriteScreen "x", row, 4
				transmit
			End if
		Else 'Some OSAs have C402 (limited access). This will search for that.
			row = 1
			col = 1
			EMSearch ("C4" & right(worker_county_code, 2)), row, col
			If row <> 0 then
				If row <> 1 then
					EMWriteScreen "x", row, 4
					transmit
				End if
			Else 'Some OSAs have EKIQ (limited MCRE access). This will search for that.
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 then
					If row <> 1 then
						EMWriteScreen "x", row, 4
						transmit
					End if
				Else
					script_end_procedure("C4" & right(worker_county_code, 2) & ", C3" & right(worker_county_code, 2) & ", EKIQ, or EK01 not found. Your access to MMIS may be limited. Contact your script Alpha user if you have questions about using this script.")
				End if
			End if
		End if
	END IF

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit
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
        EMWriteScreen "x", row, col - 3
        transmit

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		End Select
	End If
end function

function navigate_to_PRISM_screen(x)
'--- This function is to be used to navigate to a specific PRISM screen
'~~~~~ x: name of the PRISM screen
'===== Keywords: PRISM, navigate
  EMWriteScreen x, 21, 18
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function

function open_URL_in_browser(URL_to_open)
'--- This function is to be used to open a URL in user's default browser
'~~~~~ URL_to_open: web address to open
'===== Keywords: URL, open, web
	CreateObject("WScript.Shell").Run(URL_to_open)
end function

function PF1()
'--- This function sends or hits the PF1 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
end function

function PF2()
'--- This function sends or hits the PF2 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
end function

function PF3()
'--- This function sends or hits the PF3 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

function PF4()
'--- This function sends or hits the PF4 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
end function

function PF5()
'--- This function sends or hits the PF5 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
end function

function PF6()
'--- This function sends or hits the PF6 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
end function

function PF7()
'--- This function sends or hits the PF7 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
end function

function PF8()
'--- This function sends or hits the PF8 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9()
'--- This function sends or hits the PF9 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10()
'--- This function sends or hits the PF10 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

function PF11()
'--- This function sends or hits the PF11 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
end function

function PF12()
'--- This function sends or hits the PF12 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
end function

function PF13()
'--- This function sends or hits the PF13 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF13
  EMSendKey "<PF13>"
  EMWaitReady 0, 0
end function

function PF14()
'--- This function sends or hits the PF14 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF14
  EMSendKey "<PF14>"
  EMWaitReady 0, 0
end function

function PF15()
'--- This function sends or hits the PF15 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF15
  EMSendKey "<PF15>"
  EMWaitReady 0, 0
end function

function PF16()
'--- This function sends or hits the PF16 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF16
  EMSendKey "<PF16>"
  EMWaitReady 0, 0
end function

function PF17()
'--- This function sends or hits the PF17 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF17
  EMSendKey "<PF17>"
  EMWaitReady 0, 0
end function

function PF18()
'--- This function sends or hits the PF18 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF18
  EMSendKey "<PF18>"
  EMWaitReady 0, 0
end function

function PF19()
'--- This function sends or hits the PF19 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
end function

function PF20()
'--- This function sends or hits the PF20 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function PF21()
'--- This function sends or hits the PF21 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF21
  EMSendKey "<PF21>"
  EMWaitReady 0, 0
end function

function PF22()
'--- This function sends or hits the PF22 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF22
  EMSendKey "<PF22>"
  EMWaitReady 0, 0
end function

function PF23()
'--- This function sends or hits the PF23 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF23
  EMSendKey "<PF23>"
  EMWaitReady 0, 0
end function

function PF24()
'--- This function sends or hits the PF24 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF24
  EMSendKey "<PF24>"
  EMWaitReady 0, 0
end function

function PRISM_case_number_finder(variable_for_PRISM_case_number)
'--- This function finds the PRISM case number if listed on a PRISM screen
'~~~~~ variable_for_PRISM_case_number: this should be 'PRISM_case_number'
'===== Keywords: PRISM, case number
	PRISM_row = 1 'Searches for the case number.
	PRISM_col = 1
	EMSearch "Case: ", PRISM_row, PRISM_col
	If PRISM_row <> 0 then
		EMReadScreen variable_for_PRISM_case_number, 13, PRISM_row, PRISM_col + 6
		variable_for_PRISM_case_number = replace(variable_for_PRISM_case_number, " ", "-")
	Else	'Searches again if not found, this time for "Case/Person"
		PRISM_row = 1
		PRISM_col = 1
		EMSearch "Case/Person: ", PRISM_row, PRISM_col
		If PRISM_row <> 0 then
			EMReadScreen variable_for_PRISM_case_number, 13, PRISM_row, PRISM_col + 13
			variable_for_PRISM_case_number = replace(variable_for_PRISM_case_number, " ", "-")
		End if
	End if
	If isnumeric(left(variable_for_PRISM_case_number, 10)) = False or isnumeric(right(variable_for_PRISM_case_number, 2)) = False then variable_for_PRISM_case_number = ""
end function

FUNCTION PRISM_case_number_validation(case_number_to_validate, outcome)
'--- This function finds the PRISM case number if listed on a PRISM screen
'~~~~~ case_number_to_validate: needs to be 'PRISM_case_number'
'~~~~~ outcome: needs to be 'outcome'
'===== Keywords: PRISM, case number
case_number_to_validate = trim(case_number_to_validate)													'remove any spaces from the beginning or end of the case #
  IF Len(case_number_to_validate) <> 13 THEN 																		'if the case # is not 13 characters, it's not a valid case #
    outcome = False
	ELSEIF IsNumeric(Left(case_number_to_validate, 10)) = False THEN 							'if the first 10 digits of the case # are not numeric, it's not a valid case #
    outcome = False
	ELSEIF IsNumeric(Right(case_number_to_validate, 2)) = False THEN 							'if the last 2 digits of the case # are not numeric, it's not a valid case #
    outcome = False
	ELSEIF IsNumeric(Mid(case_number_to_validate, 11, 1)) = True THEN 						'if the 11th char is a number, it's not a valid case #
		outcome = False
	ELSEIF IsNumeric(Mid(case_number_to_validate, 11, 1)) = False THEN 						'if the 11th char is not a number, then...
		IF Mid(case_number_to_validate, 11, 1) = "." THEN														'if the 11th char is a period, it's a valid case #, and replace the period with a dash
			case_number_to_validate = Replace(case_number_to_validate, ".", "-")
			outcome = True
		ELSEIF Mid(case_number_to_validate, 11, 1) = " " THEN												'if the 11th char is a space, it's a valid case #, and replace the space with a dash
			case_number_to_validate = Replace(case_number_to_validate, " ", "-")
			outcome = True
		ELSEIF Mid(case_number_to_validate, 11, 1) = "-" THEN												'if the 11th char is a dash, it's a valid case #
			outcome = True
		ELSE																																				'if the 11th char is a non-numeric char but is not a period, space, or dash, it's not a valid case #
			outcome = False
		END IF
	ELSE																																					'if we haven't determined if it's a valid case # or not yet, it's a valid case #
    outcome = True
  END IF
END FUNCTION

function proceed_confirmation(result_of_msgbox)
'--- This function asks the user if they want to proceed.
'~~~~~ result_of_msgbox: returns TRUE if Yes is pressed, and FALSE if No is pressed.
'===== Keywords: MAXIS, MMIS, PRISM, dialog, proceed, confirmation
	If ButtonPressed = -1 then
		proceed_confirm = MsgBox("Are you sure you want to proceed? Press Yes to continue, No to return to the previous screen, and Cancel to end the script.", vbYesNoCancel)
		If proceed_confirm = vbCancel then stopscript
		If proceed_confirm = vbYes then result_of_msgbox = TRUE
		If proceed_confirm = vbNo then result_of_msgbox = FALSE
	End if
end function

function regl()
'--- This function clears out PRISM global variables
'===== Keywords: PRISM, REGL, global variables, clear
	EMWriteScreen "REGL", 21, 18		'This writes REGL to the command line
	transmit							'Sends the REGL command
	transmit							'Transmits past the REGL screen
end function

function run_another_script(script_path)
'--- This function runs another script from a specific file either stored locally or on the web.
'~~~~~ script_path: path of script to run
'===== Keywords: MAXIS, MMIS, PRISM, run, script, script path
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  ExecuteGlobal text_from_the_other_script
end function

function run_from_GitHub(url)
'--- This function runs a script file from GitHub
'~~~~~ url: web address of script file on GitHub
'===== Keywords: MAXIS, MMIS, PRISM, run, script, url, GitHub
	'Creates a list of items to remove from anything run from GitHub. This will allow for counties to use Option Explicit handling without fear.
	list_of_things_to_remove = array("OPTION EXPLICIT", _
									"option explicit", _
									"Option Explicit", _
									"dim case_number", _
									"DIM case_number", _
									"Dim case_number")
	If run_locally = "" or run_locally = False then					'Runs the script from GitHub if we're not set up to run locally.
		Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
		req.open "GET", url, False									'Attempts to open the URL
		req.send													'Sends request
		If req.Status = 200 Then									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			script_contents = req.responseText						'Empties the response into a variable called script_contents
			'Uses a for/next to remove the list_of_things_to_remove
			FOR EACH phrase IN list_of_things_to_remove
				script_contents = replace(script_contents, phrase, "")
			NEXT
			ExecuteGlobal script_contents							'Executes the remaining script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
											"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
											vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
			script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		call run_another_script(url)
	END IF
end function

function script_end_procedure(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then MsgBox closing_message '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Determining if the script was successful
        If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
            SCRIPT_success = -1
        else
            SCRIPT_success = 0
        end if

		'Determines if the value of the MAXIS case number - BULK and UTILITIES scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" or left(name_of_script, 4) = "UTIL" then
			MAXIS_CASE_NUMBER = ""
		End if

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
    		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function


function script_end_procedure_with_error_report(closing_message)
'--- This function is how all user stats are collected when a script ends.
'~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
'===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
	stop_time = timer
    send_error_message = ""
	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
        If testing_run = TRUE Then
            MsgBox(closing_message & vbNewLine & vbNewLine & "Since this script is in testing, please provide feedback")
            send_error_message = vbYes
        Else
            send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Script Run Completed")
        End If
    End If
    script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork")
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

        'Determining if the script was successful
        If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
            SCRIPT_success = -1
        else
            SCRIPT_success = 0
        end if

		'Determines if the value of the MAXIS case number - BULK and UTILITIES scripts will not have case number informaiton input into the database
		IF left(name_of_script, 4) = "BULK" or left(name_of_script, 4) = "UTIL" then
			MAXIS_CASE_NUMBER = ""
		End if

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
    		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
		'collecting case numbers counties
		Elseif collect_MAXIS_case_number = true then
			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
		 'for users of the new db
		Else
            objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
            "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
        End if

		'Closing the connection
		objConnection.Close
	End if

    If send_error_message = vbYes Then
        'dialog here to gather more detail
        error_type = ""
        If testing_run = TRUE Then error_type = "TESTING RESPONSE"

        If trim(MAXIS_case_number) = "" Then
            If trim(MMIS_case_number) <> "" Then MAXIS_case_number = MMIS_case_number
        End If

        Do
            Do
                confirm_err = ""

                BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
                  Text 60, 35, 55, 10, MAXIS_case_number
                  ComboBox 220, 30, 175, 45, error_type+chr(9)+"BUG - something happened that was wrong"+chr(9)+"ENHANCEMENT - something could be done better"+chr(9)+"TYPO - grammatical/spelling type errors"+chr(9)+"DAIL - add support for this DAIL message.", error_type
                  EditBox 65, 50, 330, 15, error_detail
                  CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
                  CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
                  CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
                  CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
                  CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
                  CheckBox 95, 115, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
                  CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
                  CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
                  EditBox 60, 140, 165, 15, worker_signature
                  ButtonGroup ButtonPressed
                    OkButton 290, 140, 50, 15
                    CancelButton 345, 140, 50, 15
                  Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
                  Text 5, 35, 50, 10, "Case Number:"
                  Text 125, 35, 95, 10, "What type of error occured?"
                  Text 5, 55, 60, 10, "Explain in detail:"
                  GroupBox 10, 75, 380, 60, "Common areas of issue"
                  Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
                  Text 10, 145, 50, 10, "Worker Name:"
                  Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
                EndDialog

                Dialog Dialog1

                If ButtonPressed = 0 Then
                    cancel_confirm_msg = MsgBox("An Error Report will NOT be sent as you pressed 'Cancel'." & vbNewLine & vbNewLine & "Is this what you would like to do?", vbQuestion + vbYesNo, "Confirm Cancel")
                    If cancel_confirm_msg = vbYes Then confirm_err = ""
                    If cancel_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If

                If ButtonPressed = -1 Then
                    full_text = "Error occurred on " & date & " at " & time
                    full_text = full_text & vbCr & "Error type - " & error_type
                    full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
                    full_text = full_text & vbCr & "Information: " & error_detail
                    If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

                    If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
                    If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
                    If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
                    If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
                    If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
                    If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
                    If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
                    If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

                    full_text = full_text & vbCr & "Closing message: " & closing_message
                    full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

                    send_confirm_msg = MsgBox("** This is what will be sent as an email to the BlueZone Script team:" & vbNewLine & vbNewLine & full_text & vbNewLine & vbNewLine & "*** Is this what you want to send? ***", vbQuestion + vbYesNo, "Confirm Error Report")

                    If send_confirm_msg = vbYes Then confirm_err = ""
                    If send_confirm_msg = vbNo Then confirm_err = "LOOP" & vbNewLine & confirm_err
                End If
            Loop until confirm_err = ""
            call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
        LOOP UNTIL are_we_passworded_out = false
        'sent email here
        full_text = ""
        If ButtonPressed = -1 Then
            bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
            subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"

            full_text = "Error occurred on " & date & " at " & time
            full_text = full_text & vbCr & "Error type - " & error_type
            full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
            full_text = full_text & vbCr & "Information: " & error_detail
            If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"

            If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
            If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
            If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
            If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
            If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
            If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
            If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
            If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"

            full_text = full_text & vbCr & "Closing message: " & closing_message
            full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature

            If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown

            Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)

            MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
        Else
            MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
        End If
    End If
	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
end function


function select_cso_caseload(ButtonPressed, cso_id, cso_name)
'--- This function is helpful for bulk scripts. This script is used to select the caseload by the 8 digit worker ID code entered in the dialog.
'~~~~~ ButtonPressed: should be 'ButtonPressed
'~~~~~ cso_id: should be 'cso_id'
'~~~~~ cso_name: should be 'cso_name'
'===== Keywords: PRISM, cso, select, caseload, BULK
	DO
		DO
			CALL navigate_to_PRISM_screen("USWT")
			err_msg = ""
			'Grabbing the CSO name for the intro dialog.
			CALL find_variable("Worker Id: ", cso_id, 8)
			EMSetCursor 20, 13
			PF1
			CALL write_value_and_transmit(cso_id, 20, 35)
			EMReadScreen cso_name, 24, 13, 55
			cso_name = trim(cso_name)
			PF3

			BeginDialog select_cso_dlg, 0, 0, 286, 145, " - Select CSO Caseload"
			EditBox 70, 55, 65, 15, cso_id
			Text 70, 80, 155, 10, cso_name
			ButtonGroup ButtonPressed
				OkButton 130, 125, 50, 15
				PushButton 180, 125, 50, 15, "UPDATE CSO", update_cso_button
				PushButton 230, 125, 50, 15, "STOP SCRIPT", stop_script_button
			Text 10, 15, 265, 30, "This script will check for worklist items coded E0014 for the following Worker ID. If you wish to change the Worker ID, enter the desired Worker ID in the box and press UPDATE CSO. When you are ready to continue, press OK."
			Text 10, 60, 50, 10, "Worker ID:"
			Text 10, 80, 55, 10, "Worker Name:"

			EndDialog

			DIALOG select_cso_dlg
				IF ButtonPressed = stop_script_button THEN script_end_procedure("The script has stopped.")
				IF ButtonPressed = update_cso_button THEN
					CALL navigate_to_PRISM_screen("USWT")
					CALL write_value_and_transmit(cso_id, 20, 13)
					EMReadScreen cso_name, 24, 13, 55
					cso_name = trim(cso_name)
				END IF
				IF cso_id = "" THEN err_msg = err_msg & vbCr & "* You must enter a Worker ID."
				IF len(cso_id) <> 8 THEN err_msg = err_msg & vbCr & "* You must enter a valid, 8-digit Worker ID."
																																				'The additional of IF ButtonPressed = -1 to the conditional statement is needed
																																		'to allow the worker to update the CSO's worker ID without getting a warning message.
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1
	LOOP UNTIL err_msg = ""
end function

function send_dord_doc(recipient, dord_doc)
'--- This function adds the document.  Some user involvement (resolving required labels, hard-copy printing) may be required.
'~~~~~ recipient: the recipient code from the DORD screen
'~~~~~ dord_doc: document code (also from the DORD screen)
'===== Keywords: PRISM, cso, select, caseload, BULK
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen dord_doc, 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
end function

function start_a_blank_CASE_NOTE()
'--- This function navigates user to a blank case note, presses PF9, and checks to make sure you're in edit mode (keeping you from writing all of the case note on an inquiry screen).
'===== Keywords: MAXIS, case note, navigate, edit
	call navigate_to_MAXIS_screen("case", "note")
	DO
		PF9
		EMReadScreen case_note_check, 17, 2, 33
		EMReadScreen mode_check, 1, 20, 09
		If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then
            'msgbox "The script can't open a case note. Reasons may include:" & vbnewline & vbnewline & "* You may be in inquiry" & vbnewline & "* You may not have authorization to case note this case (e.g.: out-of-county case)" & vbnewline & vbnewline & "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
            BeginDialog Inquiry_Dialog, 0, 0, 241, 115, "CASE NOTE Cannot be Started"
              ButtonGroup ButtonPressed
                OkButton 185, 95, 50, 15
              Text 10, 10, 190, 10, "The script can't open a case note. Reasons may include:"
              Text 20, 25, 80, 10, "* You may be in inquiry"
              Text 20, 40, 185, 20, "* You may not have authorization to case note this case (e.g.: out-of-county case)"
              Text 5, 70, 225, 20, "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
            EndDialog
            Do
                Dialog Inquiry_Dialog
            Loop until ButtonPressed = -1
        End If
	Loop until (mode_check = "A" or mode_check = "E")
end function

function start_a_new_spec_memo()
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

function validate_MAXIS_case_number(err_msg_variable, list_delimiter)
'--- This function checks the MAXIS_case_number variable to ensure it is present and valid
'~~~~~ err_msg_variable: the variable used in error handling within a dialog's do - loop
'~~~~~ list_delimiter: a single character to put in front of any error message to add to the err_msg_variable
'===== Keywords: MAXIS, PRISM, MMIS, dialogs
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" Then
        err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " A case number is required, enter a case number."
    Else
        If IsNumeric(MAXIS_case_number) = FALSE Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " The case number appears invalid, review the case number and update."
        If len(MAXIS_case_number) > 7 Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " The case number entered is too long, review the case number and update."
    End If
end function

function word_doc_open(doc_location, objWord, objDoc)
'--- This function opens a specific word document.
'~~~~~ doc_location: location of word document
'~~~~~ ObjWord: leave as 'ObjWord'
'~~~~~ objDoc: leave as 'objDoc'
'===== Keywords: MAXIS, PRISM, MMIS, Word
	'Opens Word object
	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True		'We want to see it

	'Opens the specific Word doc
	set objDoc = objWord.Documents.Add(doc_location)
end function

function word_doc_update_field(field_name, variable_for_field, objDoc)
'--- This function updates specific fields on a word document
'~~~~~ field_name: name of the field to update
'~~~~~ variable_for_field: information to be updated
'~~~~~ objDoc: leave as 'objDoc'
'===== Keywords: MAXIS, PRISM, MMIS, Word
	objDoc.FormFields(field_name).Result = variable_for_field	'Simply enters the Word document field based on these three criteria
end function

function write_bullet_and_variable_in_CAAD(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CAAD notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CAAD note
'===== Keywords: PRISM, bullet, CAAD note
	IF variable <> "" THEN
	  spaces_count = 6	'Temporary just to make it work

	  EMGetCursor row, col
	  EMReadScreen line_check, 2, 15, 2
	  If ((row = 20 and col + (len(bullet)) >= 78) or row = 21) and line_check = "26" then
	    MsgBox "You've run out of room in this case note. The script will now stop."
	    StopScript
	  End if
	  If row = 21 then
	    EMSendKey "<PF8>"
	    EMWaitReady 0, 0
	    EMSetCursor 16, 4
	  End if
	  variable_array = split(variable, " ")
	  EMSendKey "* " & bullet & ": "
	  For each variable_word in variable_array
	    EMGetCursor row, col
	    EMReadScreen line_check, 2, 15, 2
	    If ((row = 20 and col + (len(variable_word)) >= 78) or row = 21) and line_check = "26" then
	      MsgBox "You've run out of room in this case note. The script will now stop."
	      StopScript
	    End if
	    If (row = 20 and col + (len(variable_word)) >= 78) or (row = 16 and col = 4) or row = 21 then
	      EMSendKey "<PF8>"
	      EMWaitReady 0, 0
	      EMSetCursor 16, 4
	    End if
	    EMGetCursor row, col
	    If (row < 20 and col + (len(variable_word)) >= 78) then EMSendKey "<newline>" & space(spaces_count)
	'    If (row = 16 and col = 4) then EMSendKey space(spaces_count)		'<<<REPLACED WITH BELOW IN ORDER TO TEST column issue
	    If (col = 4) then EMSendKey space(spaces_count)
	    EMSendKey variable_word & " "
	    If right(variable_word, 1) = ";" then
	      EMSendKey "<backspace>" & "<backspace>"
	      EMGetCursor row, col
	      If row = 20 then
	        EMSendKey "<PF8>"
	        EMWaitReady 0, 0
	        EMSetCursor 16, 4
	        EMSendKey space(spaces_count)
	      Else
	        EMSendKey "<newline>" & space(spaces_count)
	      End if
	    End if
	  Next
	  EMSendKey "<newline>"
	  EMGetCursor row, col
	  If (row = 20 and col + (len(bullet)) >= 78) or (row = 16 and col = 4) then
	    EMSendKey "<PF8>"
	    EMWaitReady 0, 0
	    EMSetCursor 16, 4
	  End if
	END IF
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

function write_bullet_and_variable_in_CCOL_NOTE(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CCOL notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CCOL note
'===== Keywords: MAXIS, bullet, CCOL note
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

function write_date(date_variable, date_format_variable, screen_row, screen_col)
'--- This function will write a date in any format desired.
'~~~~~ date_variable: date to write
'~~~~~ date_format_variable: format of date. this should be a string with the correct spaces between month/day/year examples: MM DD YY, MM YY, MM  DD  YYYY
'~~~~~ screen_row: row to write date
'~~~~~ screen_col: column to write date
'===== Keywords: MAXIS, MMIS, PRISM, date, format
    'Figures out the format of the month. If it was "MM", "M", or not present.
    If instr(ucase(date_format_variable), "MM") <> 0 then
        month_format = "MM"
        month_position = instr(ucase(date_format_variable), "MM")
    Elseif instr(ucase(date_format_variable), "M") <> 0 then
        month_format = "M"
        month_position = instr(ucase(date_format_variable), "M")
    Else
        month_format = ""
        month_position = 0
    End if

    'Figures out the format of the day. If it was "DD", "D", or not present.
    If instr(ucase(date_format_variable), "DD") <> 0 then
        day_format = "DD"
        day_position = instr(ucase(date_format_variable), "DD")
    Elseif instr(ucase(date_format_variable), "D") <> 0 then
        day_format = "D"
        day_position = instr(ucase(date_format_variable), "D")
    Else
        day_format = ""
        day_position = 0
    End if

    'Figures out the format of the year. If it was "YYYY", "YY", or not present.
    If instr(ucase(date_format_variable), "YYYY") <> 0 then
        year_format = "YYYY"
        year_position = instr(ucase(date_format_variable), "YYYY")
    Elseif instr(ucase(date_format_variable), "YY") <> 0 then
        year_format = "YY"
        year_position = instr(ucase(date_format_variable), "YY")
    Else
        year_format = ""
        year_position = 0
    End if

    'Formats the month. Separates the month into its own variable and adds a zero if needed.
    var_month = datepart("m", date_variable)
    IF len(var_month) = 1 and month_format = "MM" THEN var_month = "0" & var_month

    'Formats the day. Separates the day into its own variable and adds a zero if needed.
    var_day = datepart("d", date_variable)
    IF len(var_day) = 1 and day_format = "DD" THEN var_day = "0" & var_day

    'Formats the year based on "YY" or "YYYY" formatting.
    If year_format = "YY" then
        var_year = right(datepart("yyyy", date_variable), 2)
    ElseIf year_format = "YYYY" then
        var_year = datepart("yyyy", date_variable)
    END IF

    If month_position <> 0 Then EMWriteScreen var_month, screen_row, screen_col + month_position - 1
    If day_position <> 0 Then EMWriteScreen var_day, screen_row, screen_col + day_position - 1
    If year_position <> 0 Then EMWriteScreen var_year, screen_row, screen_col + year_position - 1
end function

function write_MAXIS_info_to_ES_database(ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive, insert_string)
'--- This function will open the ES_statistics database, check for an existing case and edit it with new info, or add a new entry if there is no existing case in the database.
'~~~~~ dESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive, insert_string: all required parameters from script to be inputted into database
'===== Keywords: MAXIS, statistics, ES
	info_array = array(ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive)
	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'Opening DB
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & ES_database_path
		'This looks for an existing case number and edits it if needed
	set rs = objConnection.Execute("SELECT * FROM ESTrackingTbl WHERE ESCaseNbr = " & ESCaseNbr & " AND ESMembNbr = " & ESMembNbr & "") 'pulling all existing case / member info into a recordset

	IF NOT(rs.EOF) THEN 'There is an existing case, we need to update
		'we don't want to overwrite existing data that isn't updated by the script,
		'the following IF/THENs assign variables to the value from the recordset/database for variables that are empty in the script, and if already null in database,
		'set to "null" for inclusion in sql string.  Also appending quotes / hashtags for string / date variables.
		IF ESCaseNbr = "" THEN ESCaseNbr = rs("ESCaseNbr") 'no null setting, should never happen, but just in case we do not want to ever overwrite a case number / member number
		IF ESMembNbr = "" THEN ESMembNbr = rs("ESMembNbr")
		IF ESMembName <> "" THEN
			ESMembName = "'" & ESMembName & "'"
		ELSE
			ESMembName = "'" & rs("ESMembName") & "'"
			IF IsNull(rs("ESMembName")) = true THEN ESMembName = "null"
		END IF
		IF ESSanctionPercentage = "" THEN
			ESSanctionPercentage = rs("ESSanctionPercentage")
			IF IsNull(rs("ESSanctionPercentage")) = true THEN ESSanctionPercentage = "null"
		END IF
		IF ESEmpsStatus = "" THEN
			ESEmpsStatus = rs("ESEmpsStatus")
			IF IsNull(rs("ESEmpsStatus")) = true THEN ESEmpsStatus = "null"
		END IF
		IF ESTANFMosUsed = "" THEN
			ESTANFMosUsed = rs("ESTANFMosUsed")
			IF ISNull(rs("ESTANFMosUsed")) = true THEN ESTANFMosUsed = "null"
		END IF
		IF ESExtensionReason = "" THEN
			ESExtensionReason = rs("ESExtensionReason")
			IF IsNull(rs("ESExtensionReason")) = true THEN ESExtensionReason = "null"
		END IF
		IF IsDate(ESDisaEnd) = TRUE THEN
			ESDisaEnd = "#" & ESDisaEnd & "#"
		ELSE
			IF ESDisaEnd = "" THEN ESDisaEnd = "#" & rs("ESDisaEnd") & "#"
			IF IsNull(rs("ESDisaEnd")) = true THEN ESDisaEnd = "null"
		END IF
		IF ESPrimaryActivity <> "" THEN
			ESPrimaryActivity = "'" & ESPrimaryActivity & "'"
		ELSE
			ESPrimaryActivity = "'" & rs("ESPrimaryActivity") & "'"
			IF IsNull(rs("ESPrimaryActivity")) = true THEN ESPrimaryActivity = "null"
		END IF
		IF IsDate(ESDate) = True THEN
			ESDate = "#" & ESDate & "#"
		ELSE
			ESDate = "#" & rs("ESDate") & "#"
			IF IsNull(rs("ESDate")) = true THEN ESDate = "null"
		END IF
		IF ESSite <> "" THEN
			ESSite = "'" & ESSite & "'"
		ELSE
			ESSite = "'" & rs("ESSite") & "'"
			IF IsNull(rs("ESSite")) = true THEN ESSite = "null"
		END IF
		IF ESCounselor <> "" THEN
			ESCounselor = "'" & ESCounselor & "'"
		ELSE
			ESCounselor = "'" & rs("ESCounselor") & "'"
			IF IsNull(rs("ESCounselor")) = true THEN ESCounselor = "null"
		END IF
		IF ESActive <> "" THEN
			ESActive = "'" & ESActive & "'"
		ELSE
			ESActive = "'" & rs("ESActive") & "'"
			IF IsNull(rs("ESActive")) = true THEN ESActive = "null"
		END IF
		'This formats all the variables into the correct syntax
		ES_update_str = "ESMembName = " & ESMembName & ", ESSanctionPercentage = " & ESSanctionPercentage & ", ESEmpsStatus = " & ESEmpsStatus & ", ESTANFMosUsed = " & ESTANFMosUsed &_
				", ESExtensionReason = " & ESExtensionReason & ", ESDisaEnd = " & ESDisaEnd & ", ESPrimaryActivity = " & ESPrimaryActivity & ", ESDate = " & ESDate & ", ESSite = " &_
				ESSite & ", ESCounselor = " & ESCounselor & ", ESActive = " & ESActive & " WHERE ESCaseNbr = " & ESCaseNbr & " AND ESMembNbr = " & ESMembNbr & ""
		objConnection.Execute "UPDATE ESTrackingTbl SET " & ES_update_str 'Here we are actually writing to the database
		objConnection.Close
		set rs = nothing
	ELSE 'There is no existing case, add a new one using the info pulled from the script
		FOR EACH item IN info_array ' THIS loop writes the values string for the SQL statement (with correct syntax for each variable type) to write a NEW RECORD to the database
			IF values_string = "" THEN
				IF item <> "" THEN
					IF isnumeric(item) = true THEN
						values_string = """ " & item & " """
					ELSEIF isdate(item) = true Then
						values_string = " #" & item & "#"
					ELSE
						values_string = "'" & item & "'"
					END IF
				ELSE
					values_string = "null"
				END IF
			ELSE
				IF item <> "" THEN
					IF isnumeric(item) = true THEN
						values_string = values_string & ", "" " & item & " """
					ELSEIF isdate(item) = true THEN
						values_string = values_string & ", #" & item & "#"
					ELSE
						values_string = values_string & ", '" & item & "'"
					END IF
				ELSE
					values_string = values_string & ", null"
				END IF
			END IF

		NEXT
		values_string = values_string & ")"
		'Inserting the new record
		objConnection.Execute "INSERT INTO ESTrackingTbl (ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive) VALUES (" & values_string
		objConnection.Close
	END IF
	'Clearing all variables to avoid writing over records in future calls from same script
	ERASE info_array
	ESMembNbr = ""
	ESMembName = ""
	EsSanctionPercentage = ""
	ESEmpsStatus = ""
	ESTANFMosUsed = ""
	ESExtensionReason = ""
	ESDisaEnd = ""
	ESPrimaryActivity = ""
	ESDate = ""
	ESSite = ""
	ESCounselor = ""
	ESActive = ""
	insert_string = ""
end function

function write_three_columns_in_CASE_NOTE(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
'--- This function writes variables into three seperate columns into case notes
'~~~~~ col_01_start_point: column where to write the 1st variable
'~~~~~ col_01_variable: name of 1st variable to write
'~~~~~ col_02_start_point: column where to write the 2nd variable
'~~~~~ col_02_variable: name of 2nd variable to write
'~~~~~ col_03_start_point: column where to write the 3rd variable
'~~~~~ col_03_variable: name of 3rd variable to write
'===== Keywords: MAXIS, case note, three columns, format
  EMGetCursor row, col
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMGetCursor row, col
  EMWriteScreen "                                                                              ", row, 3
  EMSetCursor row, col_01_start_point
  EMSendKey col_01_variable
  EMSetCursor row, col_02_start_point
  EMSendKey col_02_variable
  EMSetCursor row, col_03_start_point
  EMSendKey col_03_variable
  EMSendKey "<newline>"
  EMGetCursor row, col
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
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

function write_variable_in_CAAD(variable)
'--- This function writes a variable in CAAD note
'~~~~~ variable: information to be entered into CAAD note from script/edit box
'===== Keywords: PRISM, CAAD note
    IF variable <> "" THEN
        EMGetCursor row, col
        EMReadScreen line_check, 2, 15, 2
        If ((row = 20 and col + (len(x)) >= 78) or row = 21) and line_check = "26" then
            MsgBox "You've run out of room in this case note. The script will now stop."
            StopScript
        End if
        If (row = 20 and col + (len(x)) >= 78 + 1 ) or row = 21 then
            EMSendKey "<PF8>"
            EMWaitReady 0, 0
            EMSetCursor 16, 4
        End if
        EMSendKey variable & "<newline>"
        EMGetCursor row, col
        If (row = 20 and col + (len(x)) >= 78) or (row = 21) then
            EMSendKey "<PF8>"
            EMWaitReady 0, 0
            EMSetCursor 16, 4
        End if
    END IF
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

function write_variable_in_CCOL_NOTE(variable)
''--- This function writes a variable in CCOL note
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

function write_variable_in_DORD(string_to_write, recipient)
'--- This function writes a variable in DORD document
'~~~~~ string_to_write: information to be entered into document
'~~~~~ recipient: recipeint of DORD document
'===== Keywords: PRISM, DORD
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0104", 6, 36
	EMWriteScreen recipient, 11, 51
	transmit

	'This function will add a string to DORD docs.
	IF len(string_to_write) > 1080 THEN
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text below is longer than the script can handle in one DORD document. The script will not add the text to the document." & vbCr & vbCr & _
				string_to_write
		EXIT function
	END IF

	dord_rows_of_text = Int(len(string_to_write) / 60) + 1

	ReDim write_array(dord_rows_of_text)
	'Splitting the text
	string_to_write = split(string_to_write)
	array_position = 1
	FOR EACH word IN string_to_write
		IF len(write_array(array_position)) + len(word) <= 60 THEN
			write_array(array_position) = write_array(array_position) & word & " "
		ELSE
			array_position = array_position + 1
			write_array(array_position) = write_array(array_position) & word & " "
		END IF
	NEXT

	PF14

	'Selecting the "U" label type
	CALL write_value_and_transmit("U", 20, 14)

	'Writing the values
	dord_row = 7
	FOR i = 1 TO dord_rows_of_text
		CALL write_value_and_transmit("S", dord_row, 5)
		CALL write_value_and_transmit(write_array(i), 16, 15)

		dord_row = dord_row + 1
		IF i = 12 THEN
			PF8
			dord_row = 7
		END IF
	NEXT
	PF3
	EMWriteScreen "M", 3, 29
	transmit
end function

function write_variable_in_SPEC_MEMO(variable)
'--- This function writes a variable in SPEC/MEMO
'~~~~~ variable: information to be entered into SPEC/MEMO
'===== Keywords: MAXIS, SPEC, MEMO
    EMGetCursor memo_row, memo_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    memo_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    'The following figures out if we need a new page
    Do
        EMReadScreen line_test, 60, memo_row, memo_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
        line_test = trim(line_test)
        'MsgBox line_test
        If line_test <> "" OR memo_row >= 18 Then
            memo_row = memo_row + 1

            'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
            If memo_row >= 18 then
                PF8
                memo_row = 3					'Resets this variable to 3
            End if
        End If

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then script_end_procedure("The WCOM/MEMO area is already full and no additional informaion can be added. This script should be run prior to adding manual wording.")

    Loop until line_test = ""

    'Each word becomes its own member of the array called variable_array.
    variable_array = split(variable, " ")

    For each word in variable_array
        'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
        If len(word) + memo_col > 74 then
            memo_row = memo_row + 1
            memo_col = 15
        End if

        'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
        If memo_row >= 18 then
            PF8
            memo_row = 3					'Resets this variable to 3
        End if

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then
            PF10
            end_msg = "The WCOM/MEMO area is already full and no additional informaion can be added. The wording that was not added and the script ended on is:" & vbNewLine & vbNewLine & variable & vbNewLine & vbNewLine & "**This script should be run prior to adding manual wording.**"
            script_end_procedure(end_msg)
        End If
        'Writes the word and a space using EMWriteScreen
        EMWriteScreen word & " ", memo_row, memo_col

        'Increases memo_col the length of the word + 1 (for the space)
        memo_col = memo_col + (len(word) + 1)
    Next

    'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand).
    EMSetCursor memo_row + 1, 15
end function

function write_variable_in_TIKL(variable)
'--- This function writes a variable in TIKL
'~~~~~ variable: information to be entered into TIKL
'===== Keywords: MAXIS, TIKL
	IF len(variable) <= 60 THEN
		tikl_line_one = variable
	ELSE
		tikl_line_one_len = 61
		tikl_line_one = left(variable, tikl_line_one_len)
		IF right(tikl_line_one, 1) = " " THEN
			whats_left_after_one = right(variable, (len(variable) - tikl_line_one_len))
		ELSE
			DO
				tikl_line_one = left(variable, (tikl_line_one_len - 1))
				IF right(tikl_line_one, 1) <> " " THEN tikl_line_one_len = tikl_line_one_len - 1
			LOOP UNTIL right(tikl_line_one, 1) = " "
			whats_left_after_one = right(variable, (len(variable) - (tikl_line_one_len - 1)))
		END IF
	END IF

	IF (whats_left_after_one <> "" AND len(whats_left_after_one) <= 60) THEN
		tikl_line_two = whats_left_after_one
	ELSEIF (whats_left_after_one <> "" AND len(whats_left_after_one) > 60) THEN
		tikl_line_two_len = 61
		tikl_line_two = left(whats_left_after_one, tikl_line_two_len)
		IF right(tikl_line_two, 1) = " " THEN
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - tikl_line_two_len))
		ELSE
			DO
				tikl_line_two = left(whats_left_after_one, (tikl_line_two_len - 1))
				IF right(tikl_line_two, 1) <> " " THEN tikl_line_two_len = tikl_line_two_len - 1
			LOOP UNTIL right(tikl_line_two, 1) = " "
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - (tikl_line_two_len - 1)))
		END IF
	END IF

	IF (whats_left_after_two <> "" AND len(whats_left_after_two) <= 60) THEN
		tikl_line_three = whats_left_after_two
	ELSEIF (whats_left_after_two <> "" AND len(whats_left_after_two) > 60) THEN
		tikl_line_three_len = 61
		tikl_line_three = right(whats_left_after_two, tikl_line_three_len)
		IF right(tikl_line_three, 1) = " " THEN
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - tikl_line_three_len))
		ELSE
			DO
				tikl_line_three = left(whats_left_after_two, (tikl_line_three_len - 1))
				IF right(tikl_line_three, 1) <> " " THEN tikl_line_three_len = tikl_line_three_len - 1
			LOOP UNTIL right(tikl_line_three, 1) = " "
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - (tikl_line_three_len - 1)))
		END IF
	END IF

	IF (whats_left_after_three <> "" AND len(whats_left_after_three) <= 60) THEN
		tikl_line_four = whats_left_after_three
	ELSEIF (whats_left_after_three <> "" AND len(whats_left_after_three) > 60) THEN
		tikl_line_four_len = 61
		tikl_line_four = left(whats_left_after_three, tikl_line_four_len)
		IF right(tikl_line_four, 1) = " " THEN
			tikl_line_five = right(whats_left_after_three, (len(whats_left_after_three) - tikl_line_four_len))
		ELSE
			DO
				tikl_line_four = left(whats_left_after_three, (tikl_line_four_len - 1))
				IF right(tikl_line_four, 1) <> " " THEN tikl_line_four_len = tikl_line_four_len - 1
			LOOP UNTIL right(tikl_line_four, 1) = " "
			tikl_line_five = right(whats_left_after_three, (tikl_line_four_len - 1))
		END IF
	END IF

	EMWriteScreen tikl_line_one, 9, 3
	IF tikl_line_two <> "" THEN EMWriteScreen tikl_line_two, 10, 3
	IF tikl_line_three <> "" THEN EMWriteScreen tikl_line_three, 11, 3
	IF tikl_line_four <> "" THEN EMWriteScreen tikl_line_four, 12, 3
	IF tikl_line_five <> "" THEN EMWriteScreen tikl_line_five, 13, 3
	transmit
end function


function write_variable_with_indent_in_CASE_NOTE(variable)
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

        indent_length = 5

		'Writes the bullet
		EMWriteScreen "  - ", noting_row, noting_col

		'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + 4

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

'END OF MAIN FUNCTIONS LIBRARY========================================================================================================================================================================================














'Functions for PROJECT KRABAPPEL (UTILITIES - TRAINING CASE CREATOR)====================================================================================================================================================
'writing in placeholder information for now re: the functions and parameters to be in line with the new documentation process.

function write_panel_to_MAXIS_ABPS(abps_supp_coop,abps_gc_status)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ abps_supp_coop,abps_gc_status: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","PARE")							'Starts by creating an array of all the kids on PARE
	EMReadScreen abps_pare_check, 1, 2, 78
	If abps_pare_check = "0" then
		MsgBox "No PARE exists. Exiting Creating ABPS."
	ElseIf abps_pare_check <> "0" then
		child_list = ""
		row = 8
		Do
			EMReadScreen child_check, 2, row, 24
			If child_check <> "__" then
				If child_list = "" then
					child_list = child_check
				ElseIf child_list <> "" then
					child_list = child_list & "," & child_check
				End If
			End If
			row = row + 1
			If row = 18 then
				PF8
				row = 8
			End If
		Loop until child_check = "__"
		call navigate_to_MAXIS_screen("STAT","ABPS")						'Navigates to ABPS to enter kids in
		call create_panel_if_nonexistent
		abps_child_list = split(child_list, ",")
		row = 15
		for each abps_child in abps_child_list
			EMWriteScreen abps_child, row, 35
			EMWriteScreen "2", row, 53
			EMWriteScreen "1", row, 67
			row = row + 1
			If row = 18 then
				PF8
				row = 15
			End If
		next
		IF abps_act_date <> "" THEN call create_MAXIS_friendly_date_with_YYYY(date, 0, 18, 38)
		EMWriteScreen reference_number, 4, 47		'Enters the reference_number
		If abps_supp_coop <> "" then
			abps_supp_coop = ucase(abps_supp_coop)
			abps_supp_coop = left(abps_supp_coop,1)
			EMWriteScreen abps_supp_coop, 4, 73
		End If
		If abps_gc_status <> "" then
			EMWriteScreen abps_gc_status, 5, 47
		End If
		transmit
	End If
end function

function write_panel_to_MAXIS_ACCT(acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "ACCT")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen acct_type, 6, 44  'enters the account type code
	Emwritescreen acct_numb, 7, 44  'enters the account number
	Emwritescreen acct_location, 8, 44  'enters the account location

	' >>>>> Comment: Updated 06/22/2016 <<<<<
	' >>>>> Looking for the acct_bal_ver location. This changed with asset unification...
	' >>>>> ... but the location is not the same across all months. It needs to be variable...
	' >>>>> ... so Krabappel knows where to write stuff and junk or whatever ...
	' >>>>> This has been tested on training case 226398 for the benefit months 05/16 and 06/16...
	' >>>>> ... in 05/16 the acct_bal_ver coordinates are 10, 63 and in 06/16, they are 10, 64...
	' >>>>> ... and the code is working in both months.
	' >>> Looking for the balance field and then we will write the verif code on the same line...
	acct_row = 1
	acct_col = 1
	EMSearch "Balance: ", acct_row, acct_col
	EMWriteScreen acct_balance, acct_row, acct_col + 11  'enters the balance
	acct_col = 1
	EMSearch "Ver: ", acct_row, acct_col
	EMWriteScreen acct_bal_ver, acct_row, acct_col + 5  'enters the balance verification

	IF acct_date <> "" THEN call create_MAXIS_friendly_date(acct_date, 0, 11, 44)  'enters the account balance date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen acct_withdraw, 12, 46  'enters the withdrawl penalty
	Emwritescreen acct_cash_count, 14, 50  'enters y/n if counted for cash
	Emwritescreen acct_snap_count, 14, 57  'enters y/n if counted for snap
	Emwritescreen acct_HC_count, 14, 64  'enters y/n if counted for HC
	Emwritescreen acct_GRH_count, 14, 72  'enters y/n if counted for grh
	Emwritescreen acct_IV_count, 14, 80  'enters y/n if counted for IV
	Emwritescreen acct_joint_owner, 15, 44  'enters if it is a jointly owned acct
	Emwritescreen left(acct_share_ratio, 1), 15, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(acct_share_ratio, 1), 15, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	Emwritescreen acct_interest_date_mo, 17, 57  'enters the next interest date MM format
	Emwritescreen acct_interest_date_yr, 17, 60  'enters the next interest date YY format
	transmit
	transmit
end function

function write_panel_to_MAXIS_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "ACUT")
	call create_panel_if_nonexistent
		EMWritescreen ACUT_shared, 6, 42
		EMWritescreen ACUT_heat, 10, 61
		EMWritescreen ACUT_air, 11, 61
		EMWritescreen ACUT_electric, 12, 61
		EMWritescreen ACUT_fuel, 13, 61
		EMWritescreen ACUT_garbage, 14, 61
		EMWritescreen ACUT_water, 15, 61
		EMWritescreen ACUT_sewer, 16, 61
		EMWritescreen ACUT_other, 17, 61
		EMWritescreen ACUT_heat_verif, 10, 55
		EMWritescreen ACUT_air_verif, 11, 55
		EMWritescreen ACUT_electric_verif, 12, 55
		EMWritescreen ACUT_fuel_verif, 13, 55
		EMWritescreen ACUT_garbage_verif, 14, 55
		EMWritescreen ACUT_water_verif, 15, 55
		EMWritescreen ACUT_sewer_verif, 16, 55
		EMWritescreen ACUT_other_verif, 17, 55
		EMWritescreen Left(ACUT_phone, 1), 18, 55
	transmit
end function

'---This function writes the information for BILS.
function write_panel_to_MAXIS_BILS(bils_1_ref_num, bils_1_serv_date, bils_1_serv_type, bils_1_gross_amt, bils_1_third_party, bils_1_verif, bils_1_bils_type, bils_2_ref_num, bils_2_serv_date, bils_2_serv_type, bils_2_gross_amt, bils_2_third_party, bils_2_verif, bils_2_bils_type, bils_3_ref_num, bils_3_serv_date, bils_3_serv_type, bils_3_gross_amt, bils_3_third_party, bils_3_verif, bils_3_bils_type, bils_4_ref_num, bils_4_serv_date, bils_4_serv_type, bils_4_gross_amt, bils_4_third_party, bils_4_verif, bils_4_bils_type, bils_5_ref_num, bils_5_serv_date, bils_5_serv_type, bils_5_gross_amt, bils_5_third_party, bils_5_verif, bils_5_bils_type, bils_6_ref_num, bils_6_serv_date, bils_6_serv_type, bils_6_gross_amt, bils_6_third_party, bils_6_verif, bils_6_bils_type, bils_7_ref_num, bils_7_serv_date, bils_7_serv_type, bils_7_gross_amt, bils_7_third_party, bils_7_verif, bils_7_bils_type, bils_8_ref_num, bils_8_serv_date, bils_8_serv_type, bils_8_gross_amt, bils_8_third_party, bils_8_verif, bils_8_bils_type, bils_9_ref_num, bils_9_serv_date, bils_9_serv_type, bils_9_gross_amt, bils_9_third_party, bils_9_verif, bils_9_bils_type)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ bils_1_ref_num, bils_1_serv_date, bils_1_serv_type, bils_1_gross_amt, bils_1_third_party, bils_1_verif, bils_1_bils_type, bils_2_ref_num, bils_2_serv_date, bils_2_serv_type, bils_2_gross_amt, bils_2_third_party, bils_2_verif, bils_2_bils_type, bils_3_ref_num, bils_3_serv_date, bils_3_serv_type, bils_3_gross_amt, bils_3_third_party, bils_3_verif, bils_3_bils_type, bils_4_ref_num, bils_4_serv_date, bils_4_serv_type, bils_4_gross_amt, bils_4_third_party, bils_4_verif, bils_4_bils_type, bils_5_ref_num, bils_5_serv_date, bils_5_serv_type, bils_5_gross_amt, bils_5_third_party, bils_5_verif, bils_5_bils_type, bils_6_ref_num, bils_6_serv_date, bils_6_serv_type, bils_6_gross_amt, bils_6_third_party, bils_6_verif, bils_6_bils_type, bils_7_ref_num, bils_7_serv_date, bils_7_serv_type, bils_7_gross_amt, bils_7_third_party, bils_7_verif, bils_7_bils_type, bils_8_ref_num, bils_8_serv_date, bils_8_serv_type, bils_8_gross_amt, bils_8_third_party, bils_8_verif, bils_8_bils_type, bils_9_ref_num, bils_9_serv_date, bils_9_serv_type, bils_9_gross_amt, bils_9_third_party, bils_9_verif, bils_9_bils_type: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "BILS")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMReadScreen num_of_BILS, 1, 2, 78
	IF num_of_BILS = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	'---MAXIS will not allow BILS to be updated if HC is inactive. Exiting the function if HC is inactive.
	EMReadScreen hc_inactive, 21, 24, 2
	IF hc_inactive = "HC STATUS IS INACTIVE" THEN Exit function

	BILS_row = 6
	DO
		EMReadScreen available_row, 2, BILS_row, 26
		IF available_row <> "__" THEN BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	LOOP UNTIL available_row = "__"

	IF bils_1_ref_num <> "" THEN
		IF len(bils_1_ref_num) = 1 THEN bils_1_ref_num = "0" & bils_1_ref_num
		EMWriteScreen bils_1_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_1_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_1_serv_type, BILS_row, 40
		EMWriteScreen bils_1_gross_amt, BILS_row, 45
		EMWriteScreen bils_1_third_party, BILS_row, 57
		IF bils_1_verif = "03" AND bils_1_serv_type <> "22" THEN bils_1_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_1_verif, BILS_row, 67
		EMWriteScreen bils_1_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_2_ref_num <> "" THEN
		IF len(bils_2_ref_num) = 1 THEN bils_2_ref_num = "0" & bils_2_ref_num
		EMWriteScreen bils_2_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_2_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_2_serv_type, BILS_row, 40
		EMWriteScreen bils_2_gross_amt, BILS_row, 45
		EMWriteScreen bils_2_third_party, BILS_row, 57
		IF bils_2_verif = "03" AND bils_2_serv_type <> "22" THEN bils_2_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_2_verif, BILS_row, 67
		EMWriteScreen bils_2_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_3_ref_num <> "" THEN
		IF len(bils_3_ref_num) = 1 THEN bils_3_ref_num = "0" & bils_3_ref_num
		EMWriteScreen bils_3_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_3_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_3_serv_type, BILS_row, 40
		EMWriteScreen bils_3_gross_amt, BILS_row, 45
		EMWriteScreen bils_3_third_party, BILS_row, 57
		IF bils_3_verif = "03" AND bils_3_serv_type <> "22" THEN bils_3_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_3_verif, BILS_row, 67
		EMWriteScreen bils_3_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_4_ref_num <> "" THEN
		IF len(bils_4_ref_num) = 1 THEN bils_4_ref_num = "0" & bils_4_ref_num
		EMWriteScreen bils_4_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_4_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_4_serv_type, BILS_row, 40
		EMWriteScreen bils_4_gross_amt, BILS_row, 45
		EMWriteScreen bils_4_third_party, BILS_row, 57
		IF bils_4_verif = "03" AND bils_4_serv_type <> "22" THEN bils_4_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_4_verif, BILS_row, 67
		EMWriteScreen bils_4_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_5_ref_num <> "" THEN
		IF len(bils_5_ref_num) = 1 THEN bils_5_ref_num = "0" & bils_5_ref_num
		EMWriteScreen bils_5_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_5_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_5_serv_type, BILS_row, 40
		EMWriteScreen bils_5_gross_amt, BILS_row, 45
		EMWriteScreen bils_5_third_party, BILS_row, 57
		IF bils_5_verif = "03" AND bils_5_serv_type <> "22" THEN bils_5_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_5_verif, BILS_row, 67
		EMWriteScreen bils_5_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_6_ref_num <> "" THEN
		IF len(bils_6_ref_num) = 1 THEN bils_6_ref_num = "0" & bils_6_ref_num
		EMWriteScreen bils_6_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_6_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_6_serv_type, BILS_row, 40
		EMWriteScreen bils_6_gross_amt, BILS_row, 45
		EMWriteScreen bils_6_third_party, BILS_row, 57
		IF bils_6_verif = "03" AND bils_6_serv_type <> "22" THEN bils_6_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_6_verif, BILS_row, 67
		EMWriteScreen bils_6_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_7_ref_num <> "" THEN
		IF len(bils_7_ref_num) = 1 THEN bils_7_ref_num = "0" & bils_7_ref_num
		EMWriteScreen bils_7_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_7_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_7_serv_type, BILS_row, 40
		EMWriteScreen bils_7_gross_amt, BILS_row, 45
		EMWriteScreen bils_7_third_party, BILS_row, 57
		IF bils_7_verif = "03" AND bils_7_serv_type <> "22" THEN bils_7_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_7_verif, BILS_row, 67
		EMWriteScreen bils_7_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_8_ref_num <> "" THEN
		IF len(bils_8_ref_num) = 1 THEN bils_8_ref_num = "0" & bils_8_ref_num
		EMWriteScreen bils_8_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_8_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_8_serv_type, BILS_row, 40
		EMWriteScreen bils_8_gross_amt, BILS_row, 45
		EMWriteScreen bils_8_third_party, BILS_row, 57
		IF bils_8_verif = "03" AND bils_8_serv_type <> "22" THEN bils_8_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_8_verif, BILS_row, 67
		EMWriteScreen bils_8_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_9_ref_num <> "" THEN
		IF len(bils_9_ref_num) = 1 THEN bils_9_ref_num = "0" & bils_9_ref_num
		EMWriteScreen bils_9_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_9_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_9_serv_type, BILS_row, 40
		EMWriteScreen bils_9_gross_amt, BILS_row, 45
		EMWriteScreen bils_9_third_party, BILS_row, 57
		IF bils_9_verif = "03" AND bils_9_serv_type <> "22" THEN bils_9_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_9_verif, BILS_row, 67
		EMWriteScreen bils_9_bils_type, BILS_row, 71
	END IF
end function

function write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the busi panel in MAXIS)
'~~~~~ busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "BUSI")  'navigates to the stat panel
	Emwritescreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_BUSI, 1, 2, 78
	IF num_of_BUSI = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit

		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 54)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 71)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 7, 26  'this enters into the gross income calculator
			Transmit
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened.
			LOOP UNTIL busi_gross_income_check = "Gross Income"
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 14, 59  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 14, 73  'enters the prospective hours

		ELSE				'This is the NEW logic for all months after 02/2015
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 55)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 72)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 6, 26  'this enters into the gross income calculator
			Transmit
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened.
			LOOP UNTIL busi_gross_income_check = "Gross Income"
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 13, 60  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 13, 74  'enters the prospective hours
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
		END IF
	ELSEIF num_of_BUSI <> "0" THEN
		PF9
		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) >= 0 then
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
			'---Going into the HC Income Estimate
			EMWriteScreen "X", 17, 27
			transmit
			DO
				EMReadScreen hc_income, 9, 4, 42
			LOOP UNTIL hc_income = "HC Income"
			EMReadScreen current_month_plus_one, 17, 21, 59
			IF current_month_plus_one = "CURRENT MONTH + 1" THEN
				PF3
			ELSE
				Emwritescreen busi_hc_total_est_a, 7, 54                'enters hc total income estimation for method A
				Emwritescreen busi_hc_total_est_b, 8, 54                'enters hc total income estimation for method B
				Emwritescreen busi_hc_exp_est_a, 11, 54                 'enters hc expense estimation for method A
				Emwritescreen busi_hc_exp_est_b, 12, 54                 'enters hc expense estimation for method B
				Emwritescreen busi_hc_hours_est, 18, 58                 'enters hc hours estimation
				transmit
				PF3
			END IF
		END IF
	END IF
end function

function write_panel_to_MAXIS_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "CARS")  'navigates to the stat screen
	call create_panel_if_nonexistent
	Emwritescreen cars_type, 6, 43  'enters the vehicle type
	Emwritescreen cars_year, 8, 31  'enters the vehicle year
	Emwritescreen cars_make, 8, 43  'enters the vehicle make
	Emwritescreen cars_model, 8, 66  'enters the vehicle model
	Emwritescreen cars_trade_in, 9, 45  'enters the trade in value
	Emwritescreen cars_loan, 9, 62  'enters the loan value
	Emwritescreen cars_value_source, 9, 80  'enters the source of value information
	Emwritescreen cars_ownership_ver, 10, 60  'enters the ownership verification code
	Emwritescreen cars_amount_owed, 12, 45  'enters the amount owed on vehicle
	Emwritescreen cars_amount_owed_ver, 12, 60  'enters the amount owed verification code
	IF cars_date <> "" THEN call create_MAXIS_friendly_date(cars_date, 0, 13, 43)  'enters the amounted owed as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen cars_use, 15, 43  'enters the use code for the vehicle
	Emwritescreen cars_HC_benefit, 15, 76  'enters if the vehicle is for client benefit
	Emwritescreen cars_joint_owner, 16, 43  'enters if it is a jointly owned car
	Emwritescreen left(cars_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(cars_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_CASH(cash_amount)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the cash panel in MAXIS)
'~~~~~ cash_amount: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "CASH")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen cash_amount, 8, 39
end function

function write_panel_to_MAXIS_COEX(retro_support, prosp_support, support_verif, retro_alimony, prosp_alimony, alimony_verif, retro_tax_dep, prosp_tax_dep, tax_dep_verif, retro_other, prosp_other, other_verif, change_in_circum, hc_exp_support, hc_exp_alimony, hc_exp_tax_dep, hc_exp_other)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the COEX panel in MAXIS.)
'~~~~~ retro_support, prosp_support, support_verif, retro_alimony, prosp_alimony, alimony_verif, retro_tax_dep, prosp_tax_dep, tax_dep_verif, retro_other, prosp_other, other_verif, change_in_circum, hc_exp_support, hc_exp_alimony, hc_exp_tax_dep, hc_exp_other: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "COEX")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_COEX, 1, 2, 78
	IF num_of_COEX = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
		'---If the script is creating a new COEX panel, it will enter this information...
		EMWriteScreen support_verif, 10, 36
		EMWriteScreen retro_support, 10, 45
		EMWriteScreen prosp_support, 10, 63
		EMWriteScreen alimony_verif, 11, 36
		EMWriteScreen retro_alimony, 11, 45
		EMWriteScreen prosp_alimony, 11, 63
		EMWriteScreen tax_dep_verif, 12, 36
		EMWriteScreen retro_tax_dep, 12, 45
		EMWriteScreen prosp_tax_dep, 12, 63
		EMWriteScreen other_verif, 13, 36
		EMWriteScreen retro_other, 13, 45
		EMWriteScreen prosp_other, 13, 63
		EMWriteScreen change_in_circum, 17, 61
	ELSEIF num_of_COEX <> "0" THEN
		PF9
		'---...if the script is PF9'ing, it is doing so to enter information into the HC Expense sub-menu
		'Opening the HC Expenses Sub-menu
		EMWriteScreen "X", 18, 44
		transmit

		DO
			EMReadScreen hc_expense_est, 14, 4, 30
		LOOP UNTIL hc_expense_est = "HC Expense Est"

		EMReadScreen current_month_plus_one, 17, 13, 51
		IF current_month_plus_one <> "CURRENT MONTH + 1" THEN
			EMWriteScreen hc_exp_support, 6, 38
			EMWriteScreen hc_exp_alimony, 7, 38
			EMWriteScreen hc_exp_tax_dep, 8, 38
			EMWriteScreen hc_exp_other, 9, 38
			transmit
		END IF
		PF3
	END IF
	transmit
end function

function write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DCEX")
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_DCEX, 1, 2, 78
	IF num_of_DCEX = "0" THEN
		EMWriteScreen "__", 20, 76
		Emwritescreen "NN", 20, 79
		transmit

		'---If the script is creating a new DCEX panel, it is going to enter this information into the DCEX main screen...
		EMWritescreen DCEX_provider, 6, 47
		EMWritescreen DCEX_reason, 7, 44
		EMWritescreen DCEX_subsidy, 8, 44
		IF len(DCEX_child_number1) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number1
		EMWritescreen DCEX_child_number1, 11, 29
		IF len(DCEX_child_number2) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number2
		EMWritescreen DCEX_child_number2, 12, 29
		IF len(DCEX_child_number3) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number3
		EMWritescreen DCEX_child_number3, 13, 29
		IF len(DCEX_child_number4) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number4
		EMWritescreen DCEX_child_number4, 14, 29
		IF len(DCEX_child_number5) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number5
		EMWritescreen DCEX_child_number5, 15, 29
		IF len(DCEX_child_number6) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number6
		EMWritescreen DCEX_child_number6, 16, 29
		EMWritescreen DCEX_child_number1_ver, 11, 41
		EMWritescreen DCEX_child_number2_ver, 12, 41
		EMWritescreen DCEX_child_number3_ver, 13, 41
		EMWritescreen DCEX_child_number4_ver, 14, 41
		EMWritescreen DCEX_child_number5_ver, 15, 41
		EMWritescreen DCEX_child_number6_ver, 16, 41
		EMWritescreen DCEX_child_number1_retro, 11, 48
		EMWritescreen DCEX_child_number2_retro, 12, 48
		EMWritescreen DCEX_child_number3_retro, 13, 48
		EMWritescreen DCEX_child_number4_retro, 14, 48
		EMWritescreen DCEX_child_number5_retro, 15, 48
		EMWritescreen DCEX_child_number6_retro, 16, 48
		EMWritescreen DCEX_child_number1_pro, 11, 63
		EMWritescreen DCEX_child_number2_pro, 12, 63
		EMWritescreen DCEX_child_number3_pro, 13, 63
		EMWritescreen DCEX_child_number4_pro, 14, 63
		EMWritescreen DCEX_child_number5_pro, 15, 63
		EMWritescreen DCEX_child_number6_pro, 16, 63
	ELSE
		PF9
		'---...if the script is PF9'ing, it is ONLY because it is going to enter information in the HC Expense sub-menu.
		'---Writing in the HC Expenses Est
		EMWriteScreen "X", 17, 55
		transmit

		DO			'---Waiting to make sure the HC Expense Est window has opened.
			EMReadScreen hc_expense_est, 10, 4, 41
		LOOP UNTIL hc_expense_est = "HC Expense"

		EMReadScreen hc_month, 17, 18, 62
		IF hc_month = "CURRENT MONTH + 1" THEN
			PF3
		ELSE
			IF len(DCEX_child_number1) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number1
			EMWritescreen DCEX_child_number1, 8, 39
			IF len(DCEX_child_number2) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number2
			EMWritescreen DCEX_child_number2, 9, 39
			IF len(DCEX_child_number3) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number3
			EMWritescreen DCEX_child_number3, 10, 39
			IF len(DCEX_child_number4) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number4
			EMWritescreen DCEX_child_number4, 11, 39
			IF len(DCEX_child_number5) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number5
			EMWritescreen DCEX_child_number5, 12, 39
			IF len(DCEX_child_number6) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number6
			EMWritescreen DCEX_child_number6, 13, 39
			EMWritescreen DCEX_child_number1_pro, 8, 49
			EMWritescreen DCEX_child_number2_pro, 9, 49
			EMWritescreen DCEX_child_number3_pro, 10, 49
			EMWritescreen DCEX_child_number4_pro, 11, 49
			EMWritescreen DCEX_child_number5_pro, 12, 49
			EMWritescreen DCEX_child_number6_pro, 13, 49
			transmit
			PF3
		END IF
	END IF
	transmit
end function

function write_panel_to_MAXIS_DFLN(conv_dt_1, conv_juris_1, conv_st_1, conv_dt_2, conv_juris_2, conv_st_2, rnd_test_dt_1, rnd_test_provider_1, rnd_test_result_1, rnd_test_dt_2, rnd_test_provider_2, rnd_test_result_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ conv_dt_1, conv_juris_1, conv_st_1, conv_dt_2, conv_juris_2, conv_st_2, rnd_test_dt_1, rnd_test_provider_1, rnd_test_result_1, rnd_test_dt_2, rnd_test_provider_2, rnd_test_result_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "DFLN")
	EMReadScreen num_of_DFLN, 1, 2, 78
	IF num_of_DFLN = "0" THEN
		EMWriteScreen reference_number, 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	CALL create_MAXIS_friendly_date(conv_dt_1, 0, 6, 27)
	EMWriteScreen conv_juris_1, 6, 41
	EMWriteScreen conv_st_1, 6, 75
	IF conv_dt_2 <> "" THEN
		CALL create_MAXIS_friendly_date(conv_dt_2, 0, 7, 27)
		EMWriteScreen conv_juris_2, 7, 41
		EMWriteScreen conv_st_2, 7, 75
	END IF
	IF rnd_test_dt_1 <> "" THEN
		CALL create_MAXIS_friendly_date(rnd_test_dt_1, 0, 14, 27)
		EMWriteScreen rnd_test_provider_1, 14, 41
		EMWriteScreen rnd_test_result_1, 14, 75
		IF rnd_test_dt_2 <> "" THEN
			CALL create_MAXIS_friendly_date(rnd_test_dt_2, 0, 15, 27)
			EMWriteScreen rnd_test_provider_2, 15, 41
			EMWriteScreen rnd_test_result_2, 15, 75
		END IF
	END IF
end function

function write_panel_to_MAXIS_DIET(DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DIET")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen DIET_mfip_1, 8, 40
	EMWriteScreen DIET_mfip_1_ver, 8, 51
	EMWriteScreen DIET_mfip_2, 9, 40
	EMWriteScreen DIET_mfip_2_ver, 9, 51
	EMWriteScreen DIET_msa_1, 11, 40
	EMWriteScreen DIET_msa_1_ver, 11, 51
	EMWriteScreen DIET_msa_2, 12, 40
	EMWriteScreen DIET_msa_2_ver, 12, 51
	EMWriteScreen DIET_msa_3, 13, 40
	EMWriteScreen DIET_msa_3_ver, 13, 51
	EMWriteScreen DIET_msa_4, 14, 40
	EMWriteScreen DIET_msa_4_ver, 14, 51
	transmit
end function

function write_panel_to_MAXIS_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_1619, disa_drug_alcohol)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the disa panel in MAXIS)
'~~~~~ disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_1619, disa_drug_alcohol: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "DISA")  'navigates to the stat panel
	call create_panel_if_nonexistent
	IF disa_begin_date <> "" THEN
		call create_MAXIS_friendly_date(disa_begin_date, 0, 6, 47)  'enters the disability begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_begin_date), 6, 53
	END IF
	IF disa_end_date <> "" THEN
		call create_MAXIS_friendly_date(disa_end_date, 0, 6, 69)  'enters the disability end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_end_date), 6, 75
	END IF
	IF disa_cert_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_begin, 0, 7, 47)  'enters the disability certification begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_begin), 7, 53
	END IF
	IF disa_cert_end <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_end, 0, 7, 69)  'enters the disability certification end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_end), 7, 75
	END IF
	IF disa_wavr_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_wavr_begin, 0, 8, 47)  'enters the disability waiver begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_begin), 8, 53
	END IF
	IF disa_wavr_end <> "" THEN
		call create_MAXIS_friendly_date(disa_wavr_end, 0, 8, 69)  'enters the disability waiver end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_end), 8, 75
	END IF
	IF disa_grh_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_grh_begin, 0, 9, 47)  'enters the disability grh begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_begin), 9, 53
	END IF
	IF disa_grh_end <> "" THEN
		call create_MAXIS_friendly_date(disa_grh_end, 0, 9, 69)  'enters the disability grh end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_end), 9, 75
	END IF
	Emwritescreen disa_cash_status, 11, 59  'enters status code for cash disa status
	Emwritescreen disa_cash_status_ver, 11, 69  'enters verification code for cash disa status
	Emwritescreen disa_snap_status, 12, 59  'enters status code for snap disa status
	Emwritescreen disa_snap_status_ver, 12, 69  'enters verification code for snap disa status
	Emwritescreen disa_hc_status, 13, 59  'enters status code for hc disa status
	Emwritescreen disa_hc_status_ver, 13, 69  'enters verification code for hc disa status
	Emwritescreen disa_waiver, 14, 59  'enters home and comminuty waiver code
	Emwritescreen disa_1619, 16, 59  'enters 1619 status
	Emwritescreen disa_drug_alcohol, 18, 69  'enters material drug & alcohol verification
end function

function write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DSTT")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen DSTT_ongoing_income, 6, 69
	IF HH_income_stop_date <> "" THEN call create_MAXIS_friendly_date(HH_income_stop_date, 0, 9, 69)
	EMWriteScreen income_expected_amt, 12, 71
end function

function write_panel_to_MAXIS_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	IF reference_number = "01" THEN
		call navigate_to_MAXIS_screen("STAT", "EATS")
		call create_panel_if_nonexistent
		EMWriteScreen eats_together, 4, 72
		EMWriteScreen eats_boarder, 5, 72
		IF ucase(eats_together) = "N" THEN
			EMWriteScreen "01", 13, 28
			eats_group_one = replace(eats_group_one, " ", "")
			eats_group_one = split(eats_group_one, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_one
				EMWriteScreen eats_household_member, 13, eats_col
				eats_col = eats_col + 4
			NEXT
			EMWriteScreen "02", 14, 28
			eats_group_two = replace(eats_group_two, " ", "")
			eats_group_two = split(eats_group_two, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_two
				EMWriteScreen eats_household_member, 14, eats_col
				eats_col = eats_col + 4
			NEXT
			IF eats_group_three <> "" THEN
				EMWriteScreen "03", 15, 28
				eats_group_three = replace(eats_group_three, " ", "")
				eats_group_three = split(eats_group_three, ",")
				eats_col = 39
				FOR EACH eats_household_member IN eats_group_three
					EMWriteScreen eats_household_member, 15, eats_col
					eats_col = eats_col + 4
				NEXT
			END IF
		END IF
	transmit
	END IF
end function

function write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "EMMA")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen EMMA_medical_emergency, 6, 46
	EMWriteScreen EMMA_health_consequence, 8, 46
	EMWriteScreen EMMA_verification, 10, 46
	call create_MAXIS_friendly_date(EMMA_begin_date, 0, 12, 46)
	IF EMMA_end_date <> "" THEN call create_MAXIS_friendly_date(EMMA_end_date, 0, 14, 46)
end function

function write_panel_to_MAXIS_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "EMPS")
	call create_panel_if_nonexistent
	If EMPS_orientation_date <> "" then call create_MAXIS_friendly_date(EMPS_orientation_date, 0, 5, 39) 'enter orientation date
	EMWritescreen left(EMPS_orientation_attended, 1), 5, 65
	EMWritescreen EMPS_good_cause, 5, 79
	If EMPS_sanc_begin <> "" then call create_MAXIS_friendly_date(EMPS_sanc_begin, 1, 6, 39) 'Sanction begin date
	If EMPS_sanc_end <> "" then call create_MAXIS_friendly_date(EMPS_sanc_end, 1, 6, 65) 'Sanction end date
	EMWritescreen left(EMPS_memb_at_home, 1), 8, 76
	EMWritescreen left(EMPS_care_family, 1), 9, 76
	EMWritescreen left(EMPS_crisis, 1), 10, 76
	EMWritescreen EMPS_hard_employ, 11, 76
	EMWritescreen left(EMPS_under1, 1), 12, 76 'child under 1 exemption
	EMWritescreen "n", 13, 76 'enters n for child under 12 weeks
	If EMPS_DWP_date <> "" then call create_MAXIS_friendly_date(EMPS_DWP_date, 1, 17, 40) 'DWP plan date
	'This populates the child under 1 popup if needed
	IF ucase(left(EMPS_under1, 1)) = "Y" THEN
		EMReadScreen month_to_use, 2, 20, 55
		EMReadScreen start_year, 2, 20, 58
		Emwritescreen "x", 12, 39
		Transmit
		EMReadScreen check_for_blank, 2, 7, 22 'makes sure the popup isn't already filled out
		month_to_use = cint(month_to_use)
		start_year = cint("20" & start_year)
		popup_row = 7 'setting initial starting point for the popup
		popup_col = 22
		IF check_for_blank <> "  " THEN 'blank popup, fill it out!
			FOR i = 1 to 12
				IF month_to_use > 12 THEN 'handling the year change
					popup_month = month_to_use - 12
					year_to_use = start_year +1
				ELSE
					popup_month = month_to_use
					year_to_use = start_year
				END IF
				IF len(popup_month) = 1 THEN popup_month = "0" & popup_month 'formatting to two digit month
				Emwritescreen popup_month, popup_row, popup_col
				Emwritescreen year_to_use, popup_row, popup_col + 5
				popup_col = popup_col + 11
				month_to_use = month_to_use + 1
				IF popup_col > 55 THEN 'This moves to the next row if necessary
					popup_col = 22
					popup_row = popup_row + 1
				END IF
			NEXT
			PF3 'closing the popup
		END IF
	END IF
end function

function write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "FACI")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen FACI_vendor_number, 5, 43
	EMWriteScreen FACI_name, 6, 43
	EMWriteScreen FACI_type, 7, 43
	EMWriteScreen FACI_FS_eligible, 8, 43
	If FACI_date_in <> "" then
		call create_MAXIS_friendly_date(FACI_date_in, 0, 14, 47)
		EMWriteScreen datepart("YYYY", FACI_date_in), 14, 53
	End if
	If FACI_date_out <> "" then
		call create_MAXIS_friendly_date(FACI_date_out, 0, 14, 71)
		EMWriteScreen datepart("YYYY", FACI_date_out), 14, 77
	End if
	transmit
	transmit
end function

function write_panel_to_MAXIS_FMED(FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount)
'--- This function writes to MAXIS in Krabappel only (pulls FMED information from the Excel file. This function can handle up to 4 FMED rows per client.)
'~~~~~ FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "FMED")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMReadScreen num_of_FMED, 1, 2, 78
	IF num_of_FMED = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	'Determining where to start writing...
	FMED_row = 9
	DO
		EMReadScreen FMED_available, 2, FMED_row, 25
		IF FMED_available <> "__" THEN FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	LOOP UNTIL FMED_available = "__"

	IF FMED_1_type <> "" THEN
		EMWriteScreen FMED_1_type, FMED_row, 25
			IF FMED_1_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_1_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_1_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_1_verif, FMED_row, 32
		EMWriteScreen FMED_1_ref_num, FMED_row, 38
		EMWriteScreen FMED_1_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_1_begin)			'Turning the value in FMED_1_begin and FMED_1_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_1_begin), 2), FMED_row, 53
		IF FMED_1_end <> "" THEN
			FMED_month = DatePart("M", FMED_1_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_1_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_1_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_2_type <> "" THEN
		EMWriteScreen FMED_2_type, FMED_row, 25
			IF FMED_2_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_2_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_2_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_2_verif, FMED_row, 32
		EMWriteScreen FMED_2_ref_num, FMED_row, 38
		EMWriteScreen FMED_2_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_2_begin)			'Turning the value in FMED_2_begin and FMED_2_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_2_begin), 2), FMED_row, 53
		IF FMED_2_end <> "" THEN
			FMED_month = DatePart("M", FMED_2_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_2_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_2_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_3_type <> "" THEN
		EMWriteScreen FMED_3_type, FMED_row, 25
			IF FMED_3_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_3_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_3_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_3_verif, FMED_row, 32
		EMWriteScreen FMED_3_ref_num, FMED_row, 38
		EMWriteScreen FMED_3_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_3_begin)			'Turning the value in FMED_3_begin and FMED_3_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_3_begin), 2), FMED_row, 53
		IF FMED_3_end <> "" THEN
			FMED_month = DatePart("M", FMED_3_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_3_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_3_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_4_type <> "" THEN
		EMWriteScreen FMED_4_type, FMED_row, 25
			IF FMED_4_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_4_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_4_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_4_verif, FMED_row, 32
		EMWriteScreen FMED_4_ref_num, FMED_row, 38
		EMWriteScreen FMED_4_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_4_begin)			'Turning the value in FMED_4_begin and FMED_4_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_4_begin), 2), FMED_row, 53
		IF FMED_4_end <> "" THEN
			FMED_month = DatePart("M", FMED_4_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_4_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_4_amount, FMED_row, 70
	END IF

	transmit
end function

function write_panel_to_MAXIS_HCRE(hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","HCRE")
	call create_panel_if_nonexistent
	'Converting the Appl Addendum Date into a usable format
	call MAXIS_dater(hcre_appl_addnd_date_input, hcre_appl_addnd_date_output, "HCRE Addendum Date")
	'Converting the Received by service date into a usable format
	call MAXIS_dater(hcre_recvd_by_service_date_input, hcre_recvd_by_service_date_output, "received by Service Date")
	'Converts Retro Months Input into a negative
	hcre_retro_months_input = (Abs(hcre_retro_months_input)*(-1))
	call add_months(hcre_retro_months_input,hcre_appl_addnd_date_output,hcre_retro_date_output)
	row = 1
	col = 1
	EMSearch "* " & reference_number, row, col
		'Appl Addendum Request Date
	EMWriteScreen left(hcre_appl_addnd_date_output,2)		, row, col + 29
	EMWriteScreen mid(hcre_recvd_by_service_date_input,4,2)	, row, col + 32
	EMWriteScreen right(hcre_appl_addnd_date_output,2)		, row, col + 35
		'Coverage Request Date
	EMWriteScreen left(hcre_retro_date_output,2)	, row, col + 42
	EMWriteScreen right(hcre_retro_date_output,2)	, row, col + 45
		'Recv By Sv Date
	EMWriteScreen left(hcre_recvd_by_service_date_output,2)	, row, col + 51
	EMWriteScreen mid(hcre_recvd_by_service_date_output,4,2), row, col + 54
	EMWriteScreen right(hcre_recvd_by_service_date_output,2), row, col + 57
	transmit
end function

function write_panel_to_MAXIS_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "HEST")
	call create_panel_if_nonexistent
	Emwritescreen "01", 6, 40
	call create_MAXIS_friendly_date(HEST_FS_choice_date, 0, 7, 40)
	EMWritescreen HEST_first_month, 8, 61
	'Filling in the #/FS units field (always 01)
	If ucase(left(HEST_heat_air_retro, 1)) = "Y" then EMWritescreen "01", 13, 42
	If ucase(left(HEST_heat_air_pro, 1)) = "Y" then EMWritescreen "01", 13, 68
	If ucase(left(HEST_electric_retro, 1)) = "Y" then EMWritescreen "01", 14, 42
	If ucase(left(HEST_electric_pro, 1)) = "Y" then EMWritescreen "01", 14, 68
	If ucase(left(HEST_phone_retro, 1)) = "Y" then EMWritescreen "01", 15, 42
	If ucase(left(HEST_phone_pro, 1)) = "Y" then EMWritescreen "01", 15, 68
	EMWritescreen left(HEST_heat_air_retro, 1), 13, 34
	EMWritescreen left(HEST_electric_retro, 1), 14, 34
	EMWritescreen left(HEST_phone_retro, 1), 15, 34
	EMWritescreen left(HEST_heat_air_pro, 1), 13, 60
	EMWritescreen left(HEST_electric_pro, 1), 14, 60
	EMWritescreen left(HEST_phone_pro, 1), 15, 60
	transmit
end function

function write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality, IMIG_40_soc_sec, IMIG_40_soc_sec_verif, IMIG_battered_spouse_child, IMIG_battered_spouse_child_verif, IMIG_military_status, IMIG_military_status_verif, IMIG_hmong_lao_nat_amer, IMIG_st_prog_esl_ctzn_coop, IMIG_st_prog_esl_ctzn_coop_verif, IMIG_fss_esl_skills_training)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality, IMIG_40_soc_sec, IMIG_40_soc_sec_verif, IMIG_battered_spouse_child, IMIG_battered_spouse_child_verif, IMIG_military_status, IMIG_military_status_verif, IMIG_hmong_lao_nat_amer, IMIG_st_prog_esl_ctzn_coop, IMIG_st_prog_esl_ctzn_coop_verif, IMIG_fss_esl_skills_training: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "IMIG")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	call create_MAXIS_friendly_date(APPL_date, 0, 5, 45)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", APPL_date), 5, 51
	EMWriteScreen IMIG_imigration_status, 6, 45							'Writes imig status
	IF IMIG_entry_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_entry_date, 0, 7, 45)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_entry_date), 7, 51
	END IF
	IF IMIG_status_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_status_date, 0, 7, 71)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_status_date), 7, 77
	END IF
	EMWriteScreen IMIG_status_ver, 8, 45								'Enters status ver
	EMWriteScreen IMIG_status_LPR_adj_from, 9, 45						'Enters status LPR adj from
	EMWriteScreen IMIG_nationality, 10, 45								'Enters nationality
	EMwritescreen IMIG_40_soc_sec, 13, 56								'Enters info about Social Security Credits
	EMwritescreen IMIG_40_soc_sec_verif, 13, 71
	EMwritescreen IMIG_battered_spouse_child, 14, 56					'Enters info about Battered Child/Spouse claims
	EMwritescreen IMIG_battered_spouse_child_verif, 14, 71
	EMwritescreen IMIG_military_status, 15, 56 							'Enters info about possible military status
	EMwritescreen IMIG_military_status_verif, 15, 71
	EMwritescreen IMIG_hmong_lao_nat_amer, 16, 56 						'Enters status of particular nationalities/identity
	EMwritescreen IMIG_st_prog_esl_ctzn_coop, 17, 56 					'Enters information about ESL/Citizen cooperation status
	EMwritescreen IMIG_st_prog_esl_ctzn_coop_verif, 17, 71
	EMwritescreen IMIG_fss_esl_skills_training, 18, 56 					'Enters information about ESL Skills course
	transmit
	transmit
end function

function write_panel_to_MAXIS_INSA(insa_pers_coop_ohi, insa_good_cause_status, insa_good_cause_cliam_date, insa_good_cause_evidence, insa_coop_cost_effect, insa_insur_name, insa_prescrip_drug_cover, insa_prescrip_end_date, insa_persons_covered)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ insa_pers_coop_ohi, insa_good_cause_status, insa_good_cause_cliam_date, insa_good_cause_evidence, insa_coop_cost_effect, insa_insur_name, insa_prescrip_drug_cover, insa_prescrip_end_date, insa_persons_covered: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","INSA")
	call create_panel_if_nonexistent

	EMWriteScreen insa_pers_coop_ohi, 4, 62
	EMWriteScreen insa_good_cause_status, 5, 62
	If insa_good_cause_cliam_date <> "" then CALL create_MAXIS_friendly_date(insa_good_cause_cliam_date, 0, 6, 62)
	EMWriteScreen insa_good_cause_evidence, 7, 62
	EMWriteScreen insa_coop_cost_effect, 8, 62
	EMWriteScreen insa_insur_name, 10, 38
	EMWriteScreen insa_prescrip_drug_cover, 11, 62
	If insa_prescrip_end_date <> "" then CALL create_MAXIS_friendly_date(insa_prescrip_end_date, 0, 12, 62)

	'Adding persons covered
	insa_row = 15
	insa_col = 30

	insa_persons_covered = replace(insa_persons_covered, " ", "")
	insa_persons_covered = split(insa_persons_covered, ",")

	FOR EACH insa_peep IN insa_persons_covered
		EMWriteScreen insa_peep, insa_row, insa_col
		insa_col = insa_col + 4
		IF insa_col = 70 THEN
			insa_col = 30
			insa_row = 16
		END IF
	NEXT
	transmit
end function

function write_panel_to_MAXIS_JOBS(jobs_number, jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ jobs_number, jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "JOBS")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen jobs_number, 20, 79
	transmit

	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	EMWriteScreen jobs_inc_type, 5, 34
	EMWriteScreen jobs_inc_verif, 6, 34

	EMWriteScreen jobs_employer_name, 7, 42
	call create_MAXIS_friendly_date(jobs_inc_start, 0, 9, 35)
	EMWriteScreen jobs_pay_freq, 18, 35

	'===== navigates to the SNAP PIC to update the PIC =====
	EMWriteScreen "X", 19, 38
	transmit
	DO
		EMReadScreen at_snap_pic, 12, 3, 22
	LOOP UNTIL at_snap_pic = "Food Support"
	EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	EMReadScreen pic_info_exists, 8, 18, 57
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN
		call create_MAXIS_friendly_date(date, 0, 5, 34)
		EMWriteScreen jobs_pay_freq, 5, 64
		EMWriteScreen jobs_wkly_hrs, 8, 64
		EMWriteScreen jobs_hrly_wage, 9, 66
		transmit
		transmit
		EMReadScreen jobs_pic_hrs_per_pp, 6, 16, 51
		EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	END IF
	transmit		'<=====navigates out of the PIC

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	benefit_month = bene_month & "/01/" & bene_year
	retro_month = DatePart("M", DateAdd("M", -2, benefit_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(DatePart("YYYY", DateAdd("M", -2, benefit_month)), 2)

	EMWriteScreen retro_month, 12, 25
	EMWriteScreen retro_year, 12, 31
	EMWriteScreen bene_month, 12, 54
	EMWriteScreen bene_year, 12, 60

	IF pic_info_exists = "" THEN 		'---If the PIC is blank, the information needs to be added to the main JOBS panel as well.
		EMWriteScreen "05", 12, 28
		EMWriteScreen jobs_pic_wages_per_pp, 12, 38
		EMWriteScreen "05", 12, 57
		EMWriteScreen jobs_pic_wages_per_pp, 12, 67
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 43
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 72
	END IF

	IF jobs_pay_freq = 2 OR jobs_pay_freq = 3 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "19", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	ELSEIF jobs_pay_freq = 4 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen bene_year, 15, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "12", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 14, 28
			EMWriteScreen jobs_pic_wages_per_pp, 14, 38
			EMWriteScreen "26", 15, 28
			EMWriteScreen jobs_pic_wages_per_pp, 15, 38
			EMWriteScreen "12", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen "19", 14, 57
			EMWriteScreen jobs_pic_wages_per_pp, 14, 67
			EMWriteScreen "26", 15, 57
			EMWriteScreen jobs_pic_wages_per_pp, 15, 67
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", DATE) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to numeric.
		EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
		IF HC_income_est_check = "Est" Then 'this is the old position
			EMWriteScreen "x", 19, 54
		ELSE								'this is the new position
			EMWriteScreen "x", 19, 48
		END IF
		transmit

		DO
			EMReadScreen hc_inc_est, 9, 9, 43
		LOOP UNTIL hc_inc_est = "HC Income"

		EMWriteScreen jobs_pic_wages_per_pp, 11, 63
		transmit
		transmit
	END IF
end function

function write_panel_to_MAXIS_MEDI(SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date, MEDI_apply_prem_to_spdn, MEDI_apply_prem_end_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date, MEDI_apply_prem_to_spdn, MEDI_apply_prem_end_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "MEDI")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen SSN_first, 6, 39				'Next three lines pulled
	EMWriteScreen SSN_mid, 6, 43
	EMWriteScreen SSN_last, 6, 46
	EMWriteScreen MEDI_claim_number_suffix, 6, 51
	EMWriteScreen MEDI_part_A_premium, 7, 46
	EMWriteScreen MEDI_part_B_premium, 7, 73
	If MEDI_part_A_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_A_begin_date, 0, 15, 24)
	If MEDI_part_B_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_B_begin_date, 0, 15, 54)
	EMWriteScreen MEDI_apply_prem_to_spdn, 11, 71
	IF MEDI_apply_prem_end_date <> "" THEN
		EMWriteScreen left(MEDI_apply_prem_end_date, 2), 12, 71
		EMWriteScreen right(MEDI_apply_prem_end_date, 2), 12, 74
	END IF
	transmit
	transmit
end function

function write_panel_to_MAXIS_MMSA(mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	IF mmsa_liv_arr <> "" THEN
		call navigate_to_MAXIS_screen("STAT", "MMSA")
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen mmsa_liv_arr, 7, 54
		EMWriteScreen mmsa_cont_elig, 9, 54
		EMWriteScreen mmsa_spous_inc, 12, 62
		EMWriteScreen mmsa_shared_hous, 14, 62
		transmit
	END IF
end function

function write_panel_to_MAXIS_MSUR(msur_begin_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ msur_begin_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","MSUR")
	call create_panel_if_nonexistent

	'msur_begin_date This is the date MSUR began for this client
	row = 7
	DO
		EMReadScreen available_space, 2, row, 36
		IF available_space = "__" THEN
			row = row + 1
		ELSE
			EXIT DO
		END IF
	LOOP UNTIL available_space <> "__"

	CALL create_MAXIS_friendly_date(msur_begin_date, 0, row, 36)
	Emwritescreen DatePart("YYYY", msure_begin_date), row, 42
	transmit
end function

function write_panel_to_MAXIS_OTHR(othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint_owner, othr_share_ratio)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the othr panel in MAXIS)
'~~~~~ othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint_owner, othr_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "OTHR")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen othr_type, 6, 40  'enters other asset type
	IF othr_cash_value = "" THEN othr_cash_value = 0
	Emwritescreen othr_cash_value, 8, 40  'enters cash value of asset
	Emwritescreen othr_cash_value_ver, 8, 57  'enters cash value verification code
	IF othr_owed = "" THEN othr_owed = 0
	Emwritescreen othr_owed, 9, 40  'enters amount owed value
	Emwritescreen othr_owed_ver, 9, 57  'enters amount owed verification code
	call create_MAXIS_friendly_date(othr_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen othr_cash_count, 12, 50  'enters y/n if counted for cash
	Emwritescreen othr_SNAP_count, 12, 57  'enters y/n if counted for snap
	Emwritescreen othr_HC_count, 12, 64  'enters y/n if counted for hc
	Emwritescreen othr_IV_count, 12, 73  'enters y/n if counted for iv
	Emwritescreen othr_joint_owner, 13, 44  'enters if it is a jointly owned other asset
	Emwritescreen left(othr_share_ratio, 1), 15, 50  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(othr_share_ratio, 1), 15, 54  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_PARE(appl_date, reference_number, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, reference_number, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "PARE")
	CALL write_value_and_transmit(reference_number, 20, 76)
	EMReadScreen num_of_PARE, 1, 2, 78
	IF num_of_PARE = "0" THEN
		CALL write_value_and_transmit("NN", 20, 79)
	ELSE
		PF9
	END IF
	CALL create_MAXIS_friendly_date(appl_date, 0, 5, 37)
	EMWriteScreen DatePart("YYYY", appl_date), 5, 43

	IF len(PARE_child_1) = 1 THEN PARE_child_1 = "0" & PARE_child_1
	IF len(PARE_child_2) = 1 THEN PARE_child_1 = "0" & PARE_child_2
	IF len(PARE_child_3) = 1 THEN PARE_child_1 = "0" & PARE_child_3
	IF len(PARE_child_4) = 1 THEN PARE_child_1 = "0" & PARE_child_4
	IF len(PARE_child_5) = 1 THEN PARE_child_1 = "0" & PARE_child_5
	IF len(PARE_child_6) = 1 THEN PARE_child_1 = "0" & PARE_child_6
	EMWritescreen PARE_child_1, 8, 24
	EMWritescreen PARE_child_1_relation, 8, 53
	EMWritescreen PARE_child_1_verif, 8, 71
	EMWritescreen PARE_child_2, 9, 24
	EMWritescreen PARE_child_2_relation, 9, 53
	EMWritescreen PARE_child_2_verif, 9, 71
	EMWritescreen PARE_child_3, 10, 24
	EMWritescreen PARE_child_3_relation, 10, 53
	EMWritescreen PARE_child_3_verif, 10, 71
	EMWritescreen PARE_child_4, 11, 24
	EMWritescreen PARE_child_4_relation, 11, 53
	EMWritescreen PARE_child_4_verif, 11, 71
	EMWritescreen PARE_child_5, 12, 24
	EMWritescreen PARE_child_5_relation, 12, 53
	EMWritescreen PARE_child_5_verif, 12, 71
	EMWritescreen PARE_child_6, 13, 24
	EMWritescreen PARE_child_6_relation, 13, 53
	EMWritescreen PARE_child_6_verif, 13, 71
	transmit
end function

function write_panel_to_MAXIS_PBEN(pben_referal_date, pben_type, pben_appl_date, pben_appl_ver, pben_IAA_date, pben_disp)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the pben panel in MAXIS)
'~~~~~ pben_referal_date, pben_type, pben_appl_date, pben_appl_ver, pben_IAA_date, pben_disp: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "PBEN")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emreadscreen pben_row_check, 2, 8, 24  'reads the MAXIS screen to find out if the PBEN row has already been used.
	If pben_row_check = "__" THEN   'if the row is blank it enters it in the 8th row.
		Emwritescreen pben_type, 8, 24  'enters pben type code
		call create_MAXIS_friendly_date(pben_referal_date, 0, 8, 40)  'enters referal date in MAXIS friendly format mm/dd/yy
		call create_MAXIS_friendly_date(pben_appl_date, 0, 8, 51)  'enters appl date in  MAXIS friendly format mm/dd/yy
		Emwritescreen pben_appl_ver, 8, 62  'enters appl verification code
		call create_MAXIS_friendly_date(pben_IAA_date, 0, 8, 66)  'enters IAA date in MAXIS friendly format mm/dd/yy
		Emwritescreen pben_disp, 8, 77  'enters the status of pben application
	else
		EMreadscreen pben_row_check, 2, 9, 24  'if row 8 is filled already it will move to row 9 and see if it has been used.
		IF pben_row_check = "__" THEN  'if the 9th row is blank it enters the information there.
		'second pben row
			Emwritescreen pben_type, 9, 24
			call create_MAXIS_friendly_date(pben_referal_date, 0, 9, 40)
			call create_MAXIS_friendly_date(pben_appl_date, 0, 9, 51)
			Emwritescreen pben_appl_ver, 9, 62
			call create_MAXIS_friendly_date(pben_IAA_date, 0, 9, 66)
			Emwritescreen pben_disp, 9, 77
		else
		Emreadscreen pben_row_check, 2, 10, 24  'if row 8 is filled already it will move to row 9 and see if it has been used.
			IF pben-row_check = "__" THEN  'if the 9th row is blank it enters the information there.
			'third pben row
				Emwritescreen pben_type, 10, 24
				call create_MAXIS_friendly_date(pben_referal_date, 0, 10, 40)
				call create_MAXIS_friendly_date(pben_appl_date, 0, 10, 51)
				Emwritescreen pben_appl_ver, 10, 62
				call create_MAXIS_friendly_date(pben_IAA_date, 0, 10, 66)
				Emwritescreen pben_disp, 10, 77
			END IF
		END IF
	END IF
end function

function write_panel_to_MAXIS_PDED(PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","PDED")
	call create_panel_if_nonexistent

	'Disa Widow/ers Deductionpded_shel_spcl_needs
	If pded_wid_deduction <> "" then
		pded_wid_deduction = ucase(pded_wid_deduction)
		pded_wid_deduction = left(pded_wid_deduction,1)
		EMWriteScreen pded_wid_deduction, 7, 60
	End If

	'Disa Adult Child Disregard
	If pded_adult_child_disregard <> "" then
		pded_adult_child_disregard = ucase(pded_adult_child_disregard)
		pded_adult_child_disregard = left(pded_adult_child_disregard,1)
		EMWriteScreen pded_adult_child_disregard, 8, 60
	End If

	'Widow/ers Disregard
	If pded_wid_disregard <> "" then
		pded_wid_disregard = ucase(pded_wid_disregard)
		pded_wid_disregard = left(pded_wid_disregard,1)
		EMWriteScreen pded_wid_disregard, 9, 60
	End If

	'Other Unearned Income Deduction
	If pded_unea_income_deduction_reason <> "" and pded_unea_income_deduction_value <> "" then
		EMWriteScreen pded_unea_income_deduction_value, 10, 62
		EMWriteScreen "X", 10, 25
		Transmit
		EMWriteScreen pded_unea_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If

	'Other Earned Income Deduction
	If pded_earned_income_deduction_reason <> "" and pded_earned_income_deduction_value <> "" then
		EMWriteScreen pded_earned_income_deduction_value, 11, 62
		EMWriteScreen "X", 11, 27
		Transmit
		EMWriteScreen pded_earned_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If

	'Extend MA-EPD Income/Asset Limits
	If pded_ma_epd_inc_asset_limit <> "" then
		pded_ma_epd_inc_asset_limit = ucase(pded_ma_epd_inc_asset_limit)
		pded_ma_epd_inc_asset_limit = left(pded_ma_epd_inc_asset_limit,1)
		EMWriteScreen pded_ma_epd_inc_asset_limit, 12, 65
	End If

	'Guardianship Fee
	If pded_guard_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 44
	End If

	'Rep Payee Fee
	If pded_rep_payee_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 70
	End If

	'Other Expense
	If pded_other_expense <> "" then
		EMWriteScreen pded_other_expense, 18, 41
	End If

	'Shelter Special Needs
	If pded_shel_spcl_needs <> "" then
		pded_shel_spcl_needs = ucase(pded_shel_spcl_needs)
		pded_shel_spcl_needs = left(pded_shel_spcl_needs,1)
		EMWriteScreen pded_shel_spcl_needs, 18, 78
	End If

	'Excess Need
	If pded_excess_need <> "" then
		EMWriteScreen pded_excess_need, 19, 41
	End If

	'Restaurant Meals
	If pded_restaurant_meals <> "" then
		pded_restaurant_meals = ucase(pded_restaurant_meals)
		pded_restaurant_meals = left(pded_restaurant_meals,1)
		EMWriteScreen pded_restaurant_meals, 19, 78
	End If
	Transmit
end function

function write_panel_to_MAXIS_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver, PREG_due_date, PREG_multiple_birth)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver, PREG_due_date, PREG_multiple_birth: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "PREG")
	call create_panel_if_nonexistent
	EMWritescreen "NN", 20, 79
	transmit
	call create_MAXIS_friendly_date(PREG_conception_date, 0, 6, 53)
	third_trimester_date = dateadd("M", 6, PREG_conception_date)
	CALL create_MAXIS_friendly_date(third_trimester_date, 0, 8, 53)
	call create_MAXIS_friendly_date(PREG_due_date, 1, 10, 53)
	EMWritescreen PREG_conception_date_ver, 6, 75
	EMWritescreen PREG_third_trimester_ver, 8, 75
	EMWritescreen PREG_multiple_birth, 14, 53
	transmit
end function

function write_panel_to_MAXIS_RBIC(rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the rbic panel in MAXIS)
'~~~~~ rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "RBIC")  'navigates to the stat panel
	call create_panel_if_nonexistent
	EMwritescreen rbic_type, 5, 44  'enters rbic type code
	call create_MAXIS_friendly_date(rbic_start_date, 0, 6, 44)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic start date
	IF rbic_end_date <> "" THEN call create_MAXIS_friendly_date(rbic_end_date, 6, 68)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic end date
	rbic_group_1 = replace(rbic_group_1, " ", "")  'this will replace any spaces in the array with nothing removing the spaces.
	rbic_group_1 = split(rbic_group_1, ",")  'this will split up the reference numbers in the array based on commas
	rbic_col = 25                            'this will set the starting column to enter rbic reference numbers
	For each rbic_hh_memb in rbic_group_1    'for each reference number that is in the array for group 1 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_1, 10, 47  'enters the rbic retro income for group 1
	EMwritescreen rbic_prosp_income_group_1, 10, 62  'enters the rbic prospective income for group 1
	EMwritescreen rbic_ver_income_group_1, 10, 76    'enters the income verification code for group 1
	rbic_group_2 = replace(rbic_group_2, " ", "")
	rbic_group_2 = split(rbic_group_2, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_2    'for each reference number that is in the array for group 2 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 11, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_2, 11, 47  'enters the rbic retro income for group 2
	EMwritescreen rbic_prosp_income_group_2, 11, 62  'enters the rbic prospective income for group 2
	EMwritescreen rbic_ver_income_group_2, 11, 76    'enters the income verification code for group 2
	rbic_group_3 = replace(rbic_group_3, " ", "")
	rbic_group_3 = split(rbic_group_3, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_3    'for each reference number that is in the array for group 3 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_3, 12, 47  'enters the rbic retro income for group 3
	EMwritescreen rbic_prosp_income_group_3, 12, 62  'enters the rbic prospective income for group 3
	EMwritescreen rbic_ver_income_group_3, 12, 76    'enters the income verification code for group 3
	EMwritescreen rbic_retro_hours, 13, 52  'enters the retro hours
	EMwritescreen rbic_prosp_hours, 13, 67  'enters the prospective hours
	EMwritescreen rbic_exp_type_1, 15, 25   'enters the expenses type for group 1
	EMwritescreen rbic_exp_retro_1, 15, 47  'enters the expenses retro for group 1
	EMwritescreen rbic_exp_prosp_1, 15, 62  'enters the expenses prospective for group 1
	EMwritescreen rbic_exp_ver_1, 15, 76    'enters the expenses verification code for group 1
	EMwritescreen rbic_exp_type_2, 16, 25   'enters the expenses type for group 2
	EMwritescreen rbic_exp_retro_2, 16, 47  'enters the expenses retro for group 2
	EMwritescreen rbic_exp_prosp_2, 16, 62  'enters the expenses prospective for group 2
	EMwritescreen rbic_exp_ver_2, 16, 76    'enters the expenses verification code for group 2
end function

function write_panel_to_MAXIS_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the rest panel in MAXIS)
'~~~~~ rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "REST")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen rest_type, 6, 39  'enters residence type
	Emwritescreen rest_type_ver, 6, 62  'enters verification of residence type
	Emwritescreen rest_market, 8, 41  'enters market value of residence
	Emwritescreen rest_market_ver, 8, 62  'enters market value verification code
	Emwritescreen rest_owed, 9, 41  'enters amount owned on residence
	Emwritescreen rest_owed_ver, 9, 62  'enters amount owed verification code
	call create_MAXIS_friendly_date(rest_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen rest_status, 12, 54  'enters property status code
	Emwritescreen rest_joint, 13, 54  'enters if it is a jointly owned home
	Emwritescreen left(rest_share_ratio, 1), 14, 54  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(rest_share_ratio, 1), 14, 58  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	IF rest_agreement_date <> "" THEN call create_MAXIS_friendly_date(rest_agreement_date, 0, 16, 62)
end function

function write_panel_to_MAXIS_SCHL(appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	EMWriteScreen "SCHL", 20, 71
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_SCHL, 1, 2, 78
	IF num_of_SCHL = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		call create_MAXIS_friendly_date(appl_date, 0, 5, 40)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
		EMWriteScreen datepart("yyyy", appl_date), 5, 46
		EMWriteScreen SCHL_status, 6, 40
		EMWriteScreen SCHL_ver, 6, 63
		EMWriteScreen SCHL_type, 7, 40
		IF len(SCHL_district_nbr) <> 4 THEN
			DO
				SCHL_district_nbr = "0" & SCHL_district_nbr
			LOOP UNTIL len(SCHL_district_nbr) = 4
		END IF
		EMWriteScreen SCHL_district_nbr, 8, 40
		If SCHL_kindergarten_start_date <> "" then call create_MAXIS_friendly_date(SCHL_kindergarten_start_date, 0, 10, 63)
		EMWriteScreen left(SCHL_grad_date, 2), 11, 63
		EMWriteScreen right(SCHL_grad_date, 2), 11, 66
		EMWriteScreen SCHL_grad_date_ver, 12, 63
		EMWriteScreen SCHL_primary_secondary_funding, 14, 63
		EMWriteScreen SCHL_FS_eligibility_status, 16, 63
		EMWriteScreen SCHL_higher_ed, 18, 63
		transmit
	END IF
end function

function write_panel_to_MAXIS_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the secu panel in MAXIS)
'~~~~~ secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "SECU")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen secu_type, 6, 50  'enters security type
	Emwritescreen secu_pol_numb, 7, 50  'enters policy number
	Emwritescreen secu_name, 8, 50  'enters name of policy
	Emwritescreen secu_cash_val, 10, 52  'enters cash value of policy
	call create_MAXIS_friendly_date(secu_date, 0, 11, 35)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen secu_cash_ver, 11, 50  'enters cash value verification code
	Emwritescreen secu_face_val, 12, 52  'enters face value of policy
	Emwritescreen secu_withdraw, 13, 52  'enters withdrawl penalty
	Emwritescreen secu_cash_count, 15, 50  'enters y/n if counted for cash
	Emwritescreen secu_SNAP_count, 15, 57  'enters y/n if counted for snap
	Emwritescreen secu_HC_count, 15, 64  'enters y/n if counted for hc
	Emwritescreen secu_GRH_count, 15, 72  'enters y/n if counted for grh
	Emwritescreen secu_IV_count, 15, 80  'enters y/n if counted for iv
	Emwritescreen secu_joint, 16, 44  'enters if it is a jointly owned security
	Emwritescreen left(secu_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(secu_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SHEL")
	call create_panel_if_nonexistent
	EMWritescreen SHEL_subsidized, 6, 46
	EMWritescreen SHEL_shared, 6, 64
	EMWritescreen SHEL_paid_to, 7, 50
	EMWritescreen SHEL_rent_retro, 11, 37
	EMWritescreen SHEL_rent_retro_ver, 11, 48
	EMWritescreen SHEL_rent_pro, 11, 56
	EMWritescreen SHEL_rent_pro_ver, 11, 67
	EMWritescreen SHEL_lot_rent_retro, 12, 37
	EMWritescreen SHEL_lot_rent_retro_ver, 12, 48
	EMWritescreen SHEL_lot_rent_pro, 12, 56
	EMWritescreen SHEL_lot_rent_pro_ver, 12, 67
	EMWritescreen SHEL_mortgage_retro, 13, 37
	EMWritescreen SHEL_mortgage_retro_ver, 13, 48
	EMWritescreen SHEL_mortgage_pro, 13, 56
	EMwritescreen SHEL_mortgage_pro_ver, 13, 67
	EMWritescreen SHEL_insur_retro, 14, 37
	EMWritescreen SHEL_insur_retro_ver, 14, 48
	EMWritescreen SHEL_insur_pro, 14, 56
	EMWritescreen SHEL_insur_pro_ver, 14, 67
	EMWritescreen SHEL_taxes_retro, 15, 37
	EMWritescreen SHEL_taxes_retro_ver, 15, 48
	EMWritescreen SHEL_taxes_pro, 15, 56
	EMWritescreen SHEL_taxes_pro_ver, 15, 67
	EMWritescreen SHEL_room_retro, 16, 37
	EMWritescreen SHEL_room_retro_ver, 16, 48
	EMWritescreen SHEL_room_pro, 16, 56
	EMWritescreen SHEL_room_pro_ver, 16, 67
	EMWritescreen SHEL_garage_retro, 17, 37
	EMWritescreen SHEL_garage_retro_ver, 17, 48
	EMWritescreen SHEL_garage_pro, 17, 56
	EMWritescreen SHEL_garage_pro_ver, 17, 67
	EMWritescreen SHEL_subsidy_retro, 18, 37
	EMWritescreen SHEL_subsidy_retro_ver, 18, 48
	EMWritescreen SHEL_subsidy_pro, 18, 56
	EMWritescreen SHEL_subsidy_pro_ver, 18, 67
	transmit
end function

function write_panel_to_MAXIS_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SIBL_group_1, SIBL_group_2, SIBL_group_3: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SIBL")
	EMReadScreen num_of_SIBL, 1, 2, 78
	IF num_of_SIBL = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	END IF

	If SIBL_group_1 <> "" then
		EMWritescreen "01", 7, 28
		SIBL_group_1 = replace(SIBL_group_1, " ", "") 'Removing spaces
		SIBL_group_1 = split(SIBL_group_1, ",") 'Splits the sibling group value into an array by commas
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_1 'Writes the member numbers onto the group line
			EMWritescreen SIBL_group_member, 7, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if

	If SIBL_group_2 <> "" then
		EMWritescreen "02", 8, 28
		SIBL_group_2 = replace(SIBL_group_2, " ", "")
		SIBL_group_2 = split(SIBL_group_2, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_2
			EMWritescreen SIBL_group_member, 8, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if

	If SIBL_group_3 <> "" then
		EMWritescreen "03", 9, 28
		SIBL_group_2 = replace(SIBL_group_3, " ", "")
		SIBL_group_2 = split(SIBL_group_3, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_3
			EMWritescreen SIBL_group_member, 9, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if
	transmit
end function

function write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SPON_type, SPON_ver, SPON_name, SPON_state: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SPON")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen SPON_type, 6, 38
	EMWriteScreen SPON_ver, 6, 62
	EMWriteScreen SPON_name, 8, 38
	EMWriteScreen SPON_state, 10, 62
	transmit
end function

function write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	EMWriteScreen "STEC", 20, 71
	EMWriteSCreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_STEC, 1, 2, 78
	IF num_of_STEC = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		EMWriteScreen STEC_type_1, 8, 25				'STEC 1
		EMWriteScreen STEC_amt_1, 8, 31
		STEC_actual_from_thru_months_1 = replace(STEC_actual_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_1, 2), 8, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 4, 2), 8, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 7, 2), 8, 48
		EMWriteScreen right(STEC_actual_from_thru_months_1, 2), 8, 51
		EMWriteScreen STEC_ver_1, 8, 55
		EMWriteScreen STEC_earmarked_amt_1, 8, 59
		STEC_earmarked_from_thru_months_1 = replace(STEC_earmarked_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_1, 2), 8, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 4, 2), 8, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 7, 2), 8, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_1, 2), 8, 79
		EMWriteScreen STEC_type_2, 9, 25				'STEC 1
		EMWriteScreen STEC_amt_2, 9, 31
		STEC_actual_from_thru_months_2 = replace(STEC_actual_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_2, 2), 9, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 4, 2), 9, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 7, 2), 9, 48
		EMWriteScreen right(STEC_actual_from_thru_months_2, 2), 9, 51
		EMWriteScreen STEC_ver_2, 9, 55
		EMWriteScreen STEC_earmarked_amt_2, 9, 59
		STEC_earmarked_from_thru_months_2 = replace(STEC_earmarked_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_2, 2), 9, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 4, 2), 9, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 7, 2), 9, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_2, 2), 9, 79
		transmit
	END IF
end function

function write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	EMWriteScreen "STIN", 20, 71
	EMWriteSCreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_STIN, 1, 2, 78
	IF num_of_STIN = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		EMWriteScreen STIN_type_1, 8, 27				'STIN 1
		EMWriteScreen STIN_amt_1, 8, 34
		call create_MAXIS_friendly_date(STIN_avail_date_1, 0, 8, 46)
		STIN_months_covered_1 = replace(STIN_months_covered_1, " ", "")
		EMWriteScreen left(STIN_months_covered_1, 2), 8, 58
		EMWriteScreen mid(STIN_months_covered_1, 4, 2), 8, 61
		EMWriteScreen mid(STIN_months_covered_1, 7, 2), 8, 67
		EMWriteScreen right(STIN_months_covered_1, 2), 8, 70
		EMWriteScreen STIN_ver_1, 8, 76
		EMWriteScreen STIN_type_2, 9, 27				'STIN 2
		EMWriteScreen STIN_amt_2, 9, 34
		STIN_avail_date_2 = replace(STIN_avail_date_2, " ", "")
		IF STIN_avail_date_2 <> "" THEN call create_MAXIS_friendly_date(STIN_avail_date_2, 0, 9, 46)
		EMWriteScreen left(STIN_months_covered_2, 2), 9, 58
		EMWriteScreen mid(STIN_months_covered_2, 4, 2), 9, 61
		EMWriteScreen mid(STIN_months_covered_2, 7, 2), 9, 67
		EMWriteScreen right(STIN_months_covered_2, 2), 9, 70
		EMWriteScreen STIN_ver_2, 9, 76
		transmit
	END IF
end function

function write_panel_to_MAXIS_STWK(STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","STWK")
	call create_panel_if_nonexistent

	EMWriteScreen stwk_empl_name, 6, 46
	If stwk_wrk_stop_date <> "" then CALL create_MAXIS_friendly_date(stwk_wrk_stop_date, 0, 7, 46)
	EMWriteScreen stwk_wrk_stop_date_verif, 7, 63
	IF stwk_inc_stop_date <> "" THEN CALL create_MAXIS_friendly_date(stwk_inc_stop_date, 0, 8, 46)
	EMWriteScreen stwk_refused_empl_yn, 8, 78
	EMWriteScreen stwk_vol_quit, 10, 46
	If stwk_ref_empl_date <> "" then CALL create_MAXIS_friendly_date(stwk_ref_empl_date, 0, 10, 72)
	EMWriteScreen stwk_gc_cash, 12, 52
	EMWriteScreen stwk_gc_grh, 12, 60
	EMWriteScreen stwk_gc_fs, 12, 67
	EMWriteScreen stwk_fs_pwe, 14, 46
	EMWriteScreen stwk_maepd_ext, 16, 46
	Transmit
end function

function write_panel_to_MAXIS_TYPE_PROG_REVW(appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "TYPE")
	IF reference_number = "01" THEN
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen type_cash_yn, 6, 28
		EMWriteScreen type_hc_yn, 6, 37
		EMWriteScreen type_fs_yn, 6, 46
		EMWriteScreen "N", 6, 55
		EMWriteScreen "N", 6, 64
		EMWriteScreen "N", 6, 73
		type_row = 7
		DO				'<=====this DO/LOOP populates "N" for all other HH members on TYPE so the script can get past TYPE when the reference number = "01"
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist <> "  " THEN
				EMWriteScreen "N", type_row, 28
				EMWriteScreen "N", type_row, 37
				EMWriteScreen "N", type_row, 46
				EMWriteScreen "N", type_row, 55
				type_row = type_row + 1
			ELSE
				EXIT DO
			END IF
		LOOP WHILE type_does_hh_memb_exist <> "  "
	ELSE
		PF9
		type_row = 7
		DO
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist = reference_number THEN
				EMWriteScreen type_cash_yn, type_row, 28
				EMWriteScreen type_hc_yn, type_row, 37
				EMWriteScreen type_fs_yn, type_row, 46
				EMWriteScreen "N", type_row, 55
				exit do
			ELSE
				type_row = type_row + 1
			END IF
		LOOP UNTIL type_does_hh_memb_exist = reference_number
	END IF
	transmit		'<===== when reference_number = "01" this transmit will navigate to PROG, else, it will navigate to STAT/WRAP

	IF reference_number = "01" THEN		'<===== only accesses PROG & REVW if reference_number = "01"
		call navigate_to_MAXIS_screen("STAT", "PROG")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 6, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 55)
			END IF
			IF type_fs_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 10, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 55)
			END IF
			IF type_hc_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 12, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 12, 55)
			END IF
			EMWriteScreen mig_worker, 18, 67
			transmit
			EMWriteScreen mig_worker, 18, 67
			transmit

		call navigate_to_MAXIS_screen("STAT", "REVW")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				cash_review_date = dateadd("YYYY", 1, appl_date)
				call create_MAXIS_friendly_date(cash_review_date, 0, 9, 37)
			END IF
			IF type_fs_yn = "Y" THEN
				EMWriteScreen "X", 5, 58
				transmit
				DO
					EMReadScreen food_support_reports, 20, 5, 30
				LOOP UNTIL food_support_reports = "FOOD SUPPORT REPORTS"
				fs_csr_date = dateadd("M", 6, appl_date)
				fs_er_date = dateadd("M", 12, appl_date)
				call create_MAXIS_friendly_date(fs_csr_date, 0, 9, 26)
				call create_MAXIS_friendly_date(fs_er_date, 0, 9, 64)
				transmit
			END IF
			IF type_hc_yn = "Y" THEN
				EMWriteScreen "X", 5, 71
				transmit
				DO
					EMReadScreen health_care_renewals, 20, 4, 32
				LOOP UNTIL health_care_renewals = "HEALTH CARE RENEWALS"
				IF revw_ar_or_ir = "AR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 71)
				ELSEIF revw_ar_or_ir = "IR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 27)
				END IF
				call create_MAXIS_friendly_date((dateadd("M", 12, appl_date)), 0, 9, 27)
				EMWriteScreen revw_exempt, 9, 71
				transmit
			END IF
	END IF
end function

function write_panel_to_MAXIS_UNEA(unea_number, unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ unea_number, unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "UNEA")
	PF10
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen unea_number, 20, 79
	transmit

	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		'Putting this part in with the NN because otherwise the script will update it in later months and change claim number information.
		EMWriteScreen unea_inc_type, 5, 37
		EMWriteScreen unea_inc_verif, 5, 65
		EMWriteScreen (ssn_first & ssn_mid & ssn_last & unea_claim_suffix), 6, 37
		call create_MAXIS_friendly_date(unea_start_date, 0, 7, 37)
	ELSE
		PF9
	END IF

	'=====Navigates to the PIC for UNEA=====
	EMWriteScreen "X", 10, 26
	transmit
	EMReadScreen pic_info_exists, 6, 18, 58		'---Deteremining if PIC info exists. If it does, the script will just back out.
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN
		EMWriteScreen unea_pay_freq, 5, 64
		EMWriteScreen unea_inc_amount, 8, 66
		calc_month = datepart("M", date)
		IF len(calc_month) = 1 THEN calc_month = "0" & calc_month
		calc_day = datepart("D", date)
		IF len(calc_day) = 1 THEN calc_day = "0" & calc_day
        calc_year = right( DatePart("yyyy",date), 2)
		EMWriteScreen calc_month, 5, 34
		EMWriteScreen calc_day, 5, 37
		EMWriteScreen calc_year, 5, 40
        Do              '<=====navigates out of the PIC
            transmit
            EmReadscreen PIC_Check, 16, 3, 28
            IF PIC_check <> "SNAP Prospective" then exit do
        Loop
	ELSE
		PF3
	END IF

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	current_bene_month = bene_month & "/01/" & bene_year
	retro_month = datepart("M", DateAdd("M", -2, current_bene_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(datepart("YYYY", DateAdd("M", -2, current_bene_month)), 2)

	EMWriteScreen retro_month, 13, 25
	EMWriteScreen retro_year, 13, 31
	EMWriteScreen bene_month, 13, 54
	EMWriteScreen bene_year, 13, 60

	IF pic_info_exists = "" THEN 	'---Meaning, the case has PIC info...which is to say that this is a PF9 and not a NN
		EMWriteScreen "05", 13, 28
		EMWriteScreen unea_inc_amount, 13, 39
		EMWriteScreen "05", 13, 57
		EMWriteScreen unea_inc_amount, 13, 68
	END IF

	IF unea_pay_freq = "2" OR unea_pay_freq = "3" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "19", 14, 28
			EMWriteScreen "19", 14, 57
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen unea_inc_amount, 14, 68
		END IF
	ELSEIF unea_pay_freq = "4" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen retro_month, 16, 25
		EMWriteScreen retro_year, 16, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen bene_year, 15, 60
		EMWriteScreen bene_month, 16, 54
		EMWriteScreen bene_year, 16, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "12", 14, 28
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen "19", 15, 28
			EMWriteScreen unea_inc_amount, 15, 39
			EMWriteScreen "26", 16, 28
			EMWriteScreen unea_inc_amount, 16, 39
			EMWriteScreen "12", 14, 57
			EMWriteScreen unea_inc_amount, 14, 68
			EMWriteScreen "19", 15, 57
			EMWriteScreen unea_inc_amount, 15, 68
			EMWriteScreen "26", 16, 57
			EMWriteScreen unea_inc_amount, 16, 68
		END IF
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", date) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to a useable number
		EMWriteScreen "X", 6, 56
		transmit
		EMWriteScreen "________", 9, 65
		EMWriteScreen unea_inc_amount, 9, 65
		EMWriteScreen unea_pay_freq, 10, 63
		transmit
		transmit
	END IF
	Transmit
  	EMReadScreen warning_check, 7, 24, 2 'This checks for an error with COLA field being blank
  	IF warning_check = "WARNING" THEN transmit
end function

function write_panel_to_MAXIS_WKEX(program, fed_tax_retro, fed_tax_prosp, fed_tax_verif, state_tax_retro, state_tax_prosp, state_tax_verif, fica_retro, fica_prosp, fica_verif, tran_retro, tran_prosp, tran_verif, tran_imp_rel, meals_retro, meals_prosp, meals_verif, meals_imp_rel, uniforms_retro, uniforms_prosp, uniforms_verif, uniforms_imp_rel, tools_retro, tools_prosp, tools_verif, tools_imp_rel, dues_retro, dues_prosp, dues_verif, dues_imp_rel, othr_retro, othr_prosp, othr_verif, othr_imp_rel, HC_Exp_Fed_Tax, HC_Exp_State_Tax, HC_Exp_FICA, HC_Exp_Tran, HC_Exp_Tran_imp_rel, HC_Exp_Meals, HC_Exp_Meals_Imp_Rel, HC_Exp_Uniforms, HC_Exp_Uniforms_Imp_Rel, HC_Exp_Tools, HC_Exp_Tools_Imp_Rel, HC_Exp_Dues, HC_Exp_Dues_Imp_Rel, HC_Exp_Othr, HC_Exp_Othr_Imp_Rel)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ program, fed_tax_retro, fed_tax_prosp, fed_tax_verif, state_tax_retro, state_tax_prosp, state_tax_verif, fica_retro, fica_prosp, fica_verif, tran_retro, tran_prosp, tran_verif, tran_imp_rel, meals_retro, meals_prosp, meals_verif, meals_imp_rel, uniforms_retro, uniforms_prosp, uniforms_verif, uniforms_imp_rel, tools_retro, tools_prosp, tools_verif, tools_imp_rel, dues_retro, dues_prosp, dues_verif, dues_imp_rel, othr_retro, othr_prosp, othr_verif, othr_imp_rel, HC_Exp_Fed_Tax, HC_Exp_State_Tax, HC_Exp_FICA, HC_Exp_Tran, HC_Exp_Tran_imp_rel, HC_Exp_Meals, HC_Exp_Meals_Imp_Rel, HC_Exp_Uniforms, HC_Exp_Uniforms_Imp_Rel, HC_Exp_Tools, HC_Exp_Tools_Imp_Rel, HC_Exp_Dues, HC_Exp_Dues_Imp_Rel, HC_Exp_Othr, HC_Exp_Othr_Imp_Rel: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "WKEX")
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits0

	EMWriteScreen reference_number, 20, 76
	transmit

	'Determining the number of WKEX panels so the script knows how to handle the incoming information.
	EMReadScreen num_of_WKEX_panels, 1, 2, 78
	IF num_of_WKEX_panels = "5" THEN		'If there are already 5 WKEX panels, the script will not create a new panel.
		EXIT function
	ELSEIF num_of_WKEX_panels = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit

		'---When the script needs to generate a new WKEX, it will enter the information for that panel...
		EMWriteScreen program, 5, 33
		EMWriteScreen fed_tax_retro, 7, 43
		EMWriteScreen fed_tax_prosp, 7, 57
		EMWriteScreen fed_tax_verif, 7, 69
		EMWriteScreen state_tax_retro, 8, 43
		EMWriteScreen state_tax_prosp, 8, 57
		EMWriteScreen state_tax_verif, 8, 69
		EMWriteScreen fica_retro, 9, 43
		EMWriteScreen fica_prosp, 9, 57
		EMWriteScreen fica_verif, 9, 69
		EMWriteScreen tran_retro, 10, 43
		EMWriteScreen tran_prosp, 10, 57
		EMWriteScreen tran_verif, 10, 69
		EMWriteScreen tran_imp_rel, 10, 75
		EMWriteScreen meals_retro, 11, 43
		EMWriteScreen meals_prosp, 11, 57
		EMWriteScreen meals_verif, 11, 69
		EMWriteScreen meals_imp_rel, 11, 75
		EMWriteScreen uniforms_retro, 12, 43
		EMWriteScreen uniforms_prosp, 12, 57
		EMWriteScreen uniforms_verif, 12, 69
		EMWriteScreen uniforms_imp_rel, 12, 75
		EMWriteScreen tools_retro, 13, 43
		EMWriteScreen tools_prosp, 13, 57
		EMWriteScreen tools_verif, 13, 69
		EMWriteScreen tools_imp_rel, 13, 75
		EMWriteScreen dues_retro, 14, 43
		EMWriteScreen dues_prosp, 14, 57
		EMWriteScreen dues_verif, 14, 69
		EMWriteScreen dues_imp_rel, 14, 75
		EMWriteScreen othr_retro, 15, 43
		EMWriteScreen othr_prosp, 15, 57
		EMWriteScreen othr_verif, 15, 69
		EMWriteScreen othr_imp_rel, 15, 75
	ELSE
		PF9
		'---If the script is editing an existing WKEX page, it would be doing so ONLY to update the HC Expense sub-menu.
		'---Adding to the HC Expenses
		EMWriteScreen "X", 18, 57
		transmit

		EMReadScreen current_month, 17, 20, 51
		IF current_month = "CURRENT MONTH + 1" THEN
			PF3
		ELSE
			EMWriteScreen HC_Exp_Fed_Tax, 8, 36
			EMWriteScreen HC_Exp_State_Tax, 9, 36
			EMWriteScreen HC_Exp_FICA, 10, 36
			EMWriteScreen HC_Exp_Tran, 11, 36
			EMWriteScreen HC_Exp_Tran_imp_rel, 11, 51
			EMWriteScreen HC_Exp_Meals, 12, 36
			EMWriteScreen HC_Exp_Meals_Imp_Rel, 12, 51
			EMWriteScreen HC_Exp_Uniforms, 13, 36
			EMWriteScreen HC_Exp_Uniforms_Imp_Rel, 13, 51
			EMWriteScreen HC_Exp_Tools, 14, 36
			EMWriteScreen HC_Exp_Tools_Imp_Rel, 14, 51
			EMWriteScreen HC_Exp_Dues, 15, 36
			EMWriteScreen HC_Exp_Dues_Imp_Rel, 15, 51
			EMWriteScreen HC_Exp_Othr, 16, 36
			EMWriteScreen HC_Exp_Othr_Imp_Rel, 16, 51
			transmit
			PF3
		END IF
	END IF
	transmit
end function

function write_panel_to_MAXIS_WREG(wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_sanction_reason, wreg_abawd_status, wreg_ga_basis)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_abawd_status, wreg_ga_basis: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "WREG")
	call create_panel_if_nonexistent

	EMWriteScreen wreg_fs_pwe, 6, 68
	EMWriteScreen wreg_fset_status, 8, 50
	EMWriteScreen wreg_defer_fs, 8, 80
	IF wreg_fset_orientation_date <> "" THEN call create_MAXIS_friendly_date(wreg_fset_orientation_date, 0, 9, 50)
    IF wreg_fset_sanction_date <> "" then
        sanc_mo = right("0" & DatePart("m",    wreg_fset_sanction_date), 2)
        sanc_yr = right(      DatePart("yyyy", wreg_fset_sanction_date), 2)
        EmWriteScreen sanc_mo, 10, 50
        EmWriteScreen sanc_yr, 10, 56
    End if

	IF wreg_num_sanctions <> "" THEN EMWriteScreen wreg_num_sanctions, 11, 50
    If wreg_sanction_reason <> "" THEN EmWriteScreen wreg_sanction_reason, 12, 50
	EMWriteScreen wreg_abawd_status, 13, 50
	EMWriteScreen wreg_ga_basis, 15, 50
	transmit
end function
