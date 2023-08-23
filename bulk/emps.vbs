'--------------------------------------------------------------------------------------STAT
name_of_script = "BULK - EMPS.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "100"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("07/27/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
EMConnect ""		'connecting to MAXIS

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 251, 120, "Pull EMPS data into Excel"
  	EditBox 75, 40, 160, 15, worker_number
  	CheckBox 10, 70, 150, 10, "Check here to run this query county-wide.", all_workers_check
  		ButtonGroup ButtonPressed
    	OkButton 140, 100, 50, 15
    CancelButton 195, 100, 50, 15
  	Text 10, 45, 65, 10, "Worker(s) to check:"
  	Text 10, 85, 235, 10, "NOTE: running queries county-wide may take several hours to complete"
  	Text 10, 10, 170, 10, "Enter workers' x1 numbers, separated by a comma."
  	Text 10, 25, 100, 10, "EX: X_ _ _ _ _ _, X_ _ _ _ _ _"
  	GroupBox 5, 0, 235, 60, ""
EndDialog
'Shows dialog
Do
	Do
		err_msg = ""
		Dialog Dialog1
		Cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then err_msg = err_msg & vbNewLine & "* Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
		If (all_workers_check = 1 AND trim(worker_number) <> "") then err_msg = err_msg & vbNewLine & "* Please enter x numbers OR the county-wide query."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case Number"
ObjExcel.Cells(1, 3).Value = "Client Name"
ObjExcel.Cells(1, 4).Value = "REF #"
ObjExcel.Cells(1, 5).Value = "Fin Orient Date"
ObjExcel.Cells(1, 6).Value = "Attended (Y/N)"
objExcel.cells(1, 7).Value = "Good Cause"
objExcel.cells(1, 8).Value = "Fin Orient Sanc Dt"
objExcel.cells(1, 9).Value = "Fin Orient Sanc Dt"
objExcel.cells(1, 10).Value = "Special Med Criteria"
objExcel.cells(1, 11).Value = "Ill/Incap Family Mbr"
objExcel.cells(1, 12).Value = "Personal/Family Crisis"
objExcel.cells(1, 13).Value = "Hard To Employ Cat"
objExcel.cells(1, 14).Value = "Full-Time Care Of Child < 1"
objExcel.cells(1, 15).Value = "Child < 1 Exemption Date"
objExcel.cells(1, 16).Value = "Reg MFIP-ES"
objExcel.cells(1, 17).Value = "ES Status"
objExcel.cells(1, 18).Value = "ES Referral Dt"
objExcel.cells(1, 19).Value = "18/19 Year Old"
objExcel.cells(1, 20).Value = "DWP Plan Date"
objExcel.cells(1, 21).Value = "Hrs/Week Work Act"
objExcel.cells(1, 22).Value = "Sanction Rsn"
objExcel.cells(1, 23).Value = "Sanc Beg Date"
objExcel.cells(1, 24).Value = "Sanc End Date"
objExcel.cells(1, 25).Value = "Other Provider Info"
objExcel.cells(1, 26).Value = "Tribal Code"

FOR i = 1 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 	'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'

NEXT

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	CALL create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas
	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self
  	Call navigate_to_MAXIS_screen("REPT", "MFCM")			'navigates to MFCM in the current footer month/year'
	EMWriteScreen worker, 21, 13
	transmit
	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason----'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
    has_content_check = trim(has_content_check)
	If has_content_check <> "" then
		Do
			MAXIS_row = 7	'Sets the row to start searching in MAXIS for
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6  	'Reading case number
				EMReadScreen client_name, 18, MAXIS_row, 16
				                'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
				If trim(MAXIS_case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number
					EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
                    MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'
                END IF

                If trim(MAXIS_case_number) = "" and trim(client_name) = "" then exit do			'Exits do if we reach the end

				'add case/case information to Excel
        		ObjExcel.Cells(excel_row, 1).Value = worker
        		ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
				ObjExcel.Cells(excel_row, 3).Value = trim(client_name)

			    excel_row = excel_row + 1	'moving excel row to next row'
				MAXIS_case_number = ""          'Blanking out variable
				MAXIS_row = MAXIS_row + 1	'adding one row to search for in MAXIS
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
NEXT

excel_row = 2           're-establishing the row to start checking the members for

Do
	MAXIS_case_number  = objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    client_name        = objExcel.cells(excel_row, 3).Value	're-establishing the client name to use for the case
	client_name = trim(client_name)
	If MAXIS_case_number = "" then exit do						'exits do if the case number is ""
	Call navigate_to_MAXIS_screen("REPT", "MFCM")

	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete				'row gets deleted since it will get added to the priv case list at end of script
		IF excel_row = 3 then
			excel_row = excel_row
		Else
			excel_row = excel_row - 1
		End if
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	Else
        EMReadScreen case_content, 7, 8, 7
	    If trim(case_content) = "" then
	       	'making sure we are getting the right person for cases where there are more than one case.
        	row = 7
        	Do
            	EMReadScreen case_name, 18, row, 16
	    		'msgbox case_name & vbcr & row
	       		case_name = trim(case_name)
            	If case_name <> client_name then row = row + 1
        	LOOP until case_name = client_name
	    	EMWriteScreen "x", row, 50
	    Else
	    	EMWriteScreen "x", 7, 50
	    End if

	    transmit

	    Do
	    	EMReadScreen EMPS_screen, 4, 2, 50
	    	If EMPS_screen <> "EMPS" then Transmit
	    Loop until EMPS_screen = "EMPS"

	    EMReadScreen memb_number, 2, 4, 33

	    'Attended Financial orientation code
	    EMReadScreen attended_orient, 1, 5, 65
	    If attended_orient = "_" then attended_orient = ""
	    'msgbox attended_orient

	    'Financial orientation date
	    EMReadScreen orient_date, 8, 5, 39
	    If orient_date = "__ __ __" then
	    	orient_date = ""
	    Else
	    	orient_date = replace(orient_date, " ", "/")
	    End if
	    'msgbox orient_date

	    'Good cause (EMPS_info variable)
	    EMReadScreen EMPS_good_cause, 2, 5, 79
	    IF 	EMPS_good_cause <> "__" then
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
	    ELSE
	    	EMPS_good_cause = replace(EMPS_good_cause, "__", "")
	    END IF
	    'msgbox EMPS_good_cause

	    'sanction dates (EMPS_info variable)
	    EMReadScreen EMPS_sanc_begin_date, 8, 18, 51
        If EMPS_sanc_begin_date = "__ 01 __" then
	    	EMPS_sanc_begin_date = ""
	    Else
	    	EMPS_sanc_begin_date = replace(EMPS_sanc_begin_date, " ", "/")
	    End if
	    'msgbox EMPS_sanc_begin_date

	    'sanction end date
	    EMReadScreen EMPS_sanc_end_date, 8, 18, 70
        If EMPS_sanc_end_date = "__ 01 __" then
	    	EMPS_sanc_end_date = ""
	    Else
	    	EMPS_sanc_end_date = replace(EMPS_sanc_end_date, " ", "/")
	    End if
	    'msgbox EMPS_sanc_end_date

	    'other sanction dates (ES_exemptions variable)--------------------------------------------------------------------------------
	    'special medical criteria
	    EMReadScreen EMPS_memb_at_home, 1, 8, 76
	    IF EMPS_memb_at_home <> "N" then
	    	If EMPS_memb_at_home = "1" then EMPS_memb_at_home = "Home-Health/Waiver service"
	    	IF EMPS_memb_at_home = "2" then EMPS_memb_at_home = "Child w/ severe emotional dist"
	    	IF EMPS_memb_at_home = "3" then EMPS_memb_at_home = "Adult/Serious Persistent MI"
	    END IF

	    EMReadScreen EMPS_care_family, 1, 9, 76
	    EMReadScreen EMPS_crisis, 1, 10, 76

	    'hard to employ
	    EMReadScreen EMPS_hard_employ, 2, 11, 76
	    IF EMPS_hard_employ <> "NO" then
	    	IF EMPS_hard_employ = "IQ" then EMPS_hard_employ = "IQ tested at < 80"
	    	IF EMPS_hard_employ = "LD" then EMPS_hard_employ = "Learning Disabled"
	    	IF EMPS_hard_employ = "MI" then EMPS_hard_employ = "Mentally ill"
	    	IF EMPS_hard_employ = "DD" then EMPS_hard_employ = "Dev Disabled"
	    	IF EMPS_hard_employ = "UN" then EMPS_hard_employ = "Unemployable"
	    END IF

	    'EMPS under 1 coding and dates used(ES_exemptions variable)
	    EMReadScreen EMPS_under1, 1, 12, 76
	    IF EMPS_under1 = "Y" then
	    	EMWriteScreen "x", 12, 39
	    	transmit
	    	MAXIS_row = 7
	    	MAXIS_col = 22
	      	DO
	    		EMReadScreen exemption_date, 9, MAXIS_row, MAXIS_col
	    		If trim(exemption_date) = "" then exit do
	      		If exemption_date <> "__ / ____" then
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
	    END IF

	    'Reading ES Information (for ES_info variable)
	    EMReadScreen ES_status, 40, 15, 40
	    ES_status = trim(ES_status)

	    EMReadScreen ES_referral_date, 8, 16, 40
	    If ES_referral_date = "__ __ __" then
	    	ES_referral_date = ""
	    Else
	    	ES_referral_date = replace(ES_referral_date, " ", "/")
	    End if

	    EMReadScreen DWP_plan_date, 8, 17, 40
	    If DWP_plan_date = "__ __ __" then
	    	DWP_plan_date = ""
	    Else
	    	DWP_plan_date = replace(DWP_plan_date, " ", "/")
	    End if

	    EMReadScreen minor_ES_option, 2, 16, 76
	    IF minor_ES_option <> "__" then
	    		IF minor_ES_option = "SC" then minor_ES_option = "Secondary Education"
	    		IF minor_ES_option = "EM" then minor_ES_option = "Employment"
	    ELSE
	    		IF minor_ES_option = "__" then minor_ES_option = ""
	    END if

	    EMReadScreen Tribal_Code, 2, 19, 70
	    If Tribal_Code = "__" then
	    	Tribal_Code = ""
	    Else
	    	Tribal_Code = replace(Tribal_Code, " ", "/")
	    End if
	    'reading for Provider'
	    Do
	    	EMWriteScreen "x", 19, 25
	    	Transmit
	    	EMReadScreen OT_screen, 5, 4, 30
	    	If OT_screen = "Other" then exit do
	    Loop until OT_screen = "Other"

	    EMReadScreen Other_Prov_Info, 35, 6, 37
	    Other_Prov_Info = replace(Other_Prov_Info, "_", "")
	    If trim(Other_Prov_Info) = ""  then Other_Prov_Info = ""
	    PF3

	    ObjExcel.Cells(excel_row, 4).Value = trim(memb_number)
	    ObjExcel.Cells(excel_row, 5).Value = trim(orient_date)
	    ObjExcel.Cells(excel_row, 6).Value = trim(attended_orient)
	    objExcel.cells(excel_row, 7).Value = trim(EMPS_good_cause)
	    objExcel.cells(excel_row, 8).Value = trim(EMPS_sanc_begin_date)
	    objExcel.cells(excel_row, 9).Value = trim(EMPS_sanc_end)
	    objExcel.cells(excel_row, 10).Value = trim(EMPS_memb_at_home)
	    objExcel.cells(excel_row, 11).Value = trim(EMPS_care_family)
	    objExcel.cells(excel_row, 12).Value = trim(EMPS_crisis)
	    objExcel.cells(excel_row, 13).Value = trim(EMPS_hard_employ)
	    objExcel.cells(excel_row, 14).Value = trim(EMPS_under1)
	    objExcel.cells(excel_row, 15).Value = trim(exemption_date)
	    objExcel.cells(excel_row, 16).Value = trim(Return_Regular_MFIP_ES)
	    objExcel.cells(excel_row, 17).Value = trim(ES_Status)
	    objExcel.cells(excel_row, 18).Value = trim(ES_referral_date)
	    objExcel.cells(excel_row, 19).Value = trim(minor_ES_option)
	    objExcel.cells(excel_row, 20).Value = trim(DWP_plan_date)
	    objExcel.cells(excel_row, 21).Value = trim(Hrs_Week_Work_Activity)
	    objExcel.cells(excel_row, 22).Value = trim(Sanction_Rsn)
	    objExcel.cells(excel_row, 23).Value = trim(Sanc_Beg_Date)
	    objExcel.cells(excel_row, 24).Value = trim(Sanc_End_Date)
	    objExcel.cells(excel_row, 25).Value = trim(Other_Prov_Info)
	    objExcel.cells(excel_row, 26).Value = trim(Tribal_Code)

	    excel_row = excel_row + 1
	    STATS_counter = STATS_counter + 1
	End if
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

IF priv_case_list <> "" then
	'Creating the list of privileged cases and adding to the spreadsheet
	excel_row = 2				'establishes the row to start writing the PRIV cases to
	objExcel.cells(1, 27).Value = "PRIV cases"

	prived_case_array = split(priv_case_list, "|")

	FOR EACH MAXIS_case_number in prived_case_array
		If trim(MAXIS_case_number) <> "" then
			objExcel.cells(excel_row, 8).value = MAXIS_case_number		'inputs cases into Excel
			excel_row = excel_row + 1								'increases the row
		End if
	NEXT
End if

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
'Query date/time/runtime info
ObjExcel.Cells(1, 28).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, 28).Font.Bold = TRUE
ObjExcel.Cells(2, 28).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, 28).Font.Bold = TRUE
ObjExcel.Cells(3, 28).Value = "Case count:"	'Goes back one, as this is on the next row
objExcel.Cells(3, 28).Font.Bold = TRUE
ObjExcel.Cells(1, 29).Value = now
ObjExcel.Cells(2, 29).Value = timer - query_start_time
ObjExcel.Cells(3, 29).Value = STATS_counter

FOR i = 1 to 29		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

script_end_procedure("Success! Please review the list generated.")
