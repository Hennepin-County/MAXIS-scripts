'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - INTERVIEW TEAM CASES WORKLIST.vbs"
start_time = timer
STATS_counter = 0			     'sets the stats counter at one
STATS_manualtime = 	90			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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

If script_repository = "" THEN script_repository = "C:\MAXIS-Scripts\"

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/22/2025", "Worklist can be created for cases interviewed the current day.##~## ##~##If the same day selection is made the file will be saved with a number at the end and there will be multiple worklists for the interview day(s) selected.##~##", "Casey Love, Hennepin County")
call changelog_update("01/13/2025", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================

Call run_from_GitHub(script_repository & "misc\caseload-directory.vbs")

curr_month = MonthName(DatePart("M",date))
curr_day = DatePart("d",date)
curr_year = DatePart("yyyy", date)

manager_log_file_path 				= "https://hennepin.sharepoint.com/teams/InterviewPhoneHSRs-Supervisors/Shared%20Documents/Management%20Team/Manager%20Log%20" & curr_year & "%20" & curr_month & ".xlsx"
staff_assignment_log_file_path 		= "https://hennepin.sharepoint.com/teams/InterviewPhoneHSRs-EligibilityHSRs/Shared%20Documents/Eligibility%20HSRs/Eligibility%20Staff%20" & curr_year & "%20" & curr_month & ".xlsx"

template_manager_log_file_path 			= t_drive & "\Eligibility Support\Assignments\Interview Team Cases Worklists\Support Documents\Template Manager Log.xlsx"
template_staff_assignment_log_file_path = t_drive & "\Eligibility Support\Assignments\Interview Team Cases Worklists\Support Documents\Eligibility Staff Template.xlsx"

Const name_const            = 00
Const email_const           = 01
Const sheet_name_const      = 02
Const case_count_const      = 03
Const mx_name_const         = 04
Const mx_id_const           = 05
Const secondary_pm          = 06

Const pm_end_const             = 10

Dim PM_ARRAY()
ReDim PM_ARRAY(pm_end_const, 0)

pm_counter = 0
ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Doryan Clifton"
PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
PM_ARRAY(email_const, pm_counter)           = "Doryan.Clifton@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "D Clifton"
PM_ARRAY(mx_name_const, pm_counter)         = "CLIFTON,DORYAN C."
PM_ARRAY(mx_id_const, pm_counter)           = "X127U40"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Tammy Coenen"
PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
PM_ARRAY(email_const, pm_counter)           = "Tammy.Coenen@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "T Coenen"
PM_ARRAY(mx_name_const, pm_counter)         = "COENEN,TAMMY L."
PM_ARRAY(mx_id_const, pm_counter)           = "X127AY1"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

'THIS PROGRAM MANAGER IS REMOVED BECAUSE INTERVIEW PROCESSING IS NOT ASSIGNED TO THEM DIRECTLY - Cases will be listed on UNKNOWN
' ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
' PM_ARRAY(name_const, pm_counter)            = "Brianna Fleeman Schneider"
' PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
' PM_ARRAY(email_const, pm_counter)           = "Brianna.FleemanSchneider@hennepin.us"
' PM_ARRAY(sheet_name_const, pm_counter)      = "B Fleeman Schneider"
' PM_ARRAY(mx_name_const, pm_counter)         = "FLEEMAN SCHNEIDER,BRIANNA J."
' PM_ARRAY(mx_id_const, pm_counter)           = "X1275K4"
' PM_ARRAY(case_count_const, pm_counter)      = 0
' pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Twanda Garrett"
PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
PM_ARRAY(email_const, pm_counter)           = "Twanda.Garrett@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "T Garrett"
PM_ARRAY(mx_name_const, pm_counter)         = "GARRETT,TWANDA T."
PM_ARRAY(mx_id_const, pm_counter)           = "X1273L5"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Lisa Groves"
PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
PM_ARRAY(email_const, pm_counter)           = "Lisa.Groves@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "L Groves"
PM_ARRAY(mx_name_const, pm_counter)         = "GROVES,LISA K."
PM_ARRAY(mx_id_const, pm_counter)           = "X127Y99"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Monique Moore"
PM_ARRAY(secondary_pm, pm_counter)          = "Tenzing Yarphel"
PM_ARRAY(email_const, pm_counter)           = "Monique.Moore@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "M Moore"
PM_ARRAY(mx_name_const, pm_counter)         = "MOORE,MONIQUE L."
PM_ARRAY(mx_id_const, pm_counter)           = "X127573"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Ann Noeker"
PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
PM_ARRAY(email_const, pm_counter)           = "Ann.Noeker@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "A Noeker"
PM_ARRAY(mx_name_const, pm_counter)         = "NOEKER,ANN M."
PM_ARRAY(mx_id_const, pm_counter)           = "X127C1K"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

'THIS PROGRAM MANAGER IS REMOVED BECAUSE INTERVIEW PROCESSING IS NOT ASSIGNED TO THEM DIRECTLY - Cases will be listed on UNKNOWN
' ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
' PM_ARRAY(name_const, pm_counter)            = "Jackie Poidinger"
' PM_ARRAY(secondary_pm, pm_counter)          = "NONE"
' PM_ARRAY(email_const, pm_counter)           = "Jackie.Poidinger@hennepin.us"
' PM_ARRAY(sheet_name_const, pm_counter)      = "J Poidinger"
' PM_ARRAY(mx_name_const, pm_counter)         = "POIDINGER,JACKIE A."
' PM_ARRAY(mx_id_const, pm_counter)           = "X127R61"
' PM_ARRAY(case_count_const, pm_counter)      = 0
' pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "Faughn Ramisch-Church"
PM_ARRAY(email_const, pm_counter)           = "Faughn.Ramisch-Church@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "F Ramisch-Church"
PM_ARRAY(mx_name_const, pm_counter)         = "RAMISCH-CHURCH,FAUGHN"
PM_ARRAY(mx_id_const, pm_counter)           = "X127Z84"
PM_ARRAY(case_count_const, pm_counter)      = 0
pm_counter = pm_counter + 1

'THIS PROGRAM MANAGER IS REMOVED BECAUSE INTERVIEW PROCESSING IS NOT ASSIGNED TO THEM DIRECTLY - Cases are assigned by Monique Moore
' ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
' PM_ARRAY(name_const, pm_counter)            = "Tenzing Yarphel"
' PM_ARRAY(email_const, pm_counter)           = "Tenzing.Yarphel@hennepin.us"
' PM_ARRAY(sheet_name_const, pm_counter)      = "T Yarphel"
' PM_ARRAY(mx_name_const, pm_counter)         = "YARPHEL,TENZING N."
' PM_ARRAY(mx_id_const, pm_counter)           = "X127Y97"
' PM_ARRAY(case_count_const, pm_counter)      = 0
' pm_counter = pm_counter + 1

ReDim Preserve PM_ARRAY(pm_end_const, pm_counter)
PM_ARRAY(name_const, pm_counter)            = "UNKNOWN"
PM_ARRAY(email_const, pm_counter)           = "HSPH.EWS.BlueZoneScripts@hennepin.us; HSPH.EWS.Unit.Coenen@hennepin.us"
PM_ARRAY(sheet_name_const, pm_counter)      = "UNKNOWN"
PM_ARRAY(mx_name_const, pm_counter)         = "UNKNOWN"
PM_ARRAY(mx_id_const, pm_counter)           = ""
PM_ARRAY(case_count_const, pm_counter)      = 0

end_message = ""
'END DECLARATIONS ==========================================================================================================


'THE SCRIPT ===========================================================================================================
'Now we are ready to start the script

'Dialog is just to alert user that the script starts with information gathering and takes time.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 115, "Interview Team Cases Worklist"
  ButtonGroup ButtonPressed
    OkButton 125, 90, 50, 15
  Text 10, 10, 175, 25, "This script will create a list of cases with interviews completed by the Interview Team. It will also update the files used by the processing team to assign work. "
  Text 10, 45, 165, 20, "First the script needs to gather some information, which may take a minute or two. "
  Text 10, 75, 145, 10, "Please be patient, the script is running!"
  Text 10, 90, 105, 10, "Press OK to start info gather."
EndDialog

'no options on the dialog so no loop is necessary
dialog Dialog1
cancel_confirmation

'FILE MANAGEMENT ===========================================================================================================
'This section ensures the files needed for the run are up available and up to date.

'Setting the folder paths and objects to handle folder and file manipulation
interview_team_cases_folder = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage"
Set objFolder = objFSO.GetFolder(interview_team_cases_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

interview_team_cases_already_on_worklist = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List"
worklist_folder = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Interview Team Cases Worklists"
pm_assign_folder = t_drive & "\Eligibility Support\Assignments\Interview Team PM Assignments"

'constants for the array of all dates - this allows selection of cases by date
Const date_const 	= 01
Const month_const 	= 02
Const day_const		= 03
Const count_const	= 04
Const checkbox_const = 05
const added_count	= 06
const fam_count 	= 07
const adul_count	= 08
Const dates_last_const	= 20

'This array gathers all the dates but is not sorted
Dim TEMP_ARRAY()
ReDim TEMP_ARRAY(dates_last_const, 0)

'We will use this array to sort the original array in date order
Dim DATES_WITH_INTERVIEWS_ARRAY()
ReDim DATES_WITH_INTERVIEWS_ARRAY(dates_last_const, 0)

'setting some initial variables for array use
each_month = 0
Total_count = 0
earliest_date = ""

'Looking at each txt file in the assignments folder to capture Expedited Determination information
For Each objFile in colFiles							'looping through each file
	file_created_date = objFile.DateCreated				'Reading the date created
	Total_count = Total_count + 1						'count all cases

	'creating some better formatting of dates (date created parameter has time associated and causes issues)
	intvw_month = DatePart("m", file_created_date)
	intvw_day = DatePart("d", file_created_date)
	intvw_year = DatePart("yyyy", file_created_date)
	file_date = intvw_month & "/" & intvw_day & "/" & intvw_year
	file_date = DateAdd("d", 0, file_date)

	'seeing if we already have this date in our list of dates
	date_found = False
	For chkn_wg = 0 to UBound(TEMP_ARRAY, 2)
		If file_date = TEMP_ARRAY(date_const, chkn_wg) Then
			TEMP_ARRAY(count_const, chkn_wg) = TEMP_ARRAY(count_const, chkn_wg) + 1
			date_found = True
			Exit For
		End If
	Next

	'If the date was not found, we will add it to the list of dates
	If date_found = False Then
		ReDim Preserve TEMP_ARRAY(dates_last_const, each_month)
		TEMP_ARRAY(date_const, 	each_month) = file_date
		TEMP_ARRAY(count_const, each_month) = 1
		TEMP_ARRAY(month_const, each_month) = intvw_month
		TEMP_ARRAY(day_const, 	each_month) = intvw_day

		each_month = each_month + 1

		'Need to define the first date in the list for easy sorting later
		If earliest_date = "" Then
			earliest_date = file_date
		Else
			If DateDiff("d", earliest_date, file_date) < 0 Then earliest_date = file_date
		End If
	End If
Next

'checking to ensure there are some cases to create a worklist with from a day prior to today
If earliest_date = "" Then Call script_end_procedure("There does not appear to be any outstanding cases from interviews on previous days. To check for previous worklists reference the script instructions for the worklist folder.")

'Loop through each day from the earliest day forward to put the dates in order from oldest to newest
date_to_assess = earliest_date
tomorrow = DateAdd("d", 1, date)
cow = 0
Do While DateDiff("d", tomorrow, date_to_assess) <> 0			'we stop when we get to today and DO NOT include today
	For pig = 0 to UBound(TEMP_ARRAY, 2)					'go through the temp array to find a match
		If DateDiff("d", TEMP_ARRAY(date_const, pig), date_to_assess) = 0 Then
			ReDim Preserve DATES_WITH_INTERVIEWS_ARRAY(dates_last_const, cow)					'Add the date with all parameters to the new array
			DATES_WITH_INTERVIEWS_ARRAY(date_const, 	cow) = TEMP_ARRAY(date_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(count_const, 	cow) = TEMP_ARRAY(count_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(month_const, 	cow) = TEMP_ARRAY(month_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(day_const, 		cow) = TEMP_ARRAY(day_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, cow) = unchecked
			cow = cow + 1
		End If
	Next
	date_to_assess = DateAdd("d", 1, date_to_assess)		'incrementing the date we are assessing for the loop
Loop

'Now we can display all of the dates with case count in a dialog for selection.
dlg_len = 80
For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
	dlg_len = dlg_len + 15
Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 316, dlg_len, "Select Dates for Worklist"
	Text 10, 35, 195, 10, "To create a worklist, select the date(s) to include:"
	Text 10, 10, 250, 20, "All interview team recorded cases have been checked and there appear to be " & Total_count & " cases that have not yet been added to a worklist."
	Text 20, 50, 40, 15, "Check Here to Select"
	Text 80, 60, 25, 10, "Date"
	Text 140, 60, 35, 10, "Count"
	y_pos = 70
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		CheckBox 30, y_pos, 30, 10, DATES_WITH_INTERVIEWS_ARRAY(month_const, chkn_wg) & "/" & DATES_WITH_INTERVIEWS_ARRAY(day_const, chkn_wg), DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg)
		Text 80, y_pos, 50, 10, DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg)
		Text 145, y_pos, 50, 10, DATES_WITH_INTERVIEWS_ARRAY(count_const, chkn_wg)
		y_pos = y_pos + 15
	Next
	ButtonGroup ButtonPressed
		OkButton 200, dlg_len-25, 50, 15
		CancelButton 255, dlg_len-25, 50, 15
EndDialog

Do
	err_msg = ""
	dialog Dialog1
	cancel_confirmation

	'Only requirement is to select at least one date.
	date_selected = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked then date_selected = True
	Next

	If date_selected = False Then err_msg = err_msg & vbCr & "* Select at least 1 date."
	If err_msg <> "" Then MsgBox "* * * NOTICE * * * " & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
Loop until err_msg = ""

developer_mode = False
admin_run = False
If windows_user_ID = "CALO001" Then admin_run = True
If windows_user_ID = "DACO003" Then admin_run = True
If windows_user_ID = "MARI001" Then admin_run = True
If windows_user_ID = "TRFA001" Then admin_run = True

If admin_run Then
	run_in_dev = MsgBox("Do you want to run in developer mode?", vbQuestion + vbYesNo, "Developer Mode")
	If run_in_dev = vbYes Then developer_mode = True
End If

const x_numb_const		= 00
const case_numb_const	= 01
const population_const	= 02
const pm_const			= 03
const pm_index_const	= 04
const intvw_date_const	= 05
const appears_xfs_const	= 06
const cash_select_const	= 07
const cash_prog_const	= 08
const grh_select_const	= 09
const snap_select_const	= 10
const emer_select_const	= 11
const assigned_to_const	= 13
const mngr_excel_row	= 14
const file_path_const 	= 15
const file_name_const 	= 16
const end_const			= 20

DIM CASES_TO_ASSIGN_ARRAY()
ReDIM CASES_TO_ASSIGN_ARRAY(end_const, 0)
intvws_count = 0
case_count = 0

'creating some objects needed for XML handling
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set xml = CreateObject("Msxml2.DOMDocument")

xml_case_numbs = "*"
'Creating an object for the stream of text which we'll use frequently
Dim objTextStream

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_name = objFile.Name
	file_created_date = objFile.DateCreated					'Reading the date created

	'determining if the XML is for a file in the dates selected by the user
	save_file = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DateDiff("d", DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg), file_created_date) = 0 Then
			If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then save_file = True
		End If
	Next

	'if this is from a date selected, we need to read details of the file for the worklist
	If save_file = True Then
		STATS_counter = STATS_counter + 1
		xmlPath = objFile.Path												'identifying the current file
		With (CreateObject("Scripting.FileSystemObject"))
			Set file_object = CreateObject("Scripting.FileSystemObject")										'Create another FSO
			Set xml_sig_command = file_object.OpenTextFile(xmlPath)			'Open the text file

			name_of_xml = file_object.GetFileName(xmlPath)
			If InStr(name_of_xml, "details") Then
				If xml_sig_command.AtEndOfStream Then
					attachment_here = ""
					Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", "XML in Interview Team Tracking is BLANK", 1, False, "", "", False, "", "XML File Path: " & xmlPath, True, attachment_here, True)
					xml_sig_command.Close
				ElseIf .FileExists(xmlPath) = True then
					xml_sig_command.Close
					xmlDoc.Async = False

					' Load the XML file
					xmlDoc.load(xmlPath)

					'reads data about the case from the XML
					set node = xmlDoc.SelectSingleNode("//CaseBasket")
					case_basket_numb = node.text

					set node = xmlDoc.SelectSingleNode("//CaseNumber")
					MAXIS_case_number = node.text
					xml_case_numbs = xml_case_numbs & trim(MAXIS_case_number) & "*"

					set node = xmlDoc.SelectSingleNode("//InterviewDate")
					interview_date = node.text
					interview_date = DateAdd("d", 0, interview_date)

					set node = xmlDoc.SelectSingleNode("//ScriptRunDate")
					script_date = node.text
					script_date = DateAdd("d", 0, script_date)

					set node = xmlDoc.SelectSingleNode("//CASHRequest")
					req_cash = node.text
					req_cash = req_cash * 1

					cash_type = ""
					If req_cash = True Then
						set node = xmlDoc.SelectSingleNode("//TypeOfCASH")
						cash_type = node.text
					End If

					If DateDiff("d", #1/13/2025#, script_date) > 0 Then
						set node = xmlDoc.SelectSingleNode("//GRHRequest")
						req_grh = node.text
						req_grh = req_grh * 1
					End If

					set node = xmlDoc.SelectSingleNode("//SNAPRequest")
					req_snap = node.text
					req_snap = req_snap * 1

					set node = xmlDoc.SelectSingleNode("//EMERRequest")
					req_emer = node.text
					req_emer = req_emer * 1

					set node = xmlDoc.SelectSingleNode("//ExpeditedDetermination")
					exp_det = node.text
					If exp_det <> "" Then exp_det = exp_det * 1

					case_basket_numb = UCase(case_basket_numb)
                    population = ""
					If caseload_info.Exists(case_basket_numb) Then
						population = caseload_info.Item(case_basket_numb)
					Else
						population = "UNKNOWN"
					End If

                    'Compare case basket to caseload list to identify Program Manager
                    MX_Program_Manager = "UNKNOWN"
					If caseload_manager.Exists(population) Then
						MX_Program_Manager = caseload_manager.Item(population)
					End If

                    If IsNumeric(right(population, 1)) and left(right(population, 2),1) = " " Then
                        population = left(population, len(population) - 2)
                    End If

					ReDim preserve CASES_TO_ASSIGN_ARRAY(end_const, intvws_count)
					CASES_TO_ASSIGN_ARRAY(x_numb_const, intvws_count) 		= case_basket_numb
                    CASES_TO_ASSIGN_ARRAY(pm_const, intvws_count) 		    = MX_Program_Manager
                    CASES_TO_ASSIGN_ARRAY(pm_index_const, intvws_count)     = ""
					CASES_TO_ASSIGN_ARRAY(case_numb_const, intvws_count) 	= MAXIS_case_number
					CASES_TO_ASSIGN_ARRAY(population_const, intvws_count) 	= population
					CASES_TO_ASSIGN_ARRAY(intvw_date_const, intvws_count) 	= interview_date
					If exp_det = True Then CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count) 	= True
					If exp_det <> True Then CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count) 	= False
					CASES_TO_ASSIGN_ARRAY(cash_select_const, intvws_count) 	= req_cash
					If req_cash = True Then CASES_TO_ASSIGN_ARRAY(cash_prog_const, intvws_count) 	= cash_type
					CASES_TO_ASSIGN_ARRAY(grh_select_const, intvws_count) 	= req_grh
					CASES_TO_ASSIGN_ARRAY(snap_select_const, intvws_count) 	= req_snap
					CASES_TO_ASSIGN_ARRAY(emer_select_const, intvws_count) 	= req_emer
					CASES_TO_ASSIGN_ARRAY(assigned_to_const, intvws_count) 	= ""
					CASES_TO_ASSIGN_ARRAY(file_path_const, intvws_count) 	= xmlPath
					CASES_TO_ASSIGN_ARRAY(file_name_const, intvws_count) 	= file_name

					'TODO - If we want to add information to the XML, we need to read it all and then rewrite it all to the xml doc - when we save it, the file will be overwritten

					case_count = case_count + 1
					intvws_count = intvws_count + 1

				End If
			End If
		End With
	End If
Next

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_name = objFile.Name
	file_created_date = objFile.DateCreated					'Reading the date created

	'determining if the XML is for a file in the dates selected by the user
	save_file = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DateDiff("d", DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg), file_created_date) = 0 Then
			If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then save_file = True
		End If
	Next

	'if this is from a date selected, we need to read details of the file for the worklist
	If save_file = True Then
		STATS_counter = STATS_counter + 1
		xmlPath = objFile.Path												'identifying the current file
		With (CreateObject("Scripting.FileSystemObject"))
			' 'Creating an object for the stream of text which we'll use frequently
			' Dim objTextStream

			Set file_object = CreateObject("Scripting.FileSystemObject")										'Create another FSO
			Set xml_sig_command = file_object.OpenTextFile(xmlPath)			'Open the text file

			name_of_xml = file_object.GetFileName(xmlPath)
			If InStr(name_of_xml, "started") Then
				If xml_sig_command.AtEndOfStream Then
					attachment_here = ""
					Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", "XML in Interview Team Tracking is BLANK", 1, False, "", "", False, "", "XML File Path: " & xmlPath, True, attachment_here, True)
					xml_sig_command.Close
				ElseIf .FileExists(xmlPath) = True then
					xml_sig_command.Close
					xmlDoc.Async = False

					' Load the XML file
					xmlDoc.load(xmlPath)

					'reads data about the case from the XML
					set node = xmlDoc.SelectSingleNode("//CaseBasket")
					case_basket_numb = node.text

					set node = xmlDoc.SelectSingleNode("//CaseNumber")
					MAXIS_case_number = node.text
					search_numb = "*" & trim(MAXIS_case_number) & "*"

					set node = xmlDoc.SelectSingleNode("//ScriptRunDate")
					script_date = node.text
					script_date = DateAdd("d", 0, script_date)

					case_basket_numb = UCase(case_basket_numb)
                    population = ""
					If caseload_info.Exists(case_basket_numb) Then
						population = caseload_info.Item(case_basket_numb)
					Else
						population = "UNKNOWN"
					End If

                    'Compare case basket to caseload list to identify Program Manager
                    MX_Program_Manager = "UNKNOWN"
					If caseload_manager.Exists(population) Then
						MX_Program_Manager = caseload_manager.Item(population)
					End If

                    If IsNumeric(right(population, 1)) and left(right(population, 2),1) = " " Then
                        population = left(population, len(population) - 2)
                    End If

					If InStr(xml_case_numbs, search_numb) = 0 Then
						ReDim preserve CASES_TO_ASSIGN_ARRAY(end_const, intvws_count)
						CASES_TO_ASSIGN_ARRAY(x_numb_const, intvws_count) 		= case_basket_numb
                        CASES_TO_ASSIGN_ARRAY(pm_const, intvws_count) 		    = MX_Program_Manager
                        CASES_TO_ASSIGN_ARRAY(pm_index_const, intvws_count)     = ""
						CASES_TO_ASSIGN_ARRAY(case_numb_const, intvws_count) 	= MAXIS_case_number
						CASES_TO_ASSIGN_ARRAY(population_const, intvws_count) 	= population
						CASES_TO_ASSIGN_ARRAY(intvw_date_const, intvws_count) 	= script_date
						CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count)	= "?"
    					CASES_TO_ASSIGN_ARRAY(assigned_to_const, intvws_count) 	= ""
						CASES_TO_ASSIGN_ARRAY(file_path_const, intvws_count) 	= xmlPath
						CASES_TO_ASSIGN_ARRAY(file_name_const, intvws_count) 	= file_name

						'TODO - If we want to add information to the XML, we need to read it all and then rewrite it all to the xml doc - when we save it, the file will be overwritten

						case_count = case_count + 1
						intvws_count = intvws_count + 1
					ElseIf developer_mode = False Then
						'moving each file to the folder for cases already in a worklist
						.MoveFile xmlPath , interview_team_cases_already_on_worklist & "\" & file_name
						' .DeleteFile(xmlPath)
					End If
				End If
			End If
		End With
	End If
Next
set xmlDoc = nothing

'Saves and closes the Excel Worklist - Naming convention is 'Interview Team Cases Worklist from MM-DD_MM-DD_MM-DD.xlsx' with all interview dates listed
' formatted_date = replace(date, "/", "-")
list_of_dates = " "
For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
    If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then list_of_dates = list_of_dates & DATES_WITH_INTERVIEWS_ARRAY(month_const, chkn_wg) & "-" & DATES_WITH_INTERVIEWS_ARRAY(day_const, chkn_wg) & " "
Next
list_of_dates = replace(trim(list_of_dates), " ", "_")


end_message = "Interviews from " & list_of_dates & " have been added to a worklist."
end_message = end_message & vbCr & "Cases found with an interview completed that needs processing: " & case_count

For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
    MX_Program_Manager = CASES_TO_ASSIGN_ARRAY(pm_const, cow)
    PM_Found = False
    For duck = 0 to Ubound(PM_ARRAY, 2)
        If MX_Program_Manager = PM_ARRAY(name_const, duck) or MX_Program_Manager = PM_ARRAY(secondary_pm, duck) Then
            CASES_TO_ASSIGN_ARRAY(pm_index_const, cow) = duck
            PM_ARRAY(case_count_const, duck) = PM_ARRAY(case_count_const, duck) + 1
            PM_Found = True
            Exit For
        End If
    Next
    If PM_Found = False Then
        'assign to UNKNOWN PM
        For duck = 0 to Ubound(PM_ARRAY, 2)
            If PM_ARRAY(name_const, duck) = "UNKNOWN" Then
                CASES_TO_ASSIGN_ARRAY(pm_index_const, cow) = duck
                PM_ARRAY(case_count_const, duck) = PM_ARRAY(case_count_const, duck) + 1
                Exit For
            End If
        Next
    End If
Next

'Here is the worklist creation section
'these constants are to document the columns  - using a constant supports future changes
Const basket_col		= 1
Const case_numb_col 	= 2
Const population_col 	= 3
Const prog_mngr_col 	= 4
Const intvw_date_col 	= 5
Const exp_det_col 		= 6
Const assignment_col	= 7
Const cash_col 			= 8
Const cash_type_col		= 9
Const grh_col 			= 10
Const snap_col 			= 11
Const emer_col 			= 12


'Creating the Excel file
Set objPMExcel = CreateObject("Excel.Application")
objPMExcel.Visible = True
Set objPMWorkbook = objPMExcel.Workbooks.Add()
objPMExcel.DisplayAlerts = True

'Setting the first row with header information
objPMExcel.Cells(1, basket_col).Value = "CASELOAD"
objPMExcel.Cells(1, case_numb_col).Value = "CASE NUMBER"
objPMExcel.Cells(1, population_col).Value = "POPULATION"
objPMExcel.Cells(1, intvw_date_col).Value = "INTERVIEW COMPLETED"
objPMExcel.Cells(1, exp_det_col).Value = "APPEARS EXPEDITED"
objPMExcel.Cells(1, assignment_col).Value = "ASSIGNED TO"

For duck = 0 to UBound(PM_ARRAY, 2)
    If duck = 0 Then objPMExcel.ActiveSheet.Name = PM_ARRAY(sheet_name_const, duck)
    If duck <> 0 Then objPMExcel.Worksheets.Add().Name = PM_ARRAY(sheet_name_const, duck)

    'Setting the first row with header information
    objPMExcel.Cells(1, basket_col).Value = "CASELOAD"
    objPMExcel.Cells(1, case_numb_col).Value = "CASE NUMBER"
    objPMExcel.Cells(1, population_col).Value = "POPULATION"
    objPMExcel.Cells(1, prog_mngr_col).Value = "PROGRAM MANAGER"
    objPMExcel.Cells(1, intvw_date_col).Value = "INTERVIEW COMPLETED"
    objPMExcel.Cells(1, exp_det_col).Value = "APPEARS EXPEDITED"
    objPMExcel.Cells(1, assignment_col).Value = "ASSIGNED TO"


    excel_row = 2
    For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
        If CASES_TO_ASSIGN_ARRAY(pm_index_const, cow) = duck Then

            objPMExcel.Cells(excel_row, basket_col).Value 		= CASES_TO_ASSIGN_ARRAY(x_numb_const, cow)
            objPMExcel.Cells(excel_row, case_numb_col).Value 	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
            objPMExcel.Cells(excel_row, population_col).Value 	= CASES_TO_ASSIGN_ARRAY(population_const, cow)
            objPMExcel.Cells(excel_row, prog_mngr_col).Value    = CASES_TO_ASSIGN_ARRAY(pm_const, cow)
            objPMExcel.Cells(excel_row, intvw_date_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)

            If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "details") Then
                If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then objPMExcel.Cells(excel_row, exp_det_col).Value = "Yes"
            End If

            If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "started") Then
                objPMExcel.Cells(excel_row, exp_det_col).Value = "?"
            End If

            excel_row = excel_row + 1
        End If
    Next
    If excel_row = 2 Then
        objPMExcel.Cells(2, 1).Value = "No cases assigned to this Program Manager."
    End If
    'format the Worklist Excel
    For col_to_autofit = 1 to emer_col
        objPMExcel.columns(col_to_autofit).AutoFit()
    Next

Next

Set FSOxl = CreateObject("Scripting.FileSystemObject")
base_file_name = pm_assign_folder & "\Assignment List of Interviews from " & list_of_dates
worklist_file_name = pm_assign_folder & "\Assignment List of Interviews from " & list_of_dates
pm_worklist_file_name = worklist_file_name & ".xlsx"
file_numb_count = 1
Do
    file_is_already_here = False
    file_is_already_here = FSOxl.FileExists(pm_worklist_file_name)
    If file_is_already_here Then
        worklist_file_name = base_file_name & "_" & file_numb_count
        pm_worklist_file_name = worklist_file_name & ".xlsx"
        file_numb_count = file_numb_count + 1
    End If
Loop until file_is_already_here = False
pm_worklist_display_name = replace(pm_worklist_file_name, t_drive, "T:\")

If developer_mode = False Then
    objPMExcel.ActiveWorkbook.SaveAs pm_worklist_file_name
    objPMExcel.ActiveWorkbook.Close
    objPMExcel.Application.Quit
    objPMExcel.Quit
End If


'Creating the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first row with header information
ObjExcel.Cells(1, basket_col).Value = "CASELOAD"
ObjExcel.Cells(1, case_numb_col).Value = "CASE NUMBER"
ObjExcel.Cells(1, population_col).Value = "POPULATION"
ObjExcel.Cells(1, prog_mngr_col).Value = "PROGRAM MANAGER"
ObjExcel.Cells(1, intvw_date_col).Value = "INTERVIEW COMPLETED"
ObjExcel.Cells(1, exp_det_col).Value = "APPEARS EXPEDITED"
ObjExcel.Cells(1, assignment_col).Value = "ASSIGNED TO"

ObjExcel.Cells(1, cash_col).Value = "CASH"
ObjExcel.Cells(1, cash_type_col).Value = "CASH TYPE"
ObjExcel.Cells(1, grh_col).Value = "GRH"
ObjExcel.Cells(1, snap_col).Value = "SNAP"
ObjExcel.Cells(1, emer_col).Value = "EMER"
FOR cow = 1 to emer_col							'formatting the cells'
	objExcel.Cells(1, cow).Font.Bold = True		'bold font'
NEXT

Set mvFSO = CreateObject("Scripting.FileSystemObject")

excel_row = 2
For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
	If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "details") Then
		'Add the file information to the Excel document for the worklist
		ObjExcel.Cells(excel_row, basket_col).Value 		= CASES_TO_ASSIGN_ARRAY(x_numb_const, cow)
		ObjExcel.Cells(excel_row, case_numb_col).Value 		= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
		ObjExcel.Cells(excel_row, population_col).Value 	= CASES_TO_ASSIGN_ARRAY(population_const, cow)
        ObjExcel.Cells(excel_row, prog_mngr_col).Value      = CASES_TO_ASSIGN_ARRAY(pm_const, cow)
		ObjExcel.Cells(excel_row, intvw_date_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
		If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then ObjExcel.Cells(excel_row, exp_det_col).Value = "Yes"
		If CASES_TO_ASSIGN_ARRAY(cash_select_const, cow) = True Then
			ObjExcel.Cells(excel_row, cash_col).Value = "True"
			ObjExcel.Cells(excel_row, cash_type_col).Value = CASES_TO_ASSIGN_ARRAY(cash_prog_const, cow)
		End If
		If CASES_TO_ASSIGN_ARRAY(grh_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, grh_col).Value = "True"
		If CASES_TO_ASSIGN_ARRAY(snap_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, snap_col).Value = "True"
		If CASES_TO_ASSIGN_ARRAY(emer_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, emer_col).Value = "True"
	End If

	If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "started") Then
		'Add the file information to the Excel document for the worklist
		ObjExcel.Cells(excel_row, basket_col).Value 	= CASES_TO_ASSIGN_ARRAY(x_numb_const, cow)
		ObjExcel.Cells(excel_row, case_numb_col).Value 	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
		ObjExcel.Cells(excel_row, population_col).Value = CASES_TO_ASSIGN_ARRAY(population_const, cow)
        ObjExcel.Cells(excel_row, prog_mngr_col).Value  = CASES_TO_ASSIGN_ARRAY(pm_const, cow)
		ObjExcel.Cells(excel_row, intvw_date_col).Value = CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
		ObjExcel.Cells(excel_row, exp_det_col).Value = "?"
	End If
	excel_row = excel_row + 1		'increment the excel row to add more
Next

'format the Worklist Excel
For col_to_autofit = 1 to emer_col
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

If developer_mode = False Then
	On Error Resume Next
	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		Err.Clear
		mvFSO.MoveFile CASES_TO_ASSIGN_ARRAY(file_path_const, cow), interview_team_cases_already_on_worklist & "\" & CASES_TO_ASSIGN_ARRAY(file_name_const, cow)

		If Err.Number = 70 Then
			attempt = 0
			Do
				EMWaitReady 0,0
				Err.Clear
				mvFSO.MoveFile CASES_TO_ASSIGN_ARRAY(file_path_const, cow), interview_team_cases_already_on_worklist & "\" & CASES_TO_ASSIGN_ARRAY(file_name_const, cow)
				attempt = attempt + 1
			Loop until Err.Number = 0 or attempt = 6
			If Err.Number <> 0 Then
				email_subject = "LD Interview Worklist MOVE FILE ERROR"
				email_body = "Go manually move file: " & CASES_TO_ASSIGN_ARRAY(file_path_const, cow)
				call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
			End If
		End If
	Next
	Err.Clear
	On Error Goto 0

	Set FSOxl = CreateObject("Scripting.FileSystemObject")
	base_file_name = worklist_folder & "\Interview Team Cases Worklist from " & list_of_dates
	worklist_file_name = worklist_folder & "\Interview Team Cases Worklist from " & list_of_dates
	full_worklist_file_name = worklist_file_name & ".xlsx"
	file_numb_count = 1
	Do
		file_is_already_here = False
		file_is_already_here = FSOxl.FileExists(full_worklist_file_name)
		If file_is_already_here Then
			worklist_file_name = base_file_name & "_" & file_numb_count
			full_worklist_file_name = worklist_file_name & ".xlsx"
			file_numb_count = file_numb_count + 1
		End If
	Loop until file_is_already_here = False

	objExcel.ActiveWorkbook.SaveAs full_worklist_file_name
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

End If

end_message = end_message & vbCr & "Worklist by Program Manager Excel file created: " & pm_worklist_display_name
If developer_mode = False Then  end_message = end_message & vbCr & "File has been saved and closed." & vbCr & vbCr
end_message = end_message & "Case Counts:"

'Email to PMs of worklist completion
email_from = ""
If admin_run Then email_from = "HSPH.EWS.BlueZoneScripts@hennepin.us"
email_recip = ""
email_recip_CC = ""
email_subject = "Worklist Created - Interview Team Cases from " & list_of_dates
email_body = "The Interview Team Cases Worklist for interviews completed on " & list_of_dates & " has been created." & "<br>"
email_body = email_body & "The Worklist Excel File is here: "
email_body = email_body & "<a href=" & chr(34) & pm_worklist_file_name & chr(34) & ">Interview Worklist by PM for " & list_of_dates & "</a><br>"
email_body = email_body & "Please note that this file is only editable by one person at a time, so leaving it open will limit others' access.<br><br>"
email_body = email_body & "<br>" & "Case Count Overview:" & "<br>"
email_body = email_body & "Total Cases added to Worklist: " & case_count & "<br><br>"
email_body = email_body & "Count by Program Manager:<br>"

running_count = 0
For duck = 0 to UBound(PM_ARRAY, 2)
    email_recip = email_recip & PM_ARRAY(email_const, duck) & "; "
    email_body = email_body & PM_ARRAY(name_const, duck) & ": " & PM_ARRAY(case_count_const, duck) & " case"
    end_message = end_message & vbCr & PM_ARRAY(name_const, duck) & ": " & PM_ARRAY(case_count_const, duck) & " cases"
    If PM_ARRAY(case_count_const, duck) <> 1 Then email_body = email_body & "s"
    email_body = email_body & "<br>"
    running_count = running_count + PM_ARRAY(case_count_const, duck)
Next
email_body = email_body & "<br>" & "<b>Please assign cases on worklist.</b>"
send_email = True
If developer_mode = True Then send_email = False
email_recip = "hsph.ews.unit.coenen@hennepin.us"                        'Initial deployment email assignment. Once this process is confirmed these may be updated.
email_recip_CC = "HSPH.EWS.BlueZoneScripts@hennepin.us"                 'Initial deployment email assignment. Once this process is confirmed these may be updated.

'function labels  		  email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email
Call create_outlook_email(email_from, email_recip, email_recip_CC, "", 			    email_subject, 1, 				 False, 	   email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, False, 				      email_attachment_array, send_email)

case_diff = case_count - running_count
If case_diff <> 0 Then
    Call create_outlook_email("", "HSPH.EWS.BlueZoneScripts@hennepin.us", "", "", "LD Interview Worklist CASE COUNT MISMATCH", 1, False, "", "", False, "", "Case count mismatch detected. Cases counted in worklist: " & running_count & ". Cases expected: " & case_count & ". Difference: " & case_diff & "." & vbCr & "AUTOMATED EMAIL", False, "", True)
End If

end_message = end_message & vbCr & vbCr & "Excel created to save the information of the interviews created today." & vbCr

'Creates an oject of the whole assignments folder
Set objFolder = objFSO.GetFolder(t_drive & "\Eligibility Support\Assignments\Interview Team PM Assignments")
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder
For Each objFile in colFiles																'looping through each file
	this_file_created_date = objFile.DateCreated											'Reading the date created
	this_file_name = objFile.Name															'Grabing the file name
	this_file_path = objFile.Path															'Grabing the path for the file
	this_file_type = objFile.Type															'Grabing the file type

    ' If a worklist is over 3 motnhs old, it will be moved to the Archive folder
    If DateDiff("m", this_file_created_date, date) > 3 and this_file_type = "Microsoft Excel Worksheet" Then
        date_on_file = left(replace(this_file_name, "Assignment List of Interviews from ", ""), 2)                                          'Finding the folder name based on the date on the file
        archive_folder = t_drive & "\Eligibility Support\Assignments\Interview Team PM Assignments\Archive\" & date_on_file & "-2025"

        If NOT (FSOxl.FolderExists(archive_folder)) Then                                    'Creating the archive folder if it does not already exist
            FSOxl.CreateFolder archive_folder
        End If
        FSOxl.MoveFile this_file_path , archive_folder & "\" & this_file_name               'Moving the file to the archive folder
    End If
Next

internal_run_time = timer - start_time
internal_run_min = int(internal_run_time/60)
internal_run_sec = internal_run_time MOD 60

end_message = end_message & vbCr & "Script run time: " & internal_run_min & " min, " & internal_run_sec & " sec."

Call script_end_procedure(end_message)
