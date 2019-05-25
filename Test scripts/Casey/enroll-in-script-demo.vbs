'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - Enroll in a Script Demo.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                     'manual run time in seconds
STATS_denomination = "I"                   'C is for each CASE
'END OF stats block=========================================================================================================

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
call changelog_update("05/24/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
class script_bowie

    'Stuff the user indicates
	public script_name             	'The familiar name of the script (file name without file extension or category, and using familiar case)
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
	public SIR_instructions_button	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
	public workflows               	'The script workflows associated with this script (Changes Reported, Applications, etc)
    public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
	public release_date				'This allows the user to indicate when the script goes live (controls NEW!!! messaging)

    'Details the menus will figure out (does not need to be explicitly declared)
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run

    'Details the class itself figures out
	public property get script_URL
		If run_locally = true then
			script_repository = "C:\MAXIS-Scripts\"
			script_URL = script_repository & lcase(category) & "\" & lcase(replace(script_name, " ", "-") & ".vbs")
		Else
        	If script_repository = "" then script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"    'Assumes we're scriptwriters
        	script_URL = script_repository & lcase(category) & "/" & replace(lcase(script_name) & ".vbs", " ", "-")
		End if
    end property

    'public property get SIR_instructions_URL 'The instructions URL in SIR
    '    SIR_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
    'end property

end class

class script_demo

    public script_name
    public category
    public tags
    public instructions
    public demo_dates
    public demo_length
    public future_dates
    public demo_checkbox

end class

script_num = 0
ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
SCRIPT_DEMO_ARRAY(script_num).script_name   = "Earned Income Budgeting"
SCRIPT_DEMO_ARRAY(script_num).category      = "ACTIONS"
SCRIPT_DEMO_ARRAY(script_num).tags          = ""
SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20EARNED%20INCOME%20BUDGETING.docx"
SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#7/10/19 3:00 PM#, #7/16/19 8:30 AM#, #7/25/19 10:00 AM#, #7/31/19 2:00 PM#)
SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
SCRIPT_DEMO_ARRAY(script_num).demo_checkbox = ARRAY(unchecked, unchecked, unchecked, unchecked)

script_num = script_num + 1
ReDim Preserve SCRIPT_DEMO_ARRAY(script_num)
Set SCRIPT_DEMO_ARRAY(script_num) = new script_demo
SCRIPT_DEMO_ARRAY(script_num).script_name   = "CAF Script"
SCRIPT_DEMO_ARRAY(script_num).category      = "NOTES"
SCRIPT_DEMO_ARRAY(script_num).tags          = ""
SCRIPT_DEMO_ARRAY(script_num).instructions  = "https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CAF.docx"
SCRIPT_DEMO_ARRAY(script_num).demo_dates    = ARRAY(#8/7/19 3:00 PM#, #8/13/19 8:30 AM#, #8/22/19 10:00 AM#)
SCRIPT_DEMO_ARRAY(script_num).demo_length    = 45
SCRIPT_DEMO_ARRAY(script_num).future_dates  = FALSE
SCRIPT_DEMO_ARRAY(script_num).demo_checkbox = ARRAY(unchecked, unchecked, unchecked)

' script_num = 0
' ReDim Preserve script_array(script_num)
' Set script_array(script_num) = new script_bowie
' script_array(script_num).script_name 			= "ABAWD Exemption"																		'Script name
' script_array(script_num).description 			= "Updates FSET/ABAWD coding on STAT/WREG and case notes ABAWD exemptions."
' script_array(script_num).category               = "ACTIONS"
' script_array(script_num).workflows              = ""
' script_array(script_num).subcategory            = array("ABAWD")
' script_array(script_num).release_date           = #09/25/2017#

EMConnect ""

unique_scripts = 0
total_dates = 0

Dim CHECKBOX_ARRAY()
ReDim CHECKBOX_ARRAY(0)

For each scheduled_script in SCRIPT_DEMO_ARRAY
    no_future_dates = TRUE
    For each scheduled_date in scheduled_script.demo_dates
        ReDim Preserve CHECKBOX_ARRAY(checkbox_counter)
        If DateDiff("d", date, scheduled_date) > -1 Then
            no_future_dates = FALSE
            total_dates = total_dates + 1
            scheduled_script.future_dates  = TRUE
        End If
        checkbox_counter = checkbox_counter + 1
    Next
    If no-no_future_dates = FALSE Then unique_scripts = unique_scripts + 1
Next
dlg_len = 170 + (unique_scripts * 20) + (total_dates * 20)
y_pos = 165

BeginDialog Dialog1, 0, 0, 391, dlg_len, "Dialog"
  Text 95, 10, 145, 10, "Welcome to the BlueZone Script Roadshow!"
  GroupBox 5, 25, 380, 90, "About Script Demos"
  Text 15, 40, 350, 15, "As our project is constantly growing and changing, we want to show you how best to use the tools we create."
  Text 15, 60, 360, 25, "Since we serve all populations and regions, and because our presence is mostly virtual, our trainings will be the same. The focus of our demos and information is on the tool, how it acts, and how you can use it. You don't need to see our faces, just MAXIS and the scripts. "
  Text 15, 90, 360, 20, "Due to all of these reasons, our Demos and Trainings are scheduled as remote Skype meetings. These meetings can be found on our SharePoint site, all you have to do is click on them to join!"
  Text 10, 120, 85, 10, "Upcoming Script Demos"
  Text 25, 135, 235, 20, "Check the box by any session to enroll in that Demo. This will schedule it in your Outlook and give us a notice that you will be joining."

  checkbox_counter = 0
  For each scheduled_script in SCRIPT_DEMO_ARRAY
      If scheduled_script.future_dates = TRUE Then GroupBox 10, y_pos, 375, 20 + UBound(scheduled_script.demo_dates) *20, scheduled_script.category & " - " & scheduled_script.script_name
      'checkbox_counter = 0
      y_pos = y_pos + 15
      'For each scheduled_date in scheduled_script.demo_dates
      For array_counter = 0 to UBound(scheduled_script.demo_dates)
          scheduled_date = scheduled_script.demo_dates(array_counter)
          If DateDiff("d", date, scheduled_date) > -1 Then

              CheckBox 25, y_pos, 345, 10, FormatDateTime(scheduled_date, 1) & " at " & FormatDateTime(scheduled_date, 3) & " - " & scheduled_script.script_name & "(" & scheduled_script.demo_length & " minutes)", CHECKBOX_ARRAY(checkbox_counter)
              y_pos = y_pos + 15

          End If
          checkbox_counter = checkbox_counter + 1
      Next
      y_pos = y_pos + 10
  Next
  'y_pos = y_pos + 10


  ' GroupBox 10, 175, 375, 60, "ACTIONS - Earned Income Budgeting"
  ' CheckBox 25, 190, 345, 10, "Wednesday July 10th at 3:00 PM - Earned Income Budgeting (45 minutes)", checkBoxOne
  ' CheckBox 25, 205, 345, 10, "Tuesday July 16th at 8:30 AM - Earned Income Budgeting (45 minutes)", Check2
  ' CheckBox 25, 220, 345, 10, "Thursday July 25th at 10:00 AM - Earned Income Budgeting (45 minutes)", Check3
  ' GroupBox 5, 245, 375, 60, "NOTES - CAF"
  ' CheckBox 25, 260, 345, 10, "Wednesday August 7th at 3:00 PM - CAF Script (45 minutes)", Check4
  ' CheckBox 25, 275, 345, 10, "Tuesday August 13th at 8:30 AM - CAF Script (45 minutes)", Check5
  ' CheckBox 25, 290, 345, 10, "Thursday August 22nd at 10:00 AM - CAF Script (45 minutes)", Check6
  Text 15, y_pos + 5, 100, 10, "Enter your Name for Enrollment:"
  EditBox 120, y_pos , 100, 15, worker_name
  ButtonGroup ButtonPressed
    OkButton 280, y_pos, 50, 15
    CancelButton 335, y_pos, 50, 15
EndDialog

Do
    err_msg = ""

    dialog Dialog1
    cancel_without_confirmation

    If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

Loop until err_msg = ""

checkbox_counter = 0
For each scheduled_script in SCRIPT_DEMO_ARRAY
    For array_counter = 0 to UBound(scheduled_script.demo_dates)
        MsgBox array_counter & vbNewLine & CHECKBOX_ARRAY(checkbox_counter)
        If CHECKBOX_ARRAY(checkbox_counter) = checked Then
        ' If scheduled_script.demo_checkbox(array_counter) = checked Then
            MsgBox "EMAIL TO SEND" & vbNewLine & scheduled_script.category & " - " & scheduled_script.script_name & vbNewLine & scheduled_script.demo_dates(array_counter)

            'create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)

            body_text = "Join the BlueZone Script team remotely to see a script demo on " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & vbCr & "Go here and double click on the link to join the Skype meeting - URL GOES HERE" & vbCr & vbCr & "Instructions for this script can be found here - " & scheduled_script.instructions

            confirm_demo_schedule = MsgBox("This is the script demo you have selected:" & vbCr & vbCr & body_text & vbCr & vbCr & "Do you wish to enroll in this Demo and schedule it?", vbQuestion + vbYesNo, "Confirm Enrollment in DEMO")

            If confirm_demo_schedule = vbYes Then
                Call create_outlook_appointment(FormatDateTime(scheduled_script.demo_dates(array_counter), 2), FormatDateTime(scheduled_script.demo_dates(array_counter), 3), FormatDateTime(DateAdd("n", scheduled_script.demo_length, scheduled_script.demo_dates(array_counter)), 3), "Script Demo - " & scheduled_script.script_name, body_text, "Skype", TRUE, 60, "")

                bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
                email_text = worker_name & " has enrolled in a DEMO for " & scheduled_script.category & " - " & scheduled_script.script_name & vbCr & "On: " & scheduled_script.demo_dates(array_counter)
                Call create_outlook_email(bzt_email, "", "DEMO Enrollment", email_text, "", TRUE)
            End If

        End If
        checkbox_counter = checkbox_counter + 1
    Next
Next


script_end_procedure("Demo Scheduled")
