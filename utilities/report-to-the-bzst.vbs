'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - Report to the BZST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "I"       		'C is for each CASE
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
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("07/13/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function script_search(name_of_the_scripts)
	btn_counter = 5000
	For the_script = 0 to UBound(script_array, 1)
		script_array(the_script).script_btn_one = btn_counter
		btn_counter = btn_counter + 1

	Next
	Do
		dlg_len = 105
		If search_words <> "" Then search_words_array = split(search_words, " ")
		x_pos = 15
		y_pos = 65
		count = 1

		For the_script = 0 to UBound(script_array, 1)
			' script_array(the_script).script_checkbox_one = checked
			script_array(the_script).show_script = FALSE
			If search_category <> "" AND search_category <> "Select One..." AND search_category <> "SPECIALITY" Then
				If script_array(the_script).category = search_category OR (search_category = "NOTICES" AND script_array(the_script).category = "" ) Then
					If search_words <> "" Then
						For each the_word in search_words_array
							If InStr(script_array(the_script).script_name, the_word) <> 0 Then
								' MsgBox "1"
								script_array(the_script).show_script = TRUE
								count = count + 1
								If count = 5 Then
									dlg_len = dlg_len + 15
									count = 1
								End If
							End If
						Next
					Else
						' MsgBox "2"
						script_array(the_script).show_script = TRUE
						count = count + 1
						If count = 5 Then
							dlg_len = dlg_len + 15
							count = 1
						End If
					End If
				End If
			ElseIf search_category = "SPECIALITY" Then

			ElseIf search_words <> "" Then
				For each the_word in search_words_array
					If InStr(script_array(the_script).script_name, the_word) <> 0 Then
						' MsgBox "3"
						script_array(the_script).show_script = TRUE
						count = count + 1
						If count = 5 Then
							dlg_len = dlg_len + 15
							count = 1
						End If
					End If
				Next
			End If
			' script_array(the_script).script_checkbox_one = unchecked
		Next
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 720, dlg_len, "Search for Scripts"
		  ButtonGroup ButtonPressed
			DropListBox 50, 25, 125, 45, "Select One..."+chr(9)+"NOTES"+chr(9)+"ADMIN"+chr(9)+"ACTIONS"+chr(9)+"BULK"+chr(9)+"NOTICES"+chr(9)+"UTILITIES"+chr(9)+"SPECIALTY", search_category
			EditBox 260, 25, 220, 15, search_words
			PushButton 490, 25, 75, 15, "SEARCH", search_go_btn
			Text 10, 10, 475, 10, "This dialog can display a number of scripts that we have available, so you can select the script(s) that apply to the information you are providing."
			Text 10, 30, 35, 10, "Category:"
			Text 190, 30, 70, 10, "Part of script name:"
			Text 10, 50, 160, 10, "Check the script(s) that you are reporting about:"
			For the_script = 0 to UBound(script_array, 1)
			  If script_array(the_script).show_script = TRUE Then
			  	  PushButton x_pos, y_pos, 20, 10, "ADD", script_array(the_script).script_btn_one
				  Text x_pos + 23, y_pos, 160, 10, script_array(the_script).category & " - " & script_array(the_script).script_name
				  x_pos = x_pos + 170
				  If x_pos = 695 Then
					  y_pos = y_pos + 15
					  x_pos = 15
				  End If
			  End If
			Next
			Text 15, y_pos + 15, 540, 10, name_of_the_scripts
				' OkButton 555, y_pos + 5, 50, 15
			PushButton 665, y_pos + 20, 50, 15, "DONE", done_btn
		EndDialog

		dialog Dialog1

		For the_script = 0 to UBound(script_array, 1)
			If ButtonPressed = script_array(the_script).script_btn_one Then
				name_of_the_scripts = name_of_the_scripts & ", " & script_array(the_script).category & " - " & script_array(the_script).script_name
				If script_array(the_script).category = "NOTES" AND script_array(the_script).script_name = "CAF" Then caf_script_reported = True
				If script_array(the_script).category = "NOTES" AND script_array(the_script).script_name = "Interview" Then interview_script_reported = True
			End If
		Next

	Loop until ButtonPressed = done_btn

	If left(name_of_the_scripts, 2) = ", " Then name_of_the_scripts = right(name_of_the_scripts, len(name_of_the_scripts)-2)
	' MsgBox name_of_the_scripts

	ButtonPressed = search_btn

end function
'===========================================================================================================================
EMConnect ""											'connecting to MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)		'Grabbing the case number if it can find one
call find_user_name(email_signature)

Do
    err_msg = ""

	BeginDialog Dialog1, 0, 0, 436, 300, "What Type of Report Do you Have"
	  DropListBox 10, 100, 420, 45, "Select One..."+chr(9)+"Something went wrong with a script run (Bug or Error Report)"+chr(9)+"Error Occured on the NOTES - Interview script"+chr(9)+"Error Occured on the NOTES - CAF script"+chr(9)+"Idea for improving a current script (Enhancement Request)"+chr(9)+"Idea for a New Script"+chr(9)+"There is a process that needs automation"+chr(9)+"Data or Lists needed"+chr(9)+"Script Instructions or Documentation Needed"+chr(9)+"Unsure", report_type
	  ButtonGroup ButtonPressed
	    OkButton 325, 280, 50, 15
	    CancelButton 380, 280, 50, 15
	  GroupBox 10, 10, 420, 70, "About Script Reporting"
	  Text 20, 25, 350, 10, "Thank you for taking the time to send us your thoughts and information about the BlueZone Script Project."
	  Text 20, 40, 395, 20, "This will guide you through some of the information that we need. We want to be able to respond to all of your ideas and reports as quickly as possible, this information will help. "
	  Text 20, 65, 410, 10, "The BlueZone Scripts can only be built to provide the support you let us know you need, we look forward to hearing from you."
	  Text 10, 90, 50, 10, "Type of Report:"
	  Text 10, 125, 85, 10, "About the Report Types:"
	  Text 20, 140, 405, 20, "Something went wrong with a script run (Bug or Error Report)  --  When something fails, it can be a small thing or a larger error. We want to know about these right away and will often have follow up questions."
	  Text 20, 165, 405, 20, "Idea for improving a current script (Enhancement Request)  --  A current script works but could be changed in some way to meet the requirements of a process, policy, or system."
	  Text 20, 190, 400, 10, "Idea for a New Script  --  There is something that a script could complete or assist with that is not already in a script."
	  Text 20, 205, 405, 20, "There is a process that needs automation  --  There is a task or set of tasks that you do regularly for which you need a script or script support."
	  Text 20, 230, 405, 20, "Data or Lists needed  --  We have a data need or have a report to update. This may be a project we refer to another data area, such as IPA."
	  Text 20, 255, 405, 20, "Script Instructions or Documentation Needed  --  You are looking for some information about scripts or a specific script, either how it works or what it does."
	EndDialog



    dialog Dialog1
    cancel_without_confirmation

	If report_type = "Select One..." Then err_msg = err_msg & vbNewLine & "* Please select the type of report you want to make." & vbNewLine & "This will help us gather the right type of information for the report." & vbNewLine & vbNewLine & "* If you do not know, select 'Unsure' and the script will ask a couple of guiding questions."

	If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

Loop until err_msg = ""
interview_script_reported = False
caf_script_reported = False
If report_type = "Error Occured on the NOTES - Interview script" Then
	report_type = "Something went wrong with a script run (Bug or Error Report)"
	error_script_name = "NOTES - Interview"
	interview_script_reported = True
End If

If report_type = "Error Occured on the NOTES - CAF script" Then
	report_type = "Something went wrong with a script run (Bug or Error Report)"
	error_script_name = "NOTES - CAF"
	caf_script_reported = True
End If

If report_type = "Unsure" Then
	BeginDialog Dialog1, 0, 0, 461, 155, "Identify Script Report Type"
	  DropListBox 295, 35, 160, 45, "Select One"+chr(9)+"Yes - Script already Exists"+chr(9)+"No - Script would be NEW", existing_case
	  DropListBox 295, 55, 160, 45, "Select One"+chr(9)+"Yes - The Script Failed"+chr(9)+"No - It Works but Could be Better", script_failure
	  DropListBox 295, 75, 160, 45, "Select One"+chr(9)+"Yes - the Policy Exists"+chr(9)+"No - This is more of an idea", process_exists
	  DropListBox 295, 95, 160, 45, "Select One"+chr(9)+"This would be for a SINGLE CASE"+chr(9)+"This would be on or for A LIST OF CASES", single_or_bulk
	  DropListBox 295, 115, 160, 45, "Select One"+chr(9)+"No this is about the SCRIPT FUNCTIONING"+chr(9)+"Yes - just need updates to information", documentation_update
	  ButtonGroup ButtonPressed
	    OkButton 355, 135, 50, 15
        CancelButton 405, 135, 50, 15
	  Text 10, 10, 260, 20, "You indicated that youa re unsure about the kind of report you need to make. These questions will pull the correct type or report."
	  Text 55, 40, 235, 10, "Do you need to report something about a script that ALREADY EXISTS?"
	  Text 80, 60, 205, 10, "Is something WRONG, did it fail or do something INCORRECT?"
	  Text 10, 80, 280, 10, "Do you have documented PROCESS or POLICY that you want Script Functionality for?"
	  Text 45, 100, 245, 10, "Do you need the script to work on only a SINGLE CASE or a LIST of CASES?"
	  Text 55, 120, 235, 10, "Does the script run fine but need new INFORMATION on SHAREPOINT?"
	EndDialog

	Do
		err_msg = ""
		dialog dialog1
		cancel_without_confirmation

		If existing_case = "Select One" Then err_msg = err_msg & vbNewLine & "* Answer Question 1 - New Script or Existing Script"
		If script_failure = "Select One" Then err_msg = err_msg & vbNewLine & "* Answer Question 2 - Something failed or just improved"
		If process_exists = "Select One" Then err_msg = err_msg & vbNewLine & "* Answer Question 3 - Is there a process to support request"
		If single_or_bulk = "Select One" Then err_msg = err_msg & vbNewLine & "* Answer Question 4 - Is this for a single case or list"
		If documentation_update = "Select One" Then err_msg = err_msg & vbNewLine & "* Answer Question 5 - Does the documentation need update"

	Loop until err_msg = ""

	If existing_case = "No - Script would be NEW" Then report_type = "Idea for a New Script"
	If existing_case = "Yes - Script already Exists" AND script_failure = "Yes - The Script Failed" Then report_type = "Something went wrong with a script run (Bug or Error Report)"
	If existing_case = "Yes - Script already Exists" AND script_failure = "No - It Works but Could be Better" Then report_type = "Idea for improving a current script (Enhancement Request)"
	If process_exists = "Yes - the Policy Exists" Then report_type = "There is a process that needs automation"
	If single_or_bulk = "This would be on or for A LIST OF CASES" Then report_type = "Data or Lists needed"
	If documentation_update = "Yes - just need updates to information" Then report_type = "Script Instructions or Documentation Needed"

	MsgBox "Based on the information you provided in these questions, it appears the best report type that we can start with is: " & vbCr & vbCr & report_type
End If

Select Case report_type
	Case "Something went wrong with a script run (Bug or Error Report)"

		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 230, "Error or Issue Report"
				  EditBox 125, 50, 75, 15, MAXIS_case_number
				  EditBox 125, 70, 240, 15, error_script_name
				  ButtonGroup ButtonPressed
				    PushButton 370, 75, 50, 10, "SEARCH", search_btn
				  ComboBox 125, 105, 295, 45, "Type or Select"+chr(9)+"CASE:NOTE is wrong or missing information"+chr(9)+"STAT panel that was supposed to be updated is wrong or did not update"+chr(9)+"There is a TYPO somewhere"+chr(9)+"You got an ERROR MESSAGE (usually says 'Line 1 Column 1' and something else)"+chr(9)+"A NOTICE didn't happen or was wrong (SPEC:MEMO or SPEC:WCOM)"+chr(9)+"Some of the script functionality did not happen or was wrong."+chr(9)+"The Dialog was messed up or wrong somehow."+chr(9)+error_info, error_info
				  EditBox 10, 135, 410, 15, error_details
				  EditBox 10, 170, 410, 15, error_notes
				  DropListBox 90, 190, 45, 45, "No"+chr(9)+"Yes", error_urgent
				  EditBox 70, 210, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 210, 50, 15
				    CancelButton 370, 210, 50, 15
				  Text 10, 10, 350, 10, "This is the information that will allow us to research and hopefully fix whatever error you are seeing."
				  Text 10, 25, 345, 20, "Though none of this information (or fields) are required, the more information you provide, the easier and quicker it will be for us to resolve the error."
				  Text 10, 55, 110, 10, "The case the error happened on:"
				  Text 40, 75, 80, 10, "The script that errored:"
				  Text 125, 90, 300, 10, "Format: CATEGORY - Name if possible, or use the 'SEARCH' button to find the correct script."
				  Text 15, 110, 105, 10, "General error/issue you found:"
				  Text 10, 125, 130, 10, "Error Details (explain what happened):"
				  Text 10, 160, 80, 10, "Other notes/comments:"
				  Text 10, 195, 80, 10, "Is this error URGENT?"
				  Text 145, 195, 250, 10, "(Errors are typically urgent if they prevent you from completing your work.)"
				  Text 10, 215, 60, 10, "Sign Your Email:"
				EndDialog

			    dialog Dialog1
				cancel_confirmation

				MAXIS_case_number = trim(MAXIS_case_number)
				error_script_name = trim(error_script_name)
				error_info = trim(error_info)
				error_details = trim(error_details)
				error_notes = trim(error_notes)
				email_signature = trim(email_signature)

				If ButtonPressed = search_btn Then call script_search(error_script_name)
			Loop until ButtonPressed = -1

			email_subject = "Error Report - BUG in scripts (Script Assisted Report)"
			If error_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If MAXIS_case_number <> "" Then
				email_body = email_body & "Case Number relevant/referenced: " & MAXIS_case_number & vbCr
			Else
				email_body = email_body & "~~ NO CASE NUMBER PROVIDED ~~" & vbCr
			End If
			If error_script_name <> "" Then
				email_body = email_body & "Error happened on script: " & error_script_name & vbCr
			Else
				email_body = email_body & "~~ NO SCRIPT NAME PROVIDED ~~" & vbCr
			End If
			If error_info <> "" AND error_info <> "Type or Select" Then email_body = email_body & "General Error Information: " & error_info & vbCr & vbCr
			If error_details <> "" Then email_body = email_body & "Detail of the Error: " & error_details & vbCr & vbCr
			If error_notes <> "" Then email_body = email_body & "Additional Error Notes: " & error_notes & vbCr & vbCr

			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes

	Case "Idea for improving a current script (Enhancement Request)"


		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 290, "Enhancment Idea"
				  EditBox 115, 50, 140, 15, MAXIS_case_number
				  EditBox 115, 70, 250, 15, enhancement_script_name
				  ButtonGroup ButtonPressed
				    PushButton 370, 75, 50, 10, "SEARCH", search_btn
				  ComboBox 115, 105, 305, 45, "Type or Select"+chr(9)+"Add New Information to CASE:NOTE"+chr(9)+"Change the Verbiage of a Notice"+chr(9)+"Look for More Information in MAXIS"+chr(9)+"Add a 'Double Check' before Proceeding"+chr(9)+"Add an Automation"+chr(9)+enhancement_info, enhancement_info
				  EditBox 10, 135, 410, 15, enhancement_details
				  EditBox 10, 165, 410, 15, enhancement_policy
				  EditBox 10, 205, 410, 15, enhancement_notes
				  DropListBox 90, 225, 45, 45, "No"+chr(9)+"Yes", enhancement_urgent
				  EditBox 175, 240, 90, 15, enhancement_effective_date
				  EditBox 70, 265, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 265, 50, 15
				    CancelButton 370, 265, 50, 15
				  Text 10, 10, 375, 10, "We always want to make the BlueZone Scripts work for you, but we can't know what you need unless you tell us!"
				  Text 10, 25, 345, 20, "In order to determine the possibility of adding an enhancement, or figuring out where it fits in our current project load, we need a lot of detail, clarity and supporting documentation."
				  Text 10, 55, 105, 10, "A case this may be useful on:"
				  Text 260, 55, 160, 10, "(multiple cases or a whole caseload is okay too)"
				  Text 40, 75, 70, 10, "The script to update:"
				  Text 120, 90, 300, 10, "Format: CATEGORY - Name if possible, or use the 'SEARCH' button to find the correct script."
				  Text 45, 110, 65, 10, "Enhancement type:"
				  Text 10, 125, 175, 10, "Enhancment Idea details (be as specific as possible):"
				  Text 10, 155, 185, 10, "Policy or Procedure References to support this update:"
				  Text 15, 180, 355, 10, "This is very important if changes are based on policy or procedure, enter web links or manual references."
				  Text 10, 195, 80, 10, "Other notes/comments:"
				  Text 10, 230, 80, 10, "Is this update URGENT?"
				  Text 15, 240, 135, 20, "(Updates are typically urgent if they prevent you from completing your work.)"
				  Text 175, 230, 195, 10, "If this is a change or new process, when did/will it change?"
				  Text 10, 270, 60, 10, "Sign Your Email:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				MAXIS_case_number = trim(MAXIS_case_number)
				enhancement_script_name = trim(enhancement_script_name)
				enhancement_info = trim(enhancement_info)
				enhancement_details = trim(enhancement_details)
				enhancement_policy = trim(enhancement_policy)
				enhancement_notes = trim(enhancement_notes)
				enhancement_effective_date = trim(enhancement_effective_date)
				email_signature = trim(email_signature)

				If ButtonPressed = search_btn Then call script_search(enhancement_script_name)
			Loop until ButtonPressed = -1

			email_subject = "Enhancement Idea for a Script (Script Assisted Report)"
			If enhancement_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If MAXIS_case_number <> "" Then
				email_body = email_body & "Case Number relevant/referenced: " & MAXIS_case_number & vbCr
			Else
				email_body = email_body & "~~ NO CASE NUMBER PROVIDED ~~" & vbCr
			End If
			If enhancement_script_name <> "" Then
				email_body = email_body & "Script to enhance: " & enhancement_script_name & vbCr
			Else
				email_body = email_body & "~~ NO SCRIPT NAME PROVIDED ~~" & vbCr
			End If
			If enhancement_info <> "" AND enhancement_info <> "Type or Select" Then email_body = email_body & "General Enhancement Info: " & enhancement_info & vbCr & vbCr
			If enhancement_details <> "" Then email_body = email_body & "Details about Enhacement Needed: " & enhancement_details & vbCr & vbCr
			If enhancement_policy <> "" Then email_body = email_body & "Policy to support this enhancement: " & enhancement_policy & vbCr & vbCr
			If enhancement_notes <> "" Then email_body = email_body & "Additional Notes: " & enhancement_notes & vbCr & vbCr
			If enhancement_effective_date <> "" Then email_body = email_body & "Change happened: " & enhancement_effective_date & vbCr & vbCr

			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes

	Case "Idea for a New Script"
		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 240, "New Script Idea"
				  EditBox 160, 50, 95, 15, MAXIS_case_number
				  DropListBox 375, 50, 45, 45, "No"+chr(9)+"Yes", new_script_urgent
				  ComboBox 160, 70, 260, 45, "Type or Select"+chr(9)+"CASE:NOTE to Enter"+chr(9)+"FIAT to Complete"+chr(9)+"MEMO or WCOM to Send"+chr(9)+"DAIL to Support"+chr(9)+"Tool or Utilitiy"+chr(9)+"Screening"+chr(9)+"Update a STAT Panel"+chr(9)+new_script_info, new_script_info
				  EditBox 10, 100, 410, 15, new_script_details
				  EditBox 10, 135, 410, 15, new_script_policy
				  EditBox 10, 175, 410, 15, new_script_notes
				  EditBox 210, 200, 90, 15, new_script_effective_date
				  EditBox 70, 220, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 220, 50, 15
				    CancelButton 370, 220, 50, 15
				  Text 10, 10, 250, 10, "Have an idea for anew script? We desperately want to hear from you!"
				  Text 10, 25, 400, 20, "Be aware that your ideas may come through in other functionality, or within an update to a script. We will let you know as we review and make our plans. Also be aware that there are some things we functionally or policy wise we cannot do."
				  Text 15, 55, 145, 10, "Case this process has been completed on:"
				  Text 290, 55, 80, 10, "Is this script URGENT?"
				  Text 80, 75, 75, 10, "Script Type to Create:"
				  Text 10, 90, 175, 10, "New Script Idea Detail:"
				  Text 10, 125, 185, 10, "Policy and Procedure to support this new script:"
				  Text 15, 150, 355, 10, "This is very important if changes are based on policy or procedure, enter web links or manual references."
				  Text 10, 165, 80, 10, "Other notes/comments:"
				  Text 10, 205, 195, 10, "If this is a change or new process, when did/will it change?"
				  Text 10, 225, 60, 10, "Sign Your Email:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				MAXIS_case_number = trim(MAXIS_case_number)
				new_script_info = trim(new_script_info)
				new_script_details = trim(new_script_details)
				new_script_policy = trim(new_script_policy)
				new_script_notes = trim(new_script_notes)
				new_script_effective_date = trim(new_script_effective_date)
				email_signature = trim(email_signature)
			Loop until ButtonPressed = -1

			email_subject = "New Script Idea (Script Assisted Report)"
			If new_script_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If MAXIS_case_number <> "" Then
				email_body = email_body & "Case Number relevant/referenced: " & MAXIS_case_number & vbCr
			Else
				email_body = email_body & "~~ NO CASE NUMBER PROVIDED ~~" & vbCr
			End If
			If new_script_info <> "" AND new_script_info <> "Type or Select" Then email_body = email_body & "New Script Information: " & new_script_info & vbCr & vbCr
			If new_script_details <> "" Then email_body = email_body & "Detail of New Script to Create: " & new_script_details & vbCr & vbCr
			If new_script_policy <> "" Then email_body = email_body & "Policy to Support New Script: " & new_script_policy & vbCr & vbCr
			If new_script_notes <> "" Then email_body = email_body & "Additional Notes: " & new_script_notes & vbCr & vbCr
			If new_script_effective_date <> "" Then email_body = email_body & "Effective Date of Change in Process/Policy: " & new_script_effective_date & vbCr & vbCr

			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes

	Case "There is a process that needs automation"
		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 240, "Automate Policy/Process"
				  EditBox 160, 50, 95, 15, MAXIS_case_number
				  DropListBox 375, 50, 45, 45, "No"+chr(9)+"Yes", process_change_urgent
				  ComboBox 160, 70, 260, 45, "Type or Select"+chr(9)+"New Policy"+chr(9)+"New DHS Procedure"+chr(9)+"Change to Work Structure"+chr(9)+"Change to Work Assignment Process"+chr(9)+"Existing Lengthy Process"+chr(9)+"Existing Error Prone Process"+chr(9)+"Existing Process that is Not Well Known"+chr(9)+"MAXIS Workaround"+chr(9)+process_change_type, process_change_type
				  EditBox 10, 100, 410, 15, process_change_details
				  EditBox 10, 135, 410, 15, process_change_policy
				  EditBox 10, 175, 410, 15, process_change_notes
				  EditBox 210, 200, 90, 15, process_change_effective_date
				  EditBox 70, 220, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 220, 50, 15
				    CancelButton 370, 220, 50, 15
				  Text 10, 10, 250, 10, "Is there a process that is done regularly that you think could be automated?"
				  Text 10, 25, 400, 20, "This is where we can have very high impact on our work in ES."
				  Text 15, 55, 145, 10, "Case this process has been completed on:"
				  Text 275, 55, 95, 10, "Is this process URGENT?"
				  Text 80, 75, 75, 10, "Process Type to Automate:"
				  Text 10, 90, 175, 10, "Process Automation Idea Detail:"
				  Text 10, 125, 185, 10, "Policy and Procedure References:"
				  Text 15, 150, 355, 10, "This is very important if changes are based on policy or procedure, enter web links or manual references."
				  Text 10, 165, 80, 10, "Other notes/comments:"
				  Text 10, 205, 195, 10, "If this is a change or new process, when did/will it change?"
				  Text 10, 225, 60, 10, "Sign Your Email:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				MAXIS_case_number = trim(MAXIS_case_number)
				process_change_type = trim(process_change_type)
				process_change_details = trim(process_change_details)
				process_change_policy = trim(process_change_policy)
				process_change_notes = trim(process_change_notes)
				process_change_effective_date = trim(process_change_effective_date)
				email_signature = trim(email_signature)

			Loop until ButtonPressed = -1

			email_subject = "Process or Policy to Automate (Script Assisted Report)"
			If process_change_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If MAXIS_case_number <> "" Then
				email_body = email_body & "Case Number relevant/referenced: " & MAXIS_case_number & vbCr
			Else
				email_body = email_body & "~~ NO CASE NUMBER PROVIDED ~~" & vbCr
			End If
			If process_change_type <> "" AND process_change_type <> "Type or Select" Then email_body = email_body & "Process to Automate Information: " & process_change_type & vbCr & vbCr
			If process_change_details <> "" Then email_body = email_body & "Detail about a Process to Automate: " & process_change_details & vbCr & vbCr
			If process_change_policy <> "" Then email_body = email_body & "Policy/Procedure References: " & process_change_policy & vbCr & vbCr
			If process_change_notes <> "" Then email_body = email_body & "Additional Notes: " & process_change_notes & vbCr & vbCr
			If process_change_effective_date <> "" Then email_body = email_body & "Effective Date of Process to Automate: " & process_change_effective_date & vbCr & vbCr


			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes

	Case "Data or Lists needed"
		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 265, "Data or Report Needed"
				  ComboBox 80, 25, 340, 45, "Type or Select"+chr(9)+"List of Cases Based on Case Status"+chr(9)+"List of Cases with Certain Criteria"+chr(9)+"List of Actions by Type"+chr(9)+"List of Actions by Date"+chr(9)+"List of Cases by time frame or date"+chr(9)+data_info, data_info
				  EditBox 10, 55, 410, 15, data_detail
				  EditBox 10, 90, 410, 15, data_policy
				  EditBox 10, 130, 410, 15, data_time
				  EditBox 10, 165, 410, 15, data_who
				  EditBox 10, 200, 410, 15, data_notes
				  DropListBox 115, 225, 45, 45, "No"+chr(9)+"Yes", data_urgent
				  EditBox 360, 225, 60, 15, data_effective_date
				  EditBox 70, 245, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 245, 50, 15
				    CancelButton 370, 245, 50, 15
				  Text 10, 10, 350, 10, "Need a list or report of information that is in MAXIS? We may be able to help. Provide us with details here."
				  Text 15, 30, 65, 10, "Data Type to Pull:"
				  Text 10, 45, 70, 10, "Data/Report Detail:"
				  Text 10, 80, 185, 10, "Policy and Procedure to support pulling this information:"
				  Text 20, 105, 355, 10, "This is very important if changes are based on policy or procedure, enter web links or manual references."
				  Text 10, 120, 155, 10, "How often will this data be needed, and when?"
				  Text 10, 155, 105, 10, "Who Needs this Data/Report?"
				  Text 10, 190, 80, 10, "Other notes/comments:"
				  Text 10, 230, 100, 10, "Is this data request URGENT?"
				  Text 165, 230, 195, 10, "If this is a change or new process, when did/will it change?"
				  Text 10, 250, 60, 10, "Sign Your Email:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				data_info = trim(data_info)
				data_detail = trim(data_detail)
				data_policy = trim(data_policy)
				data_time = trim(data_time)
				data_who = trim(data_who)
				data_notes = trim(data_notes)
				data_effective_date = trim(data_effective_date)
				email_signature = trim(email_signature)

			Loop until ButtonPressed = -1

			email_subject = "Data or a Report Needed (Script Assisted Report)"
			If data_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If data_info <> "" AND data_info <> "Type or Select" Then email_body = email_body & "Data or Report Information: " & data_info & vbCr & vbCr
			If data_detail <> "" Then email_body = email_body & "Report Detail: " & data_detail & vbCr & vbCr
			If data_policy <> "" Then email_body = email_body & "Existing Policy or Procedure: " & data_policy & vbCr & vbCr
			If data_time <> "" Then email_body = email_body & "Data Timeframe: " & data_time & vbCr & vbCr
			If data_who <> "" Then email_body = email_body & "Who Needs/Will Pull the Report/Data: " & data_who & vbCr & vbCr
			If data_notes <> "" Then email_body = email_body & "Additional Notes: " & data_notes & vbCr & vbCr
			If data_effective_date <> "" Then email_body = email_body & "If this is a Change, the change is effective: " & data_effective_date & vbCr & vbCr


			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes

	Case "Script Instructions or Documentation Needed"

		Do
			Do
				email_body = ""

				BeginDialog Dialog1, 0, 0, 426, 290, "Documentation Update"
				  EditBox 115, 50, 140, 15, MAXIS_case_number
				  EditBox 115, 70, 250, 15, documentation_script
				  ButtonGroup ButtonPressed
				    PushButton 370, 75, 50, 10, "SEARCH", search_btn
				  ComboBox 115, 105, 305, 45, "Type or Select"+chr(9)+"Change or Update to Instructions"+chr(9)+"Missing Instructions"+chr(9)+"Hot Topic Documentation"+chr(9)+"Announcement or Communication for Units/Committees"+chr(9)+documentation_type, documentation_type
				  EditBox 10, 135, 410, 15, documentation_details
				  EditBox 10, 165, 410, 15, documentation_existing
				  EditBox 10, 205, 410, 15, documentation_notes
				  DropListBox 90, 225, 45, 45, "No"+chr(9)+"Yes", documentation_urgent
				  EditBox 175, 240, 90, 15, documentation_effective_date
				  EditBox 70, 265, 150, 15, email_signature
				  ButtonGroup ButtonPressed
				    OkButton 315, 265, 50, 15
				    CancelButton 370, 265, 50, 15
				  Text 10, 10, 405, 20, "We always try to provide sufficient instruction and detail about each of the BlueZone Scripts, but if there is something you think should have additional detail, please let us know!"
				  Text 10, 30, 400, 10, "Please explain how we can add clarity to our documentation."
				  Text 25, 55, 85, 10, "A case this may apply to:"
				  Text 260, 55, 160, 10, "(multiple cases or a whole caseload is okay too)"
				  Text 15, 75, 95, 10, "The script that needs clarity:"
				  Text 120, 90, 300, 10, "Format: CATEGORY - Name if possible, or use the 'SEARCH' button to find the correct script."
				  Text 35, 110, 75, 10, "Documentation Needs:"
				  Text 10, 125, 175, 10, "Documentation Detail:"
				  Text 10, 155, 185, 10, "If documentation already exists, please add the link to it here:"
				  ' Text 15, 180, 355, 10, "This is very important if changes are based on policy or procedure, enter web links or manual references."
				  Text 10, 195, 80, 10, "Other notes/comments:"
				  Text 10, 230, 80, 10, "Is this update URGENT?"
				  Text 15, 240, 135, 20, "(Updates are typically urgent if they prevent you from completing your work.)"
				  Text 175, 230, 195, 10, "If this is a change or new process, when did/will it change?"
				  Text 10, 270, 60, 10, "Sign Your Email:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				MAXIS_case_number = trim(MAXIS_case_number)
				documentation_script = trim(documentation_script)
				documentation_details = trim(documentation_details)
				documentation_existing = trim(documentation_existing)
				documentation_notes = trim(documentation_notes)
				documentation_effective_date = trim(documentation_effective_date)
				email_signature = trim(email_signature)

				If ButtonPressed = search_btn Then call script_search(documentation_script)
			Loop until ButtonPressed = -1

			email_subject = "Instructions or Documentation Needed (Script Assisted Report)"
			If documentation_urgent = "Yes" Then email_subject = "URGENT! " & email_subject

			If MAXIS_case_number <> "" Then
				email_body = email_body & "Case Number relevant/referenced: " & MAXIS_case_number & vbCr
			Else
				email_body = email_body & "~~ NO CASE NUMBER PROVIDED ~~" & vbCr
			End If
			If documentation_script <> "" Then
				email_body = email_body & "Documentation needed for the script: " & documentation_script & vbCr
			Else
				email_body = email_body & "~~ NO SCRIPT NAME PROVIDED ~~" & vbCr
			End If
			If documentation_type <> "" AND documentation_type <> "Type or Select" Then email_body = email_body & "General Information about Documentation Needed: " & documentation_type & vbCr & vbCr
			If documentation_details <> "" Then email_body = email_body & "Details about Documentation Needed: " & documentation_details & vbCr & vbCr
			If documentation_existing <> "" Then email_body = email_body & "Current Documentation: " & documentation_existing & vbCr & vbCr
			If documentation_notes <> "" Then email_body = email_body & "Additional Notes: " & documentation_notes & vbCr & vbCr
			If documentation_effective_date <> "" Then email_body = email_body & "Documentation Update Neeeded based on Change Eff: " & documentation_effective_date & vbCr & vbCr


			email_body = email_body & "---" & vbCr
			If email_signature <> "" Then email_body = email_body & "Signed, " & vbCr & email_signature

			message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO THE BZST:" & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
		Loop until message_confirmed = vbYes
End Select
attachment_here = ""
If interview_script_reported = True Then
	local_interview_save_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
	With objFSO
		If .FileExists(local_interview_save_work_path) = True then
			attachment_here = local_interview_save_work_path
		End if
	End With
End If
If caf_script_reported = True Then
	local_CAF_save_work_path = user_myDocs_folder & "caf-variables-" & MAXIS_case_number & "-info.txt"
	With objFSO
		If .FileExists(local_CAF_save_work_path) = True then
			attachment_here = local_CAF_save_work_path
		End if
	End With
End If

email_body = "~~This email is generated from completion of the 'Report to the BZST' Script.~~" & vbCr & vbCr & email_body
call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", email_subject, email_body, attachment_here, TRUE)

STATS_manualtime = (timer-start_time) + 90
end_msg = "Thank you!" & vbCr & "The Script to Report to BZST is Complete." & vbCr & vbCr & "Your Report has been submitted to the BlueZone Script Team. We will respond within a week. This response may not a resolution as some requests take longer for the team to discuss, plan and schedule."
end_msg = end_msg & vbCr & vbCr & "Content of your Email to the BZST:" & vbCr & "----------------------------------------------------------" & vbCr
end_msg = end_msg & "Subject: " & email_subject & vbCr & vbCr
end_msg = end_msg & email_body

Call script_end_procedure(end_msg)
