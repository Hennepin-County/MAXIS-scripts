name_of_script = "UTILITIES - Contact Knowledge Now.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 45                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("08/19/2020", "Initial version.", "Casey Love, Hennepin County")

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
			If ButtonPressed = script_array(the_script).script_btn_one Then name_of_the_scripts = name_of_the_scripts & ", " & script_array(the_script).category & " - " & script_array(the_script).script_name
		Next

	Loop until ButtonPressed = done_btn

	If left(name_of_the_scripts, 2) = ", " Then name_of_the_scripts = right(name_of_the_scripts, len(name_of_the_scripts)-2)
	' MsgBox name_of_the_scripts

	ButtonPressed = search_btn

end function
'===========================================================================================================================
'Connecting to BlueZone
EMConnect ""

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
' Call check_for_MAXIS(True)								'make sure we are in MAXIS
CALL MAXIS_case_number_finder (MAXIS_case_number)		'try to find the case number


subject_list = "Type or Select"+chr(9)+"SNAP - Policy Question"+chr(9)+"MFIP - Policy Question"+chr(9)+"GA - Policy Question"+chr(9)+"MSA - Policy Question"+chr(9)+"GHR/HS - Policy Question"+chr(9)+"HC - Policy"+chr(9)+"Immigration Question"+chr(9)+"Expedited SNAP Question"+chr(9)+"SNAP Procedure"+chr(9)+"MFIP Procedure"+chr(9)+"GA Procedure"+chr(9)+"MSA Procedure"+chr(9)+"GRH/GA Procedure"+chr(9)+"HC Procedure"

'One and only dialog for this script
DO
	Do
	    email_body = ""
		email_subject = ""

		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 646, 240, "Knowledge Now Question"
		  EditBox 10, 35, 630, 15, kn_question_info
		  EditBox 10, 75, 80, 15, MAXIS_case_number
		  ComboBox 10, 110, 290, 45, subject_list+chr(9)+kn_subject, kn_subject
		  EditBox 10, 160, 290, 15, kn_script_name
		  EditBox 415, 220, 115, 15, worker_name
		  ButtonGroup ButtonPressed
		    PushButton 305, 165, 95, 10, "SEARCH for a script name", search_btn
		    OkButton 535, 220, 50, 15
		    CancelButton 590, 220, 50, 15
		  Text 10, 10, 215, 10, "Tell us your question and a Knowledge Now staff will contact you."
		  Text 10, 25, 100, 10, "Question for Knowledge Now:"
		  Text 15, 50, 300, 10, "(To change lines in the email type a semi-colon followed by a space '; ' in the field above.)"
		  Text 10, 65, 50, 10, "Case Number:"
		  Text 95, 80, 295, 10, "(Providing a case number can help us provide the quickest and most thorough response.)"
		  Text 10, 100, 90, 10, "General Type of Question:"
		  Text 15, 125, 260, 20, "(This is for the subject line of the email and can help us find a SME, you can select one of the options we have provided or type whatever fits best.)"
		  Text 10, 150, 160, 10, "If a script is a part of your question enter it here:"
		  Text 5, 190, 290, 30, "This information will be sent to the Quality Improvement email for a response by one of the QI team members assigned to  Knowledge Now. We usually try to provide a response to the email through Teams and you can provide any additional detail."
		  Text 5, 225, 330, 10, "Nothing is required but the more information you provide, the more detail we can get to you quickly."
		  Text 355, 225, 55, 10, "Sign your Email"
		EndDialog


	    Dialog Dialog1
	    cancel_without_confirmation

		If ButtonPressed = search_btn Then call script_search(kn_script_name)
	Loop until ButtonPressed = -1

	email_subject = "Knowledge Now Question"
	kn_subject = trim(kn_subject)
	If kn_subject <> "" AND kn_subject <> "Type or Select" Then email_subject = email_subject & " RE: " & kn_subject

	kn_question_info = trim(kn_question_info)
	MAXIS_case_number = trim(MAXIS_case_number)
	kn_script_name = trim(kn_script_name)
	worker_name = trim(worker_name)

	email_body = "I need assistance from Knowledge Now." & vbCr & vbCr

	If kn_question_info <> "" Then email_body = email_body & "Question:" & vbCr & vbCr & replace(kn_question_info, "; ", vbCr) & vbCr & vbCr
	If MAXIS_case_number <> "" Then email_body = email_body & "About case: " & MAXIS_case_number & vbCr & vbCr
	If kn_script_name <> "" Then email_body = email_body & "Script(s) involved: " & kn_script_name & vbCr & vbCr
	email_body = email_body & "---" & vbCr
	If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

	message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO KNOWLEDGE NOW:" & vbCr & vbCr & email_subject & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
Loop until message_confirmed = vbYes

email_body = "~~This email is generated from completion of the 'Contact Knowledge Now' Script.~~" & vbCr & vbCr & email_body
call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", TRUE)

STATS_manualtime = STATS_manualtime + (timer - start_time)
end_msg = "Thank you!" & vbCr & "The Contact Knowledge Now Script is Complete." & vbCr & vbCr & "Your Report has been submitted to the QI Team. Someone will reach out to you shortly."
end_msg = end_msg & vbCr & vbCr & "Content of your Email to Knowledge Now:" & vbCr & "----------------------------------------------------------" & vbCr
end_msg = end_msg & "Subject: " & email_subject & vbCr & vbCr
end_msg = end_msg & email_body

call script_end_procedure_with_error_report(end_msg)
