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

class form_questions
	public number
	public dialog_phrasing
	public note_phrasing
	public doc_phrasing
	public info_type
	public caf_answer
	public write_in_info
	public interview_notes

	public verif_status
	public verif_notes

	public guide_btn
	public verif_btn

	public dialog_page_numb
	public dialog_order
	public dialog_height

	public sub display_in_dialog(y_pos, question_yn, question_notes, question_interview_notes)
		question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
		question_yn = caf_answer
		question_notes = write_in_info
		question_interview_notes = interview_notes

		If info_type = "standard" Then
			'funcitonality here
			GroupBox 5, y_pos, 475, 55, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If verif_status = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_notes
				Text 360, y_pos, 110, 10, "Q1 - Verification - " & verif_status
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 20
		End If
	end sub

	public sub store_dialog_entry(question_yn, question_notes, question_interview_notes)
		caf_answer = question_yn
		write_in_info = question_notes
		interview_notes = question_interview_notes
	end sub

	public sub capture_verif_detail()
		'funcitonality here

		BeginDialog Dialog1, 0, 0, 396, 95, "Add Verification"
		DropListBox 60, 35, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", verif_status
		EditBox 60, 55, 330, 15, verif_notes
		ButtonGroup ButtonPressed
			PushButton 340, 75, 50, 15, "Return", return_btn
			PushButton 145, 35, 50, 10, "CLEAR", clear_btn
		Text 10, 10, 380, 20, number & "." & dialog_phrasing
		Text 10, 40, 45, 10, "Verification: "
		Text 20, 60, 30, 10, "Details:"
		EndDialog

		Do
			dialog Dialog1
			If ButtonPressed = -1 Then ButtonPressed = return_btn
			If ButtonPressed = clear_btn Then
				verif_status = "Not Needed"
				verif_notes = ""
			End If
		Loop until ButtonPressed = return_btn
	end sub

	public sub enter_case_note()
		If info_type = "standard" Then
			'funcitonality here
			If caf_answer <> "" OR trim(write_in_info) <> "" OR verif_status <> "" OR trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE(note_phrasing)
			q_input = "    CAF Answer - " & caf_answer
			If caf_answer <> "" OR trim(write_in_info) <> "" Then q_input = q_input & " (Confirmed)"
			If q_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_input)
			If trim(write_in_info) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & write_in_info)
			If verif_status <> "" Then
				If trim(verif_notes) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status)
				If trim(verif_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status & ": " & verif_notes)
			End If
			If trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & interview_notes)
		End If
	end sub

	public sub add_to_wif()
		If info_type = "standard" Then
			'funcitonality here
			objSelection.TypeText doc_phrasing & vbCr
			objSelection.TypeText chr(9) & "CAF Answer: " & caf_answer & vbCr
			If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
			If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
			If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr
			If caf_answer <> "" OR trim(write_in_info) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
			If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		End If
	end sub

end class
		' ReDim preserve FORM_QUESTION_ARRAY(question_num)
		' FORM_QUESTION_ARRAY(question_num).number 				= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= ""
		' FORM_QUESTION_ARRAY(question_num).note_phrasing			= ""
		' FORM_QUESTION_ARRAY(question_num).doc_phrasing			= ""
		' FORM_QUESTION_ARRAY(question_num).info_type				= ""
		' FORM_QUESTION_ARRAY(question_num).caf_answer 			= ""
		' FORM_QUESTION_ARRAY(question_num).write_in_info 		= ""
		' FORM_QUESTION_ARRAY(question_num).interview_notes		= ""

		' FORM_QUESTION_ARRAY(question_num).verif_status 			= ""
		' FORM_QUESTION_ARRAY(question_num).verif_notes			= ""

		' FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		' FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		' FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		' question_num = question_num + 1



CAF_form = "CAF (DHS-5223)"

question_num = 0
Dim FORM_QUESTION_ARRAY()
ReDim FORM_QUESTION_ARRAY(0)

numb_of_quest = 0
Select Case CAF_form
	Case "CAF (DHS-5223)"
		numb_of_quest = 2

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 1
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does everyone in your household buy, fix or eat food with you?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q1. Does everyone buy, fix, or eat food together?"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 1. Does everyone in your household buy, fix or eat food with you?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
		FORM_QUESTION_ARRAY(question_num).caf_answer 			= ""
		FORM_QUESTION_ARRAY(question_num).write_in_info 		= ""
		FORM_QUESTION_ARRAY(question_num).interview_notes		= ""

		FORM_QUESTION_ARRAY(question_num).verif_status 			= ""
		FORM_QUESTION_ARRAY(question_num).verif_notes			= ""

		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 2
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2. Is anyone (60+) disabled or unable to prepare food?"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
		FORM_QUESTION_ARRAY(question_num).caf_answer 			= ""
		FORM_QUESTION_ARRAY(question_num).write_in_info 		= ""
		FORM_QUESTION_ARRAY(question_num).interview_notes		= ""

		FORM_QUESTION_ARRAY(question_num).verif_status 			= ""
		FORM_QUESTION_ARRAY(question_num).verif_notes			= ""

		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 3
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household attending school?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3. Is anyone attending school?"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Is anyone in the household attending school?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
		FORM_QUESTION_ARRAY(question_num).caf_answer 			= ""
		FORM_QUESTION_ARRAY(question_num).write_in_info 		= ""
		FORM_QUESTION_ARRAY(question_num).interview_notes		= ""

		FORM_QUESTION_ARRAY(question_num).verif_status 			= ""
		FORM_QUESTION_ARRAY(question_num).verif_notes			= ""

		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

	Case "HUF (DHS-8107)"
	Case "SNAP App for Srs (DHS-5223F)"
	Case "MNbenefits"
	Case "Combined AR for Certain Pops (DHS-3727)"
End Select

const form_yn_const			= 0
const form_write_in_const	= 1
const intv_notes_const 		= 2
const verif_yn_const 		= 3
const verif_notes_const		= 4
const q_last_const			= 10

' numb_of_quest = UBound(FORM_QUESTION_ARRAY)
Dim TEMP_INFO_ARRAY()
ReDim TEMP_INFO_ARRAY(q_last_const, numb_of_quest)

MAXIS_case_number = "344839"
MsgBox "CAREFUL! This will CASE/NOTE in " & MAXIS_case_number & " without any real warning." & vbCr & vbCr & "USE IN TRAINING REGION."

page_display = 1
Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"

		ButtonGroup ButtonPressed
			If page_display = 1 Then
				ComboBox 120, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_interview_with, who_are_we_completing_the_interview_with
				ComboBox 120, 30, 75, 45, "Select or Type"+chr(9)+"Phone"+chr(9)+"In Office"+chr(9)+how_are_we_completing_the_interview, how_are_we_completing_the_interview
				EditBox 120, 50, 50, 15, interview_date
				ComboBox 120, 70, 340, 45, "No Interpreter Used"+chr(9)+"Language Line Interpreter Used"+chr(9)+"Interpreter through Henn Co. OMS (Office of Multi-Cultural Services)"+chr(9)+"Interviewer speaks Resident Language"+chr(9)+interpreter_information, interpreter_information
				ComboBox 120, 90, 205, 45, "English"+chr(9)+"Somali"+chr(9)+"Spanish"+chr(9)+"Hmong"+chr(9)+"Russian"+chr(9)+"Oromo"+chr(9)+"Vietnamese"+chr(9)+interpreter_language, interpreter_language
				PushButton 330, 90, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
				EditBox 120, 110, 340, 15, arep_interview_id_information
				EditBox 10, 155, 450, 15, non_applicant_interview_info
			ElseIf page_display = 2 Then
				GroupBox 10, 35, 375, 95, "Residence Address"
				Text 20, 55, 45, 10, "House/Street"
				Text 45, 75, 20, 10, "City"
				Text 185, 75, 20, 10, "State"
				Text 325, 75, 15, 10, "Zip"
				Text 20, 95, 100, 10, "Do you live on a Reservation?"
				Text 180, 95, 60, 10, "If yes, which one?"
				Text 20, 115, 100, 10, "Resident Indicates Homeless:"
				Text 185, 115, 60, 10, "Living Situation?"
				GroupBox 10, 135, 375, 70, "Mailing Address"
				Text 20, 165, 45, 10, "House/Street"
				Text 45, 185, 20, 10, "City"
				Text 185, 185, 20, 10, "State"
				Text 325, 185, 15, 10, "Zip"
				GroupBox 10, 210, 235, 90, "Phone Number"
				Text 20, 225, 50, 10, "Number"
				Text 125, 225, 25, 10, "Type"
				Text 255, 225, 60, 10, "Date of Change:"
				Text 255, 245, 75, 10, "County of Residence:"
			ElseIf page_display = 3 Then
				Text 70, 35, 50, 10, "Last Name"
				Text 165, 35, 50, 10, "First Name"
				Text 245, 35, 50, 10, "Middle Name"
				Text 300, 35, 50, 10, "Other Names"
				Text 70, 65, 55, 10, "Soc Sec Number"
				Text 145, 65, 45, 10, "Date of Birth"
				Text 220, 65, 45, 10, "Gender"
				Text 275, 65, 90, 10, "Relationship to MEMB 01"
				Text 370, 65, 50, 10, "Marital Status"
				Text 70, 95, 75, 10, "Last Grade Completed"
				Text 185, 95, 55, 10, "Moved to MN on"
				Text 260, 95, 65, 10, "Moved to MN from"
				Text 400, 95, 75, 10, "US Citizen or National"
				Text 70, 125, 40, 10, "Interpreter?"
				Text 140, 125, 95, 10, "Preferred Spoken Language"
				Text 140, 155, 95, 10, "Preferred Written Language"
				Text 70, 175, 65, 10, "Identity Verification"
				GroupBox 325, 125, 155, 100, "Demographics"
				Text 330, 135, 35, 10, "Hispanic?"
				Text 330, 160, 50, 10, "Race"
				Text 70, 200, 145, 10, "Which programs is this person requesting?"
				Text 70, 255, 80, 10, "Intends to reside in MN"
				Text 155, 255, 65, 10, "Immigration Status"
				Text 365, 255, 50, 10, "Sponsor?"
				Text 70, 285, 50, 10, "Verification"
				Text 155, 285, 65, 10, "Verification Details"
				Text 70, 315, 50, 10, "Notes:"
			ElseIf page_display = 4 Then
				' display_count = 1
				y_pos = 10
				For quest = 0 to UBound(FORM_QUESTION_ARRAY)
					If FORM_QUESTION_ARRAY(quest).dialog_page_numb = page_display Then
						' If FORM_QUESTION_ARRAY(quest).dialog_order = display_count Then
						call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest))
						' y_pos = y_pos + FORM_QUESTION_ARRAY(quest).dialog_height
						' MsgBox "y_pos - " & y_pos
						' 	display_count = display_count + 1
						' End If
					End If
				Next
			ElseIf page_display = 5 Then
				GroupBox 10, 35, 375, 95, "PAGE 5"

			End If

			Text 485, 5, 75, 10, "---   DIALOGS   ---"
			Text 485, 17, 10, 10, "1"
			Text 485, 32, 10, 10, "2"
			Text 485, 47, 10, 10, "3"
			Text 485, 62, 10, 10, "4"
			Text 485, 77, 10, 10, "5"

			If page_display <> 1 Then PushButton 495, 15, 55, 13, "INTVW / CAF 1", caf_page_one_btn
			If page_display <> 2 Then PushButton 495, 30, 55, 13, "CAF ADDR", caf_addr_btn
			If page_display <> 3 Then PushButton 495, 45, 55, 13, "CAF MEMBs", caf_membs_btn
			If page_display <> 4 Then PushButton 495, 60, 55, 13, "Q. 1 - 6", caf_q_1_6_btn
			If page_display <> 5 Then PushButton 495, 75, 55, 13, "Q. 7 - 11", caf_q_7_11_btn
	EndDialog


	err_msg = "LOOP"

	dialog Dialog1
	cancel_without_confirmation
	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest))
	Next

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If ButtonPressed = FORM_QUESTION_ARRAY(quest).verif_btn Then
			call FORM_QUESTION_ARRAY(quest).capture_verif_detail()
		End If
	Next

	If ButtonPressed = caf_page_one_btn Then page_display = 1
	If ButtonPressed = caf_addr_btn Then page_display = 2
	If ButtonPressed = caf_membs_btn Then page_display = 3
	If ButtonPressed = caf_q_1_6_btn Then page_display = 4
	If ButtonPressed = caf_q_7_11_btn Then page_display = 5
	If ButtonPressed = -1 Then err_msg = ""

Loop until err_msg = ""


'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "CAF Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "CAF Form Details - CASE #" & MAXIS_case_number			'Title of the document
' objWord.Visible = True														'Let the worker see the document
objWord.Visible = True 														'The worker should NOT see the docuement
'allow certain workers to see the document
' If user_ID_for_validation = "WFA168" or user_ID_for_validation = "LILE002" Then objWord.Visible = True

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "NOTES on INTERVIEW"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

If MAXIS_case_number <> "" Then objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information
For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
	FORM_QUESTION_ARRAY(each_question).add_to_wif()
Next


Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("TRIAL INTERVIEW")

Call write_variable_in_CASE_NOTE("Interview Date")

CALL write_variable_in_CASE_NOTE("-----  CAF Information and Notes -----")

For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
	FORM_QUESTION_ARRAY(each_question).enter_case_note()
Next

Call write_variable_in_CASE_NOTE("SCRIPT WRITER")
