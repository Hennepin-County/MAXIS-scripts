'VERSION #1: DROPDOWN & CHECKBOX
'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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
EMConnect "" 'Connects to BlueZone


'Button Defined
add_button 			= 201
all_forms 			= 202
review_selections 	= 203
clear_button		= 204

asset_btn			= 400
arep_btn			= 401
atr_btn				= 402
change_btn 			= 403
evf_btn				= 404
hospice_btn			= 405
iaa_btn				= 406
iaa_ssi_btn			= 407
mof_btn				= 408
mtaf_btn			= 409
psn_btn				= 410
shelter_btn			= 411
diet_btn			= 412





call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'FIRST DIALOG COLLECTING CASE, MO/YR===========================================================================
Do
	DO
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 181, 90, "Case number dialog"
			EditBox 70, 5, 65, 15, MAXIS_case_number
			EditBox 70, 25, 30, 15, MAXIS_footer_month
			EditBox 105, 25, 30, 15, MAXIS_footer_year
			EditBox 70, 45, 100, 15, worker_signature
			ButtonGroup ButtonPressed
				OkButton 65, 70, 50, 15
				CancelButton 120, 70, 50, 15
			Text 20, 10, 50, 10, "Case number: "
			Text 20, 30, 45, 10, "Footer month:"
			Text 5, 50, 60, 10, "Worker signature:"
		EndDialog


		dialog Dialog1	'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF IsNumeric(MAXIS_footer_month) = FALSE OR IsNumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbNewLine &  "* You must type a valid footer month and year."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'SECOND & THIRD DIALOG FORM SELECTION===========================================================================

Dim unchecked, checked		'Defining unchecked/checked 
unchecked = 0			
checked = 1


Dim form_type_array()		'Defining 1D array
ReDim form_type_array(0)	'Redefining array so we can resize it 
form_count = 0				'Counter for array should start with 0


Do							'Do Loop to cycle through dialog as many times as needed until all desired forms are added
	Do
		Do
			err_msg = ""
			Dialog1 = "" 			'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
				DropListBox 30, 30, 180, 15, ""+chr(9)+"Asset Statement"+chr(9)+"AREP (Authorized Rep)"+chr(9)+"Authorization to Release Information (ATR)"+chr(9)+"Change Report Form"+chr(9)+"Employment Verification Form (EVF)"+chr(9)+"Hospice Transaction Form"+chr(9)+"Interim Assistance Agreement (IAA)"+chr(9)+"Interim Assistance Authorization- SSI"+chr(9)+"Medical Opinion Form (MOF)"+chr(9)+"Minnesota Transition Application Form (MTAF)"+chr(9)+"Professional Statement of Need (PSN)"+chr(9)+"Residence and Shelter Expenses Release Form"+chr(9)+"Special Diet Information Request (MFIP and MSA)", Form_type
				ButtonGroup ButtonPressed
				PushButton 225, 30, 35, 10, "Add", add_button
				PushButton 225, 60, 35, 10, "All Forms", all_forms
				OkButton 205, 215, 40, 15
				CancelButton 255, 215, 40, 15
				PushButton 155, 215, 40, 15, "Clear", clear_button
				GroupBox 5, 5, 280, 70, "Directions: For each document received either:"
				Text 15, 15, 275, 10, "1. Select document from dropdown, then select Add button. Repeat for each form."
				Text 10, 45, 15, 10, "OR"
				Text 15, 60, 180, 10, "2. Select All Forms to select forms via checkboxes."
				GroupBox 45, 85, 210, 125, "Documents Selected"
				y_pos = 95			'defining y_pos so that we can dynamically add forms to the dialog as they are selected
				
				'TODO: Handle for duplicate selection
				For each form in form_type_array		'For each must be within dialog so it knows where to write the information 
					Text 55, y_pos, 195, 10, form		'Writing form name
					y_pos = y_pos + 10					'Increasing y_pos by 10 before the next form is written on the dialog
				Next

			EndDialog								'Dialog handling	
			dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
			cancel_confirmation

			If ButtonPressed = add_button and form_type <> "" Then					'If statement to know when to store the information in the array
				ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
				form_type_array(form_count) = Form_type			'form_type is the same var for the dropbox for types  
				form_count= form_count + 1 
			End If
				
			If ButtonPressed = clear_button Then 
				ReDim form_type_array(form_count)		'Clear button wipes out any selections already made so the user can reselect correct forms.
				form_count = 0							'Reset the form count to 0 so that y_pos resets to 95. 
				asset_checkbox = unchecked				'Resetting checkboxes to unchecked
				arep_checkbox = unchecked				'Resetting checkboxes to unchecked
				atr_checkbox = unchecked				'Resetting checkboxes to unchecked
				change_checkbox = unchecked				'Resetting checkboxes to unchecked
				evf_checkbox = unchecked				'Resetting checkboxes to unchecked
				hospice_checkbox = unchecked			'Resetting checkboxes to unchecked
				iaa_checkbox = unchecked				'Resetting checkboxes to unchecked
				iaa_ssi_checkbox = unchecked			'Resetting checkboxes to unchecked
				mof_checkbox = unchecked				'Resetting checkboxes to unchecked
				mtaf_checkbox = unchecked				'Resetting checkboxes to unchecked
				psn_checkbox = unchecked				'Resetting checkboxes to unchecked
				shelter_checkbox = unchecked			'Resetting checkboxes to unchecked
				diet_checkbox = unchecked				'Resetting checkboxes to unchecked
				form_type = ""							'Resetting dropdown to blank
				MsgBox "Form selections cleared." & vbNewLine & "Please make new form selections."	'Notify end user that entries were cleared.
			End If
			
			If form_count = 0 and ButtonPressed = Ok Then err_msg = "-Add forms to process or select cancel to exit script"		'If form_count = 0, then no forms have been added to doc rec to be processed.	
			If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	
	If ButtonPressed = all_forms Then		'Opens Dialog with checkbox selection for each form
		Do
			Do
				ReDim form_type_array(form_count)		'Reseting any selections already made so the user can reselect correct forms using different format.
				form_count = 0							'Reseting the form count to 0 so that y_pos resets to 95. 
				
				
			
				err_msg = ""
				Dialog1 = "" 'Blanking out previous dialog detail
				BeginDialog Dialog1, 0, 0, 196, 180, "Document Selection"
					CheckBox 15, 20, 160, 10, "Asset Statement", asset_checkbox
					CheckBox 15, 30, 160, 10, "AREP (Authorized Rep)", arep_checkbox
					CheckBox 15, 40, 160, 10, "Authorization to Release Information (ATR)", atr_checkbox
					CheckBox 15, 50, 160, 10, "Change Report Form", change_checkbox
					CheckBox 15, 60, 160, 10, "Employment Verification Form (EVF)", evf_checkbox
					CheckBox 15, 70, 160, 10, "Hospice Transaction Form", hospice_checkbox
					CheckBox 15, 80, 160, 10, "Interim Assistance Agreement (IAA)", iaa_checkbox
					CheckBox 15, 90, 160, 10, "Interim Assistance Authorization- SSI", iaa_ssi_checkbox
					CheckBox 15, 100, 160, 10, "Medical Opinion Form (MOF)", mof_checkbox
					CheckBox 15, 110, 160, 10, "Minnesota Transition Application Form (MTAF)", mtaf_checkbox
					CheckBox 15, 120, 160, 10, "Professional Statement of Need (PSN)", psn_checkbox
					CheckBox 15, 130, 170, 10, "Residence and Shelter Expenses Release Form", shelter_checkbox
					CheckBox 15, 140, 175, 10, "Special Diet Information Request (MFIP and MSA)", diet_checkbox
					ButtonGroup ButtonPressed
						PushButton 65, 160, 70, 15, "Review Selections", review_selections
						CancelButton 145, 160, 40, 15
					Text 5, 5, 200, 10, "Select documents received, then Review Selections."
				EndDialog
				dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation
				
		
				'Capturing forms with checked checkboxes in array, which will then be listed on the Select Documents Received dialog.
				If asset_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Asset Statement" 
					form_count= form_count + 1 
				End If

				If arep_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "AREP (Authorized Rep)"
					form_count= form_count + 1 
				End If

				If atr_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Authorization to Release Information (ATR)"
					form_count= form_count + 1 
				End If

				If change_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Change Report Form"
					form_count= form_count + 1 
				End If

				If evf_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Employment Verification Form (EVF)"
					form_count= form_count + 1 
				End If

				If hospice_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Hospice Transaction Form"
					form_count= form_count + 1 
				End If

				If iaa_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Interim Assistance Agreement (IAA)"
					form_count= form_count + 1 
				End If

				If iaa_ssi_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Interim Assistance Authorization- SSI"
					form_count= form_count + 1 
				End If
				
				If mof_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Medical Opinion Form (MOF)"
					form_count= form_count + 1 
				End If

				If mtaf_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Minnesota Transition Application Form (MTAF)"
					form_count= form_count + 1 
				End If

				If psn_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Professional Statement of Need (PSN)"
					form_count= form_count + 1 
				End If

				If shelter_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Residence and Shelter Expenses Release Form"
					form_count= form_count + 1 
				End If

				If diet_checkbox = checked Then 
					ReDim Preserve form_type_array(form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_count) = "Special Diet Information Request (MFIP and MSA)"
					form_count= form_count + 1 
				End If
				
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and iaa_ssi_checkbox = unchecked and mof_checkbox = unchecked and mtaf_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""	
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If			
Loop Until ButtonPressed = Ok

'MsgBox "Form dialogs are not built yet." & vbNewLine & "These are placeholders to give you an idea of the look and feel."

'Do while/Loop display individual dialogs for each form selected via checkbox or dropdown. Do/Loops allows us to jump around/are more flexible than For/Next 

form_count = 0
Do 
	If form_type_array(form_count) = "Asset Statement" then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Asset Statement"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If

	If form_type_array(form_count) = "AREP (Authorized Rep)" then 
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "AREP (Authorized Rep)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 
	
	If form_type_array(form_count) = "Authorization to Release Information (ATR)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Authorization to Release Information (ATR)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Change Report Form" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Change Report Form"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog

		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If

	If form_type_array(form_count) = "Employment Verification Form (EVF)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Employment Verification Form (EVF)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Hospice Transaction Form" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Hospice Transaction Form"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog

		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If

	If form_type_array(form_count) = "Interim Assistance Agreement (IAA)" Then
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Interim Assistance Agreement (IAA)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Interim Assistance Authorization- SSI" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Interim Assistance Authorization- SSI"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Medical Opinion Form (MOF)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Medical Opinion Form (MOF)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Minnesota Transition Application Form (MTAF)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Minnesota Transition Application Form (MTAF)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Professional Statement of Need (PSN)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Professional Statement of Need (PSN)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 

	If form_type_array(form_count) = "Residence and Shelter Expenses Release Form" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Residence and Shelter Expenses Release Form"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If

	If form_type_array(form_count) = "Special Diet Information Request (MFIP and MSA)" Then
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 376, 300, "Special Diet Information Request (MFIP and MSA)"
			EditBox 60, 5, 40, 15, MAXIS_case_number
			EditBox 160, 5, 45, 15, effective_date
			EditBox 285, 5, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			EditBox 75, 275, 85, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 330, 45, 45, 15, "Form #1", Button9
				PushButton 330, 65, 45, 15, "Form #2", Button11
				PushButton 330, 85, 45, 15, "Form #3", Button7
				PushButton 260, 275, 50, 15, "Previous", previous_btn
				PushButton 315, 275, 50, 15, "Next", next_btn
			Text 110, 10, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 220, 10, 60, 10, "Document Date:"
			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
			Text 5, 10, 50, 10, "Case Number:"
			Text 10, 280, 60, 10, "Worker Signature:"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, "..."
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	End If 
form_count = form_count + 1
Loop until form_count > Ubound(form_type_array)


'For/Next displays individual dialogs for each form selected via checkbox or dropdown
' For form_count = 0 to Ubound(form_type_array)			
' 	If form_type_array(form_count) = "Asset Statement" then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Asset Statement"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If

' 	If form_type_array(form_count) = "AREP (Authorized Rep)" then 
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "AREP (Authorized Rep)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 
	
' 	If form_type_array(form_count) = "Authorization to Release Information (ATR)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Authorization to Release Information (ATR)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Change Report Form" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Change Report Form"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog

' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If

' 	If form_type_array(form_count) = "Employment Verification Form (EVF)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Employment Verification Form (EVF)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Hospice Transaction Form" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Hospice Transaction Form"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog

' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If

' 	If form_type_array(form_count) = "Interim Assistance Agreement (IAA)" Then
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Interim Assistance Agreement (IAA)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Interim Assistance Authorization- SSI" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Interim Assistance Authorization- SSI"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Medical Opinion Form (MOF)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Medical Opinion Form (MOF)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Minnesota Transition Application Form (MTAF)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Minnesota Transition Application Form (MTAF)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Professional Statement of Need (PSN)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Professional Statement of Need (PSN)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 

' 	If form_type_array(form_count) = "Residence and Shelter Expenses Release Form" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Residence and Shelter Expenses Release Form"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If

' 	If form_type_array(form_count) = "Special Diet Information Request (MFIP and MSA)" Then
' 		err_msg = ""
' 		Dialog1 = "" 'Blanking out previous dialog detail
' 		BeginDialog Dialog1, 0, 0, 376, 300, "Special Diet Information Request (MFIP and MSA)"
' 			EditBox 60, 5, 40, 15, MAXIS_case_number
' 			EditBox 160, 5, 45, 15, effective_date
' 			EditBox 285, 5, 45, 15, date_received
' 			EditBox 30, 65, 270, 15, address_notes
' 			EditBox 30, 85, 270, 15, household_notes
' 			EditBox 30, 105, 270, 15, Edit14
' 			EditBox 30, 125, 270, 15, Edit15
' 			EditBox 30, 145, 270, 15, Edit16
' 			EditBox 75, 275, 85, 15, worker_signature
' 			ButtonGroup ButtonPressed
' 				PushButton 330, 45, 45, 15, "Form #1", Button9
' 				PushButton 330, 65, 45, 15, "Form #2", Button11
' 				PushButton 330, 85, 45, 15, "Form #3", Button7
' 				PushButton 260, 275, 50, 15, "Previous", previous_btn
' 				PushButton 315, 275, 50, 15, "Next", next_btn
' 			Text 110, 10, 50, 10, "Effective Date:"
' 			Text 15, 70, 10, 10, "Q1"
' 			Text 220, 10, 60, 10, "Document Date:"
' 			GroupBox 5, 50, 320, 195, "Reponses to form questions captured here"
' 			Text 5, 10, 50, 10, "Case Number:"
' 			Text 10, 280, 60, 10, "Worker Signature:"
' 			Text 15, 110, 10, 10, "Q3"
' 			Text 15, 130, 15, 10, "Q4"
' 			Text 15, 90, 15, 10, "Q2"
' 			Text 15, 150, 15, 10, "..."
' 		EndDialog
' 		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
' 		cancel_confirmation
' 	End If 
' Next

