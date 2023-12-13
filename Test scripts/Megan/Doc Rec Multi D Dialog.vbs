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
next_btn			= 205

asset_btn			= 400
atr_btn				= 401
arep_btn			= 402
change_btn 			= 403
evf_btn				= 404
hospice_btn			= 405
iaa_btn				= 406
iaa_ssi_btn			= 407
mof_btn				= 408
mtaf_btn			= 409
psn_btn				= 410
sf_btn				= 411
diet_btn			= 412


'Check for case number & footer
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'DIALOG COLLECTING CASE, FOOTER MO/YR===========================================================================
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


'DEFINING CONSTANTS & ARRAY===========================================================================
'Define Constants
const form_type_const   = 0
const btn_name_const    = 1
const btn_number_const	= 2
const the_last_const	= 3

'Defining array capturing form names, button names, button numbers
Dim form_type_array()		'Defining 1D array
ReDim form_type_array(the_last_const, 0)	'Redefining array so we can resize it 
form_count = 0				'Counter for array should start with 0

'Dim/ReDim Array for form checkbox selections
Dim unchecked, checked		'Defining unchecked/checked 
unchecked = 0			
checked = 1


'Defining count of forms
asset_count 	= 0 
atr_count 		= 0 
arep_count 		= 0 
change_count	= 0 
evf_count		= 0 
hosp_count		= 0 
iaa_count		= 0 
iaa_ssi_count	= 0
mof_count		= 0 
mtaf_count		= 0 
psn_count		= 0 
sf_count		= 0 
diet_count		= 0

'DIALOGS COLLECTING FORM SELECTION===========================================================================
'TODO: Handle for duplicate selection
Do							'Do Loop to cycle through dialog as many times as needed until all desired forms are added
	Do
		Do
			err_msg = ""
			Dialog1 = "" 			'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
				DropListBox 30, 30, 180, 15, ""+chr(9)+"Asset Statement"+chr(9)+"Authorization to Release Information (ATR)"+chr(9)+"AREP (Authorized Rep)"+chr(9)+"Change Report Form"+chr(9)+"Employment Verification Form (EVF)"+chr(9)+"Hospice Transaction Form"+chr(9)+"Interim Assistance Agreement (IAA)"+chr(9)+"Interim Assistance Authorization- SSI"+chr(9)+"Medical Opinion Form (MOF)"+chr(9)+"Minnesota Transition Application Form (MTAF)"+chr(9)+"Professional Statement of Need (PSN)"+chr(9)+"Residence and Shelter Expenses Release Form"+chr(9)+"Special Diet Information Request (MFIP and MSA)", Form_type
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
				
				For form = 0 to UBound(form_type_array, 2) 'Writing form name by incrementing to the next value in the array. For/next must be within dialog so it knows where to write the information. 
					'MsgBox form_type_array(form_type_const, form) 'TEST
					'MsgBox form_type_array(btn_name_const, form) 'TEST
					'MsgBox form_type_array(btn_number_const, form) 'TEST
					'MsgBox "Ubound" & UBound(form_type_array, 2) 'TEST
					Text 55, y_pos, 195, 10, form_type_array(form_type_const, form)
					y_pos = y_pos + 10					'Increasing y_pos by 10 before the next form is written on the dialog
				Next
				EndDialog								'Dialog handling	
				dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation
				
			If ButtonPressed = add_button and form_type <> "" Then					'If statement to know when to store the information in the array
				ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
				form_type_array(form_type_const, form_count) = Form_type			
		
				'Capturing form button name/label and button number in array based on drop down selections
				If form_type = "Asset Statement" Then 
					form_type_array(btn_name_const, form_count) = "ASSET"
					form_type_array(btn_number_const, form_count) = 400
				End If
				If form_type = "Authorization to Release Information (ATR)" Then
					form_type_array(btn_name_const, form_count) = "ATR"
					form_type_array(btn_number_const, form_count) = 401
				End If	
				If form_type = "AREP (Authorized Rep)" Then
					form_type_array(btn_name_const, form_count) = "AREP"
					form_type_array(btn_number_const, form_count) = 402
				End If
				If form_type = "Change Report Form" Then
					form_type_array(btn_name_const, form_count) = "CHNG"
					form_type_array(btn_number_const, form_count) = 403
				End If
				If form_type = "Employment Verification Form (EVF)" Then
					form_type_array(btn_name_const, form_count) = "EVF"
					form_type_array(btn_number_const, form_count) = 404
				End If
				If form_type = "Hospice Transaction Form" Then
					form_type_array(btn_name_const, form_count) = "HOSP"
					form_type_array(btn_number_const, form_count) = 405
				End If
				If form_type = "Interim Assistance Agreement (IAA)" Then
					form_type_array(btn_name_const, form_count) = "IAA"
					form_type_array(btn_number_const, form_count) = 406
				End If
				If form_type = "Interim Assistance Authorization- SSI" Then
					form_type_array(btn_name_const, form_count) = "IAA-SSI"
					form_type_array(btn_number_const, form_count) = 407
				End If
				If form_type = "Medical Opinion Form (MOF)" Then
					form_type_array(btn_name_const, form_count) = "MOF"
					form_type_array(btn_number_const, form_count) = 408
				End If 
				IF form_type = "Minnesota Transition Application Form (MTAF)" Then
					form_type_array(btn_name_const, form_count) = "MTAF"
					form_type_array(btn_number_const, form_count) = 409
				End If
				If form_type = "Professional Statement of Need (PSN)" Then
						form_type_array(btn_name_const, form_count) = "PSN"
						form_type_array(btn_number_const, form_count) = 410
				End If
				If form_type = "Residence and Shelter Expenses Release Form" Then
						form_type_array(btn_name_const, form_count) = "SF"
						form_type_array(btn_number_const, form_count) = 411
				End If
				If form_type = "Special Diet Information Request (MFIP and MSA)" Then
						form_type_array(btn_name_const, form_count) = "DIET"
						form_type_array(btn_number_const, form_count) = 412
				End If
				form_count= form_count + 1 'incrementing in the array
			End If
			If ButtonPressed = clear_button Then 'Clear button wipes out any selections already made so the user can reselect correct forms.
				ReDim form_type_array(form_count)		
				form_count = 0							'Reset the form count to 0 so that y_pos resets to 95. 
				asset_checkbox = unchecked				'Resetting checkboxes to unchecked
				atr_checkbox = unchecked				'Resetting checkboxes to unchecked
				arep_checkbox = unchecked				'Resetting checkboxes to unchecked
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

				asset_count 	= 0 
				atr_count 		= 0 
				arep_count 		= 0 
				change_count 	= 0
				evf_count		= 0 
				hosp_count		= 0 
				iaa_count		= 0 
				iaa_ssi_count	= 0
				mof_count		= 0 
				mtaf_count		= 0 
				psn_count		= 0 
				sf_count		= 0 
				diet_count		= 0
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
				ReDim form_type_array(the_last_const, form_count)		'Resetting any selections already made so the user can reselect correct forms using different format.
				form_type_array(form_type_const, form_count) = Form_type
                form_count = 0							'Resetting the form count to 0 so that y_pos resets to 95. 
		
				err_msg = ""
				Dialog1 = "" 'Blanking out previous dialog detail
				BeginDialog Dialog1, 0, 0, 196, 180, "Document Selection"
					CheckBox 15, 20, 160, 10, "Asset Statement", asset_checkbox
					CheckBox 15, 30, 160, 10, "Authorization to Release Information (ATR)", atr_checkbox
					CheckBox 15, 40, 160, 10, "AREP (Authorized Rep)", arep_checkbox
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

				'Capturing form button name/label and button number in array based on checkboxes selected 
				If asset_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Asset Statement" 
					form_type_array(btn_name_const, form_count) = "ASSET"
					form_type_array(btn_number_const, form_count) = 400
					form_count= form_count + 1 
				End If
				If atr_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Authorization to Release Information (ATR)"
					form_type_array(btn_name_const, form_count) = "ATR"
					form_type_array(btn_number_const, form_count) = 401
					form_count= form_count + 1 
				End If
				If arep_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "AREP (Authorized Rep)"
					form_type_array(btn_name_const, form_count) = "AREP"
					form_type_array(btn_number_const, form_count) = 402
					form_count= form_count + 1 
				End If
				If change_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Change Report Form"
					form_type_array(btn_name_const, form_count) = "CHNG"
					form_type_array(btn_number_const, form_count) = 403
					form_count= form_count + 1 
				End If
				If evf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Employment Verification Form (EVF)"
					form_type_array(btn_name_const, form_count) = "EVF"
					form_type_array(btn_number_const, form_count) = 404
					form_count= form_count + 1 
				End If
				If hospice_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Hospice Transaction Form"
					form_type_array(btn_name_const, form_count) = "HOSP"
					form_type_array(btn_number_const, form_count) = 405
					form_count= form_count + 1 
				End If
				If iaa_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Interim Assistance Agreement (IAA)"
					form_type_array(btn_name_const, form_count) = "IAA"
					form_type_array(btn_number_const, form_count) = 406
					form_count= form_count + 1 
				End If
				If iaa_ssi_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Interim Assistance Authorization- SSI"
					form_type_array(btn_name_const, form_count) = "IAA-SSI"
					form_type_array(btn_number_const, form_count) = 407
					form_count= form_count + 1 
				End If
				If mof_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Medical Opinion Form (MOF)"
					form_type_array(btn_name_const, form_count) = "MOF"
					form_type_array(btn_number_const, form_count) = 408
					form_count= form_count + 1 
				End If
				If mtaf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Minnesota Transition Application Form (MTAF)"
					form_type_array(btn_name_const, form_count) = "MTAF"
					form_type_array(btn_number_const, form_count) = 409
					form_count= form_count + 1 
				End If
				If psn_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Professional Statement of Need (PSN)"
					form_type_array(btn_name_const, form_count) = "PSN"
					form_type_array(btn_number_const, form_count) = 410
					form_count= form_count + 1 
				End If
				If shelter_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Residence and Shelter Expenses Release Form"
					form_type_array(btn_name_const, form_count) = "SF"
					form_type_array(btn_number_const, form_count) = 411
					form_count= form_count + 1 
				End If
				If diet_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Special Diet Information Request (MFIP and MSA)"
					form_type_array(btn_name_const, form_count) = "DIET"
					form_type_array(btn_number_const, form_count) = 412
					form_count= form_count + 1 
				End If
								
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and iaa_ssi_checkbox = unchecked and mof_checkbox = unchecked and mtaf_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""	
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If		
Loop Until ButtonPressed = Ok

'Capturing count of each form so we can iterate the necessary forms
For form_added = 0 to Ubound(form_type_array, 2)
	If form_type_array(form_type_const, form_added) = "Asset Statement" Then asset_count = asset_count + 1 
	If form_type_array(form_type_const, form_added) = "Authorization to Release Information (ATR)" Then atr_count = atr_count + 1
	If form_type_array(form_type_const, form_added) = "AREP (Authorized Rep)" Then arep_count = arep_count + 1
	If form_type_array(form_type_const, form_added) = "Change Report Form" Then change_count = change_count + 1 
	If form_type_array(form_type_const, form_added) = "Employment Verification Form (EVF)" Then evf_count = evf_count + 1  
	If form_type_array(form_type_const, form_added) = "Hospice Transaction Form" Then hosp_count = hosp_count + 1 
	If form_type_array(form_type_const, form_added) = "Interim Assistance Agreement (IAA)" Then iaa_count = iaa_count + 1 
	If form_type_array(form_type_const, form_added) = "Interim Assistance Authorization- SSI" Then iaa_ssi_count = iaa_ssi_count + 1 
	If form_type_array(form_type_const, form_added) = "Medical Opinion Form (MOF)" Then mof_count = mof_count + 1 
	If form_type_array(form_type_const, form_added) = "Minnesota Transition Application Form (MTAF)" Then mtaf_count = mtaf_count + 1 
	If form_type_array(form_type_const, form_added) = "Professional Statement of Need (PSN)" Then psn_count = psn_count + 1 
	If form_type_array(form_type_const, form_added) = "Residence and Shelter Expenses Release Form" Then sf_count = sf_count + 1 
	If form_type_array(form_type_const, form_added) = "Special Diet Information Request (MFIP and MSA)" Then diet_count = diet_count + 1 
	MsgBox "Asset count" & asset_count & vbcr & "ATR count" & atr_count & vbcr & "AREP" & arep_count & vbcr & "chng" & change_count & vbcr & "evf" & evf_count & vbcr & "hosp" & hosp_count & vbcr & "iaa" & iaa_count & vbcr & "iaa-ssi" & iaa_ssi_count & vbcr & "mof" & mof_count & vbcr & "mtaf" & mtaf_count & vbcr & "psn" & psn_count & vbcr & "sf" & sf_count & vbcr & "diet" & diet_count	'TEST
Next

'DIALOG DISPLAYING FORM SPECIFIC INFORMATION===========================================================================
'Displays individual dialogs for each form selected via checkbox or dropdown. Do/Loops allows us to jump around/are more flexible than For/Next 
form_count = 0
Do
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 456, 300, "Documents Received"
		If form_type_array(form_type_const, form_count) = "Asset Statement" then
				Text 60, 25, 45, 10, MAXIS_case_number
				EditBox 175, 20, 45, 15, effective_date
				EditBox 310, 20, 45, 15, date_received
				EditBox 30, 65, 270, 15, address_notes
				EditBox 30, 85, 270, 15, household_notes
				EditBox 30, 105, 270, 15, Edit14
				EditBox 30, 125, 270, 15, Edit15
				EditBox 30, 145, 270, 15, Edit16
				Text 5, 5, 220, 10, "ASSET STATEMENT"
				Text 125, 25, 50, 10, "Effective Date:"
				Text 15, 70, 10, 10, "Q1"
				Text 245, 25, 60, 10, "Document Date:"
				GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
				Text 5, 25, 50, 10, "Case Number:"
				Text 395, 35, 45, 10, "    --Forms--"
				Text 15, 110, 10, 10, "Q3"
				Text 15, 130, 15, 10, "Q4"
				Text 15, 90, 15, 10, "Q2"
				Text 15, 150, 15, 10, ""
		
		ElseIf form_type_array(form_type_const, form_count) = "Authorization to Release Information (ATR)" Then
				Text 60, 25, 45, 10, MAXIS_case_number
				EditBox 175, 20, 45, 15, effective_date
				EditBox 310, 20, 45, 15, date_received
				EditBox 30, 65, 270, 15, address_notes
				EditBox 30, 85, 270, 15, household_notes
				EditBox 30, 105, 270, 15, Edit14
				EditBox 30, 125, 270, 15, Edit15
				EditBox 30, 145, 270, 15, Edit16
				Text 5, 5, 220, 10, "AUTHORIZATION TO RELEASE INFORMATION (ATR)"
				Text 125, 25, 50, 10, "Effective Date:"
				Text 15, 70, 10, 10, "Q1"
				Text 245, 25, 60, 10, "Document Date:"
				GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
				Text 5, 25, 50, 10, "Case Number:"
				Text 395, 35, 45, 10, "    --Forms--"
				Text 15, 110, 10, 10, "Q3"
				Text 15, 130, 15, 10, "Q4"
				Text 15, 90, 15, 10, "Q2"
				Text 15, 150, 15, 10, ""
		
		ElseIf form_type_array(form_type_const, form_count) = "AREP (Authorized Rep)" then 
				Text 60, 25, 45, 10, MAXIS_case_number
				EditBox 175, 20, 45, 15, effective_date
				EditBox 310, 20, 45, 15, date_received
				EditBox 30, 65, 270, 15, address_notes
				EditBox 30, 85, 270, 15, household_notes
				EditBox 30, 105, 270, 15, Edit14
				EditBox 30, 125, 270, 15, Edit15
				EditBox 30, 145, 270, 15, Edit16
				Text 5, 5, 220, 10, "AREP (Authorized Rep)"
				Text 125, 25, 50, 10, "Effective Date:"
				Text 15, 70, 10, 10, "Q1"
				Text 245, 25, 60, 10, "Document Date:"
				GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
				Text 5, 25, 50, 10, "Case Number:"
				Text 15, 110, 10, 10, "Q3"
				Text 15, 130, 15, 10, "Q4"
				Text 15, 90, 15, 10, "Q2"
				Text 15, 150, 15, 10, ""
		
		ElseIf form_type_array(form_type_const, form_count) = "Change Report Form" Then
			EditBox 175, 15, 45, 15, effective_date
			EditBox 310, 15, 45, 15, date_received		
			EditBox 50, 45, 320, 15, address_notes
			EditBox 50, 65, 320, 15, household_notes
			EditBox 50, 125, 320, 15, income_notes
			EditBox 50, 145, 320, 15, shelter_notes
			EditBox 110, 85, 260, 15, asset_notes
			EditBox 50, 105, 320, 15, vehicles_notes
			EditBox 50, 165, 320, 15, other_change_notes
			EditBox 65, 200, 305, 15, actions_taken
			EditBox 65, 220, 305, 15, other_notes
			EditBox 75, 240, 295, 15, verifs_requested
			CheckBox 10, 285, 140, 10, "Check here to navigate to DAIL/WRIT", tikl_nav_check
			DropListBox 270, 280, 95, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
			Text 5, 5, 220, 10, "CHANGE REPORT FORM"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 5, 20, 50, 10, "Case Number:"
			Text 60, 20, 45, 10, MAXIS_case_number
			Text 125, 20, 50, 10, "Effective Date:"
			Text 245, 20, 60, 10, "Document Date:"
			GroupBox 5, 35, 370, 150, "CHANGES REPORTED"
			Text 15, 50, 30, 10, "Address:"
			Text 15, 70, 35, 10, "HH Comp:"
			Text 15, 130, 30, 10, "Income:"
			Text 15, 150, 25, 10, "Shelter:"
			Text 15, 90, 95, 10, "Assets (savings or property):"
			Text 15, 110, 30, 10, "Vehicles:"
			Text 15, 170, 20, 10, "Other:"
			GroupBox 5, 190, 370, 70, "ACTIONS"
			Text 15, 205, 45, 10, "Action Taken:"
			Text 15, 225, 45, 10, "Other Notes:"
			Text 180, 285, 90, 10, "The changes client reports:"
			Text 15, 245, 60, 10, "Verifs Requested:"
			CheckBox 10, 270, 140, 10, "Check if no notable changes reported.", checkbox_not_notable		'TODO: Need handling around this new checkbox and case note clearly if checked
		
        ElseIf form_type_array(form_type_const, form_count) = "Employment Verification Form (EVF)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Employment Verification Form (EVF)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Hospice Transaction Form" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Hospice Transaction Form"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Interim Assistance Agreement (IAA)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Interim Assistance Agreement (IAA)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Interim Assistance Authorization- SSI" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Interim Assistance Authorization- SSI"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Medical Opinion Form (MOF)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Medical Opinion Form (MOF)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Minnesota Transition Application Form (MTAF)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Minnesota Transition Application Form (MTAF)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Professional Statement of Need (PSN)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Professional Statement of Need (PSN)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Residence and Shelter Expenses Release Form" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Residence and Shelter Expenses Release Form"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""

		ElseIf form_type_array(form_type_const, form_count) = "Special Diet Information Request (MFIP and MSA)" Then
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, effective_date
			EditBox 310, 20, 45, 15, date_received
			EditBox 30, 65, 270, 15, address_notes
			EditBox 30, 85, 270, 15, household_notes
			EditBox 30, 105, 270, 15, Edit14
			EditBox 30, 125, 270, 15, Edit15
			EditBox 30, 145, 270, 15, Edit16
			Text 5, 5, 220, 10, "Special Diet Information Request (MFIP and MSA)"
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Reponses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""
		'MsgBox "form count" & form_type_array(form_type_const, form_count) 'TEST
		End If
	
		'Button Display: Buttons only display if the respective form was selected in the initial dialog. 
		btn_pos = 45		'variable to iterate down for each necessary button
		For current_form = 0 to Ubound(form_type_array, 2) 		'This iterates through the array and creates buttons for each form selected. It also positions them from top down so there aren't weird spaces in between. 
			If form_type_array(form_type_const, current_form) = "Asset Statement" then 
				PushButton 395, btn_pos, 45, 15, "ASSET", asset_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Authorization to Release Information (ATR)" Then 
				PushButton 395, btn_pos, 45, 15, "ATR", atr_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "AREP (Authorized Rep)" then 
				PushButton 395, btn_pos, 45, 15, "AREP", arep_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Change Report Form"  then 
				PushButton 395, btn_pos, 45, 15, "CHNG", change_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Employment Verification Form (EVF)"  then 
				PushButton 395, btn_pos, 45, 15, "EVF", evf_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Hospice Transaction Form"  then 
				PushButton 395, btn_pos, 45, 15, "HOSP", hospice_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Interim Assistance Agreement (IAA)"  then 
				PushButton 395, btn_pos, 45, 15, "IAA", iaa_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Interim Assistance Authorization- SSI" then 
				PushButton 395, btn_pos, 45, 15, "IAA-SSI", iaa_ssi_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Medical Opinion Form (MOF)" then 
				PushButton 395, btn_pos, 45, 15, "MOF", mof_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Minnesota Transition Application Form (MTAF)" then 
				PushButton 395, btn_pos, 45, 15, "MTAF", mtaf_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Professional Statement of Need (PSN)" then 
				PushButton 395, btn_pos, 45, 15, "PSN", psn_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Residence and Shelter Expenses Release Form" then 
				PushButton 395, btn_pos, 45, 15, "SF", sf_btn
				btn_pos = btn_pos + 15
			End If
			If form_type_array(form_type_const, current_form) = "Special Diet Information Request (MFIP and MSA)" then 
				PushButton 395, btn_pos, 45, 15, "DIET", diet_btn
				btn_pos = btn_pos + 15
			End If
			'MsgBox "Current form" & form_type_array(form_type_const, current_form)
		Next
		
		PushButton 395, 275, 50, 15, "Next Form", next_btn	'Next button to navigate from one form to the next. TODO: Determine if we need more handling around this. 
	
		'TODO: error handling 
		
		EndDialog
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		
		'Button movement handling for buttons displayed on the individual form dialogs. 
		If ButtonPressed = -1 Then ButtonPressed = next_btn 	'If the enter button is selected the script will handle this as if Next was selected 
		If ButtonPressed = next_btn Then form_count = form_count + 1	'If next is selected, it will iterate to the next form in the array and display this dialog
		If ButtonPressed >= 400 Then 'All forms have been defined at 400 or greater 
			For i = 0 to Ubound(form_type_array, 2) 	'For/Next used to iterate through the array to display the buttons stored in the array.
				If ButtonPressed = asset_btn and form_type_array(form_type_const, i) = "Asset Statement" Then form_count = i 
				If ButtonPressed = atr_btn and form_type_array(form_type_const, i) = "Authorization to Release Information (ATR)" Then form_count = i 
				If ButtonPressed = arep_btn and form_type_array(form_type_const, i) = "AREP (Authorized Rep)" Then form_count = i 
				If ButtonPressed = change_btn and form_type_array(form_type_const, i) = "Change Report Form" Then form_count = i 
				If ButtonPressed = evf_btn and form_type_array(form_type_const, i) = "Employment Verification Form (EVF)" Then form_count = i 
				If ButtonPressed = hospice_btn and form_type_array(form_type_const, i) = "Hospice Transaction Form" Then form_count = i 
				If ButtonPressed = iaa_btn and form_type_array(form_type_const, i) = "Interim Assistance Agreement (IAA)" Then form_count = i 
				If ButtonPressed = iaa_ssi_btn and form_type_array(form_type_const, i) = "Interim Assistance Authorization- SSI" Then form_count = i 
				If ButtonPressed = mof_btn and form_type_array(form_type_const, i) = "Medical Opinion Form (MOF)" Then form_count = i 
				If ButtonPressed = mtaf_btn and form_type_array(form_type_const, i) = "Minnesota Transition Application Form (MTAF)" Then form_count = i 
				If ButtonPressed = psn_btn and form_type_array(form_type_const, i) = "Professional Statement of Need (PSN)" Then form_count = i 
				If ButtonPressed = sf_btn and form_type_array(form_type_const, i) = "Residence and Shelter Expenses Release Form" Then form_count = i 
				If ButtonPressed = diet_btn and form_type_array(form_type_const, i) = "Special Diet Information Request (MFIP and MSA)" Then form_count = i 
			Next
		End If 
		'MsgBox "i" & i  TEST
		'MsgBox "form type-form count @ end" & form_type_array(form_type_const, form_count) 'TEST
Loop until form_count > Ubound(form_type_array, 2)
		

'TODO: Case Notes
script_end_procedure ("Success! The script has ended. ")
