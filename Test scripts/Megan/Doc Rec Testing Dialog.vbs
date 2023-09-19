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

Function all_forms_checkboxes()
	BeginDialog Dialog1, 0, 0, 216, 205, "Document Selection"
	ButtonGroup ButtonPressed
		PushButton 85, 185, 70, 15, "Review Selections", review_selections
		CancelButton 165, 185, 40, 15
	GroupBox 5, 5, 205, 170, "Directions: Select all documents received, then select OK."
	CheckBox 15, 140, 160, 10, "Special Diet Information Request (MFIP & MSA )", Check1
	CheckBox 15, 20, 160, 10, "Asset Statement", Check2
	CheckBox 15, 30, 160, 10, "AREP (Authorized Rep)", Check3
	CheckBox 15, 40, 160, 10, "Authorization to Release Information (ATR)", Check4
	CheckBox 15, 50, 160, 10, "Change Report Form", Check5
	CheckBox 15, 60, 160, 10, "Employment Verification Form (EVF)", Check8
	CheckBox 15, 70, 160, 10, "Hospice Transaction Form", Check9
	CheckBox 15, 80, 160, 10, "Interim Assistance Agreement (IAA)", Check10
	CheckBox 15, 100, 160, 10, "Medical Opinion Form (MOF)", Check11
	CheckBox 15, 110, 160, 10, "Minnesota Transition Application Form (MTAF)", Check12
	CheckBox 15, 120, 160, 10, "Professional Statement of Need (PSN)", Check13
	CheckBox 15, 130, 170, 10, "Residence and Shelter Expenses Release Form", Check14
	CheckBox 15, 90, 160, 10, "Interim Assistance Authorization- SSI", Check15
	CheckBox 15, 150, 160, 10, "MISC/OTHER", Check16
	EndDialog
	dialog dialog1
End Function


'Button Defined
add_button = 201
all_forms = 202

call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
doc_date_stamp = "04/12/2023"

'FIRST DIALOG COLLECTING CASE & MONTH/YEAR===========================================================================
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 136, 95, "Case number dialog"
	EditBox 65, 10, 65, 15, MAXIS_case_number
	EditBox 65, 30, 30, 15, MAXIS_footer_month
	EditBox 100, 30, 30, 15, MAXIS_footer_year
	EditBox 85, 50, 45, 15, doc_date_stamp
	ButtonGroup ButtonPressed
	OkButton 25, 75, 50, 15
	CancelButton 80, 75, 50, 15
	Text 10, 15, 50, 10, "Case number: "
	Text 10, 35, 50, 10, "Footer month:"
	Text 10, 55, 75, 10, "Document date stamp:"
EndDialog

Do
	DO
	err_msg = ""
	dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
	cancel_confirmation
	Call validate_MAXIS_case_number(err_msg, "*")
	IF IsNumeric(MAXIS_footer_month) = FALSE OR IsNumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbNewLine &  "* You must type a valid footer month and year."
	If IsDate(doc_date_stamp) = FALSE Then err_msg = err_msg & vbNewLine & "* Please enter a valid document date."
	If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'SECOND DIALOG FORM SELECTION===========================================================================
'Defined y_pos so that I can use y_pos to add forms selected to the dialog


' 'PHASE 1 DIALOG DOCUMENT SELECTION
' Do	'Phase 1: Currently this do loop bring the user back to the Select Documents Received after msgbox/all forms dialog
' Dialog1 = "" 'Blanking out previous dialog detail
' BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
' y_pos = 30
' ComboBox 30, y_pos, 180, 15, "...Select or Type"+chr(9)+"Asset Statement"+chr(9)+"AREP (Authorized Rep)"+chr(9)+"Authorization to Release Information (ATR)"+chr(9)+"Change Report Form"+chr(9)+"Employment Verification Form (EVF)"+chr(9)+"Hospice Transaction Form"+chr(9)+"Interim Assistance Agreement (IAA)"+chr(9)+"Medical Opinion Form (MOF)"+chr(9)+"Minnesota Transition Application Form (MTAF)"+chr(9)+"Professional Statement of Need (PSN)"+chr(9)+"Residence and Shelter Expenses Release Form"+chr(9)+"SSI Interim Assistance Authorization"+chr(9)+"Special Diet Information Request (MFIP & MSA )"+chr(9)+"MISC/OTHER", Form_type
' ButtonGroup ButtonPressed
' PushButton 225, y_pos, 35, 10, "Add", add_button
' PushButton 225, y_pos + 30, 35, 10, "All Forms", all_forms
' OkButton 205, y_pos + 185, 40, 15
' CancelButton 255, y_pos + 185, 40, 15
' GroupBox 5, 5, y_pos + 250, 70, "Directions: For each document received either:"
' Text 15, 15, y_pos + 245, 10, "1. Select document from dropdown, then select Add button. Repeat for each form."
' Text 10, y_pos + 15, 15, 10, "OR"
' Text 15, y_pos + 30, 180, 10, "2. Select All Forms to select forms via checkboxes."
' GroupBox 45, y_pos + 55, 210, 125, "Documents Selected"
' EndDialog

' dialog dialog1	'TODO: Place this in a do loop and add handling to ensure the user selected the correct entries or will be warned
' cancel_confirmation


' If ButtonPressed = add_button Then MsgBox form_type 	'Phase 1: If Add is selected, then msg box selected form ' TODO: Store selection and list selection in dialog
' If ButtonPressed = all_forms Then Call all_forms_checkboxes		'Phase 1: Brings user to next dialg. 'TODO: Need coding to go back to previous dialo
' Loop 

'PHASE 2 DOCUMENT SELECTION
'Form_Type_Array = Array("Asset Statement", "AREP (Authorized Rep)", "Authorization to Release (ATR)", "Change Report Form", "Employment Verification Form (EVF)", "Hospice Transaction Form", "Interim Assistance Agreement (IAA)", "Interim Assistance Agreement-SSI", "Medical Opinion Form (MOF)", "Minnesota Transition Application Form (MTAF)", "Professional Statement of Need (PSN)", "Residence and Shelter Expenses Release Form", "Special Diet Information Request")

Dim form_type_array()
ReDim form_type_array(0)

form_count = 0

Do	
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
	y_pos = 30
	ComboBox 30, y_pos, 180, 15, "...Select or Type"+chr(9)+"Asset Statement"+chr(9)+"AREP (Authorized Rep)"+chr(9)+"Authorization to Release Information (ATR)"+chr(9)+"Change Report Form"+chr(9)+"Employment Verification Form (EVF)"+chr(9)+"Hospice Transaction Form"+chr(9)+"Interim Assistance Agreement (IAA)"+chr(9)+"Medical Opinion Form (MOF)"+chr(9)+"Minnesota Transition Application Form (MTAF)"+chr(9)+"Professional Statement of Need (PSN)"+chr(9)+"Residence and Shelter Expenses Release Form"+chr(9)+"SSI Interim Assistance Authorization"+chr(9)+"Special Diet Information Request (MFIP & MSA )"+chr(9)+"MISC/OTHER", Form_type
	ButtonGroup ButtonPressed
	PushButton 225, y_pos, 35, 10, "Add", add_button
	PushButton 225, y_pos + 30, 35, 10, "All Forms", all_forms
	OkButton 205, y_pos + 185, 40, 15
	CancelButton 255, y_pos + 185, 40, 15
	GroupBox 5, 5, y_pos + 250, 70, "Directions: For each document received either:"
	Text 15, 15, y_pos + 245, 10, "1. Select document from dropdown, then select Add button. Repeat for each form."
	Text 10, y_pos + 15, 15, 10, "OR"
	Text 15, y_pos + 30, 180, 10, "2. Select All Forms to select forms via checkboxes."
	GroupBox 45, y_pos + 55, 210, 125, "Documents Selected"

	If ButtonPressed = add_button Then
		ReDim Preserve form_type_array (form_count)
		form_type_array(form_count) = 
		form_count= form_count + 1 
	End If 

	
		For each form in form_type_array
			Text 50, y_pos, 195, 10, form
			y_pos = y_pos + 15
		Next
	End If

	If ButtonPressed = all_forms Then Call all_forms_checkboxes

	EndDialog
	
	dialog dialog1	
	cancel_confirmation
Loop until err_msg = ""


