'Code Blocks to add to Interview for the specific Interview team

'THIS block uses an array (interviewer_array) created in the Complete List of Testers to identify workers who process as a part of a unique interview only team
If NOT IsArray(interviewer_array) Then
	tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
End If

'look to see if the worker is listed as one of the interviewer workers
run_by_interview_team = False										'Default the interview team option to false
For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
	If user_ID_for_validation = worker.tester_id_number Then		'if the worker county logon ID that is running the script matches one of the interviewer_array workers
		run_by_interview_team = True 								'the script will run the interview only option
	End If
Next

'Looking for BZ Script writers to allow them to select the option.
For each tester in tester_array                         													'looping through all of the testers
	If user_ID_for_validation = tester.tester_id_number and tester.tester_population = "BZ" Then            'If the person who is running the script is a tester
		continue_with_testing_file = MsgBox("The Interview Script has two run options."  & vbCr & vbCr & "Do you want to run the Interview Team - INTERVIEW ONLY - Option?", vbQuestion + vbYesNo, "Use Interview Team Option")
		If continue_with_testing_file = vbYes Then run_by_interview_team = True
	End If
Next


'THIS Block creates an XML File with details of the the interview
Set xmlTracDoc = CreateObject("Microsoft.XMLDOM")
xmlTracPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_details_" & MAXIS_case_number & "_at_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".xml"

xmlTracDoc.async = False

Set root = xmlTracDoc.createElement("interview")
xmlTracDoc.appendChild root

Set element = xmlTracDoc.createElement("ScriptRunDate")
root.appendChild element
Set info = xmlTracDoc.createTextNode(date)
element.appendChild info

Set element = xmlTracDoc.createElement("ScriptRunTime")
root.appendChild element
Set info = xmlTracDoc.createTextNode(time)
element.appendChild info

Set element = xmlTracDoc.createElement("WorkerName")
root.appendChild element
Set info = xmlTracDoc.createTextNode(worker_name)
element.appendChild info

Set element = xmlTracDoc.createElement("CaseNumber")
root.appendChild element
Set info = xmlTracDoc.createTextNode(MAXIS_case_number)
element.appendChild info

Set element = xmlTracDoc.createElement("CaseBasket")
root.appendChild element
Set info = xmlTracDoc.createTextNode(case_pw)
element.appendChild info

Set element = xmlTracDoc.createElement("DHSFormNumber")
root.appendChild element
Set info = xmlTracDoc.createTextNode(CAF_form_number)
element.appendChild info

Set element = xmlTracDoc.createElement("DHSFormName")
root.appendChild element
Set info = xmlTracDoc.createTextNode(CAF_form_name)
element.appendChild info

Set element = xmlTracDoc.createElement("InterviewDate")
root.appendChild element
Set info = xmlTracDoc.createTextNode(interview_date)
element.appendChild info

Set element = xmlTracDoc.createElement("CaseActive")
root.appendChild element
Set info = xmlTracDoc.createTextNode(case_active)
element.appendChild info

Set element = xmlTracDoc.createElement("CasePending")
root.appendChild element
Set info = xmlTracDoc.createTextNode(case_pending)
element.appendChild info

If case_pending = True Then
	Set element = xmlTracDoc.createElement("DaysPendingAtInterview")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(item)
	element.appendChild info
End If

Set element = xmlTracDoc.createElement("InterviewPerson")
root.appendChild element
Set info = xmlTracDoc.createTextNode(who_are_we_completing_the_interview_with)
element.appendChild info

Set element = xmlTracDoc.createElement("InterviewMethod")
root.appendChild element
Set info = xmlTracDoc.createTextNode(how_are_we_completing_the_interview)
element.appendChild info

Set element = xmlTracDoc.createElement("InterviewLength")
root.appendChild element
Set info = xmlTracDoc.createTextNode(length_of_interview)
element.appendChild info

Set element = xmlTracDoc.createElement("InterviewInterpreter")
root.appendChild element
Set info = xmlTracDoc.createTextNode(interpreter_information)
element.appendChild info

Set element = xmlTracDoc.createElement("InterviewLanguage")
root.appendChild element
Set info = xmlTracDoc.createTextNode(interpreter_language)
element.appendChild info

Set element = xmlTracDoc.createElement("FormInfo")
root.appendChild element
Set info = xmlTracDoc.createTextNode(CAF_form)
element.appendChild info

Set element = xmlTracDoc.createElement("CAFDateStamp")
root.appendChild element
Set info = xmlTracDoc.createTextNode(CAF_datestamp)
element.appendChild info

Set element = xmlTracDoc.createElement("SNAPStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(snap_status)
element.appendChild info

Set element = xmlTracDoc.createElement("GRHStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(grh_status)
element.appendChild info

Set element = xmlTracDoc.createElement("MFIPStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(mfip_status)
element.appendChild info

Set element = xmlTracDoc.createElement("DWPStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(dwp_status)
element.appendChild info

Set element = xmlTracDoc.createElement("GAStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(ga_status)
element.appendChild info

Set element = xmlTracDoc.createElement("MSAStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(msa_status)
element.appendChild info

Set element = xmlTracDoc.createElement("EMERStatus")
root.appendChild element
Set info = xmlTracDoc.createTextNode(emer_status)
element.appendChild info

Set element = xmlTracDoc.createElement("UnspecifiedCASHPending")
root.appendChild element
Set info = xmlTracDoc.createTextNode(unknown_cash_pending)
element.appendChild info

Set element = xmlTracDoc.createElement("CASHRequest")
root.appendChild element
Set info = xmlTracDoc.createTextNode(cash_request)
element.appendChild info

Set element = xmlTracDoc.createElement("SNAPRequest")
root.appendChild element
Set info = xmlTracDoc.createTextNode(snap_request)
element.appendChild info

Set element = xmlTracDoc.createElement("EMERRequest")
root.appendChild element
Set info = xmlTracDoc.createTextNode(emer_request)
element.appendChild info

If cash_request = True Then
	Set element = xmlTracDoc.createElement("CASHProcess")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(the_process_for_cash)
	element.appendChild info

	Set element = xmlTracDoc.createElement("TypeOfCASH")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(type_of_cash)
	element.appendChild info

	If the_process_for_cash = "Renewal" Then
		Set element = xmlTracDoc.createElement("CASHRenewalMonth")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(next_cash_revw_mo & "/" & next_cash_revw_yr)
		element.appendChild info
	End If
End If

If snap_request = True Then
	Set element = xmlTracDoc.createElement("SNAPProcess")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(the_process_for_snap)
	element.appendChild info
	If the_process_for_snap = "Renewal" Then
		Set element = xmlTracDoc.createElement("CASHRenewalMonth")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(next_snap_revw_mo & "/" & next_snap_revw_yr)
		element.appendChild info
	End If
End If

If emer_request = True Then
	Set element = xmlTracDoc.createElement("EMERProcess")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(the_process_for_emer)
	element.appendChild info
End If

For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
	Set element = xmlTracDoc.createElement("member")
	root.appendChild element

	Set element = xmlTracDoc.createElement("ReferenceNumber")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(ref_number, the_members))
	element.appendChild info

	Set element = xmlTracDoc.createElement("LastName")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(last_name_const, the_members))
	element.appendChild info

	Set element = xmlTracDoc.createElement("FirstName")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(first_name_const, the_members))
	element.appendChild info

	Set element = xmlTracDoc.createElement("Age")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(age, the_members))
	element.appendChild info

	Set element = xmlTracDoc.createElement("RelationshipToApplicant")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(rel_to_applcnt, the_members))
	element.appendChild info

	If HH_MEMB_ARRAY(memb_is_caregiver, caregiver) = True
		Set element = xmlTracDoc.createElement("MFIPOrientation")
		root.appendChild element
		If HH_MEMB_ARRAY(orientation_needed_const, caregiver) = True and HH_MEMB_ARRAY(orientation_done_const, caregiver) = False and HH_MEMB_ARRAY(orientation_exempt_const, caregiver) = False Then
			Set info = xmlTracDoc.createTextNode("Incomplete")
			element.appendChild info
		ElseIf  HH_MEMB_ARRAY(orientation_needed_const, caregiver) = False Then
			Set info = xmlTracDoc.createTextNode("Not Needed")
			element.appendChild info
		ElseIf HH_MEMB_ARRAY(orientation_needed_const, caregiver) = True and HH_MEMB_ARRAY(orientation_done_const, caregiver) = True Then
			Set info = xmlTracDoc.createTextNode("Completed")
			element.appendChild info
		ElseIf HH_MEMB_ARRAY(orientation_needed_const, caregiver) = True and HH_MEMB_ARRAY(orientation_exempt_const, caregiver) = True Then
			Set info = xmlTracDoc.createTextNode("Exempt")
			element.appendChild info
		End If
	End If
Next

Set element = xmlTracDoc.createElement("eDRSMatchFound")
root.appendChild element
Set info = xmlTracDoc.createTextNode(edrs_match_found)
element.appendChild info

Set element = xmlTracDoc.createElement("ExpeditedScreening")
root.appendChild element
Set info = xmlTracDoc.createTextNode(expedited_screening)
element.appendChild info

Set element = xmlTracDoc.createElement("ExpeditedDetermination")
root.appendChild element
Set info = xmlTracDoc.createTextNode(is_elig_XFS)
element.appendChild info

Set element = xmlTracDoc.createElement("DeterminedIncome")
root.appendChild element
Set info = xmlTracDoc.createTextNode(determined_income)
element.appendChild info

Set element = xmlTracDoc.createElement("DeterminedAssets")
root.appendChild element
Set info = xmlTracDoc.createTextNode(determined_assets)
element.appendChild info

Set element = xmlTracDoc.createElement("DeterminedShelter")
root.appendChild element
Set info = xmlTracDoc.createTextNode(determined_shel)
element.appendChild info

Set element = xmlTracDoc.createElement("DeterminedUtilities")
root.appendChild element
Set info = xmlTracDoc.createTextNode(determined_utilities)
element.appendChild info


' Set element = xmlTracDoc.createElement("")
' root.appendChild element
' Set info = xmlTracDoc.createTextNode(item)
' element.appendChild info

' Set element = xmlTracDoc.createElement("")
' root.appendChild element
' Set info = xmlTracDoc.createTextNode(item)
' element.appendChild info

' Set element = xmlTracDoc.createElement("")
' root.appendChild element
' Set info = xmlTracDoc.createTextNode(item)
' element.appendChild info


xmlTracDoc.save(xmlTracPath)

Set xml = CreateObject("Msxml2.DOMDocument")
Set xsl = CreateObject("Msxml2.DOMDocument")

txt = Replace(fso.OpenTextFile(xmlTracPath).ReadAll, "><", ">" & vbCrLf & "<")
stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
"<xsl:output method=""xml"" indent=""yes""/>" & _
"<xsl:template match=""/"">" & _
"<xsl:copy-of select="".""/>" & _
"</xsl:template>" & _
"</xsl:stylesheet>"

xsl.loadXML stylesheet
xml.loadXML txt

xml.transformNode xsl

xml.Save xmlTracPath


'BLOCK for a dialog page to change Expedited Determination for the Interview Team
GroupBox 5, 10, 475, 225, "Expedited Detail"
Text 15, 30, 290, 10, "How much income was received (or will be received) in the application month (MM/YY)?"
EditBox 310, 25, 50, 15, exp_det_income
y_pos = 45
For cow = 0 to UBound(JOB_ARRAY_PLACEHOLDER)
	Text 25, y_pos, 235, 10, JOB_ARRAY_PLACEHOLDER(cow)
	y_pos = y_pos + 10
Next
If BUSI_INFO_PLACEHOLDER Then
	Text 25, y_pos, 235, 10, "BUSI Detail"
	y_pos = y_pos + 10
End If
For cow = 0 to UBound(UNEA_ARRAY_PLACEHOLDER)
	Text 25, y_pos, 235, 10, UNEA_ARRAY_PLACEHOLDER(cow)
	y_pos = y_pos + 10
Next

' Text 25, 45, 235, 10, "Job Item and Information"
' Text 25, 55, 235, 10, "Job Item and Information"
' Text 25, 65, 235, 10, "BUSI Detail"
' Text 25, 75, 235, 10, "UNEA Item and Information"
' Text 25, 85, 235, 10, "UNEA Item and Information"
y_pos = y_pos + 10

Text 15, y_pos, 330, 10, "How much does the household have in assets (accounts and cash) in the application month (MM/YY)?"
EditBox 350, y_pos-5, 50, 15, exp_det_assets
y_pos = y_pos + 15
for egg = 0 to UBound(ASSET_ARRAY_PLACEHOLDER)
	Text 25, y_pos, 235, 10, ASSET_ARRAY_PLACEHOLDER(egg)
	y_pos = y_pos + 10
Next
' Text 25, 120, 235, 10, "Asset Details"
' Text 25, 130, 235, 10, "Asset Details"
' Text 25, 140, 235, 10, "Asset Details"
' Text 25, 150, 235, 10, "Asset Details"
y_pos = y_pos + 10

Text 15, y_pos, 305, 10, "How much does the household pay in housing expenses in the application month (MM/YY)?"
EditBox 320, y_pos-5, 50, 15, exp_det_housing
y_pos = y_pos + 15
If HOUSING_INFO_PLACEHOLDER Then
	Text 25, y_pos, 235, 10, "BUSI Detail"
	y_pos = y_pos + 10
End If
' Text 25, 185, 235, 10, "Housing Details"
y_pos = y_pos + 10

Text 15, y_pos, 315, 10, "Which type of utilities is the household responsible to pay in the application month (MM/YY)?"
y_pos = y_pos + 15
CheckBox 25, y_pos, 45, 10, "Heat", heat_exp_checkbox
CheckBox 90, y_pos, 70, 10, "Air Conditioning", ac_exp_checkbox
CheckBox 175, y_pos, 45, 10, "Electric", electric_exp_checkbox
CheckBox 240, y_pos, 55, 10, "Telephone", phone_exp_checkbox

'MOCK UP DIALOG CODE
' BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"
'   GroupBox 5, 10, 475, 225, "Expedited Detail"
'   Text 485, 5, 75, 10, "---   DIALOGS   ---"
'   Text 485, 15, 10, 10, "1"
'   Text 485, 30, 10, 10, "2"
'   Text 485, 45, 10, 10, "3"
'   Text 485, 60, 10, 10, "4"
'   Text 485, 75, 10, 10, "5"
'   Text 485, 90, 10, 10, "6"
'   Text 485, 105, 10, 10, "7"
'   Text 485, 120, 10, 10, "8"
'   Text 485, 135, 10, 10, "9"
'   Text 485, 150, 10, 10, "10"
'   Text 485, 165, 10, 10, "11"
'   ButtonGroup ButtonPressed
'     PushButton 10, 365, 130, 15, "Interview Ended - INCOMPLETE", incomplete_interview_btn
'     PushButton 140, 365, 130, 15, "View Verifications", verif_button
'     PushButton 415, 365, 50, 15, "NEXT", next_btn
'     PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
'   Text 15, 30, 290, 10, "How much income was received (or will be received) in the application month (MM/YY)?"
'   EditBox 310, 25, 50, 15, Edit1
'   Text 25, 45, 235, 10, "Job Item and Information"
'   Text 25, 55, 235, 10, "Job Item and Information"
'   Text 25, 65, 235, 10, "BUSI Detail"
'   Text 25, 75, 235, 10, "UNEA Item and Information"
'   Text 25, 85, 235, 10, "UNEA Item and Information"
'   Text 15, 105, 330, 10, "How much does the household have in assets (accounts and cash) in the application month (MM/YY)?"
'   EditBox 350, 100, 50, 15, Edit2
'   Text 25, 120, 235, 10, "Asset Details"
'   Text 25, 130, 235, 10, "Asset Details"
'   Text 25, 140, 235, 10, "Asset Details"
'   Text 25, 150, 235, 10, "Asset Details"
'   Text 15, 170, 305, 10, "How much does the household pay in housing expenses in the application month (MM/YY)?"
'   EditBox 320, 165, 50, 15, Edit3
'   Text 25, 185, 235, 10, "Housing Details"
'   Text 15, 205, 315, 10, "Which type of utilities is the household responsible to pay in the application month (MM/YY)?"
'   CheckBox 25, 220, 45, 10, "Heat", heat_exp_checkbox
'   CheckBox 90, 220, 70, 10, "Air Conditioning", ac_exp_checkbox
'   CheckBox 175, 220, 45, 10, "Electric", electric_exp_checkbox
'   CheckBox 240, 220, 55, 10, "Telephone", phone_exp_checkbox
' EndDialog



'BLOCK for update to Program Selection Dialog
BeginDialog Dialog1, 0, 0, 326, 310, "Programs to Interview For"
	Text 10, 10, 300, 10, "Record details from the form here for CAF_form_name being used for this interview:"
	Text 15, 30, 125, 10, "Date form was received in the county:"
	EditBox 140, 25, 45, 15, CAF_datestamp
	Text 200, 30, 110, 10, "CAF_form_name"
	Text 25, 45, 260, 10, "Active Programs: list_active_programs"
	Text 25, 55, 260, 10, "Pending Programs: list_pending_programs"
	GroupBox 10, 75, 265, 30, "Check All Programs Marked on the Form"
	CheckBox 15, 90, 30, 10, "CASH", CASH_on_CAF_checkbox
	CheckBox 55, 90, 35, 10, "SNAP", SNAP_on_CAF_checkbox
	CheckBox 95, 90, 55, 10, "EMERGENCY", EMER_on_CAF_checkbox
	CheckBox 160, 90, 100, 10, "HOUSING SUPPORT (GRH)", GRH_on_CAF_checkbox
	Text 15, 110, 180, 10, "About the different programs:"
	Text 20, 120, 245, 10, "- CASH is a monthly cash benefit."
	Text 20, 130, 245, 10, "- SNAP is a monthly benefit for the purchase of food items only."
	Text 20, 140, 245, 10, "- EMERGENCY is a one-time payment to resolve an emergency situation."
	Text 25, 150, 245, 10, "An example of emergency situation is eviction or utility disconnect."
	Text 20, 160, 265, 10, "- HOUSING SUPPORT is monthly benefit for people working with an organization"
	Text 25, 170, 125, 10, "or facility for housing supports."
	Text 15, 185, 200, 10, "Confirm with the resident these were the programs selected."
	Text 15, 195, 245, 10, "Explain to the resident they can verbally request additional programs to be"
	Text 30, 205, 125, 10, "assessed while their case is pending. "
	Text 15, 215, 270, 10, "Explain additionally to the resident they can withdraw their requests at any time."
	Text 15, 245, 85, 10, "Program Request Notes:"
	EditBox 15, 255, 300, 15, program_request_notes
	Text 115, 270, 205, 10, "(Do not document verbal program request or withdrawls here.)"
	ButtonGroup ButtonPressed
		PushButton 150, 225, 130, 15, "Press Here to Add Verbal Requests", program_requests_btn
		OkButton 210, 285, 50, 15
		CancelButton 265, 285, 50, 15
EndDialog

Call verbal_requests



'BLOCK for having a Verbal Request and Withdrawal Option - accessed after CAF forms selection OR button on Main Dialog
'TODO - add a main dialog button to access this functionality
function verbal_requests()
	program_marked_on_CAF = False
	all_programs_marked_on_CAF = True
	If CASH_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If SNAP_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If EMER_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If GRH_on_CAF_checkbox = checked Then program_marked_on_CAF = True

	If CASH_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If SNAP_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If EMER_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If GRH_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 316, 225, "Programs to Interview For"
		GroupBox 10, 10, 295, 60, "Form Details:"
		Text 20, 25, 155, 10, "CAF_form_name"
		Text 20, 40, 125, 10, "CAF Date: CAF_date"
		Text 190, 15, 95, 10, "Programs Marked on Form:"
		prog_y_pos = 25
		If CASH_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 50, 10, "- CASH"
			prog_y_pos = prog_y_pos + 10
		End If
		If SNAP_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 50, 10, "- SNAP"
			prog_y_pos = prog_y_pos + 10
		End If
		If EMER_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 70, 10, "- EMERGENCY"
			prog_y_pos = prog_y_pos + 10
		End If
		If GRH_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 100, 10, "- HOUSING SUPPORT / GRH"
			prog_y_pos = prog_y_pos + 10
		End If
		If prog_y_pos = 25 Then Text 195, prog_y_pos, 100, 10, "NONE"

		If all_programs_marked_on_CAF = False Then
			req_y_pos = 100
			If CASH_on_CAF_checkbox = unchecked Then
				Text 65, req_y_pos, 25, 10, " Cash:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", cash_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If SNAP_on_CAF_checkbox = unchecked Then
				Text 65, req_y_pos, 25, 10, "SNAP:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", snap_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If EMER_on_CAF_checkbox = unchecked Then
				Text 40, req_y_pos, 50, 10, "EMERGENCY:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", emer_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If GRH_on_CAF_checkbox = unchecked Then
				Text 15, req_y_pos, 75, 10, " HOUSING SUPPORT:"
				DropListBox 90, req_y_pos, 60, 45, "No"+chr(9)+"Yes", grh_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			GroupBox 10, 80, 145, req_y_pos-80, "VERBAL PROGRAM REQUESTS "
		End If
		If program_marked_on_CAF = True Then
			If CASH_on_CAF_checkbox = checked Then
				Text 215, 100, 25, 10, " Cash:"
				DropListBox 240, 95, 60, 45, "No"+chr(9)+"Yes", cash_verbal_withdraw
			End If
			If SNAP_on_CAF_checkbox = checked Then
				Text 215, 115, 25, 10, "SNAP:"
				DropListBox 240, 110, 60, 45, "No"+chr(9)+"Yes", snap_verbal_withdraw
			End If
			If EMER_on_CAF_checkbox = checked Then
				Text 190, 130, 50, 10, "EMERGENCY:"
				DropListBox 240, 125, 60, 45, "No"+chr(9)+"Yes", emer_verbal_withdraw
			End If
			If GRH_on_CAF_checkbox = checked Then
				Text 165, 145, 75, 10, " HOUSING SUPPORT:"
				DropListBox 240, 140, 60, 45, "No"+chr(9)+"Yes", grh_verbal_withdraw
			End If
			GroupBox 160, 80, 145, 80, "VERBAL PROGRAM WITHDRAWALS"
		End If

		Text 10, 170, 220, 10, "Additional Notes about Verbal Program Requests or Withdrawals"
		EditBox 10, 180, 295, 15, verbal_request_notes
		ButtonGroup ButtonPressed
			' OkButton 195, 200, 50, 15
			' CancelButton 255, 200, 50, 15
			PushButton 255, 200, 50, 15, "Return", return_btn
	EndDialog

	Do
		dialog Dialog1
		cancel_confirmation

	Loop until ButtonPressed = return_btn
	ButtonPressed = ""

	cash_request = False
	snap_request = False
	emer_request = False
	grh_request = False
	If CASH_on_CAF_checkbox = checked OR cash_verbal_request = "Yes" Then cash_request = True
	If SNAP_on_CAF_checkbox = checked OR snap_verbal_request = "Yes" Then snap_request = True
	If EMER_on_CAF_checkbox = checked OR emer_verbal_request = "Yes" Then emer_request = True
	If GRH_on_CAF_checkbox = checked OR grh_verbal_request = "Yes" Then grh_request = True

	run_process_selection = False
	If cash_request = True Then
		If type_of_cash = "?" or type_of_cash = "" Then run_process_selection = True
		If the_process_for_cash = "Select One..." or the_process_for_cash = "" Then run_process_selection = True
	End If
	If snap_request = True Then
		If the_process_for_snap = "Select One..." or the_process_for_snap = "" Then run_process_selection = True
	End If
	If emer_request = True Then
		If type_of_emer = "?" or type_of_emer = "" Then run_process_selection = True
		If the_process_for_emer = "Select One..." or the_process_for_emer = "" Then run_process_selection = True
	End If
	If grh_request = True Then
		If the_process_for_grh = "Select One..." or the_process_for_grh = "" Then run_process_selection = True
	End If

	If run_process_selection = True Then call program_process_selection

	ButtonPressed = ""
end function

'Dialog Code
' BeginDialog Dialog1, 0, 0, 316, 225, "Programs to Interview For"
'   GroupBox 10, 10, 295, 60, "Form Details:"
'   Text 20, 25, 155, 10, "CAF_form_name"
'   Text 20, 40, 125, 10, "CAF Date: CAF_date"
'   Text 190, 15, 95, 10, "Programs Marked on Form:"
'   Text 195, 25, 50, 10, "- CASH"
'   Text 195, 35, 50, 10, "- SNAP"
'   Text 195, 45, 70, 10, "- EMERGENCY"
'   Text 195, 55, 100, 10, "- HOUSING SUPPORT / GRH"
'   GroupBox 10, 80, 145, 80, "VERBAL PROGRAM REQUESTS "
'   Text 65, 100, 25, 10, " Cash:"
'   DropListBox 90, 95, 60, 45, "No"+chr(9)+"Yes", cash_verbal_request
'   Text 65, 115, 25, 10, "SNAP:"
'   DropListBox 90, 110, 60, 45, "No"+chr(9)+"Yes", snap_verbal_request
'   Text 40, 130, 50, 10, "EMERGENCY:"
'   DropListBox 90, 125, 60, 45, "No"+chr(9)+"Yes", emer_verbal_request
'   Text 15, 145, 75, 10, " HOUSING SUPPORT:"
'   DropListBox 90, 140, 60, 45, "No"+chr(9)+"Yes", List4
'   GroupBox 160, 80, 145, 85, "VERBAL PROGRAM WITHDRAWALS"
'   Text 215, 100, 25, 10, " Cash:"
'   DropListBox 240, 95, 60, 45, "No"+chr(9)+"Yes", List9
'   Text 215, 115, 25, 10, "SNAP:"
'   DropListBox 240, 110, 60, 45, "No"+chr(9)+"Yes", List10
'   Text 190, 130, 50, 10, "EMERGENCY:"
'   DropListBox 240, 125, 60, 45, "No"+chr(9)+"Yes", List11
'   Text 165, 145, 75, 10, " HOUSING SUPPORT:"
'   DropListBox 240, 140, 60, 45, "No"+chr(9)+"Yes", List12
'   Text 10, 170, 220, 10, "Additional Notes about Verbal Program Requests or Withdrawals"
'   EditBox 10, 180, 295, 15, verbal_request_notes
'   ButtonGroup ButtonPressed
'     OkButton 195, 200, 50, 15
'     CancelButton 255, 200, 50, 15
' EndDialog





'BLOCK for process being reviewed dialog update
function program_process_selection()
	dlg_len = 100
	y_pos = 75
	If cash_request = True Then dlg_len = dlg_len + 20
	If snap_request = True Then dlg_len = dlg_len + 20
	If emer_request = True Then dlg_len = dlg_len + 20
	If grh_request = True Then dlg_len = dlg_len + 20
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 205, dlg_len, "CAF Process"
		Text 10, 10, 205, 20, "Interviews are completed on cases when programs are initially requested and at annual renewal for SNAP and MFIP."
		Text 10, 35, 210, 20, "To correctly identify the information needed, each program needs to be associated with an Application or Renewal process."

		Text 10, 60, 35, 10, "Program"
		Text 80, 60, 50, 10, "CAF Process"
		Text 155, 60, 50, 10, "Recert MM/YY"
		If cash_request = True Then
			Text 10, y_pos + 5, 20, 10, "Cash"
			DropListBox 35, y_pos, 35, 45, "?"+chr(9)+"Family"+chr(9)+"Adult", type_of_cash
			DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_cash
			EditBox 155, y_pos, 20, 15, next_cash_revw_mo
			EditBox 180, y_pos, 20, 15, next_cash_revw_yr
			y_pos = y_pos + 20
		End If
		If snap_request = True Then
			Text 10, y_pos + 5, 20, 10, "SNAP"
			DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_snap
			EditBox 155, y_pos, 20, 15, next_snap_revw_mo
			EditBox 180, y_pos, 20, 15, next_snap_revw_yr
			y_pos = y_pos + 20
		End If
		If emer_request = True Then
			Text 10, y_pos + 5, 20, 10, "EMER"
			DropListBox 35, y_pos, 35, 45, "?"+chr(9)+"EA"+chr(9)+"EGA", type_of_emer
			DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application", the_process_for_emer
			y_pos = y_pos + 20
		End If
		If grh_request = True Then
			Text 10, y_pos + 5, 20, 10, "GRH"
			DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_grh
			EditBox 155, y_pos, 20, 15, next_grh_revw_mo
			EditBox 180, y_pos, 20, 15, next_grh_revw_yr
			y_pos = y_pos + 20
		End If
		y_pos = y_pos + 5
		Text 10, y_pos+5, 125, 10, "(The programs do not need to match.)"
		ButtonGroup ButtonPressed
			OkButton 150, y_pos, 50, 15
	EndDialog

	Do
		DO
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation

			If len(next_cash_revw_yr) = 4 AND left(next_cash_revw_yr, 2) = "20" Then next_cash_revw_yr = right(next_cash_revw_yr, 2)
			If len(next_snap_revw_yr) = 4 AND left(next_snap_revw_yr, 2) = "20" Then next_snap_revw_yr = right(next_snap_revw_yr, 2)
			If cash_request = True Then
				If the_process_for_cash = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the CASH program is at application or renewal."
				If the_process_for_cash = "Renewal" AND (len(next_cash_revw_mo) <> 2 or len(next_cash_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For CASH at renewal, enter the footer month and year the of the renewal."
			End If
			If snap_request = True Then
				If the_process_for_snap = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the SNAP program is at application or renewal."
				If the_process_for_snap = "Renewal" AND (len(next_snap_revw_mo) <> 2 or len(next_snap_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For SNAP at renewal, enter the footer month and year the of the renewal."
			End If
			If emer_request = True Then
				If type_of_emer = "?" Then r_msg = err_msg & vbNewLine & "*Indicate if EMER request in EA or EGA"
			End If
			If grh_request = True Then
				If the_process_for_grh = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the GRH program is at application or renewal."
				If the_process_for_grh = "Renewal" AND (len(next_grh_revw_mo) <> 2 or len(next_grh_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For GRH at renewal, enter the footer month and year the of the renewal."
			End If


			IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

end function
'additional info dialog code
' BeginDialog Dialog1, 0, 0, 226, 125, "CAF Process"
'   ButtonGroup ButtonPressed
'     OkButton 170, 105, 50, 15
'   Text 10, 10, 205, 20, "Interviews are completed on cases when programs are initially requested and at annual renewal for SNAP and MFIP."
'   Text 10, 35, 210, 20, "To correctly identify the information needed, each program needs to be associated with an Application or Renewal process."
'   Text 10, 60, 35, 10, "Program"
'   Text 10, 110, 125, 10, "(The programs do not need to match.)"
' EndDialog



'BLOCK for Final Wrap Dialog - currently line 10089
BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"
  Text 10, 10, 395, 10, "We have finished gathering all the information for the interview. Finish by reviewing this information with the resident."
  GroupBox 10, 30, 530, 315, "CASE INTERVIEW WRAP UP"
  Text 15, 45, 505, 10, "Explain Verifications:"
  Text 20, 55, 505, 10, "If verifications are needed, a request will be sent in the mail. Provide proofs quickly, as they are due in 10 days."
  Text 20, 65, 505, 10, "We can help you obtain these verifications if you have any difficulties. Contact us by phone or come to a service center if you need help."
  Text 15, 85, 460, 10, "Your case will be processed by another worker, there is a possibility they will need to contact you with additional clarifications."
  Text 25, 100, 150, 10, "Confirm the best Phone Number to reach you:"
  ComboBox 175, 95, 85, 45, "", phone_number_selection				'TODO - add phone comboBox functionality
  Text 270, 100, 170, 10, "Can we leave a detailed message at this number?"
  DropListBox 440, 95, 60, 45, "", leave_a_message
  Text 25, 115, 400, 10, "Do you have any questions or requests I can pass on to the processing worker or a program specialist?"
  EditBox 25, 125, 475, 15, resident_questions
  Text 15, 155, 505, 10, "Your address and phone number are our best way to contact you, let us know of these changes so you do not miss any notices or requests."
  Text 20, 165, 505, 10, "Our mail does not forward and missing notices can cause your benefits to end."
  Text 15, 185, 505, 10, "If you are unsure of program rules and requirements, the forms we reviewed earlier can always be resent, or you can call us with questions."
  GroupBox 15, 210, 505, 95, "Contact to Hennepin County by phone, in person, or online. Ask the resident if they need any more details:"
  Text 20, 220, 40, 10, "By Phone:"
  Text 60, 220, 450, 10, "612-596-1300. The phone lines are open Monday - Friday from 9:00 - 4:00"
  Text 20, 230, 40, 10, "In person: "
  Text 60, 230, 170, 10, "Northwest Human Service Center"
  Text 230, 230, 200, 10, "7051 Brooklyn Blvd Brooklyn Center 55429"
  Text 60, 240, 170, 10, "North Minneapolis Service Center"
  Text 230, 240, 200, 10, "1001 Plymouth Ave N Minneapolis 55411"
  Text 60, 250, 170, 10, "South Minneapolis Human Service Center"
  Text 230, 250, 200, 10, "2215 East Lake Street Minneapolis 55407"
  Text 60, 260, 170, 10, "Health Services Building (Downtown Minneapolis)"
  Text 230, 260, 200, 10, "525  Portland Ave S (5th floor) Minneapolis 55415"
  Text 60, 270, 170, 10, "South Suburban Human Service Center"
  Text 230, 270, 200, 10, "9600 Aldrich Ave S Bloomington 55420"
  Text 20, 280, 40, 10, "Online:"
  Text 60, 280, 400, 10, "MNBenefits  at  https://mnbenefits.mn.gov/  -  Use for submitting applications and documents."
  Text 60, 290, 465, 10, "InfoKeep  at  https://infokeep.hennepin.us/  -  Create a unique sign in to submit documents directly to your case, has a chat functionality."
  Text 15, 315, 270, 10, "Summarize any additional case details to pass on to the processing worker:"
  EditBox 15, 325, 520, 15, case_summary
  ButtonGroup ButtonPressed
    PushButton 465, 365, 80, 15, "Interview Completed", continue_btn
EndDialog


