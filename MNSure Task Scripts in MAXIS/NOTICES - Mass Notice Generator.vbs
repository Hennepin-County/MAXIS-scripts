'Call Functions--------------------------------------------------------------------------------------------

Function save_word_document(docname)
	objDoc.SaveAs("M:\Income-Maintence-Share\Pending Notices Generator\" & docname &".docx")
End Function

Function print_word_doc
	objDoc.PrintOut()
End Function

Function quit_excel
	objExcel.Quit
End Function

Function quit_word
	objWord.Quit
End Function

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog notice_generation_dialog, 0, 0, 106, 62, "Batch Notices"
  EditBox 41, 3, 60, 12, rfi_due_date
  CheckBox 5, 17, 101, 10, "Save Copies of Documents", save_copies
  CheckBox 5, 30, 50, 10, "Print Copies", print_copies
  ButtonGroup ButtonPressed
    OkButton 48, 47, 20, 12
    CancelButton 72, 47, 30, 12
  Text 6, 5, 32, 8, "Due Date"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

Dialog notice_generation_dialog
	If buttonpressed = 0 then stopscript

'Set Excel as a possible program to open
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("M:\Income-Maintence-Share\Pending Notices Generator\Pending Notice List.xls")
objExcel.Visible = True

intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""
    
	'Read Excel Doc--------------------------------------------------
	client_first_name = objExcel.Cells(intRow, 1).Value
    client_last_name = objExcel.Cells(intRow, 2).Value
	client_application_date = objExcel.Cells(intRow, 3).Value
    client_case_number = objExcel.Cells(intRow, 4).Value
    client_address_line_1 = objExcel.Cells(intRow, 5).Value
    client_address_line_2 = objExcel.Cells(intRow, 6).Value
	rsn_income_type = objExcel.Cells(intRow, 7).Value
	rsn_excluded_income = objExcel.Cells(intRow, 8).Value
	rsn_citizen_status_code = objExcel.Cells(intRow, 9).Value
	rsn_ssnid = objExcel.Cells(intRow, 10).Value
	rsn_incarceration = objExcel.Cells(intRow, 11).Value
	rsn_deductions = objExcel.Cells(intRow, 12).Value
	rsn_projected_annual_income = objExcel.Cells(intRow, 13).Value	
	rsn_applied_for_ssn = objExcel.Cells(intRow, 14).Value	
	'Create Notice---------------------------------------------------
	
		'Open New Doc
	Set objWord = CreateObject("Word.Application")
	objWord.Visible = True	
		
	Set objDoc = objWord.Documents.Add()
	Set objSelection = objWord.Selection
				
		'Write Document
	
	Set colShapes = objDoc.Shapes
	Set objShape = colShapes.AddPicture("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\Pending Notices\mnsurelogo.jpg")
    objShape.height = 120
	objShape.width = 150
	objShape.left = 255
	objShape.top = -220
		
	objSelection.ParagraphFormat.SpaceAfter = 0
	objSelection.Font.Name = "Arial"
	objSelection.Font.Size = "11"
	objSelection.ParagraphFormat.LeftIndent = 0	
	objSelection.ParagraphFormat.Alignment = 0
		'Word Line 1
	objSelection.TypeText county_name & Chr(11)

		'Word Line 2
	objSelection.TypeText county_address_line_01 & Chr(11)
	'objSelection.TypeParagraph()
	
		'Word Line 3
	objSelection.TypeText county_address_line_02 & Chr(11)

		'Word Line 4
	objSelection.TypeParagraph()
		
		'Word Line 5
	objSelection.Font.Size = "10"	
	objSelection.TypeText "Case Number: "&client_case_number & Chr(11)
	
		'Word Line 6
	objSelection.TypeParagraph()
	
		'Word Line 7
	objSelection.TypeParagraph()
	
		'Word Line 8
	objSelection.ParagraphFormat.LeftIndent = 0
	objSelection.TypeText client_first_name&" "&client_last_name & Chr(11)
	
		'Word Line 9
	objSelection.TypeText client_address_line_1 & Chr(11)
	
		'Word Line 10
	objSelection.ParagraphFormat.SpaceAfter = 15
	objSelection.TypeText client_address_line_2
	objSelection.TypeParagraph()
	objSelection.ParagraphFormat.SpaceAfter = 0
	
		'Word Line 11
	objSelection.ParagraphFormat.LeftIndent = 0
	
		'Word Line 12
		
	objSelection.ParagraphFormat.SpaceBefore = 5
	objSelection.ParagraphFormat.SpaceAfter = 3
		'HEALTH CARE NOTICE TABLE
	Set objRange = objSelection.Range
	'objRange.Collapse 0
	objDoc.Tables.Add objRange, 1, 1
	Set objTable = objSelection.Tables(1)	
	objTable.Borders.Enable = True
	objTable.Range.Font.Size = 14
	objTable.Range.Font.Bold = True
	objSelection.ParagraphFormat.Alignment = 1
	objTable.Cell(1, 1).Range.Text = "Health Care Notice"

	objSelection.EndKey 6

	objSelection.ParagraphFormat.SpaceBefore = 0
	objSelection.ParagraphFormat.SpaceAfter = 0
	
		'Word Line 13
	objSelection.TypeParagraph()
	objSelection.Font.Size = "10"

	objSelection.font.bold = true
	objSelection.ParagraphFormat.SpaceAfter = 3
	objSelection.TypeText " Health Care Results"	
	objSelection.TypeParagraph()
	objSelection.TypeText " " & client_first_name & " " & client_last_name & " - MNSure Case Number: " & client_case_number 
	objSelection.TypeParagraph()
	objSelection.font.bold = false
	objSelection.ParagraphFormat.SpaceAfter = 0
	
	'Next Table'
	Set objRange = objSelection.Range
	'objRange.Collapse 0
	objDoc.Tables.Add objRange, 2, 3
	Set objTable = objSelection.Tables(1)	
	objTable.Borders.Enable = True
	objTable.Range.Font.Size = 10
	objTable.Cell(1, 1).Range.Font.Bold = True
	objTable.Cell(1, 1).Range.Text = "Effective Date"
	objTable.Cell(1, 2).Range.Font.Bold = True
	objTable.Cell(1, 2).Range.Text = "Action"
	objTable.Cell(1, 3).Range.Font.Bold = True
	objTable.Cell(1, 3).Range.Text = "Coverage Type"
	objTable.Cell(2, 1).Range.Text = client_application_date
	objTable.Cell(2, 2).Range.Text = "Pending"
	objTable.Cell(2, 3).Range.Text = "Medical Assistance"

	objSelection.EndKey 6
	
	objSelection.TypeParagraph()
	
	
	objSelection.ParagraphFormat.LeftIndent = 24
	objSelection.TypeText "    Your application for health care assistance is pending. We need information or proofs for the items listed below to decide if you qualify. Send in the information or proofs by the due date below. Write your MNsure ID on anything you send. You may not be able to get assistance if you do not send in the information or proofs by the due date below. Call us if you have questions or need help getting any of the information or proofs. (Minnesota Rules, part 9505.0090, subpart 3; Code of Federal Regulations, title 45, section 155.315; Minnesota Statutes, section 256L.05)."
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.LeftIndent = 0
	
	objSelection.font.bold = true
	objSelection.ParagraphFormat.SpaceAfter = 3
	objSelection.TypeText " You must give us more information"	
	objSelection.TypeParagraph()
	objSelection.font.bold = false
	objSelection.TypeText " We need more information from:"	
	objSelection.TypeParagraph()
	objSelection.font.bold = True
	objSelection.TypeText " " & client_first_name & " " & client_last_name & " - MNSure Case Number: " & client_case_number 
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.font.bold = true
	objSelection.ParagraphFormat.SpaceAfter = 0
	
	'REQUEST FOR INFORMATION START--------------------------------------------------------------
	
	objSelection.TypeText "                                Due Date            Needed Information            Supporting Documents"
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = false
	
	If rsn_income_type <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Income Type                          Send copies of all pay stubs or a written statement "
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              for earnings from the employer. Send copies of"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              checks, awards letters, court orders, or other"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              documents as proof of other types of income."
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_excluded_income <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Excluded Income                  doc type one"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Line two"
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_citizen_status_code <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Citizen Status Code              U.S. passport, Certificate of Naturalization, Certificate"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              of Citizenship, PASS card, Tribal enrollment or"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              membership card, certificate of degree of Indian"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              blood issued by a federally recognized Indian tribe."
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Birth certificate or other document from the U.S."
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Department of State, U.S. citizen ID card, American"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Indian card (I-872) from the U.S. Department of"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Homeland Security, Final U.S. adoption papers,"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Papers showing U.S. government employment before"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              June 1976, Official military record of service showing"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              United States place of birth, Hospital record showing"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              birth in the United States. Examples include hospital"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              chart pages with notes of the birth or a record of a"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              baby's hospital stay after birth. Insurance company"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              record showing United States as place or birth."
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Federal or state census record showing U.S."
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              citizenship place of birth. Medical records from a"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              clinic, doctor or hospital showing United States as the"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              place of birth. Records must be from within the last"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              five years. A statement signed by a doctor or midwife"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              who was at the birth. Statement must be from within"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              the last five years. Institutional admission papers"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              showing the United States as the place of birth,"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Papers must be from within the last five years."
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_ssnid <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            SSNID                                      U.S. passport, Certificate of Naturalization, Certificate"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                               of Citizenship, PASS card, Tribal enrollment or"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                               membership card, or Certificate of degree of Indian"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                               blood issued by a federally recognized Indian Tribe."
		objSelection.TypeParagraph()

		objSelection.TypeParagraph()	
	End If
	If rsn_incarceration <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Incarceration                         Please submit documentation of your release date"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              from incarceration or documentation of your"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              anticipated release date."
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_deductions <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Deductions                            Please submit documentation supporting the"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              deductions that were stated on the submitted"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              application. The most common form of this"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              verification is the first page of your 1040"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Federal Tax Form."
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_projected_annual_income <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Annual Income                     doc type one"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              Line two"
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	If rsn_applied_for_ssn <> "" then
		objSelection.TypeText "                                " & rfi_due_date & "            Application for SSN              Please submit documentation that a SSN has been"
		objSelection.TypeParagraph()
		objSelection.TypeText "                                                                                                              applied for."
		objSelection.TypeParagraph()
		
		objSelection.TypeParagraph()
	End If
	
	
	
	
	'REQUEST FOR INFORMATION END----------------------------------------------------------------
	
	objSelection.TypeParagraph()
	'Information needed ends here
	
	objSelection.ParagraphFormat.LeftIndent = 24
	objSelection.TypeText "    If the above information is not given to us by the date listed, your health care coverage may be denied. Send copies of any listed proofs to the above agency address. You may call us to explain the proofs needed in some cases. "
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.LeftIndent = 0
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "How do I use my health care coverage?"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "Medical Assistance"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "If you do not qualify for Medical Assistance, this information does not apply to you. Contact us to obtain your Medical ID. Give your Medical ID Number to your medical providers. If you have medical bills for services received since the date you qualified for coverage, contact the medical provider and ask them to bill the State of Minnesota. The provider may be able to pay you back for bills you paid."
	objSelection.TypeParagraph()
	objSelection.TypeText "You may be enrolled in a health plan. You will get information in the mail about choosing a health plan. Once enrolled, you will get information from the health plan telling you how to get services."

	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "MinnesotaCare"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "If you do not qualify for MinnesotaCare, this information does not apply to you."
	objSelection.TypeParagraph()
	objSelection.TypeText "Your coverage starts on January 1, 2014, unless you have a premium amount due. If you must make a payment for coverage to start, your coverage starts on the first day of the month after you make your first payment but no earlier than January 1, 2014. You will receive, if you have not already, your first premium in the mail. Send the payment to us as soon as you can."
	objSelection.TypeParagraph()
	objSelection.TypeText "You must enroll in a health plan. You will get information in the mail about choosing a health plan. Once enrolled, you will get information from the health plan telling you how to get services."

	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "Advanced Premium Tax Credit and Cost Sharing Reduction"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "If you do not qualify for Advanced Premium Tax Credit or Cost Sharing Reduction, this information does not apply to you. You must choose a health plan through MNsure and pay your share of the cost to start your health care coverage. Log into MNsure at mnsure.org. Choose a plan. You can choose a plan that costs less than your tax credit amount. MNsure tells you if you need to pay for your plan, amount you must pay and where to send your payment."
	objSelection.TypeParagraph()
	objSelection.TypeText "If you qualify for cost sharing reduction, MNsure tells you which Silver-level plans give you the cost sharing reduction you qualify for."
	objSelection.TypeParagraph()
	objSelection.TypeText "Your coverage starts on January 1, 2014, if you choose a health plan and pay your first premium. You will get a premium notice in the mail. Send the payment to us as soon as you can."

	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "Qualified Health Plan"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "If you are not eligible for Qualified Health Plan, this information does not apply to you."
	objSelection.TypeParagraph()
	objSelection.TypeText "You must choose a health plan through MNsure and pay your share of the cost to start your health care coverage. Log into MNsure at mnsure.org. Choose a plan. MNsure tells you if you need to pay for your plan, amount you must pay and where to send your payment."
	objSelection.TypeParagraph()
	objSelection.TypeText "Your coverage starts on January 1, 2014, if you choose a health plan and pay your first premium. You will get a premium notice in the mail. Send the payment to us as soon as you can."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "When should I tell you if I have a change?"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "Report changes within 10 days of the change event. Tell us about all changes including:"
	objSelection.TypeParagraph()
	objSelection.TypeText "	1.	Where you live."
	objSelection.TypeParagraph()
	objSelection.TypeText "	2.	Who lives with you."
	objSelection.TypeParagraph()
	objSelection.TypeText "	3.	Who you list as a dependant on your income taxes."
	objSelection.TypeParagraph()
	objSelection.TypeText "	4.	Income changes."
	objSelection.TypeParagraph()
	objSelection.TypeText "	5.	Starting or stopping other health insurance."
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeText "If you are not sure if you should report a change, call the number below and explain what is happening."
	objSelection.TypeParagraph()
	objSelection.TypeText "If you do not tell us you moved and returned mail has no forwarding address, coverage may end."
	objSelection.TypeParagraph()	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "What if I think you made a mistake?"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
		
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "If you think a mistake has been made, you can call 1-855-366-7873 and tell us what you think is wrong. You can also appeal the action. An appeal is a meeting where you can talk to a judge about why you think we made a mistake."
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "How to appeal a decision?"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
		
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "For more details, please see the enclosed Appeals Rights document titled ""IMPORTANT APPEAL RIGHTS! READ THIS NOW!"" If you are appealing a Medical Assistance or MinnesotaCare action or change, you may need to act within 10days; read the Appeals Rights document immediately. If you did not get the Appeals Rights document or have questions about your appeal rights, call 1-855-366-7873."
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "Questions?"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
		
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "Call the MNsure Contact Center, 1-855-366-7873, if you have questions about this notice."
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11
	objSelection.font.bold = true
	objSelection.TypeText "Go Paperless"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
		
	objSelection.font.size = 10
	objSelection.font.bold = false
	
	objSelection.TypeText "Paperless notices are a great way to stay organized, save postage, and help the environment. Instead of getting a paper notice in the mail, we will email you when you have notices to view in your online MNsure account. Signing up for paperless notices is easy, fast and secure! Log into your MNsure account and select the ""Go Paperless"" option."
	objSelection.TypeParagraph()
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.size = 11.5
	objSelection.font.bold = true
	objSelection.ParagraphFormat.Alignment = 1
	objSelection.TypeText "IMPORTANT APPEAL RIGHTS! READ THIS NOW!"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.Alignment = 0
	objSelection.font.size = 10
	objSelection.font.bold = true
	objSelection.font.underline = true
	
	objSelection.TypeText "What if I disagree with the action taken on my application?"
	objSelection.TypeParagraph()
	
	objSelection.font.bold = false
	objSelection.font.underline = false
	
	objSelection.TypeText "You will get a Health Care Notice letting you know if you qualify to get coverage through MNsure. If you do not think the decision is correct, you have the right to appeal. This is a legal process where an Appeals Examiner reviews a decision made by MNsure. You can learn more about how this works at www.mnsure.org."
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()

	objSelection.TypeText "        1. Internet               2. Phone                          3. Mail                                 4. In-Person"
	objSelection.TypeParagraph()
	objSelection.TypeText "        Log in to your          MNSure Contact            MNSure                              Minnesota Department of Human"
	objSelection.TypeParagraph()
	objSelection.TypeText "        account at                Center at 1-855-366-    81 Seventh Street East    Services Information Desk"
	objSelection.TypeParagraph()
	objSelection.TypeText "        www.mnsure.org    7873                                Suite 300                            444 Lafayette Road North"
	objSelection.TypeParagraph()
	objSelection.TypeText "                                                                                     St. Paul, MN 55101          St. Paul, MN 55101"
	objSelection.TypeParagraph()	
	objSelection.TypeParagraph()
	
	objSelection.font.bold = true
	objSelection.font.underline = true
	objSelection.TypeText "What can I appeal"
	objSelection.TypeParagraph()
	
	objSelection.font.bold = false
	objSelection.font.underline = false
	
	objSelection.TypeText "If MNsure did not act on your request about health care coverage or processed your request too slowly If you do not agree with the action taken"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "Important: You must file your appeal within 90 days of the date on your Health Care Notice. If your appeal involves Medical Assistance or MinnesotaCare, you must file your appeal within 30 days of receiving your Health Care Notice. If you show good cause for not appealing a Medical Assistance or MinnesotaCare action within 30 days, you may be able to appeal up to 90 days after receiving your Health Care Notice. See more information for Medical Assistance and MinnesotaCare appeal time limits below."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "Important: An appeal decision for one household member may affect the eligibility of other household members. Household eligibility may need to be re-determined."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.Font.bold = true
	objSelection.font.underline = true
	objSelection.TypeText "What if it's an emergency?"
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = false
	objSelection.Font.Underline = false
	
	objSelection.TypeText "You have a right to request an expedited (sped up) appeal. This happens when a person's life or health or ability to get, keep, or regain maximum function is in serious danger. If this applies to you, check ""yes"" when asked if the appeal involves a medical emergency. This is on the appeal request form. Or call the MNsure Contact Center at 1-855-366-7873."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.font.bold = true
	objSelection.font.underline = true
	
	objSelection.TypeText "What happens to my benefits during an appeal involving a redetermination of eligibility?"
	
	objSelection.TypeParagraph()
	objSelection.Font.bold = false
	objSelection.font.Underline = false
	
	objSelection.TypeText "Your benefits will automatically continue at the rate of prior coverage. But if you lose your appeal, you will have to pay back the benefits that you were not eligible to receive. You may want to ask to have your benefits reduced during your appeal so you do not have to pay them back if you lose. Check ""I want to reduce or stop my benefits..."" on the appeal request form, or call the MNsure Contact Center at 1-855-366-7873."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "For Medical Assistance or MinnesotaCare, your benefits continue only if you follow these time frames. You must appeal:"
	objSelection.TypeParagraph()
	objSelection.TypeText "Within 10 days of the date on the Health Care Notice; or"
	objSelection.TypeParagraph()
	objSelection.TypeText "Before the date when the action takes place."
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	'objSelection.TypeText.font.bold = true
	'objSelection.TypeText.font.italics = true
	
	objSelection.typetext "Important: If you do not appeal within 10 days of the date on the Health Care Notice, you can still appeal within 30 days. Your benefits will only go back to your prior coverage if you win the appeal."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = true
	objSelection.Font.Underline = true
	
	objSelection.TypeText "What if I lose my appeal?"
	
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = false
	objSelection.Font.Underline = false
	
	objSelection.TypeText "If you lose your appeal, you will have to pay back the benefits you got while your appeal was pending."	
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "Important: You have the right to apply for Medical Assistance or MinnesotaCare again if your benefits stop."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.Font.bold = true
	objSelection.Font.underline = true
	
	objSelection.TypeText "Can I get help with my appeal?"
	
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = false
	objSelection.Font.Underline = false
	
	objSelection.TypeText "You may represent yourself at the hearing. You may also have someone else speak for you. You must let us know in writing that the person is that you want to speak for you. You can do that on the appeal request form. If your income is below a certain limit, you may be able to get legal advice or help with an appeal from your local legal aid office."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.Font.Bold = true
	objSelection.Font.Underline = true
	
	objSelection.TypeText "Discrimination is against the law"
	
	objSelection.TypeParagraph ()
	
	objSelection.Font.Bold = false
	objSelection.Font.Underline = false
	
	objSelection.TypeText "The U.S. Department of Health and Human Services' Office for Civil Rights prohibits discrimination in its programs because of race, color, national origin, age, disability and sex, including sex stereotypes and gender identity. If you believe you have been discriminated against, you have the right to file a complaint directly with the federal agency. Write to the U.S. Department of Health and Human Services Office for Civil Rights Region V at 233 North Michigan Avenue, Suite 240, Chicago, IL 60601 or call at (312) 886-2359 (Voice) and (800) 368-1019 (Toll-Free) (800) 537-7697 (TTY)."
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.TypeText "In Minnesota, if you believe you have been discriminated against because of race, color, national origin, religion, creed, sex, sexual orientation, public assistance status, age, or disability, you have the right to file a complaint with:"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.LeftIndent = 50
	
	objSelection.TypeText "Minnesota Department of Human Services, Equal Opportunity and Access Division, P.O. Box 64997, St. Paul, MN 55164-0997. Telephone (651) 431-3040. Minnesota Relay 711 or (800) 627-3529. "
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeText "Minnesota Department of Human Rights, Freeman Building, 625 Robert Street North, St. Paul, MN 55155. Telephone (651) 539-1100 and Toll-Free (800) 657-3704. TTY (651) 296-1283. "
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeText "MNsure Accessibility and Equal opportunity Office, 81 7th Street East, Suite 300, St. Paul, MN 55101-2211, AEO@MNsure.org, Telephone (612) 279-8955. "
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.LeftIndent = 0
	
	objSelection.InsertBreak
			
	Set objShape = objSelection.InlineShapes.AddPicture("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\Pending Notices\multi_lang.jpg")
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.Alignment = 1
	objSelection.Font.Bold = true
	
	objSelection.TypeText "ADA ADVISORY"
	
	objSelection.TypeParagraph()
	objSelection.TypeParagraph()
	
	objSelection.ParagraphFormat.Alignment = 0
	objSelection.Font.bold = false
	objSelection.font.size = 9.5
	
	objSelection.TypeText "This information is available in accessible formats for individuals with disabilities by contacting MNsure at:  AEO@MNSure.org or (612) 279-8955. For other information on disability rights and protections to access"
	objSelection.TypeParagraph()
	objSelection.TypeText "MNsure programs, contact the agency's Accessibility & Equal Opportunity office."
	
		'Using Options After Doc is Made
	If save_copies = 1 then call save_word_document(client_last_name & ", " & client_first_name & " - " & client_case_number)
	If print_copies = 1 then call print_word_doc
	
		'Close Word
	call quit_word
	
		'Move to next row in Excel
	intRow = intRow + 1
Loop

call quit_excel

script_end_procedure("")