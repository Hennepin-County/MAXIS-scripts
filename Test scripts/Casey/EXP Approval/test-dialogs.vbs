'DIALOGS to save
'Mock Ups and planning

'Initial Questions
BeginDialog Dialog1, 0, 0, 555, 385, "Expedited Determination"
  Text 140, 10, 205, 10, "SNAP is PENDING and Expedited Determination is Needed"
  GroupBox 10, 30, 455, 50, "Application Information"
  Text 20, 50, 115, 10, "Date of Application: mm/dd/yyyy"
  Text 20, 65, 115, 10, "Date of Interview: mm/dd/yyyy"
  Text 170, 50, 85, 10, "Household Composition:"
  Text 255, 50, 160, 10, "2 Adults and 3 Children - potentially eligible"
  Text 255, 60, 160, 10, "1 Adult - NOT SNAP eligible"
  GroupBox 10, 80, 455, 95, "Income"
  Text 20, 100, 230, 10, "Has anyone received a paycheck from a job apready in APP MONTH?"
  DropListBox 255, 95, 45, 45, "", List6
  Text 20, 120, 235, 10, "Does anyone expect to receive more income from a job in APP MONTH?"
  DropListBox 255, 115, 45, 45, "", List7
  Text 20, 140, 220, 10, "Does anyone have income from self employment in APP MONTH?"
  DropListBox 255, 135, 45, 45, "", List8
  Text 20, 160, 220, 10, "Does anyone have income from any other source  in APP MONTH?"
  DropListBox 255, 155, 45, 45, "", List9
  GroupBox 10, 175, 455, 35, "Assets"
  Text 20, 195, 125, 10, "Does anyone have a bank account?"
  DropListBox 145, 190, 45, 45, "", List10
  Text 220, 195, 80, 10, "Do you have any cash?"
  DropListBox 300, 190, 45, 45, "", List11
  GroupBox 10, 210, 455, 35, "Housing Expense"
  Text 20, 230, 225, 10, "Is anyone in the household responsible to pay a housing expense?"
  DropListBox 245, 225, 45, 45, "", List12
  GroupBox 10, 245, 455, 100, "Utilities Expense"
  Text 20, 260, 270, 10, "Is the household responsible to paythe Heat Expense or Air Conditioner Expense?"
  DropListBox 295, 255, 45, 45, "", List13
  Text 20, 280, 180, 10, "Is the household responsible to pay electric expense?"
  DropListBox 205, 275, 45, 45, "", List9
  Text 40, 300, 140, 10, "If yes, does this include any heat source?"
  DropListBox 185, 295, 45, 45, "", List11
  Text 245, 300, 105, 10, "Is AC plugged into this electric?"
  DropListBox 350, 295, 45, 45, "", List10
  Text 20, 320, 145, 10, "Is anyone responsible to PAY for a phone?"
  DropListBox 165, 315, 45, 45, "", List14
  Text 25, 330, 230, 10, "(Free phone plans without a payment requirement cannot be counted.)"
  ButtonGroup ButtonPressed
    PushButton 485, 10, 65, 15, "Summary", ADDR_page_btn
    PushButton 485, 25, 65, 15, "HH Comp", Button17
    PushButton 485, 40, 65, 15, "Income", SHEL_page_btn
    PushButton 485, 55, 65, 15, "Assets", HEST_page_btn
    PushButton 485, 70, 65, 15, "Housing Cost", Button9
    PushButton 485, 85, 65, 15, "Utilities Cost", Button11
    PushButton 485, 100, 65, 15, "Other Issues", Button9
    OkButton 450, 365, 50, 15
    CancelButton 500, 365, 50, 15
EndDialog
'
dialog Dialog1


BeginDialog Dialog1, 0, 0, 555, 385, "Expedited Determination"
  Text 15, 10, 115, 10, "ELIG has been created for SNAP"
  Text 25, 25, 145, 10, "Case appears EXPEDITED"
  Text 25, 35, 130, 10, "Date of Application: mm/dd/yyyy"
  Text 25, 45, 135, 10, "Expedited Package includes: MM/YY"
  Text 200, 25, 135, 10, "Postponed Verifications: TRUE"
  GroupBox 15, 60, 480, 230, "ELIG for MM/YY"
  Text 365, 60, 115, 10, "Monthly SNAP Allotment: $XXXX"
  Text 25, 70, 125, 10, "SNAP Benefit Amount $ XXX"
  ButtonGroup ButtonPressed
    PushButton 150, 70, 80, 10, "SEE CALCULATION", Button14
  Text 50, 80, 75, 10, "Prorated from DD"
  Text 30, 95, 70, 10, "Recoupment: $XXX"
  Text 30, 105, 95, 10, "Previously Issued: $XXX"
  Text 30, 120, 75, 10, "Members Counted"
  ButtonGroup ButtonPressed
    PushButton 150, 120, 35, 10, "DETAILS", Button12
  Text 40, 130, 75, 10, "2 Adults, 3 Children"
  Text 40, 140, 125, 20, "MEMBERs: 01/02/03/04/05"
  Text 40, 165, 120, 10, "Maximum Gross Income: $ XXXX"
  Text 40, 175, 120, 10, "Maximum Net Income: $ XXXX"
  Text 25, 190, 105, 10, "CASE Details Upon Approval"
  Text 30, 205, 75, 10, "Eligibility Result . . . . . ."
  Text 110, 205, 45, 10, "ELIGIBLE"
  Text 30, 215, 75, 10, "Reporting Status . . . . . . . . . ."
  Text 110, 215, 45, 10, "SIX-MONTH"
  Text 30, 225, 75, 10, "Benefit . . . . . . . . . . . . ."
  Text 110, 225, 45, 10, "INCREASE"
  Text 285, 70, 50, 10, "Budget"
  Text 295, 80, 85, 10, "Counted Earned Income"
  Text 390, 80, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 80, 35, 10, "DETAILS", Button3
  Text 295, 90, 90, 10, "Counted Unearned Income"
  Text 390, 90, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 90, 35, 10, "DETAILS", Button20
  Text 295, 105, 80, 10, "Total GROSS Income"
  Text 390, 105, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 105, 35, 10, "DETAILS", Button21
  Text 295, 125, 85, 10, "Total Deductions"
  Text 390, 125, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 125, 35, 10, "DETAILS", Button22
  Text 295, 135, 65, 10, "Housing Expenses"
  Text 360, 135, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 135, 35, 10, "DETAILS", Button23
  Text 295, 145, 65, 10, "Utility Expenses"
  Text 360, 145, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 145, 35, 10, "DETAILS", Button24
  Text 295, 160, 90, 10, "Allowed Shelter Expense"
  Text 390, 160, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 160, 35, 10, "DETAILS", Button25
  Text 295, 175, 80, 10, "Total NET Income"
  Text 390, 175, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 175, 35, 10, "DETAILS", Button26
  Text 295, 190, 90, 10, "Monthly SNAP Allotment:"
  Text 390, 190, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 190, 35, 10, "DETAILS", Button27
  Text 295, 205, 70, 10, "Recoupment:"
  Text 390, 205, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 205, 35, 10, "DETAILS", Button28
  Text 295, 215, 95, 10, "Previously Issued:"
  Text 390, 215, 40, 10, "$ XXXX"
  ButtonGroup ButtonPressed
    PushButton 440, 215, 35, 10, "DETAILS", Button29
    PushButton 510, 10, 40, 10, "MM/YY", Button16
    PushButton 510, 20, 40, 10, "MM/YY", Button18
    PushButton 510, 30, 40, 10, "MM/YY", Button19
    OkButton 450, 365, 50, 15
    CancelButton 500, 365, 50, 15
  Text 25, 245, 90, 10, "EXPEDITED DETAIL"
  Text 35, 260, 150, 10, "Expedited Package - One Month"
  Text 35, 270, 240, 10, "Expedited Criteria - Resources Plus Income Less than Shelter Costs"
  Text 35, 280, 270, 10, "Verification Status - Postponed Verifications Pending"
  Text 20, 305, 60, 10, "Additional Notes:"
  EditBox 80, 300, 415, 15, Edit1
  Text 15, 325, 65, 10, "WCOM Information:"
  EditBox 80, 320, 415, 15, Edit2
EndDialog

dialog Dialog1
