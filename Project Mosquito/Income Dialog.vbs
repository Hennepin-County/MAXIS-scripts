BeginDialog income_screening, 0, 0, 311, 140, "Income Screening"
  EditBox 70, 30, 60, 15, earned_income
  DropListBox 200, 30, 70, 15, "Select One..."+chr(9)+"Weekly"+chr(9)+"Bi-Weekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_frequency
  EditBox 70, 50, 60, 15, wage
  EditBox 175, 50, 120, 15, notes
  EditBox 70, 70, 60, 15, ssi_income
  EditBox 175, 70, 60, 15, rsdi_income
  EditBox 70, 90, 60, 15, otherunea_income
  EditBox 200, 90, 60, 15, child_supp_income
  ButtonGroup ButtonPressed
    PushButton 70, 120, 50, 15, "Calculate", calc_button
    OkButton 195, 120, 50, 15
    CancelButton 250, 120, 50, 15
  Text 5, 10, 245, 10, "Enter information for household member, income expected THIS MONTH:"
  Text 15, 35, 55, 10, "Earned Income:"
  Text 145, 30, 55, 10, "Pay Frequency:"
  Text 45, 55, 25, 10, "Wage:"
  Text 145, 55, 25, 15, "Notes:"
  Text 50, 75, 20, 10, "SSI:"
  Text 145, 75, 25, 10, "RSDI:"
  Text 25, 95, 45, 10, "Other UNEA:"
  Text 145, 95, 50, 10, "Child Support:"
EndDialog

