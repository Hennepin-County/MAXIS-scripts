BeginDialog income_screening, 0, 0, 311, 140, "Income Screening"
  ButtonGroup ButtonPressed
    OkButton 195, 120, 50, 15
    CancelButton 250, 120, 50, 15
  Text 15, 35, 55, 10, "Earned Income:"
  EditBox 70, 30, 60, 15, earned_income
  Text 145, 75, 25, 10, "RSDI:"
  EditBox 175, 70, 60, 15, rsdi_income
  Text 50, 75, 20, 10, "SSI:"
  EditBox 70, 70, 60, 15, ssi_income
  Text 25, 95, 45, 10, "Other UNEA:"
  EditBox 70, 90, 60, 15, otherunea_income
  Text 5, 10, 245, 10, "Enter information for household member, income expected THIS MONTH:"
  ButtonGroup ButtonPressed
    PushButton 70, 120, 50, 15, "Calculate", calc_button
  Text 145, 30, 55, 10, "Pay Frequency:"
  DropListBox 200, 30, 70, 15, "Select One..."+chr(9)+"Weekly"+chr(9)+"Bi-Weekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_frequency
  Text 45, 55, 25, 10, "Wage:"
  EditBox 70, 50, 60, 15, wage
  Text 145, 55, 25, 15, "Notes:"
  EditBox 175, 50, 120, 15, notes
  Text 145, 95, 50, 10, "Child Support:"
  EditBox 200, 90, 60, 15, child_supp_income
EndDialog
