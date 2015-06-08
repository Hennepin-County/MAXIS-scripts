BeginDialog shelter_screening, 0, 0, 271, 135, "Shelter Screening"
  EditBox 190, 30, 65, 15, shelter_costs
  CheckBox 20, 75, 55, 15, "Heat (or AC)", heat_ac
  CheckBox 90, 75, 45, 15, "Electricity", electricity
  CheckBox 160, 75, 35, 15, "Phone", 
  ButtonGroup ButtonPressed
    OkButton 145, 110, 50, 15
    CancelButton 205, 110, 50, 15
  Text 5, 10, 170, 10, "Enter information on the household's shelter costs:"
  Text 10, 35, 180, 10, "How much does client pay in shelter costs this month?"
  GroupBox 5, 60, 205, 40, "What utilities does client pay this month?"
EndDialog
