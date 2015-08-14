BeginDialog Asset_Screening, 0, 0, 226, 205, "Asset Screening "
  EditBox 50, 40, 50, 15, Accounts_Total
  EditBox 145, 40, 55, 15, Accounts_Excluded
  EditBox 50, 80, 50, 15, Cash_Total
  EditBox 145, 80, 55, 15, Cash_Excluded
  EditBox 50, 115, 50, 15, Cars_Total
  EditBox 145, 115, 55, 15, Cars_Excluded
  ButtonGroup ButtonPressed
    PushButton 0, 150, 50, 15, "Calculate", Calc_button
    OkButton 25, 180, 50, 15
    CancelButton 90, 180, 50, 15
  Text 5, 5, 195, 15, "Enter information (for present time) for household member: "
  Text 50, 20, 60, 15, "Total"
  Text 145, 20, 70, 15, "Excluded"
  Text 5, 40, 35, 15, "Accounts"
  Text 5, 80, 40, 10, "Cash"
  Text 5, 115, 40, 10, "Cars"
EndDialog
