BeginDialog MTAF_dialog, 0, 0, 526, 340, "MTAF dialog"
 EditBox 45, 5, 60, 15, MTAF_date
  EditBox 160, 5, 60, 15, MFIP_elig_date
  EditBox 275, 5, 60, 15, interview_date
  DropListBox 275, 30, 60, 15, "complete"+chr(9)+"incomplete", MTAF_status_dropdown
  EditBox 75, 45, 260, 15, ADDR_change
  EditBox 75, 65, 260, 15, HHcomp_change_checkbox
  EditBox 75, 85, 260, 15, asset_change
  EditBox 105, 105, 230, 15, earned_income_change
  EditBox 105, 125, 230, 15, unearned_income_change
  EditBox 105, 145, 230, 15, shelter_costs_change
  EditBox 175, 165, 160, 15, subsidized_housing
  DropListBox 175, 185, 160, 15, "Select one..."+chr(9)+"Not subsidized"+chr(9)+"Verification provided"+chr(9)+"Verification pending", sub_housing_droplist
  EditBox 110, 200, 225, 15, child_adult_care_costs
  EditBox 110, 220, 225, 15, relationship_proof
  EditBox 175, 240, 160, 15, referred_to_OMB_PBEN
  EditBox 125, 260, 210, 15, elig_results_fiated
  EditBox 75, 280, 260, 15, other_notes
  EditBox 75, 300, 260, 15, verifications_needed
  CheckBox 350, 45, 135, 10, "Rights and responsibilities explained.", RR_explained_checkbox
  CheckBox 350, 60, 55, 10, "MTAF signed.", mtaf_signed_checkbox
  CheckBox 350, 75, 150, 10, "MFIP/financial orientation completed.", mfip_financial_orientation_checkbox
  CheckBox 350, 90, 200, 10, "Client exempt from cooperation with ES.", ES_exemption_checkbox
  CheckBox 5, 325, 115, 10, "Open approved programs script", open_approved_programs_checkbox
  CheckBox 130, 325, 110, 10, "Open denied programs script", open_denied_programs_checkbox
  EditBox 340, 320, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 425, 320, 50, 15
    CancelButton 475, 320, 50, 15
  Text 5, 10, 40, 10, "MTAF date:"
  Text 110, 10, 50, 10, "MFIP elig date:"
  Text 225, 10, 55, 15, "Interview date:"
  Text 5, 30, 225, 15, "**Changes reported on MTAF**  (Complete boxes as applicable.)"
  Text 225, 30, 45, 15, "MTAF status:"
  Text 5, 50, 70, 15, "Change in address:"
  Text 5, 70, 70, 15, "Change in HH comp:"
  Text 5, 90, 70, 15, "Change in assets:"
  Text 5, 110, 90, 15, "*Change in earned income:"
  Text 5, 130, 95, 15, "Change in unearned income:"
  Text 5, 150, 95, 15, "Change in shelter costs:"
  Text 5, 170, 170, 15, "Is housing subsidized? If so, what is the amount?"
  Text 75, 185, 90, 10, "**Subsidized housing status:"
  Text 5, 200, 85, 15, "Child or adult care costs:"
  Text 5, 220, 95, 15, "Proof of relationship on file:"
  Text 5, 240, 160, 15, "Client has been referred to apply for OMB/PBEN:"
  Text 5, 265, 115, 15, "Eligibility results fiated? If so, why:"
  Text 5, 285, 45, 10, "Other notes:"
  Text 5, 305, 70, 10, "Verifications needed:"
  GroupBox 350, 105, 150, 100, ""
  Text 360, 115, 135, 35, "*STOP WORK - Verification only necessary to verify income in the month of application/eligibility. (CM 0010.18.01)"
  Text 360, 155, 135, 45, "**SUBSIDY - Verification of housing subsidy is a mandatory verification for MFIP. STAT must be appropriately updated to ensure accurate approval of housing grant. (CM 0010.18.01)"
  Text 275, 325, 60, 10, "Worker signature:"
EndDialog
