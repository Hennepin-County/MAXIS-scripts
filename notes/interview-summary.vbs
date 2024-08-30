'This dialog will replace the initial dialog for the interview
BeginDialog Dialog1, 0, 0, 371, 315, "Interview Script Case number dialog"
  EditBox 75, 25, 60, 15, MAXIS_case_number
  DropListBox 75, 45, 145, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"MNbenefits"+chr(9)+"Combined AR for Certain Pops (DHS-3727)", CAF_form
  EditBox 75, 65, 145, 15, worker_signature
  DropListBox 10, 270, 350, 45, "Alert at the time you attempt to save each page of the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
  ButtonGroup ButtonPressed
    OkButton 260, 295, 50, 15
    CancelButton 315, 295, 50, 15
    PushButton 235, 20, 125, 15, "Open Interpreter Services Link", interpreter_servicves_btn
    PushButton 235, 35, 125, 15, "HSR Manual - Interview", hrs_manual_interview
    PushButton 235, 50, 125, 15, "SIR - SNAP Phone Interview Guide", sir_snap_interview
    PushButton 235, 65, 60, 15, "Script Overview", msg_what_script_does_btn
    PushButton 295, 65, 65, 15, "Script How to Use", msg_script_interaction_btn
    PushButton 10, 160, 120, 15, "Interview Summary", run_interview_summary_btn
    PushButton 240, 200, 120, 15, "More about 'SAVE YOUR WORK'", msg_save_your_work_btn
    PushButton 240, 235, 120, 15, "Details on Dialog Correction", msg_script_messaging_btn
    PushButton 10, 295, 50, 15, "Instructions", msg_show_instructions_btn
    PushButton 60, 295, 70, 15, "Quick Start Guide", msg_show_quick_start_guide_btn
    PushButton 130, 295, 30, 15, "FAQ", msg_show_faq_btn
  GroupBox 5, 10, 220, 75, "Case Information"
  Text 20, 30, 50, 10, "Case number:"
  Text 10, 50, 60, 10, "Actual CAF Form:"
  Text 10, 70, 60, 10, "Worker Signature:"
  GroupBox 230, 10, 135, 75, "Policy and Resources"
  GroupBox 5, 90, 360, 90, "Important Points"
  Text 10, 105, 240, 10, "* * * THIS  SCRIPT  SHOULD  BE  RUN  DURING  THE  INTERVIEW * * *"
  Text 25, 115, 315, 10, "Start this script at the beginning of the interview and use it to record the interview as it happens."
  Text 10, 130, 205, 10, "* Capture info from the form AND info from the conversation."
  Text 10, 150, 315, 10, "If the interview is already over, we have a temporary option to record the interview information:"
  GroupBox 5, 190, 360, 95, "Script Functionality"
  Text 10, 205, 185, 10, "This script SAVES the information you enter as it runs!"
  Text 10, 215, 345, 10, "IF the script errors, fails, is cancelled, the network goes down. YOU CAN GET YOUR WORK BACK!!!"
  Text 10, 240, 215, 10, "Dialog correction messages can be handled in two different ways."
  Text 10, 255, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
EndDialog




'Replacement Dialog for the Interview Question dialogs
BeginDialog Dialog1, 0, 0, 555, 385, "Summarized Interview Questions"
  GroupBox 5, 5, 475, 255, "Case Information"
  Text 10, 20, 190, 10, "Household Information (School, Disability, Absense):"
  EditBox 10, 30, 460, 15, household_info
  Text 10, 55, 195, 10, "Earned Income (Jobs, Self-Employment, start or stop work):"
  EditBox 10, 65, 460, 15, earned_income_info
  Text 10, 90, 255, 10, "Unearned Income (Social Security, Child Support, Unemployment, VA, etc):"
  EditBox 10, 100, 460, 15, unearned_income_info
  Text 10, 125, 245, 10, "Housing Expenses (Rent, Mortgage, Subsidy, Utilities):"
  EditBox 10, 135, 460, 15, housing_expanese_info
  Text 10, 160, 330, 10, "Other Expenses (Child Support Payments, Child/Adult Care Expenses, Medical Expesnse, etc.):"
  EditBox 10, 170, 460, 15, other_expenses_info
  Text 10, 195, 245, 10, "Assets (Accounts, Real Estate, Vehicles, Securities, etc):"
  EditBox 10, 205, 460, 15, assets_info
  Text 10, 230, 245, 10, "Other Notes (Unique Case Scenarios, Changes, Additional Details, etc.):"
  EditBox 10, 240, 460, 15, other_notes_info
  GroupBox 5, 265, 475, 95, "Interview Specifics"
  Text 10, 280, 245, 10, "Detail any Verbal Changes from the FORM:"
  EditBox 10, 290, 460, 15, form_changes_info
  Text 10, 310, 250, 10, "Check ONLY the boxes of information you discussed (not all are required):"
  CheckBox 15, 320, 110, 10, "Rights and Responsibilities", rights_responsibilities_checkbox
  CheckBox 15, 330, 110, 10, "Privacy Practices", privacy_practices_checkbox
  CheckBox 15, 340, 130, 10, "EBT Card Process and Information", ebt_card_checkbox
  CheckBox 140, 320, 110, 10, "Complaints and Civil Rights", complaints_civil_rights_checkbox
  CheckBox 140, 330, 110, 10, "Reporting Responsibilities", reporting_resp_checkbox
  CheckBox 140, 340, 110, 10, "Program Information", program_info_checkbox
  CheckBox 250, 320, 120, 10, "Child Support Referral and Coop", child_support_info_checkbox
  CheckBox 250, 330, 120, 10, "MFIP Minor Caregiver Information", mfip_minor_caregiver_checkbox
  CheckBox 250, 340, 120, 10, "Renewal Process and Information", renewal_info_checkbox
  CheckBox 380, 320, 110, 10, "IEVS Information", ievs_info_checkbox
  CheckBox 380, 330, 110, 10, "Appeal Rights", appeal_rights_checkbox
  ButtonGroup ButtonPressed
    PushButton 5, 365, 130, 15, "View Verifications", verif_button
    PushButton 415, 365, 50, 15, "NEXT", next_btn
  Text 485, 5, 75, 10, "---   DIALOGS   ---"
  Text 485, 15, 10, 10, "1"
  Text 485, 30, 10, 10, "2"
  Text 485, 45, 10, 10, "3"
  Text 485, 60, 10, 10, "4"
  Text 485, 75, 10, 10, "5"
  Text 485, 90, 10, 10, "6"
  ButtonGroup ButtonPressed
    PushButton 495, 15, 55, 10, "INTVW / CAF 1", caf_page_one_btn
    PushButton 495, 30, 55, 10, "CAF ADDR", caf_addr_btn
    PushButton 495, 45, 55, 10, "CAF MEMBs", caf_membs_btn
  Text 510, 60, 60, 10, "DETAILS"
  ButtonGroup ButtonPressed
    PushButton 495, 75, 55, 10, "CAF QUAL Q", caf_qual_q_btn
    PushButton 495, 90, 55, 10, "CAF Last Page", caf_last_page_btn
    PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
EndDialog
