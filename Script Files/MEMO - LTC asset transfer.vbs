'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - LTC asset transfer"
start_time = timer

''LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

BeginDialog LTC_asset_tranfer_dialog, 0, 0, 126, 82, "LTC asset tranfer dialog"
  EditBox 35, 0, 85, 15, client
  EditBox 35, 20, 85, 15, spouse
  EditBox 70, 40, 50, 15, renewal_footer_month_year
  ButtonGroup LTC_asset_tranfer_dialog_ButtonPressed
    OkButton 10, 60, 50, 15
    CancelButton 65, 60, 50, 15
  Text 5, 5, 30, 10, "Client:"
  Text 5, 25, 30, 10, "Spouse:"
  Text 5, 45, 65, 10, "ER date (MM/YY):"
EndDialog


Do
  Dialog LTC_asset_tranfer_dialog
  If LTC_asset_tranfer_dialog_ButtonPressed = 0 then stopscript
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen WCOM_input_check, 27, 2, 28
  If WCOM_input_check <> "Worker Comment Input Screen" and WCOM_input_check <> "  Client Memo Input Screen " then MsgBox "You need to be on a notice in SPEC/WCOM or SPEC/MEMO for this to work. Please try again."
Loop until WCOM_input_check = "Worker Comment Input Screen" or WCOM_input_check = "  Client Memo Input Screen "

EMSendKey "<home>" + "The ownership of " + client + "'s assets must be transferred to " + spouse + " to avoid having them counted in future eligibility determinations. You are encouraged to do this as soon as possible. This transfer of assets must be done before " + client + "'s first annual renewal for " + renewal_footer_month_year + ". Verification of the transfer can be provided at any time. " + "<newline>" + "<newline>" 
EMSendKey "At the first annual renewal in " + renewal_footer_month_year + " the value of all assets that list " + client + " as an owner or co-owner will be applied towards the Medical Assistance Asset limit of $3,000.00.  If the total value of all countable assets for " + client + " is more than $3,000.00, Medical Assistance may be closed for " + renewal_footer_month_year + "."

script_end_procedure("")






