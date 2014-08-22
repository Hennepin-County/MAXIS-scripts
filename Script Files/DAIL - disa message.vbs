'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - disa message"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.


EMConnect ""

EMSendKey "s"
transmit

EMSendKey "disa"
transmit

'HH member dialog to select who's job this is.
BeginDialog HH_memb_dialog, 0, 0, 191, 52, "HH member"
  EditBox 50, 25, 25, 15, HH_memb
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 5, 10, 125, 15, "Which HH member is this for? (ex: 01)"
EndDialog
HH_memb = "01"
dialog HH_memb_dialog
If ButtonPressed = 0 then stopscript

EMWriteScreen HH_memb, 20, 76
transmit

EMReadScreen cash_disa_status, 1, 11, 69
If cash_disa_status <> "1" then
  MsgBox "This type of DISA status is not yet supported. It could be a SMRT or some other type of verif needed. Process manually at this time."
  stopscript
End if

PF4

PF9

EMSendKey "<home>" + "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" + "<newline>"
If cash_disa_status = 1 then EMSendKey "* Client needs a new Medical Opinion Form. Created using " & EDMS_choice & " and sent to client. TIKLed for 30-day return." & "<newline>"
EMSendKey "---" + "<newline>"

BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 25, 10, 75, 10, "Sign your case note."
EndDialog

dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMSendKey worker_sig
PF3
PF3
PF3

EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 30 days from now.
TIKL_month = datepart("m", dateadd("d", 30, date))
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_day = datepart("d", dateadd("d", 30, date))
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_year = datepart("yyyy", dateadd("d", 30, date))
TIKL_year = TIKL_year - 2000

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year
EMSetCursor 9, 3
EMSendKey "Medical Opinion Form sent 30 days ago. If not responded to, send another, and TIKL to close in 30 additional days."
transmit
PF3


MsgBox "Case note and TIKL made. Send a Medical Opinion Form using " & EDMS_choice & "."
script_end_procedure("")






