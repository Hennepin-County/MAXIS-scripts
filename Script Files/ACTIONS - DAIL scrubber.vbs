'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    HOW THE DAIL SCRUBER WORKS:
'
'    This script opens up other script files, using a custom function (run_DAIL_scrubber_script), followed by the path to the script file. It's done this
'      way because there could be hundreds of DAIL messages, and to work all of the combinations into one script would be incredibly tedious and long.
'
'    This script works by moving the message (where the cursor is located) to the top of the screen, and then reading the message text. Whatever the
'      message text says dictates which script loads up.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'CUSTOM FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS--------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'This is a custom function <<<<<evaluate including in main functions
function run_DAIL_scrubber_script(scrubber_script_path)
  Set run_another_DAIL_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_DAIL_command = run_another_DAIL_script_fso.OpenTextFile(scrubber_script_path)
  text_from_the_other_DAIL_script = fso_DAIL_command.ReadAll
  fso_DAIL_command.Close
  Execute text_from_the_other_DAIL_script
  stopscript
end function



'CONNECTS TO DEFAULT SCREEN
EMConnect ""

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t" + "<enter>"
EMWaitReady 0, 0

'THE FOLLOWING CODES ARE THE INDIVIDUAL MESSAGES. IT READS THE MESSAGE, THEN CALLS A NEW SCRIPT.----------------------------------------------------------------------------------------------------

'RSDI/BENDEX info received by agency.
EMReadScreen BENDEX_check, 47, 6, 30
If BENDEX_check = "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - BENDEX INFORMATION HAS BEEN STORED - CHECK INFC.vbs")

'CIT/ID has been verified through the SSA.
EMReadScreen CIT_check, 46, 6, 20
If CIT_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - citizenship verified.vbs")

'CS reports a new employer to the worker.
EMReadScreen CS_new_emp_check, 25, 6, 20
If CS_new_emp_check = "CS REPORTED: NEW EMPLOYER" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - CS reported new employer.vbs")

'Child support messages.
EMReadScreen CSES_check, 4, 6, 6
If CSES_check = "CSES" then
  EMReadScreen CSES_DISB_check, 4, 6, 20
  If CSES_DISB_check = "DISB" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - CSES processing.vbs")
End if

'Disability certification ends in 60 days.
EMReadScreen DISA_check, 58, 6, 20
If DISA_check = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - disa message.vbs")

'Client can receive an FMED deduction for SNAP.
EMReadScreen FMED_check, 59, 6, 20
If FMED_check = "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - FMED deduction.vbs")

'New HIRE messages, client started a new job.
EMReadScreen HIRE_check, 15, 6, 20
If HIRE_check = "NEW JOB DETAILS" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - new hire.vbs")

'Remedial care messages. May only happen at COLA.
EMReadScreen remedial_care_check, 34, 6, 20
If remedial_care_check = "PERSON HAS REMEDIAL CARE DEDUCTION" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - LTC remedial care.vbs")

'Student income is ending.
EMReadScreen SCHL_check, 58, 6, 20
If SCHL_check = "STUDENT INCOME HAS ENDED - REVIEW FS AND/OR HC RESULTS/APP" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - student income.vbs")

'SSI info received by agency.
EMReadScreen SDX_check, 44, 6, 30
If SDX_check = "SDX INFORMATION HAS BEEN STORED - CHECK INFC" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - SDX info has been stored.vbs")

'Random messages generated from an affiliated case.
EMReadScreen stat_check, 4, 6, 6
If stat_check = "FS  " or stat_check = "HC  " or stat_check = "GA  " or stat_check = "MSA " or stat_check = "STAT" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - affiliated case lookup.vbs")

'SSA info received by agency.
EMReadScreen SVES_check, 31, 6, 30
If SVES_check = "TPQY RESPONSE RECEIVED FROM SSA" then run_DAIL_scrubber_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\DAIL - TPQY response.vbs")

'NOW IF NO SCRIPT HAS BEEN WRITTEN FOR IT, THE DAIL SCRUBBER STOPS AND GENERATES A MESSAGE TO THE WORKER.----------------------------------------------------------------------------------------------------
script_end_procedure("You are not on a supported DAIL message. The script will now stop.")






