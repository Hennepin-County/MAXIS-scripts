'This is Part Deux of the effort to create a training case generator. The desired structure of this script is as follows...
'----------(1) The script prompts the user to enter the x###### of the workers getting a fake case load. The script uses the number of MAXIS workers
'----------    to create an array for the number of case loads (FOR EACH maxis_worker IN worker_array ... create case load)
'----------(2) The script APPLs the desired cases and enters the STAT information based on the programs requested.
'----------(3) The script approves the results of every case created if user opts for that option.
'----------(4) The script transfers the cases to the MAXIS worker. 
'----------    worker_array is a combination of worker ID, case number, and Excel row... such that the first case created from Excel row 01 is saved
'----------    if it is going to worker x102132 - as x1021320022222201 (worker ID = x102132, case number = 222222, Excel row = 01)


'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-Maxis-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------DIALOGS----------
BeginDialog Dialog1, 0, 0, 221, 230, "Dialog"
  EditBox 15, 35, 45, 15, worker01
  EditBox 15, 55, 45, 15, worker02
  EditBox 15, 75, 45, 15, worker03
  EditBox 15, 95, 45, 15, worker04
  EditBox 15, 115, 45, 15, worker05
  EditBox 15, 135, 45, 15, worker06
  EditBox 65, 35, 45, 15, worker07
  EditBox 65, 55, 45, 15, worker08
  EditBox 65, 75, 45, 15, worker09
  EditBox 65, 95, 45, 15, Edit8
  EditBox 65, 115, 45, 15, Edit14
  EditBox 65, 135, 45, 15, Edit17
  EditBox 115, 35, 45, 15, Edit9
  EditBox 115, 55, 45, 15, Edit10
  EditBox 115, 75, 45, 15, Edit11
  EditBox 115, 95, 45, 15, Edit12
  EditBox 115, 115, 45, 15, Edit15
  EditBox 115, 135, 45, 15, Edit18
  EditBox 165, 35, 45, 15, Edit19
  EditBox 165, 55, 45, 15, Edit20
  EditBox 165, 75, 45, 15, Edit21
  EditBox 165, 95, 45, 15, Edit22
  EditBox 165, 115, 45, 15, Edit23
  EditBox 165, 135, 45, 15, Edit24
  CheckBox 10, 160, 200, 10, "Check here to create cases with SNAP", snap_cases_check
  CheckBox 10, 175, 210, 10, "Check here to create cases with CASH", cash_check
  CheckBox 10, 190, 200, 10, "Check here to have the script approve all cases.", approve_cases_check
  ButtonGroup ButtonPressed
    OkButton 60, 210, 50, 15
    CancelButton 115, 210, 50, 15
  Text 10, 10, 205, 15, "Please enter the MAXIS worker numbers of the workers getting training case loads..."
EndDialog

'----------THE SCRIPT----------

EMConnect ""

EMwritescreen "        ", 18, 43
current_month_plus_one = (left(date, 2)) + 1

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("Q:\Blue Zone Scripts\Script Files\sandbox\APPL Cases\Copy of Test Cases.xlsx")
objExcel.DisplayAlerts = True

excel_row = 4
'ObjExcel.Cells(excel_row, 2).Value = ""
DO
  x_row_array = ""          '========== This array is necessary to store the Excel row that the script is reading from when it APPLs the cases as the script will then have to re-read the specific
                            '========== lines when it goes back in to enter STAT information into the case.
  'APPL
  appl_date = ObjExcel.Cells(excel_row, 1).Value
    appl_date = replace(appl_date, "/", "")
    appl_date_month = left(appl_date, 2)
    appl_date_day = right(left(appl_date, 4), 2)
    appl_date_year = right(appl_date, 2)
  last_name = ObjExcel.Cells(excel_row, 3).Value
  first_name = ObjExcel.Cells(excel_row, 4).Value
  mid_init = ObjExcel.Cells(excel_row, 5).Value

  EMWriteScreen "APPL", 16, 43
  EMWriteScreen appl_date_month, 20, 43 
  EMWriteScreen appl_date_year, 20, 46
  transmit
  EMWriteScreen appl_date_month, 04, 63
  EMWriteScreen appl_date_day, 04, 66
  EMWriteScreen appl_date_year, 04, 69
  EMWriteScreen last_name, 07, 30
  EMWriteScreen first_name, 07, 63
  EMWriteScreen mid_init, 07, 79
  transmit

  DO
  'Reads the individual information on the case
  reference_number = ObjExcel.Cells(excel_row, 45).Value 
    IF len(reference_number) = 1 THEN reference_number = "0" & reference_number
  last_name = ObjExcel.Cells(excel_row, 3).Value
  first_name = ObjExcel.Cells(excel_row, 4).Value
  mid_init = ObjExcel.Cells(excel_row, 5).Value
  client_age = ObjExcel.Cells(excel_row, 6).Value
   IF client_age <> "" THEN
      dob_month = "01"
      dob_day = "01"
        date_year = right(date, 4)
        date_year = cint(date_year)
        client_age = cint(client_age)
      dob_year = date_year - client_age
   ELSE
      dob_month = "__"
      dob_day = "__"
      dob_year = "__"
   END IF
  DOB_verif = ObjExcel.Cells(excel_row, 7).Value
  gender = ObjExcel.Cells(excel_row, 8).Value
  ID_verif = ObjExcel.Cells(excel_row, 9).Value
  rel_to_appl = ObjExcel.Cells(excel_row, 10).Value
    IF len(rel_to_appl) = 1 THEN rel_to_appl = "0" & rel_to_appl
  spoken_lang = ObjExcel.Cells(excel_row, 11).Value
    IF len(spoken_lang) = 1 THEN spoken_lang = "0" & spoken_lang
  written_lang = ObjExcel.Cells(excel_row, 12).Value
    IF len(written_lang) = 1 THEN written_lang = "0" & written_lang
  interpreter_yn = ObjExcel.Cells(excel_row, 13).Value
  alias_yn = ObjExcel.Cells(excel_row, 14).Value
  alien_ID = ObjExcel.Cells(excel_row, 15).Value
  hisp_lat_yn = ObjExcel.Cells(excel_row, 16).Value
  race = ObjExcel.Cells(excel_row, 17).Value
    race = ucase(race)
  marital_status = ObjExcel.Cells(excel_row, 18).Value
  spouse = ObjExcel.Cells(excel_row, 19).Value
    IF len(spouse) = 1 THEN spouse = "0" & spouse
  last_grade_completed = ObjExcel.Cells(excel_row, 20).Value
    IF len(last_grade_completed) = 1 THEN last_grade_completed = "0" & last_grade_completed
    IF last_grade_completed = "" THEN last_grade_completed = "00"
  citizen = ObjExcel.Cells(excel_row, 21).Value
  cit_verif = ObjExcel.Cells(excel_row, 22).Value
  MN_12_mos = ObjExcel.Cells(excel_row, 23).Value
  res_verif = ObjExcel.Cells(excel_row, 24).Value
  addr_line_one = ObjExcel.Cells(excel_row, 25).Value
  addr_line_two = ObjExcel.Cells(excel_row, 26).Value
  city = ObjExcel.Cells(excel_row, 27).Value
  state = ObjExcel.Cells(excel_row, 28).Value
  zip = ObjExcel.Cells(excel_row, 29).Value
  county = ObjExcel.Cells(excel_row, 30).Value
    IF len(county) = 1 THEN county = "0" & county
  addr_verif = ObjExcel.Cells(excel_row, 31).Value
  homeless = ObjExcel.Cells(excel_row, 32).Value
  reservation = ObjExcel.Cells(excel_row, 33).Value
  mailing_addr_line_one = ObjExcel.Cells(excel_row, 34).Value
  mailing_addr_line_two = ObjExcel.Cells(excel_row, 35).Value
  mailing_addr_city = ObjExcel.Cells(excel_row, 36).Value
  mailing_addr_state = ObjExcel.Cells(excel_row, 37).Value
  mailing_addr_zip = ObjExcel.Cells(excel_row, 38).Value
  phone_one = ObjExcel.Cells(excel_row, 39).Value
  phone_one_type = ObjExcel.Cells(excel_row, 40).Value
  phone_two = ObjExcel.Cells(excel_row, 41).Value
  phone_two_type = ObjExcel.Cells(excel_row, 42).Value
  phone_three = ObjExcel.Cells(excel_row, 43).Value
  phone_three_type = ObjExcel.Cells(excel_row, 44).Value

  DO  'This DO-LOOP is to check that the CL's SSN created via random number generation is unique. If the SSN matches an SSN on file, the script creates a new SSN and re-enters the CL's information on MEMB
    DO
      Randomize
      ssn_first = Rnd
      ssn_first = 1000000000 * ssn_first
      ssn_first = left(ssn_first, 3)
    LOOP UNTIL left(ssn_first, 1) <> "9"
      Randomize
      ssn_mid = Rnd
      ssn_mid = 100000000 * ssn_mid
      ssn_mid = left(ssn_mid, 2)
      Randomize
      ssn_end = Rnd 
      ssn_end = 100000000 * ssn_end
      ssn_end = left(ssn_end, 4)

    'MEMB
    EMReadScreen cl_reference_number, 2, 4, 33
      IF cl_reference_number = "__" THEN EMWriteScreen reference_number, 4, 33
    EMWriteScreen last_name, 6, 30
    EMWriteScreen first_name, 6, 63
    EMWriteScreen mid_init, 6, 79
    EMWriteScreen ssn_first, 7, 42
    EMWriteScreen ssn_mid, 7, 46
    EMWriteScreen ssn_end, 7, 49
    EMWriteScreen "P", 7, 68
    EMWriteScreen dob_month, 8, 42
    EMWriteScreen dob_day, 8, 45
    EMWriteScreen dob_year, 8, 48
    EMWriteScreen DOB_verif, 8, 68
    EMWriteScreen gender, 9, 42
    EMWriteScreen ID_verif, 9, 68
    EMWriteScreen rel_to_appl, 10, 42
    EMWriteScreen spoken_lang, 12, 42
    EMWriteScreen written_lang, 13, 42
    EMWriteScreen interpreter_yn, 14, 68
    EMWriteScreen alias_yn, 15, 42
    EMWriteScreen alien_ID, 15, 68
    EMWriteScreen hisp_lat_yn, 16, 68
    EMWriteScreen "X", 17, 34
    transmit
    DO
      EMReadScreen race_mini_box, 18, 5, 12
      IF race_mini_box = "X AS MANY AS APPLY" THEN
        EMReadScreen cl_asian, 1, 7, 14
          IF race = cl_asian THEN EMWriteScreen "X", 7, 12
        EMReadScreen cl_black, 1, 8, 14
          IF race = cl_black THEN EMWriteScreen "X", 8, 12
        EMReadScreen cl_am_ind, 1, 10, 14
          IF race = cl_am_ind THEN EMWriteScreen "X", 10, 12
        EMReadScreen cl_pac_isl, 1, 12, 14
          IF race = cl_pac_isl THEN EMWriteScreen "X", 12, 12
        EMReadScreen cl_white, 1, 14, 14
          IF race = cl_white THEN EMWriteScreen "X", 14, 12
        EMReadScreen cl_unable, 1, 15, 14
          IF race = cl_unable THEN EMWriteScreen "X", 15, 12
        transmit
        transmit
      END IF
    LOOP UNTIL race_mini_box = "X AS MANY AS APPLY"
    cl_ssn = ssn_first & "-" & ssn_mid & "-" & ssn_end
    EMReadScreen ssn_match, 11, 8, 7
    IF cl_ssn <> ssn_match THEN
      PF8
      PF8
      PF5
    ELSE
      PF3
    END IF
  LOOP UNTIL cl_ssn <> ssn_match
  EMWaitReady 0, 0
  EMWriteScreen "Y", 6, 67
  transmit
  
  'MEMI
  EMWriteScreen marital_status, 7, 49
  EMWriteScreen spouse, 8, 49
  EMWriteScreen last_grade_completed, 9, 49
  EMWriteScreen citizen, 10, 49
  EMWriteScreen cit_verif, 10, 78
  EMWriteScreen mn_12_mos, 13, 49
  EMWriteScreen res_verif, 13, 78
  transmit
  
  x_row = cstr(excel_row)
  IF len(x_row) = 1 THEN x_row = "0" & x_row
  x_row_array = x_row_array & x_row & " "

  excel_row = excel_row + 1
 LOOP UNTIL addr_verif <> ""

 x_row_array = trim(x_row_array)
 x_row_array = split(x_row_array)

  'ADDR
   transmit
   EMWriteScreen addr_line_one, 6, 43
   EMWriteScreen addr_line_two, 7, 43
   EMWriteScreen city, 8, 43
   EMWriteScreen state, 8, 66
   EMWriteScreen zip, 9, 43
   EMWriteScreen county, 9, 66
   EMWriteScreen addr_verif, 9, 74
   EMWriteScreen homeless, 10, 43
   EMWriteScreen reservation, 10, 74
   EMWriteScreen mailing_addr_line_one, 13, 43
   EMWriteScreen mailing_addr_line_two, 14, 43
   EMWriteScreen mailing_addr_city, 15, 43
   EMWriteScreen mailing_addr_state, 16, 43
   EMWriteScreen mailing_addr_zip, 16, 52
 
   transmit
   EMWaitReady 0, 0
   EMReadScreen addr_warning, 7, 3, 6
   IF addr_warning = "Warning" THEN
     transmit
     EMWaitReady 0, 0
   END IF
   transmit
   EMWaitReady 0, 0
   PF3
   EMWaitReady 0, 0

'========== This section creates the information for an array that combines the case number with the excel row that can be used when entering information into STAT panels. ==========
   EMReadScreen case_number, 8, 18, 43
   case_number = replace(case_number, "_", "")

   DO
     IF len(case_number) <> 8 THEN case_number = "0" & case_number          '========== This converts the case number to 8 digits to accommodate the script later when entering information into STAT panels. The script will read the left 8 digits of to-be-determined variable.
   LOOP UNTIL len(case_number) = 8

   FOR EACH hh_x_row IN x_row_array          '========== This array smashed the case number with the Excel row associated with the case. When the script goes back to populated STAT information into the panels, it will read the 8 digits on the left as the case number and the 2 digits on the right as the Excel row to read from.
     use_case_number = case_number & hh_x_row
     case_load_array = case_load_array & use_case_number & " "
   NEXT

   EMwritescreen "        ", 18, 43

LOOP UNTIL excel_row = 13

case_load_array = trim(case_load_array)
case_load_array = split(case_load_array)


Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
FOR EACH maxis_case in case_load_array
	objselection.typetext left(maxis_case, 8) & "     " & right(maxis_case, 2)
	objselection.TypeParagraph()
NEXT


'========== This is the part of the script that puts the STAT information into the cases. ==========
FOR EACH created_case IN case_load_array
  excel_row = right(created_case, 2)
  case_number = left(created_case, 8)

  back_to_SELF
  EMWriteScreen "________", 18, 43
  EMWriteScreen appl_date_month, 20, 43
  EMWriteScreen appl_date_year, 20, 46
  TRANSMIT
 
 'These are all the variables that will be dumped into the stat panels
 'Dates need to be converted into their component parts so they can be dumped into the panels. We have discovered that BZ doesn't support dumping "/" from the dates into the panels so the month, day, and year have
 'to be separated before being dumped into the STAT panels.
  reference_number = objExcel.Cells(excel_row, 45).Value
    reference_number = cstr(reference_number)
    IF len(reference_number) = 1 THEN reference_number = "0" & reference_number
  appl_date = ObjExcel.Cells(excel_row, 1).Value
    appl_date = replace(appl_date, "/", "")
    appl_date_month = left(appl_date, 2)
    appl_date_day = right(left(appl_date, 4), 2)
    appl_date_year = right(appl_date, 2)
  client_age = ObjExcel.Cells(excel_row, 6).Value
 
'================================== TYPE ===========================
  type_cash_yn = objExcel.Cells(excel_row, 46).Value
  type_hc_yn = objExcel.Cells(excel_row, 47).Value
  type_emer_yn = objExcel.Cells(excel_row, 49).Value
  type_fs_yn = objExcel.Cells(excel_row, 48).Value
  type_grh_yn = objExcel.Cells(excel_row, 50).Value
  type_ive_yn = objExcel.Cells(excel_row, 51).Value
  
'================================== PROG ===========================
      prog_cash_one_appl_date = objExcel.Cells(excel_row, 52).Value
      prog_cash_one_appl_date = replace(prog_cash_one_appl_date, "/", "")
      prog_cash_one_appl_month = left(prog_cash_one_appl_date, 2)
      prog_cash_one_appl_day = right(left(prog_cash_one_appl_date, 4), 2)
      prog_cash_one_appl_year = right(prog_cash_one_appl_date, 2)
      prog_cash_one_elig_date = objExcel.Cells(excel_row, 53).Value
      prog_cash_one_elig_date = replace(prog_cash_one_elig_date, "/", "")
      prog_cash_one_elig_month = left(prog_cash_one_elig_date, 2)
      prog_cash_one_elig_day = right(left(prog_cash_one_elig_date, 4), 2)
      prog_cash_one_elig_year = right(prog_cash_one_elig_date, 2)
      prog_cash_one_interview_date = objExcel.Cells(excel_row, 54).Value
      prog_cash_one_interview_date = replace(prog_cash_one_interview_date, "/", "")
      prog_cash_one_interview_month = left(prog_cash_one_interview_date, 2)
      prog_cash_one_interview_day = right(left(prog_cash_one_interview_date, 4), 2)
      prog_cash_one_interview_year = right(prog_cash_one_interview_date, 2)
      cash_one_prog = objExcel.Cells(excel_row, 55).Value
      emer_appl_date = objExcel.Cells(excel_row, 56).Value
        emer_appl_date = replace(emer_appl_date, "/", "")
        emer_appl_month = left(emer_appl_date, 2)
        emer_appl_day = right(left(emer_appl_date, 4), 2)
        emer_appl_year = right(emer_appl_date, 2)
      emer_interview_date = objExcel.Cells(excel_row, 57).Value
        emer_interview_date = replace(emer_interview_date, "/", "")
        emer_interview_month = left(emer_interview_date, 2)
        emer_interview_day = right(left(emer_interview_date, 4), 2)
        emer_interview_year = right(emer_interview_date, 2)
      emer_prog = objExcel.Cells(excel_row, 58).Value
      fs_appl_date = objExcel.Cells(excel_row, 59).Value
        fs_appl_date = replace(fs_appl_date, "/", "")
        fs_appl_month = left(fs_appl_date, 2)
        fs_appl_day = right(left(fs_appl_date, 4), 2)
        fs_appl_year = right(fs_appl_date, 2)
      fs_elig_date = objExcel.Cells(excel_row, 60).Value
        fs_elig_date = replace(fs_elig_date, "/", "")
        fs_elig_month = left(fs_elig_date, 2)
        fs_elig_day = right(left(fs_elig_date, 4), 2)
        fs_elig_year = right(fs_elig_date, 2)
      fs_interview_date = objExcel.Cells(excel_row, 61).Value
        fs_interview_date = replace(fs_interview_date, "/", "")
        fs_interview_month = left(fs_interview_date, 2)
        fs_interview_day = right(left(fs_interview_date, 4), 2)
        fs_interview_year = right(fs_interview_date, 2)
      hc_appl_date = objExcel.Cells(excel_row, 62).Value
        hc_appl_date = replace(hc_appl_date, "/", "")
        hc_appl_month = left(hc_appl_date, 2)
        hc_appl_day = right(left(hc_appl_date, 4), 2)
        hc_appl_year = right(hc_appl_date, 2)
      migrant_worker = objExcel.Cells(excel_row, 63).Value 

'================================== DISA ===========================
      disa_begin_date = objExcel.Cells(excel_row, 70).Value
        disa_begin_date = replace(disa_begin_date, "/", "")
        disa_begin_month = left(disa_begin_date, 2)
        disa_begin_day = right(left(disa_begin_date, 4), 2)
        disa_begin_year = right(disa_begin_date, 4)
  
      disa_end_date = objExcel.Cells(excel_row, 71).Value
        disa_end_date = replace(disa_end_date, "/", "")
        disa_end_month = left(disa_end_date, 2)
        disa_end_day = right(left(disa_end_date, 4), 2)
        disa_end_year = right(disa_end_date, 2)

      disa_cert_begin_date = objExcel.Cells(excel_row, 72).Value
        disa_cert_begin_date = replace(disa_cert_begin_date, "/", "")
        disa_cert_begin_month = left(disa_cert_begin_date, 2)
        disa_cert_begin_day = right(left(disa_cert_begin_date, 4), 2)
        disa_cert_begin_year = right(disa_cert_begin_date, 2)

      disa_cert_end_date = objExcel.Cells(excel_row, 73).Value
        disa_cert_end_date = replace(disa_cert_end_date, "/", "")
        disa_cert_end_month = left(disa_cert_end_date, 2)
        disa_cert_end_day = right(left(disa_cert_end_date, 4), 2)
        disa_cert_end_year = right(disa_cert_end_date, 2)

      cash_disa_status = objExcel.Cells(excel_row, 74).Value
      cash_disa_verif = objExcel.Cells(excel_row, 75).Value
      fs_disa_status = objExcel.Cells(excel_row, 76).Value
      fs_disa_verif = objExcel.Cells(excel_row, 77).Value
      hc_disa_status = objExcel.Cells(excel_row, 78).Value
      hc_disa_verif = objExcel.Cells(excel_row, 79).Value

'================================ JOBS (first panel) ==============================
      jobs1_income_type = objExcel.Cells(excel_row, 92).Value
      jobs1_income_verif = objExcel.Cells(excel_row, 93).Value
      jobs1_employer = objExcel.Cells(excel_row, 94).Value
      jobs1_inc_start_date = objExcel.Cells(excel_row, 95).Value
        jobs1_inc_start_date = replace(jobs1_inc_start_date, "/", "")
        jobs1_inc_start_month = left(jobs1_inc_start_date, 2)
        jobs1_inc_start_day = right(left(jobs1_inc_start_date, 4), 2)
        jobs1_inc_start_year = right(jobs1_inc_start_date, 2)
      jobs1_pay_freq = objExcel.Cells(excel_row, 96).Value
      jobs1_avg_pay_amt = objExcel.Cells(excel_row, 97).Value
      jobs1_work_hours = objExcel.Cells(excel_row, 98).Value
      jobs1_hourly_wage = objExcel.Cells(excel_row, 99).Value

'================================ WREG ==============================
      wreg_fs_yn = objExcel.Cells(excel_row, 108).Value
      fs_pwe = objExcel.Cells(excel_row, 109).Value
      fset_wreg_status = objExcel.Cells(excel_row, 110). Value
        IF len(fset_wreg_status) = 1 THEN fset_wreg_status = "0" & fset_wreg_status
      defer_fset = objExcel.Cells(excel_row, 111).Value
      fset_orientation_date = objExcel.Cells(excel_row, 112).Value
        fset_orientation_date = replace(fset_orientation_date, "/", "")
        fset_orientation_month = left(fset_orientation_date, 2)
        fset_orientation_day = right(left(fset_orientation_date, 4), 2)
        fset_orientation_year = right(fset_orientation_date, 2)
      abawd_status = objExcel.Cells(excel_row, 113).Value
        abawd_status = cstr(abawd_status)
        IF len(abawd_status) = 1 THEN abawd_status = "0" & abawd_status
      ga_elig_basis = objExcel.Cells(excel_row, 114).Value
        IF len(ga_elig_bais) = 1 THEN ga_elig_basis = "0" & ga_elig_basis

'================================ UNEA (first panel) ============================== 
      unea1_income_type = objExcel.Cells(excel_row, 80).Value
        IF len(unea1_income_type) = 1 THEN unea1_income_type = "0" & unea1_income_type
      unea1_income_verif = objExcel.Cells(excel_row, 81).Value
      unea1_claim_number = objExcel.Cells(excel_row, 82).Value
      unea1_pay_freq = objExcel.Cells(excel_row, 84).Value
      unea1_pay_amt = objExcel.Cells(excel_row, 85).Value

'================================ UNEA (first panel) ==============================
      unea2_income_type = objExcel.Cells(excel_row, 86).Value
        IF len(unea2_income_type) = 1 THEN unea2_income_type = "0" & unea2_income_type
      unea2_income_verif = objExcel.Cells(excel_row, 87).Value
      unea2_claim_number = objExcel.Cells(excel_row, 88).Value
      unea2_pay_freq = objExcel.Cells(excel_row, 90).Value
      unea2_pay_amt = objExcel.Cells(excel_row, 91).Value 

'================================ EATS ==============================
      eats_hh_eat_together = objExcel.Cells(excel_row, 115).Value
      eats_boarder = objExcel.Cells(excel_row, 116).Value
      is_cl_eats_group_01_01 = objExcel.Cells(excel_row, 117).Value
        is_cl_eats_group_01_01 = cstr(is_cl_eats_group_01_01)
        IF LEN(is_cl_eats_group_01_01) = 1 THEN is_cl_eats_group_01_01 = "0" & is_cl_eats_group_01_01
      is_cl_eats_group_01_02 = objExcel.Cells(excel_row, 118).Value
        is_cl_eats_group_01_02 = cstr(is_cl_eats_group_01_02)
        IF LEN(is_cl_eats_group_01_02) = 1 THEN is_cl_eats_group_01_02 = "0" & is_cl_eats_group_01_02
      is_cl_eats_group_01_03 = objExcel.Cells(excel_row, 119).Value
        is_cl_eats_group_01_03 = cstr(is_cl_eats_group_01_03)
        IF LEN(is_cl_eats_group_01_03) = 1 THEN is_cl_eats_group_01_03 = "0" & is_cl_eats_group_01_03
      is_cl_eats_group_02_01 = objExcel.Cells(excel_row, 120).Value
        is_cl_eats_group_02_01 = cstr(is_cl_eats_group_02_01)
        IF LEN(is_cl_eats_group_02_01) = 1 THEN is_cl_eats_group_02_01 = "0" & is_cl_eats_group_02_01
      is_cl_eats_group_02_02 = objExcel.Cells(excel_row, 121).Value
        is_cl_eats_group_02_02 = cstr(is_cl_eats_group_02_02)
        IF LEN(is_cl_eats_group_02_02) = 1 THEN is_cl_eats_group_02_02 = "0" & is_cl_eats_group_02_02
      is_cl_eats_group_02_03 = objExcel.Cells(excel_row, 122).Value
        is_cl_eats_group_02_03 = cstr(is_cl_eats_group_02_03)
        IF LEN(is_cl_eats_group_02_03) = 1 THEN is_cl_eats_group_02_03 = "0" & is_cl_eats_group_02_03
      is_cl_eats_group_03_01 = objExcel.Cells(excel_row, 123).Value
        is_cl_eats_group_03_01 = cstr(is_cl_eats_group_03_01)
        IF LEN(is_cl_eats_group_03_01) = 1 THEN is_cl_eats_group_03_01 = "0" & is_cl_eats_group_03_01
      is_cl_eats_group_03_02 = objExcel.Cells(excel_row, 124).Value
        is_cl_eats_group_03_02 = cstr(is_cl_eats_group_03_02)
        IF LEN(is_cl_eats_group_03_02) = 1 THEN is_cl_eats_group_03_02 = "0" & is_cl_eats_group_03_02
      is_cl_eats_group_03_03 = objExcel.Cells(excel_row, 125).Value
        is_cl_eats_group_03_03 = cstr(is_cl_eats_group_03_03)
        IF LEN(is_cl_eats_group_03_03) = 1 THEN is_cl_eats_group_03_03 = "0" & is_cl_eats_group_03_03

'================================ REVW ==============================
      IF reference_number = "01" THEN
        IF cint(fs_elig_month) < 7 THEN
          fs_csr_month = fs_elig_month + 6
          fs_csr_year = fs_elig_year
        ELSE
          fs_csr_month = fs_elig_month - 6
          fs_csr_year = fs_elig_year + 1
        END IF 

        fs_elig_review_month = fs_elig_month
        fs_elig_review_year = fs_elig_year + 1
      END IF

'========== Starts entering information into the STAT panels ==========
'================================ TYPE ==============================
    DO
      back_to_SELF
      EMWriteScreen "STAT", 16, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen appl_date_month, 20, 43
      EMWriteScreen appl_date_year, 20, 46
      EMWriteScreen "TYPE", 21, 70
      transmit
      EMReadScreen at_the_type_panel, 4, 2, 48
    LOOP UNTIL at_the_type_panel = "TYPE"

    IF reference_number = "01" THEN 
      EMReadScreen number_of_type_panels_for_01, 1, 2, 78
      IF number_of_type_panels_for_01 = "1" THEN
        PF9
      ELSE
        EMWriteScreen "NN", 20, 79
        transmit
      END IF
    ELSE
      PF9
    END IF
 
    IF reference_number = "01" THEN
      EMWriteScreen type_cash_yn, 6, 28
      EMWriteScreen type_hc_yn, 6, 37
      EMWriteScreen type_fs_yn, 6, 46
      EMWriteScreen type_emer_yn, 6, 55
      EMWriteScreen type_grh_yn, 6, 64
      EMWriteScreen type_ive_yn, 6, 73
      EMReadScreen type_second_row, 2, 7, 3
      IF type_second_row <> "  " THEN 
        EMWriteScreen "N", 7, 28
        EMWriteScreen "N", 7, 37
        EMWriteScreen "N", 7, 46
        EMWriteScreen "N", 7, 55
      END IF
      EMReadScreen type_third_row, 2, 8, 3
      IF type_third_row <> "  " THEN
        EMWriteScreen "N", 8, 28
        EMWriteScreen "N", 8, 37
        EMWriteScreen "N", 8, 46
        EMWriteScreen "N", 8, 55
      END IF
    ELSE
      type_row = 7
      DO
        EMReadScreen type_reference_number, 2, type_row, 3
        IF type_reference_number = reference_number THEN 
          EMWriteScreen type_cash_yn, type_row, 28
          EMWriteScreen type_hc_yn, type_row, 37
          EMWriteScreen type_fs_yn, type_row, 46
          EMWriteScreen type_emer_yn, type_row, 55
        ELSE
          type_row = type_row + 1
        END IF
      LOOP UNTIL type_reference_number = reference_number
    END IF

'================================ PROG ==============================
    EMWriteScreen "PROG", 20, 71
    transmit

    DO
      EMReadScreen at_prog_screen, 4, 2, 50
    LOOP UNTIL at_prog_screen = "PROG"

    IF reference_number = "01" THEN
      EMReadScreen number_of_prog_panels, 1, 2, 78
      IF number_of_prog_panels = "0" THEN
        EMWriteScreen "NN", 20, 79
        transmit
      ELSE
        PF9
      END IF
      IF prog_cash_one_appl_date <> "" THEN
        EMWriteScreen prog_cash_one_appl_month, 6, 33
        EMWriteScreen prog_cash_one_appl_day, 6, 36
        EMWriteScreen prog_cash_one_appl_year, 6, 39
        EMWriteScreen prog_cash_one_elig_month, 6, 44
        EMWriteScreen prog_cash_one_elig_day, 6, 47
        EMWriteScreen prog_cash_one_elig_year, 6, 50
        EMWriteScreen prog_cash_one_interview_month, 6, 55
        EMWriteScreen prog_cash_one_interview_day, 6, 58
        EMWriteScreen prog_cash_one_interview_year, 6, 61
        EMWriteScreen cash_one_prog, 6, 67
      END IF
      IF fs_appl_date <> "" THEN
        EMWriteScreen fs_appl_month, 10, 33
        EMWriteScreen fs_appl_day, 10, 36
        EMWriteScreen fs_appl_year, 10, 39
        EMWriteScreen fs_elig_month, 10, 44
        EMWriteScreen fs_elig_day, 10, 47
        EMWriteScreen fs_elig_year, 10, 50
        EMWriteScreen fs_interview_month, 10, 55
        EMWriteScreen fs_interview_day, 10, 58
        EMWriteScreen fs_interview_year, 10, 61
      END IF
      EMWriteScreen migrant_worker, 18, 67
      transmit
    END IF

'================================ WREG ==============================    
    IF wreg_fs_yn = "Y" OR wreg_fs_yn = "y" THEN
	DO
	      EMWriteScreen "WREG", 20, 71
      	EMWriteScreen reference_number, 20, 76
	      transmit
            EMReadScreen at_wreg_panel, 4, 2, 48
		EMReadScreen is_wreg_stuck_in_bgtx, 10, 24, 29
	LOOP UNTIL is_wreg_stuck_in_background <> "BACKGROUND" and at_wreg_panel = "WREG"
      EMReadScreen does_wreg_panel_exist, 14, 24, 13
	IF does_wreg_panel_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
      END IF
      EMWriteScreen fs_pwe, 6, 68
      EMWriteScreen fset_wreg_status, 8, 50
      IF defer_fset <> "" THEN EMWriteScreen defer_fset, 8, 80
      IF fset_orientation_date <> "" THEN
        EMWriteScreen fset_orientation_month, 9, 50
        EMWriteScreen fset_orientation_day, 9, 53
        EMWriteScreen fset_orientation_year, 9, 56
      END IF
      EMWriteScreen abawd_status, 13, 50
      EMWriteScreen ga_elig_basis, 15, 50
      transmit
    END IF

'================================ EATS ==============================
    IF reference_number = "01" THEN
      IF eats_hh_eat_together <> "" THEN
	  DO    
          EMWriteScreen "eats", 20, 71
          transmit
          EMReadScreen at_eats_screen, 4, 2, 47
	    EMReadScreen is_eats_stuck_in_bgtx, 10, 24, 29
        LOOP UNTIL at_eats_screen = "EATS" and is_eats_stuck_in_BGTX <> "BACKGROUND"
        EMReadScreen number_of_eats_panels, 1, 2, 78
        IF number_of_eats_panels = "0" THEN
          EMWriteScreen "NN", 20, 79
          transmit
        ELSE
          PF9
        END IF
        EMWriteScreen eats_hh_eat_together, 4, 72
        EMWriteScreen eats_boarder, 5, 72
        IF ucase(eats_hh_eat_together) = "N" THEN
          EMWriteScreen "01", 13, 28
          EMWriteScreen is_cl_eats_group_01_01, 13, 39
          EMWriteScreen is_cl_eats_group_01_02, 13, 43
          EMWriteScreen is_cl_eats_group_01_03, 13, 47
          EMWriteScreen "02", 14, 28
          EMWriteScreen is_cl_eats_group_02_01, 14, 39
          EMWriteScreen is_cl_eats_group_02_02, 14, 43
          EMWriteScreen is_cl_eats_group_02_03, 14, 47
          IF is_cl_eats_group_03_01 <> "" THEN
            EMWriteScreen "03", 15, 28
            EMWriteScreen is_cl_eats_group_03_01, 15, 39
            EMWriteScreen is_cl_eats_group_03_02, 15, 43
            EMWriteScreen is_cl_eats_group_03_03, 15, 47
          END IF
        END IF
      END IF
    END IF

'================================ REVW ==============================
    DO
      EMWriteScreen "REVW", 20, 71
      transmit
      EMReadScreen at_revw_screen, 4, 2, 46
      EMReadScreen is_revw_stuck_in_bgtx, 10, 24, 29
    LOOP UNTIL at_revw_screen = "REVW" AND is_revw_stuck_in_bgtx <> "BACKGROUND"

    IF reference_number = "01" THEN
      EMReadScreen number_of_revw_panels, 1, 2, 78
      IF number_of_revw_panels = "0" THEN
        EMWriteScreen "NN", 20, 79
        transmit
      ELSE
        PF9
      END IF
      IF fs_appl_date <> "" THEN
        EMWriteScreen "X", 5, 58
        transmit
        EMReadScreen food_support_reports, 20, 5, 30
        DO
          IF food_support_reports = "Food Support Reports" THEN
            EMWriteScreen fs_csr_month, 9, 26
            EMWriteScreen fs_csr_year, 9, 32
            EMWriteScreen fs_elig_review_month, 9, 64
            EMWriteScreen fs_elig_review_year, 9, 70
            transmit
            PF3
          END IF
        LOOP UNTIL food_support_reports = "Food Support Reports"
        EMWriteScreen "N", 15, 75
      END IF
    END IF

'=============== JOBS ===============
    IF jobs1_income_type <> "" THEN
      DO
        EMWriteScreen "jobs", 20, 71
        EMWriteScreen reference_number, 20, 76
        transmit
        EMReadScreen at_jobs_screen, 4, 2, 45
        EMReadScreen is_jobs_stuck_in_bgtx, 10, 24, 29
      LOOP UNTIL is_jobs_stuck_in_bgtx <> "BACKGROUND" AND at_jobs_screen = "JOBS"
      EMReadScreen number_of_jobs_panels, 1, 2, 78
      IF number_of_jobs_panels = "0" THEN
        EMWriteScreen "nn", 20, 79
        transmit
      ELSE
        PF9
        EMReadScreen existing_employer, 30, 7, 42
        existing_employer = replace(existing_employer, "_", "")
        IF ucase(jobs1_employer) <> existing_employer THEN 
          msgbox "ERROR: You already have JOBS information loaded somehow! The script is stopping..."
          stopscript
        END IF
      END IF
      EMWriteScreen jobs1_income_type, 5, 38
      EMWriteScreen jobs1_income_verif, 6, 38
      EMWriteScreen jobs1_employer, 7, 42
      EMWriteScreen jobs1_inc_start_month, 9, 35
      EMWriteScreen jobs1_inc_start_day, 9, 38
      EMWriteScreen jobs1_inc_start_year, 9, 41
      EMReadScreen jobs1_benefit_month, 2, 20, 55
      EMReadScreen jobs1_benefit_year, 2, 20, 58
      IF jobs1_pay_freq = "1" THEN
        EMWriteScreen jobs1_benefit_month, 12, 54
        EMWriteScreen "01", 12, 57
        EMWriteScreen jobs1_benefit_year, 12, 60
        EMWriteScreen jobs1_avg_pay_amt, 12, 67
        EMWriteScreen jobs1_work_hours, 18, 72
        EMWriteScreen jobs1_pay_freq, 18, 35
      ELSEIF jobs1_pay_freq = "2" THEN
        EMWriteScreen jobs1_benefit_month, 12, 54
        EMWriteScreen "01", 12, 57
        EMWriteScreen jobs1_benefit_year, 12, 60
        EMWriteScreen jobs1_avg_pay_amt, 12, 67
        EMWriteScreen jobs1_benefit_month, 13, 54
        EMWriteScreen "15", 13, 57
        EMWriteScreen jobs1_benefit_year, 13, 60
        EMWriteScreen jobs1_avg_pay_amt, 13, 67
        EMWriteScreen (2 * jobs1_work_hours), 18, 72
        EMWriteScreen jobs1_pay_freq, 18, 35
      ELSEIF jobs1_pay_freq = "3" THEN 
        EMWriteScreen jobs1_benefit_month, 12, 54
        EMWriteScreen "05", 12, 57
        EMWriteScreen jobs1_benefit_year, 12, 60
        EMWriteScreen jobs1_avg_pay_amt, 12, 67
        EMWriteScreen jobs1_benefit_month, 13, 54
        EMWriteScreen "19", 13, 57
        EMWriteScreen jobs1_benefit_year, 13, 60
        EMWriteScreen jobs1_avg_pay_amt, 13, 67
        EMWriteScreen (2 * jobs1_work_hours), 18, 72
        EMWriteScreen jobs1_pay_freq, 18, 35
      ELSEIF jobs1_pay_freq = "4" THEN
        EMWriteScreen jobs1_benefit_month, 12, 54
        EMWriteScreen "05", 12, 57
        EMWriteScreen jobs1_benefit_year, 12, 60
        EMWriteScreen jobs1_avg_pay_amt, 12, 67
        EMWriteScreen jobs1_benefit_month, 13, 54
        EMWriteScreen "12", 13, 57
        EMWriteScreen jobs1_benefit_year, 13, 60
        EMWriteScreen jobs1_avg_pay_amt, 13, 67
        EMWriteScreen jobs1_benefit_month, 14, 54
        EMWriteScreen "19", 14, 57
        EMWriteScreen jobs1_benefit_year, 14, 60
        EMWriteScreen jobs1_avg_pay_amt, 14, 67
        EMWriteScreen jobs1_benefit_month, 15, 54
        EMWriteScreen "26", 15, 57
        EMWriteScreen jobs1_benefit_year, 15, 60
        EMWriteScreen jobs1_avg_pay_amt, 15, 67
        EMWriteScreen (4 * jobs1_work_hours), 18, 72
        EMWriteScreen jobs1_pay_freq, 18, 35
      END IF
      EMWriteScreen "X", 19, 38      'This navigates to the PIC=============================
      transmit
      DO
        EMReadScreen on_the_pic, 4, 3, 22
        IF on_the_pic = "Food" THEN
          EMWriteScreen appl_date_month, 5, 34
          EMWriteScreen appl_date_day, 5, 37
          EMWriteScreen appl_date_year, 5, 40
        END IF
      LOOP UNTIL on_the_pic = "Food"
      EMWriteScreen jobs1_pay_freq, 5, 64
      EMWriteScreen jobs1_work_hours, 8, 64
      EMWriteScreen jobs1_hourly_wage, 	9, 66
      transmit
      DO
        EMReadScreen pic_warning_message, 7, 20, 6
        IF pic_warning_message = "WARNING" THEN 
          transmit
          transmit
          transmit
        END IF
      LOOP until pic_warning_message = "WARNING"
      transmit
    END IF

'=============== UNEA1 ===============
    IF unea1_income_type <> "" THEN
      DO
        EMWriteScreen "unea", 20, 71
        EMWriteScreen reference_number, 20, 76
        transmit
        EMReadScreen is_unea_stuck_in_bgtx, 10, 24, 29
      LOOP UNTIL is_unea_stuck_in_bgtx <> "BACKGROUND"
      EMReadScreen number_of_unea_panels, 1, 2, 73
      IF number_of_unea_panels = "0" THEN
        EMWriteScreen "NN", 20, 79
        transmit
      ELSE
       PF9
      END IF
      EMWriteScreen unea1_income_type, 5, 37
      EMWriteScreen unea1_income_verif, 5, 65
      EMWriteScreen unea1_claim_number, 6, 37
      EMWriteScreen "01", 7, 37
      EMWriteScreen "01", 7, 40
        unea1_inc_start_yr = right((right(date, 4) - 3), 2)
      EMWriteScreen unea1_inc_start_yr, 7, 43
      EMReadScreen unea1_benefit_month, 2, 20, 55
      EMReadScreen unea1_benefit_year, 2, 20, 58
      EMWriteScreen unea1_benefit_month, 13, 54
      EMWriteScreen "01", 13, 57
      EMWriteScreen unea1_benefit_year, 13, 60
      EMWriteScreen unea1_pay_amt, 13, 68
      EMWriteScreen "X", 10, 26     'UNEA PIC
      transmit
      DO 
        EMReadScreen reading_unea1_pic, 4, 3, 28
        IF reading_unea1_pic = "SNAP" THEN
          EMWriteScreen appl_date_month, 5, 34
          EMWriteScreen appl_date_day, 5, 37
          EMWriteScreen appl_date_year, 5, 40
        END IF
      LOOP UNTIL reading_unea1_pic = "SNAP"
      EMWriteScreen unea1_pay_freq, 5, 64
      EMWriteScreen unea1_pay_amt, 8, 66
      transmit
      DO
        EMReadScreen unea1_pic_warning, 7, 20, 6
        IF unea1_pic_warning = "WARNING" THEN
          transmit
          transmit
        END IF
      LOOP UNTIL unea1_pic_warning = "WARNING"
      transmit
    END IF

'======================== UNEA2 =================
    IF unea2_income_type <> "" THEN
      DO
        DO
          EMWriteScreen "unea", 20, 71
          EMWriteScreen reference_number, 20, 76
          EMWriteScreen "NN", 20, 79
          transmit
          EMReadScreen second_unea_panel, 1, 2, 73
        LOOP UNTIL second_unea_panel = "2"
          EMReadScreen is_unea2_stuck_in_bgtx, 10, 24, 29
      LOOP UNTIL is_unea2_stuck_in_bgtx <> "BACKGROUND"
      EMWriteScreen unea2_income_type, 5, 37
      EMWriteScreen unea2_income_verif, 5, 65
      EMWriteScreen unea2_claim_number, 6, 37
      EMWriteScreen "01", 7, 37
      EMWriteScreen "01", 7, 40
        unea2_inc_start_yr = right((right(date, 4) - 3), 2)
      EMWriteScreen unea2_inc_start_yr, 7, 43
      EMReadScreen unea2_benefit_month, 2, 20, 55
      EMReadScreen unea2_benefit_year, 2, 20, 58
      EMWriteScreen unea2_benefit_month, 13, 54
      EMWriteScreen "01", 13, 57
      EMWriteScreen unea2_benefit_year, 13, 60
      EMWriteScreen unea2_pay_amt, 13, 68
      EMWriteScreen "X", 10, 26     'UNEA PIC
      transmit
      DO 
        EMReadScreen reading_unea2_pic, 4, 3, 28
        IF reading_unea2_pic = "SNAP" THEN
          EMWriteScreen appl_date_month, 5, 34
          EMWriteScreen appl_date_day, 5, 37
          EMWriteScreen appl_date_year, 5, 40
        END IF
      LOOP UNTIL reading_unea1_pic = "SNAP"
      EMWriteScreen unea2_pay_freq, 5, 64
      EMWriteScreen unea2_pay_amt, 8, 66
      transmit
      DO
        EMReadScreen unea2_pic_warning, 7, 20, 6
        IF unea2_pic_warning = "WARNING" THEN
          transmit
          transmit
        END IF
      LOOP UNTIL unea2_pic_warning = "WARNING"
      transmit
    END IF

'========== The script navigates to the STAT/WRAP panel. The script needs to determine if there are JOBS or UNEA panels on the cases that need to be updated before the case is sent through background.
    PF3
    IF jobs1_income_type = "" and unea1_income_type = "" THEN transmit
    IF jobs1_income_type <> "" or unea1_income_type <> "" THEN
      DO
        EMReadScreen wrap_bene_month, 2, 20, 55  
        comparison_month = (cstr(left(date, 2) + 1)) 
        IF len(comparison_month) = 1 THEN comparison_month = "0" & comparison_month
        IF jobs1_income_type <> "" or unea1_income_type <> "" THEN 
          IF comparison_month <> wrap_bene_month THEN
            EMWriteScreen "Y", 16, 54
            transmit
            IF jobs1_income_type <> "" THEN
              EMWriteScreen "JOBS", 20, 71
              EMWriteScreen reference_number, 20, 76
              transmit
              PF9
              EMReadScreen jobs_bene_month, 2, 20, 55
              EMWriteScreen jobs_bene_month, 12, 54
              EMReadScreen jobs_pay_date_02, 2, 13, 54
                IF jobs_pay_date_02 <> "__" THEN
                  EMWriteScreen jobs_bene_month, 13, 54
                  EMReadScreen jobs_pay_date_03, 2, 14, 54
                  IF jobs_pay_date_03 <> "__" THEN
                    EMWriteScreen jobs_bene_month, 14, 54
                    EMReadScreen jobs_pay_date_04, 2, 15, 54
                    IF jobs_pay_date_04 <> "__" THEN
                      EMWriteScreen jobs_bene_month, 15, 54
                    END IF
                  END IF
                END IF
              transmit
            END IF
            IF unea1_income_type <> "" THEN
              EMWriteScreen "UNEA", 20, 71
              EMWriteScreen reference_number, 20, 76
              EMWriteScreen "01", 20, 79
              transmit
              PF9
              EMReadScreen unea1_bene_month, 2, 20, 55
              EMWriteScreen unea1_bene_month, 13, 54
              transmit
            END IF
            IF unea2_income_type <> "" THEN
              EMWriteScreen "UNEA", 20, 71
              EMWriteScreen reference_number, 20, 76
              EMWriteScreen "02", 20, 79
              transmit
              PF9
              EMReadScreen unea2_bene_month, 2, 20, 55
              EMWriteScreen unea2_bene_month, 13, 54
              transmit
            END IF     
          ELSEIF comparison_month = wrap_bene_month THEN
            transmit
          END IF
        END IF
        PF3
      LOOP UNTIL comparison_month = wrap_bene_month
    END IF

NEXT

'========== Approval of results ==========
'========== The script is going to open another script. For whatever reason, we are having ==
'========== difficulty approving all of the cases and it is suspected that the length of ====
'========== this script is the culprit, although we cannot explain why. =====================

run_another_script("Q:\Blue Zone Scripts\Script Files\sandbox\APPL cases\Approve Training Cases.vbs")


