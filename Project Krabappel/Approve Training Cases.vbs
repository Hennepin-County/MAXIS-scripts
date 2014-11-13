'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-Maxis-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

back_to_SELF
EMWriteScreen "________", 18, 43

FOR EACH case_to_be_approved IN case_load_array
  DO
    back_to_SELF
    case_number = left(case_to_be_approved, 8)
    EMWriteScreen "ELIG", 16, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen appl_date_month, 20, 43
    EMWriteScreen appl_date_year, 20, 46
    EMWriteScreen "FS", 21, 70
'========== This TRANSMIT sends the case to the FSPR screen ==========
    transmit
    EMReadScreen no_version, 10, 24, 2
  LOOP UNTIL no_version <> "NO VERSION"
  EMReadScreen is_case_approved, 10, 3, 3
  IF is_case_approved <> "UNAPPROVED" THEN
    back_to_SELF
  ELSE
'========== This TRANSMIT sends the case to the FSCR screen ==========
    transmit
'========== Reading for EXPEDITED STATUS ==========
    EMReadScreen is_case_expedited, 9, 4, 3
'========== This TRANSMIT sends the case to the FSB1 screen ==========
    transmit
'========== This TRANSMIT sends the case to the FSB2 screen ==========
    transmit
'========== This TRANSMIT sends the case to the FSSM screen ==========
    transmit
    IF is_case_expedited <> "EXPEDITED" THEN
      DO
        EMWriteScreen "APP", 19, 70
        transmit
        EMReadScreen not_allowed, 11, 24, 18
        EMReadScreen locked_by_background, 6, 24, 19
        EMReadScreen what_is_next, 5, 16, 44
      LOOP UNTIL not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED" OR what_is_next = "(Y/N)"
      DO
        EMReadScreen please_examine, 14, 4, 25
      LOOP UNTIL please_examine = "PLEASE EXAMINE"
      EMWriteScreen "Y", 16, 51
      transmit
      transmit
    ELSE
      DO
        EMWriteScreen "APP", 19, 70
        transmit
        EMReadScreen not_allowed, 11, 24, 18
        EMReadScreen locked_by_background, 6, 24, 19
        EMReadScreen what_is_next, 5, 16, 44
      LOOP UNTIL not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED" OR what_is_next = "(Y/N)"
      DO
        EMReadScreen rei_benefit, 3, 15, 33
      LOOP UNTIL rei_benefit = "REI"
      EMWriteScreen "Y", 15, 60
      transmit
      DO
        EMReadScreen rei_confirm, 3, 14, 30
      LOOP UNTIL rei_confirm = "REI"
      EMWriteScreen "Y", 14, 62
      transmit
      DO
        EMReadScreen continue_with_approval, 5, 16, 44
      LOOP UNTIL continue_with_approval = "(Y/N)"
      EMWriteScreen "Y", 16, 51
      transmit
      transmit
    END IF
  END IF
  back_to_SELF
NEXT