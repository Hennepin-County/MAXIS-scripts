DO  'This DO-LOOP is to check that the CL's SSN created via random number generation is unique. If the SSN matches an SSN on file, the script creates a new SSN and re-enters the CL's information on MEMB
	DO	'This DO-LOOP makes sure that the first digit of the SSN is not a 9
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
      ssn_last = Rnd 
      ssn_last = 100000000 * ssn_last
      ssn_last = left(ssn_last, 4)

	'MEMB
	EMReadScreen cl_reference_number, 2, 4, 33
	IF cl_reference_number = "__" THEN EMWriteScreen reference_number, 4, 33
	EMWriteScreen last_name, 6, 30
	EMWriteScreen first_name, 6, 63
	EMWriteScreen mid_init, 6, 79
	EMWriteScreen ssn_first, 7, 42
	EMWriteScreen ssn_mid, 7, 46
	EMWriteScreen ssn_last, 7, 49
	EMWriteScreen "P", 7, 68

	dob_year = datepart("YYYY", date) - age
	dob_month = "01"
	dob_day = "01"

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
	DO	'===== This DO-LOOP checks that the race mini-box was opened
		EMReadScreen race_mini_box, 18, 5, 12
	LOOP UNTIL race_mini_box = "X AS MANY AS APPLY"
	EMWriteScreen "X", 15, 12
	transmit

	cl_ssn = ssn_first & "-" & ssn_mid & "-" & ssn_last
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
EMWriteScreen memi_citizen_yn, 10, 49
IF memi_citizen_yn = "Y" THEN EMWriteScreen "OT", 10, 78
EMWriteScreen "y", 13, 49		'<----- In MN > 12 mos?
EMWriteScreen "4", 13, 78
transmit


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
IF mailing_addr_line_one <> "" THEN
	EMWriteScreen mailing_addr_line_one, 13, 43
	EMWriteScreen mailing_addr_line_two, 14, 43
	EMWriteScreen mailing_addr_city, 15, 43
	EMWriteScreen mailing_addr_state, 16, 43
	MWriteScreen mailing_addr_zip, 16, 52
END IF
EMWriteScreen left(phone1, 3), 17, 45
EMWriteScreen right(left(phone1, 6), 3), 17, 51
EMWriteScreen right(phone1, 4), 17, 55
IF phone2 <> "" THEN
	EMWriteScreen left(phone2, 3), 18, 45
	EMWriteScreen right(left(phone2, 6), 3), 18, 51
	EMWriteScreen right(phone2, 4), 18, 55
END IF
IF phone3 <> "" THEN
	EMWriteScreen left(phone3, 3), 19, 45
	EMWriteScreen right(left(phone3, 6), 3), 19, 51
	EMWriteScreen right(phone3, 4), 19, 55
END IF

