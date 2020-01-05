'GATHERING STATS===========================================================================================
name_of_script = "NOTES - GA BASIS OF ELIGIBILITY.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 150
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("10/20/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
member_number = "01"
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 321, 125, "GA Basis of Eligibility"
  EditBox 55, 5, 50, 15, MAXIS_case_number
  EditBox 145, 5, 20, 15, member_number
  EditBox 255, 5, 55, 15, elig_date
  DropListBox 80, 30, 150, 12, "Select one..."+chr(9)+"Permanent illness"+chr(9)+"Temporary illness"+chr(9)+"Needed in the home"+chr(9)+"Placement in a Facility"+chr(9)+"Unemployable"+chr(9)+"Medically certified as having DISA"+chr(9)+"Appl/appeal pending for RSDI or SSI"+chr(9)+"Advanced age"+chr(9)+"FT Student elig displaced homemaker serv"+chr(9)+"Performing court-ordered services"+chr(9)+"Learning disabled"+chr(9)+"H.S. students age 18 and older (LES)"+chr(9)+"Drug/alcohol addiction",     basis_elig
  ButtonGroup ButtonPressed
	PushButton 240, 30, 75, 10, "Combined Manual", GA_CM_button
  EditBox 80, 50, 235, 15, verif_basis
  EditBox 80, 75, 235, 15, other_notes
  EditBox 80, 100, 130, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 215, 100, 50, 15
	CancelButton 265, 100, 50, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 10, 55, 70, 10, "Verification of basis:"
  Text 175, 10, 80, 10, "Date client meets basis:"
  Text 15, 105, 60, 10, "Worker signature:"
  Text 115, 10, 30, 10, "Memb #:"
  Text 5, 35, 75, 10, "GA basis of eligibility:"
  Text 35, 80, 40, 10, "Other notes:"
EndDialog
'the dialog
Do
	Do
	    err_msg = ""
		Do
  			Dialog ga_basis_dialog
			Cancel_confirmation
			If ButtonPressed = GA_CM_button then CreateObject("WScript.Shell").Run("http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_001315")
  		Loop until ButtonPressed <> GA_CM_button
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid member number."
		If isDate(elig_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid GA basis of eligibilty date."
		If basis_elig = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the member's GA basis of eligibility."
		If trim(verif_basis) = "" then err_msg = err_msg & vbNewLine & "* Enter the verification of the GA basis."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'specific text for case note based on the basis of eligibility
If basis_elig = "Permanent illness" then basis_text = "A person meets this basis if they have a medically certified permanent illness or incapacity which prevents them from getting and keeping suitable employment."
If basis_elig = "Temporary illness" then basis_text = "A person meets this basis if they have a medically certified temporary illness, injury, or incapacity which is expected to continue for more than 45 days and which prevents them from getting and keeping suitable employment."
If basis_elig = "Needed in the home" then basis_text = "A person meets this basis if they must be home to care for a household member on a continuous basis because of age or due to a medically certified illness, injury, or disability. Medical statements must say that people requiring care are unable to care for themselves, and they must verify that no other household member is able to provide the care."
If basis_elig = "Placement in a Facility" then basis_text = "A person meets this basis if they have been placed in, and is residing in, a licensed or certified facility for purposes of physical or mental health or rehabilitation, or in an approved chemical dependency domiciliary facility, meets a GA basis of eligibility and is eligible for a personal needs allowance if: " & _
"The placement is based on illness or incapacity, and is according to a plan developed or approved by the county agency through its director or designated representative, AND Personal needs are not already provided for in the facility per diem rates or funding package, AND The client is otherwise eligible for GA."
If basis_elig = "Unemployable" then basis_text = "A person meets this basis if they have been assessed by a vocational specialist and, in consultation with the county agency, have been determined to be unemployable."
If basis_elig = "Medically certified as having DISA" then basis_text = "A person meets this basis when diagnosed or certified by a qualified professional as being developmentally disabled or as having mental illness, and their condition prevents them from getting or keeping suitable employment."
If basis_elig = "Appl/appeal pending for RSDI or SSI" then basis_text = "A person meets this basis if they have an application pending for, or are appealing termination or denial of, Social Security Disability (through RSDI) or Supplemental Security Income (SSI) AND have a professionally certified permanent or temporary illness, injury, or incapacity which is expected to last for more than 30 days AND which prevents them from obtaining or keeping employment."
If basis_elig = "Advanced age" then basis_text = "A person meets this basis if they cannot get or keep suitable employment because they are age 55 or older and their work history shows a marked deterioration compared to their work history before age 55 as indicated by decreased occupational status, reduced hours of employment, or decreased periods of employment."
If basis_elig = "FT Student elig displaced homemaker serv" then basis_text = "A person meets this basis if they are eligible for displaced homemaker services and are full-time students."
If basis_elig = "Performing court-ordered services" then basis_text = "A person meets this basis if they are involved with protective or court-ordered services which prevent them from working at least 4 hours per day."
If basis_elig = "Learning disabled" then basis_text = "A person meets this basis if if they have a condition that qualifies under Minnesota's special education rules as a specific learning disability and are following a rehabilitation plan the county agency has developed or provided for them. Learning disabled under Minnesota's special education rules means a disorder in 1 or more of the psychological processes involved in perceiving, understanding, or using concepts through verbal language or non-verbal means."
If basis_elig = "H.S. students age 18 and older (LES)" then basis_text = "A person meets this basis if they age 18 and older whose primary language is not English and who are attending high school at least half-time, as recognized by the Minnesota Department of Education, are eligible for GA."
If basis_elig = "Drug/alcohol addiction" then basis_text = "A person meets this basis if drug or alcohol addiction have a basis of eligibility for GA when addiction is a material factor that prevents them from getting and keeping suitable employment. It must be verified through a physician's certification that the client's disability is the result of continued drug or alcohol addiction." & _
"GA benefits issued for clients with drug or alcohol addiction are subject to vendor payment for shelter and utility costs."

'----------------------------------------------------------------------------------------------------The case note
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~~GA basis of elig for M" & member_number & ": " & basis_elig & "~~")
Call write_bullet_and_variable_in_CASE_NOTE("Date client meets basis of elig", elig_date)
Call write_bullet_and_variable_in_CASE_NOTE("Verification of basis", verif_basis)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(basis_text)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(Worker_Signature)

script_end_procedure("Please review the GA basis of eligibilty text in the case note, the GA basis coding and ABAWD status on STAT/WREG, as well STAT/DISA coding prior to approval.")
