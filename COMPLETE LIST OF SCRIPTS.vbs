'THIS SCRIPT DOES NOT REQUIRE A STATS BLOCK SINCE IT'S PURELY INFORMATIONAL

''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
'IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
'	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
'		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		Else											'Everyone else should use the release branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		End if
'		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
'		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
'		req.send													'Sends request
'		IF req.Status = 200 THEN									'200 means great success
'			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'			Execute req.responseText								'Executes the script code
'		ELSE														'Error message
'			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
'                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'            StopScript
'		END IF
'	ELSE
'		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
'		text_from_the_other_script = fso_command.ReadAll
'		fso_command.Close
'		Execute text_from_the_other_script
'	END IF
'END IF
''END FUNCTIONS LIBRARY BLOCK================================================================================================



class script_bowie

    'Stuff the user indicates
	public script_name             	'The familiar name of the script (file name without file extension or category, and using familiar case)
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
	public SIR_instructions_button	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
	public workflows               	'The script workflows associated with this script (Changes Reported, Applications, etc)
    public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
    
    'Details the menus will figure out (does not need to be explicitly declared)
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run

    'Details the class itself figures out
	public property get script_URL
		If run_locally = true then
			script_repository = "C:\DHS-MAXIS-Scripts\Script Files\"
			script_URL = script_repository & ucase(category) & "\" & ucase(category & " - " & script_name) & ".vbs"
		Else
        	If script_repository = "" then script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/"    'Assumes we're scriptwriters
        	script_URL = script_repository & ucase(category) & "/" & replace(ucase(category & "%20-%20" & script_name) & ".vbs", " ", "%20")
		End if
    end property
    
    public property get SIR_instructions_URL 'The instructions URL in SIR
        SIR_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
    end property

end class

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order first by category, then by script name. Copy a block of code from above and paste your script info in. The function does the rest.




'ACTIONS SCRIPTS=====================================================================================================================================

script_num = 0
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "ABAWD Banked Months FIATer"																		'Script name
script_array(script_num).description 			= "FIATS SNAP eligibility, income, and deductions for HH members using banked months."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD FSET Exemption Check"																		'Script name
script_array(script_num).description 			= "Double checks a case to see if any possible ABAWD/FSET exemptions exist."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ABAWD Screening Tool"
script_array(script_num).description			= "A tool to walk through a screening to determine if client is ABAWD."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "BILS Updater"
script_array(script_num).description			= "Updates a BILS panel with reoccurring or actual BILS received."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Check EDRS"
script_array(script_num).description			= "Checks EDRS for HH members with disqualifications on a case."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Copy Panels to Word"
script_array(script_num).description			= "Copies MAXIS panels to Word en masse for a case for easier review."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "FSET SANCTION"
script_array(script_num).description			= "Updates the WREG panel, and case notes when imposing or resolving a FSET sanction."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "HG SUPPLEMENT"
script_array(script_num).description			= "NEW 04/2016!!! Issues a housing grant in MONY/CHCK for cases that should have been issued in prior months."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")


script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "LTC - Spousal Allocation FIATer"
script_array(script_num).description			= "FIATs a spousal allocation across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC - ICF-DD Deduction FIATer"																			'Script name
script_array(script_num).description 			= "FIATs earned income and deductions across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA-EPD EI FIAT"
script_array(script_num).description			= "FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "New Job Reported"
script_array(script_num).description			= "Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "PA Verif Request"
script_array(script_num).description			= "Creates a Word document with PA benefit totals for other agencies to determine client benefits."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Paystubs Received"
script_array(script_num).description			= "Enter in pay stubs, and puts it on JOBS (both retro & pro if applicable), as well as the PIC and HC pop-up, and case note."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Expense Verif Received"
script_array(script_num).description			= "Enter shelter expense/address information in a dialog and the script updates SHEL, HEST, and ADDR and case notes."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Send SVES"
script_array(script_num).description			= "Sends a SVES/QURY."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Transfer Case"
script_array(script_num).description			= "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "TYMA TIKLer"
script_array(script_num).description			= "TIKLS for TYMA report forms to be sent."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")











'BULK SCRIPTS=====================================================================================================================================

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Address Report"																		'Script name
script_array(script_num).description 			= "Creates a list of all addresses from a caseload(or entire county)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Banked Months Report"																		'Script name
script_array(script_num).description 			= "Creates a month specific report of banked months used, also checks these cases to confirm banked month use and creates a rejected report."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "CASE NOTE from List"																		'Script name
script_array(script_num).description 			= "Creates the same case note on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Case Transfer"																		'Script name
script_array(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "CEI Premium Noter"
script_array(script_num).description				= "Case notes recurring CEI premiums on multiple cases simultaneously."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Check SNAP for GA RCA"
script_array(script_num).description 			= "Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "COLA Auto approved Dail Noter"
script_array(script_num).description				= "Case notes all cases on DAIL/DAIL with Auto-approved COLA message, creates list of these messages, deletes the DAIL."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL Report"
script_array(script_num).description 			= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MAEPD MEDI CEI"
script_array(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find Panel Update Date"
script_array(script_num).description 			= "Creates a list of cases from a caseload(s) showing when selected panels have been updated."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Housing Grant Exemption Finder"
script_array(script_num).description 			= "Creates a list the rolling 12 months of housing grant issuances for MFIP recipients who've met an exemption."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IEVS Report"
script_array(script_num).description 			= "Pulls a list of cases in REPT/IEVC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "INAC Scrubber"
script_array(script_num).description				= "Checks cases on REPT/INAC (for criteria see SIR) case notes if passes criteria, and transfers if agency uses closed-file worker number. "
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC-GRH List Generator"
script_array(script_num).description 			= "Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload(s)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MAGI Non MAGI Report"
script_array(script_num).description 			= "NEW 06/2016!! Creates a list of cases and clients active on health care in MAXIS by MAGI/Non-MAGI."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "MEMO from List"
script_array(script_num).description				= "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Non-MAGI HC Info"
script_array(script_num).description 			= "Creates a list of cases with non-MAGI HC/PDED information."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ACTV List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/ACTV into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ARST List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/ARST into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-EOMC List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/EOMC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-INAC List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/INAC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MAMS List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MAMS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MFCM List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MFCM into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MONT List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MONT into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MRSR List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MRSR into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND1 List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/PND1 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND2 List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/PND2 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVS List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/REVS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVW List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/REVW into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Returned Mail"
script_array(script_num).description				= "Case notes that returned mail (without a forwarding address) was received for up to 60 cases, TIKLs for 10-day return."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "REVW-MONT Closures"
script_array(script_num).description				= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SWKR List Generator"
script_array(script_num).description 			= "Creates a list of SWKRs assigned to the various cases in a caseload(s)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Targeted SNAP Review Selection"
script_array(script_num).description				= "Creates a list of SNAP cases meeting review criteria and selects a random sample for review."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "TIKL from List"
script_array(script_num).description				= "Creates the same TIKL on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Update EOMC List"
script_array(script_num).description				= "Updates a saved REPT/EOMC excel file from previous month with current case status."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")











'DAIL SCRIPTS=====================================================================================================================================











'NAV SCRIPTS=====================================================================================================================================









'NOTES SCRIPTS=====================================================================================================================================


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Application Received"																		'Script name
script_array(script_num).description 			= "Template for documenting details about an application recevied."
script_array(script_num).category               = "NOTES"
script_array(script_num).subcategory            = "#-C"

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Approved programs"																		'Script name
script_array(script_num).description 			= "Template for when you approve a client's programs."
script_array(script_num).category               = "NOTES"
script_array(script_num).subcategory            = "#-C"











'NOTICES SCRIPTS=====================================================================================================================================















'UTILITIES SCRIPTS=====================================================================================================================================

