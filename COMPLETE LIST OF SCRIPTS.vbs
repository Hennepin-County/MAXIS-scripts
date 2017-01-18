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
	public release_date				'This allows the user to indicate when the script goes live (controls NEW!!! messaging)

    'Details the menus will figure out (does not need to be explicitly declared)
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run

    'Details the class itself figures out
	public property get script_URL
		If run_locally = true then
			script_repository = "C:\DHS-MAXIS-Scripts\"
			script_URL = script_repository & lcase(category) & "\" & lcase(replace(script_name, " ", "-") & ".vbs")
		Else
        	If script_repository = "" then script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/"    'Assumes we're scriptwriters
        	script_URL = script_repository & lcase(category) & "/" & replace(lcase(script_name) & ".vbs", " ", "-")
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
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "ABAWD FIATer"																		'Script name
script_array(script_num).description 			= "FIATS SNAP eligibility, income, and deductions for HH members with more than 3 counted months on the ABAWD tracking record."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #01/17/2017#


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD FSET Exemption Check"																		'Script name
script_array(script_num).description 			= "Double checks a case to see if any possible ABAWD/FSET exemptions exist."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ABAWD Minor Child Exemption FIATer"
script_array(script_num).description			= "FIATs SNAP eligibility, income and deductions for non-parents with minor children in HH."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #12/30/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ABAWD Screening Tool"
script_array(script_num).description			= "A tool to walk through a screening to determine if client is ABAWD."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "ACCT Panel Updater"
script_array(script_num).description			= "A tool which updates ACCT panels."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #07/25/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "BILS Updater"
script_array(script_num).description			= "Updates a BILS panel with reoccurring or actual BILS received."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Check EDRS"
script_array(script_num).description			= "Checks EDRS for HH members with disqualifications on a case."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "EMPS Updater"
script_array(script_num).description			= "Updates the EMPS panel, and case notes when for Child Under 12 Months Exemptions."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/03/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "FSET Sanction"
script_array(script_num).description			= "Updates the WREG panel, and case notes when imposing or resolving a FSET sanction."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("ABAWD")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "FSS Status Change"
script_array(script_num).description			= "Updates STAT with information from a Status Update."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #11/03/2016#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HG expansion MONY-CHCK"																		'Script name
script_array(script_num).description 			= "Issues a housing grant in MONY/CHCK for cases that meet HG expansion criteria."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #12/01/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "HG Supplement"
script_array(script_num).description			= "Issues a housing grant in MONY/CHCK for cases that should have been issued in prior months."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #04/25/2016#


script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "LTC Spousal Allocation FIATer"
script_array(script_num).description			= "FIATs a spousal allocation across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC ICF-DD Deduction FIATer"																			'Script name
script_array(script_num).description 			= "FIATs earned income and deductions across a budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #05/23/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA-EPD EI FIAT"
script_array(script_num).description			= "FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "New Job Reported"
script_array(script_num).description			= "Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "PA Verif Request"
script_array(script_num).description			= "Creates a Word document with PA benefit totals for other agencies to determine client benefits."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Paystubs Received"
script_array(script_num).description			= "Enter in pay stubs, and puts it on JOBS (both retro & pro if applicable), as well as the PIC and HC pop-up, and case note."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Shelter Expense Verif Received"
script_array(script_num).description			= "Enter shelter expense/address information in a dialog and the script updates SHEL, HEST, and ADDR and case notes."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Send SVES"
script_array(script_num).description			= "Sends a SVES/QURY."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Transfer Case"
script_array(script_num).description			= "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "TYMA TIKLer"
script_array(script_num).description			= "TIKLS for TYMA report forms to be sent."
script_array(script_num).category               = "ACTIONS"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #02/22/2016#











'BULK SCRIPTS=====================================================================================================================================

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Address Report"																		'Script name
script_array(script_num).description 			= "Creates a list of all addresses from a caseload(or entire county)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Banked Months Report"																		'Script name
script_array(script_num).description 			= "Creates a month specific report of banked months used, also checks these cases to confirm banked month use and creates a rejected report."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #04/25/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "CASE NOTE from List"																		'Script name
script_array(script_num).description 			= "Creates the same case note on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Case Transfer"																		'Script name
script_array(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "CEI Premium Noter"
script_array(script_num).description				= "Case notes recurring CEI premiums on multiple cases simultaneously."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Check SNAP for GA RCA"
script_array(script_num).description 			= "Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "COLA Auto approved Dail Noter"
script_array(script_num).description				= "Case notes all cases on DAIL/DAIL with Auto-approved COLA message, creates list of these messages, deletes the DAIL."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL Report"
script_array(script_num).description 			= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "EXP SNAP Review"
script_array(script_num).description 			= "Creates a list of PND1/PND2 cases that need to reviewed for EXP SNAP criteria."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #09/26/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MAEPD MEDI CEI"
script_array(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find Panel Update Date"
script_array(script_num).description 			= "Creates a list of cases from a caseload(s) showing when selected panels have been updated."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Housing Grant Exemption Finder"
script_array(script_num).description 			= "Creates a list the rolling 12 months of housing grant issuances for MFIP recipients who've met an exemption."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "INAC Scrubber"
script_array(script_num).description				= "Checks cases on REPT/INAC (for criteria see SIR) case notes if passes criteria, and transfers if agency uses closed-file worker number. "
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC-GRH List Generator"
script_array(script_num).description 			= "Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload(s)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MAGI Non MAGI Report"
script_array(script_num).description 			= "Creates a list of cases and clients active on health care in MAXIS by MAGI/Non-MAGI."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "MEMO from List"
script_array(script_num).description				= "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Non-MAGI HC Info"
script_array(script_num).description 			= "Creates a list of cases with non-MAGI HC/PDED information."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ACTV List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/ACTV into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ARST List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/ARST into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-EOMC List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/EOMC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-GRMR List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/GRMR into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-IEVC LIST"
script_array(script_num).description 			= "Pulls a list of cases in REPT/IEVC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-INAC List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/INAC into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MAMS List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MAMS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MFCM List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MFCM into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MONT List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MONT into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MRSR List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/MRSR into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND1 List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/PND1 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND2 List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/PND2 into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVS List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/REVS into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVW List"
script_array(script_num).description 			= "Pulls a list of cases in REPT/REVW into an Excel spreadsheet."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Returned Mail"
script_array(script_num).description				= "Case notes that returned mail (without a forwarding address) was received for up to 60 cases, TIKLs for 10-day return."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "REVS Scrubber"
script_array(script_num).description				= "Sends appointment letters to all interview-requiring REVS cases, and creates a spreadsheet of when each appointment is."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "REVW-MONT Closures"
script_array(script_num).description				= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Spenddown Report"
script_array(script_num).description				= "Creates a list of HC Cases from a caseload(s) with a Spenddown indicated on MOBL."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/26/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SWKR List Generator"
script_array(script_num).description 			= "Creates a list of SWKRs assigned to the various cases in a caseload(s)."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("REPORTS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Targeted SNAP Review Selection"
script_array(script_num).description				= "Creates a list of SNAP cases meeting review criteria and selects a random sample for review."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "TIKL from List"
script_array(script_num).description				= "Creates the same TIKL on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Update EOMC List"
script_array(script_num).description				= "Updates a saved REPT/EOMC excel file from previous month with current case status."
script_array(script_num).category               = "BULK"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#











'DAIL SCRIPTS=====================================================================================================================================

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ABAWD FSET Exemption Check"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Affiliated Case Lookup"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "BNDX Scrubber"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Citizenship Verified"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CS Reported New Employer"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CSES Processing"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CSES Scrubber"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #08/22/2016#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DISA Message"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ES Referral Missing"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/30/2016#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Financial Orientation Missing"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #09/30/2016#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "FMED Deduction"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LTC Remedial Care"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "New Hire NDNH"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "New Hire"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Postponed Expedited SNAP Verifications"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SDX Info Has Been Stored"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Send NOMI"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Student Income"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "TPQY Response"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "TYMA Scrubber"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#


script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Wage Match Scrubber"																		'Script name
script_array(script_num).description 			= ""
script_array(script_num).category               = "DAIL"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/04/2016#










'NAV SCRIPTS=====================================================================================================================================

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CASE-CURR"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "CASE-NOTE"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL-DAIL"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DAIL-WRIT"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-DWP"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-EMER"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-FS"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-GA"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-GRH"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-HC"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-MFIP"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ELIG-MSA"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MAXIS case in MMIS"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Find MMIS PMI in MAXIS"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "PERS"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI-TEMP"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ACTV - Bottom"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-ACTV"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-INAC"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MFCM"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-MONT"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND1"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-PND2"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-REVW"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "REPT-USER"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SELF"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SPEC-MEMO"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SPEC-WCOM"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "SPEC-XFER"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-ACCT"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-ADDR"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-AREP"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-JOBS"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-MEMB"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-MONT"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-PNLP"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-PROG"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-REVW"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "STAT-UNEA"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "View INFC"
script_array(script_num).description 			= ""
script_array(script_num).category               = "NAV"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#












'NOTES SCRIPTS=====================================================================================================================================

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Appeals"																		'Script name
script_array(script_num).description 			= "Template for documenting details about an appeal, and the appeal process."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #12/12/2016#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Application Received"																		'Script name
script_array(script_num).description 			= "Template for documenting details about an application recevied."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Approved programs"																		'Script name
script_array(script_num).description 			= "Template for when you approve a client's programs."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "AREP Form Received"
script_array(script_num).description				= "Template for when you receive an Authorized Representative (AREP) form."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Burial Assets"
script_array(script_num).description				= "Template for burial assets."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "CAF"
script_array(script_num).description				= "Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1						'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie	'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Case Discrepancy"
script_array(script_num).description			= "Template for case noting information about a case discrepancy."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/24/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Change Report Form Received"
script_array(script_num).description				= "Template for case noting information reported from a Change Report Form."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Change Reported"
script_array(script_num).description				= "Template for case noting HHLD Comp or Baby Born being reported. **More changes to be added in the future**"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Citizenship-Identity Verified"
script_array(script_num).description				= "Template for documenting citizenship/identity status for a case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Client Contact"
script_array(script_num).description				= "Template for documenting client contact, either from or to a client."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Client Transportation Costs"
script_array(script_num).description				= "Template for documenting client transportation costs."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Closed Programs"
script_array(script_num).description				= "Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "Combined AR"
script_array(script_num).description				= "Template for the Combined Annual Renewal.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "County Burial Application"
script_array(script_num).description				= "Template for the County Burial Application.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name				= "CSR"
script_array(script_num).description				= "Template for the Combined Six-month Report (CSR).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Deceased Client Summary"																		'Script name
script_array(script_num).description 			= "Adds details about a deceased client to a CASE/NOTE."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #04/25/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Denied Programs"																		'Script name
script_array(script_num).description 			= "Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Documents Received"
script_array(script_num).description 			= "Template for case noting information about documents received."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Drug Felon"
script_array(script_num).description 			= "Template for noting drug felon info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "DWP Budget"
script_array(script_num).description 			= "Template for noting DWP budgets."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")  '<<Temporarily removing first alpha split, will rebuild using function to auto-alpha-split, VKC 06/16/2016
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "EDRS DISQ Match Found"
script_array(script_num).description 			= "Template for noting the action steps when a SNAP recipient has an eDRS DISQ per TE02.08.127."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Emergency"
script_array(script_num).description 			= "Template for EA/EGA applications.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Employment Plan or Status Update"
script_array(script_num).description 			= "Template for case noting an employment plan or status update for family cash cases."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "EVF Received"
script_array(script_num).description 			= "Template for noting information about an employment verification received by the agency."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "ES Referral"
script_array(script_num).description 			= "Template for sending an MFIP or DWP referral to employment services."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Determination"
script_array(script_num).description 			= "Template for noting detail about how expedited was determined for a case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Expedited Screening"
script_array(script_num).description 			= "Template for screening a client for expedited status."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Explanation of Income Budgeted"
script_array(script_num).description 			= "Template for explaining the income budgeted for a case."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Foster Care HCAPP"
script_array(script_num).description 			= "Template for noting foster care HCAPP info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Foster Care Review"
script_array(script_num).description 			= "Template for noting foster care review info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Fraud Info"
script_array(script_num).description 			= "Template for noting fraud info."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Good Cause Claimed"
script_array(script_num).description				= "Template for requests of good cause to not receive child support."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Good Cause Results"
script_array(script_num).description				= "Template for Good Cause results for determination or renewal.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HC ICAMA"
script_array(script_num).description			= "Template for HC Interstate Compact on Adoption and Medical Assistance (HC ICAMA)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date			= #02/22/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HC Renewal"
script_array(script_num).description				= "Template for HC renewals.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HCAPP"
script_array(script_num).description				= "Template for HCAPPs.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "HRF"
script_array(script_num).description				= "Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "IEVS Match Received"
script_array(script_num).description				= "Template to case note when a IEVS notice is returned."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Incarceration"
script_array(script_num).description				= "Template to note details of an incarceration, and also updates STAT/FACI if necessary."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Interview Completed"
script_array(script_num).description				= "Template to case note an interview being completed but no stat panels updated."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Interview No Show"
script_array(script_num).description				= "Template for case noting a client's no-showing their in-office or phone appointment."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("E-L")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "Medical Opinion Form Received"
script_array(script_num).description				= "Template for case noting information about a Medical Opinion Form."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MFIP Sanction And DWP Disqualification"
script_array(script_num).description				= "Template for MFIP sanctions and DWP disqualifications, both CS and ES."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MFIP to SNAP Transition"
script_array(script_num).description				= "Template for noting when closing MFIP and opening SNAP."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MSQ"
script_array(script_num).description				= "Template for noting Medical Service Questionaires (MSQ)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MTAF"
script_array(script_num).description				= "Template for the MN Transition Application form (MTAF)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "OHP Received"
script_array(script_num).description				= "Template for noting Out of Home Placement (OHP)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Overpayment"
script_array(script_num).description				= "Template for noting basic information about overpayments."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Pregnancy Reported"
script_array(script_num).description				= "Template for case noting a pregnancy. This script can update STAT/PREG."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Proof of Relationship"
script_array(script_num).description				= "Template for documenting proof of relationship between a member 01 and someone else in the household."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "REIN Progs"
script_array(script_num).description				= "Template for noting program reinstatement information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Returned Mail Received"
script_array(script_num).description				= "Template for noting Returned Mail Received information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Significant Change"
script_array(script_num).description				= "Template for noting Significant Change information."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Verifications Needed"
script_array(script_num).description				= "Template for when verifications are needed (enters each verification clearly)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

'NOTES subcategories (placing them here to be sure buttons go in right place)-------------------------------------------------------------------------------------

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LEP - EMA"
script_array(script_num).description				= "Template for EMA applications."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LEP")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LEP - SAVE"
script_array(script_num).description				= "Template for the SAVE system for verifying immigration status."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LEP")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "LEP - Sponsor Income"
script_array(script_num).description				= "Template for the sponsor income deeming calculation (it will also help calculate it for you)."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LEP")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 				= "LTC - 1503"
script_array(script_num).description				= "Template for processing DHS-1503."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - 5181"
script_array(script_num).description				= "Template for processing DHS-5181."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - Application Received"
script_array(script_num).description				= "Template for initial details of a LTC application.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - Asset Assessment"
script_array(script_num).description				= "Template for the LTC asset assessment. Will enter both person and case notes if desired."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - COLA Summary"
script_array(script_num).description				= "Template to summarize actions for the changes due to COLA.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - Intake Approval"
script_array(script_num).description				= "Template for use when approving a LTC intake.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - MA Approval"
script_array(script_num).description				= "Template for approving LTC MA (can be used for changes, initial application, or recertification).*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - Renewal"
script_array(script_num).description				= "Template for LTC renewals.*"
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)				'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie			'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 				= "LTC - Transfer Penalty"
script_array(script_num).description				= "Template for noting a transfer penalty."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("LTC")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MNSure - Documents Requested"
script_array(script_num).description				= "Template for when MNsure documents have been requested."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "MNSure Retro HC Application"
script_array(script_num).description				= "Template for when MNsure retro HC has been requested."
script_array(script_num).category               = "NOTES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("M-Z")
script_array(script_num).release_date           = #10/01/2000#











'NOTICES SCRIPTS=====================================================================================================================================

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "12 Month Contact"																		'Script name
script_array(script_num).description 			= "Sends a MEMO to the client of their reporting responsibilities (required for SNAP 2-yr certifications, per POLI/TEMP TE02.08.165)."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Appointment Letter"																		'Script name
script_array(script_num).description 			= "Sends a MEMO containing the appointment letter (with text from POLI/TEMP TE02.05.15)."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Eligibility Notifier"																		'Script name
script_array(script_num).description 			= "Sends a MEMO informing client of possible program eligibility for SNAP, MA, MSP, MNsure or CASH."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "GRH OP CL LEFT FACI"
script_array(script_num).description			= "Sends a MEMO to a facility indicating that an overpayment is due because a client left."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "LTC - Asset Transfer"
script_array(script_num).description			= "Sends a MEMO to a LTC client regarding asset transfers. "
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA Inmate Application WCOM"
script_array(script_num).description			= "Sends a WCOM on a MA notice for Inmate Applications"
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MA-EPD No Initial Premium"
script_array(script_num).description			= "Sends a WCOM on a denial for no initial MA-EPD premium."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Method B WCOM"													'needs spaces to generate button width properly.
script_array(script_num).description			= "Makes detailed WCOM regarding spenddown vs. recipient amount for method B HC cases."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MFIP Orientation"
script_array(script_num).description			= "Sends a MEMO to a client regarding MFIP orientation."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "MNsure Memo"
script_array(script_num).description			= "Sends a MEMO to a client regarding MNsure."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "NOMI"
script_array(script_num).description			= "Sends the SNAP notice of missed interview (NOMI) letter, following rules set out in POLI/TEMP TE02.05.15."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Overdue Baby"
script_array(script_num).description			= "Sends a MEMO informing client that they need to report information regarding the status of pregnancy, within 10 days or their case may close."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "SNAP E and T Letter"
script_array(script_num).description			= "Sends a SPEC/LETR informing client that they have an Employment and Training appointment."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Verifications Still Needed"
script_array(script_num).description			= "Creates a Word document informing client of a list of verifications that are still required."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date			= #04/25/2016#



'-------------------------------------------------------------------------------------------------------------------------SNAP WCOMS LISTS
'Resetting the variable
script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script_bowie
script_array(script_num).script_name 			= "ABAWD with Child in HH WCOM"'needs spaces to generate button width properly.																'Script name
script_array(script_num).description 			= "Adds a WCOM to a notice for an ABAWD adult receiving child under 18 exemption."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Banked Month WCOMS"
script_array(script_num).description 			= "Adds various WCOMS to a notice for regarding banked month approvals/closure."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Client Death WCOM"
script_array(script_num).description 			= "Adds a WCOM to a notice regarding SNAP closure due to death of last HH member."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #01/18/2017#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Duplicate Assistance WCOM"
script_array(script_num).description 			= "Adds a WCOM to a notice for duplicate assistance explaining why the client was ineligible."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Postponed WREG Verifs"
script_array(script_num).description 			= "Sends a WCOM informing the client of postponed verifications that MAXIS won't add to notice correctly by itself."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #01/18/2017#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Returned Mail WCOM"
script_array(script_num).description 			= "Adds a WCOM to a notice for SNAP returned mail closure."
script_array(script_num).category               = "NOTICES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("SNAP WCOMS")
script_array(script_num).release_date           = #01/18/2017#











'UTILITIES SCRIPTS=====================================================================================================================================

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Banked Month Database Updater"
script_array(script_num).description 			= "Updates cases in the banked month database with actual MAXIS status."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Copy Case Data for Training"
script_array(script_num).description 			= "Copies data from a case to a spreadsheet to be run on the Training Case Generator."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/03/2016#


script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Copy Case Note Elsewhere"
script_array(script_num).description 			= "Copies a CASE/NOTE to either a claims note or a SPEC/MEMO."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #06/27/2016#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name			= "Copy Panels to Word"
script_array(script_num).description			= "Copies MAXIS panels to Word en masse for a case for easier review."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Info"
script_array(script_num).description 			= "Displays information about your BlueZone Scripts installation."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Move Production Screen to Inquiry"
script_array(script_num).description 			= "Moves a screen from MAXIS prouduction mode to MAXIS inquiry."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Phone Number or Name Look Up"
script_array(script_num).description 			= "Checks every case on REPT screens to find a case number when you have a phone number. *OR* Searches for a specific case on multiple REPT screens by last name."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "POLI TEMP List"
script_array(script_num).description 			= "Creates a list of current POLI/TEMP topics, TEMP reference and revised date."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "PRISM Screen Finder"
script_array(script_num).description 			= "Navigates to popular PRISM screens. The navigation window stays open until user closes it."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Training Case Creator"
script_array(script_num).description 			= "Creates training case scenarios en masse and XFERs them to workers."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)		'Resets the array to add one more element to it
Set script_array(script_num) = new script_bowie		'Set this array element to be a new script_bowie. Script details below...
script_array(script_num).script_name 			= "Update Worker Signature"
script_array(script_num).description 			= "Sets or updates the default worker signature for this user."
script_array(script_num).category               = "UTILITIES"
script_array(script_num).workflows              = ""
script_array(script_num).subcategory            = array("")
script_array(script_num).release_date           = #10/01/2000#
