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



class script

    'Stuff the user indicates
	public script_name             	'The familiar name of the script (file name without file extension or category, and using familiar case)
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
    public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
    
    'Details the menus will figure out (does not need to be explicitly declared)
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run

    'Details the class itself figures out
	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property
    
    public property get script_URL
        If script_repository = "" then script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script%20Files/"    'Assumes we're scriptwriters
        script_URL = script_repository & ucase(category) & "/" & replace(ucase(category & "%20-%20" & script_name) & ".vbs", " ", "%20")
    end property
    
    public property get SIR_instructions_URL 'The instructions URL in SIR
        SIR_instructions_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/" & replace(ucase(script_name) & ".aspx", " ", "%20")
    end property

    'THIS IS DEPRECIATED AND SHOULD BE REMOVED FROM MAIN MENU SCRIPTS ONCE IT'S FULLY TESTED
	public file_name               	'The actual file name

end class

script_num = 0
ReDim Preserve script_array(script_num)
Set script_array(script_num) = new script
script_array(script_num).script_name 			= "Application Received"																		'Script name
script_array(script_num).description 			= "Template for documenting details about an application recevied."
script_array(script_num).category               = "NOTES"
script_array(script_num).subcategory            = "#-C"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array(script_num)			'Resets the array to add one more element to it
Set script_array(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array(script_num).script_name 			= "Approved programs"																		'Script name
script_array(script_num).description 			= "Template for when you approve a client's programs."
script_array(script_num).category               = "NOTES"
script_array(script_num).subcategory            = "#-C"

For each script_to_test in script_array
    CreateObject("WScript.Shell").Run(script_to_test.SIR_instructions_URL)
    CreateObject("WScript.Shell").Run(script_to_test.script_URL)
Next
