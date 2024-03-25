'Required for statistical purposes==========================================================================================
name_of_script = "NOTES Maxis-to-Mets.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block================================================================================

run_locally = true

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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

'END FUNCTIONS LIBRARYBLOCK================================================================================================

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County

CALL changelog_update("10/20/2023", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'MAXIS_CASE_NUMBER = "330077"
'This class is necessary for the HH_member_enhanced_dialog. 
	Class member_data
		public member_number
		public name
		public ssn
		public birthdate
        public first_checkbox
        public second_checkbox
	End Class


function HH_member_enhanced_dialog(HH_member_array, instruction_text, display_birthdate, display_ssn, first_checkbox, first_checkbox_default, second_checkbox, second_checkbox_default)
'--- This function creates an array of all household members in a MAXIS case, and displays a dialog of HH members that allows the user to select up to two checkboxes per member.
'~~~~~ enhanced_HH_member_array: array that stores all members of the household, with attributes for each member stored in an object. 
'~~~~~ instruction_text: String variable that will appear at the top of dialog as text to give instructions or other info to the user. Limit to 400 characters????
'~~~~~ display_birthdate: true/false. True will display the birthdate after the member name for each HH member
'~~~~~ display_ssn: true/False. True will display the last 4 digits of the SSN after the member name for each HH member
'~~~~~ first_checkbox: string value that contains the text to display for the first checkbox. If no checkbox is wanted, set to ""
'~~~~~ first_checkbox_default: checked/unchecked or 0/1. Determines default state of first checkbox.
'~~~~~ second_checkbox: string value that contains the text to display for the second checkbox. If no checkbox is wanted, set to ""
'~~~~~ second_checkbox_default: checked/unchecked or 0/1. Determines default state of first checkbox.
'If both checkboxes are set to "", the dialog will not display. Use this option when populating an array of the whole household.
'The 6 attributes (member_number, name, ssn, birthdate, first_checkbox, second_checkbox) will be stored in the enhanced_hh_member_array and can be called with this syntax: enhanced_hh_member_array(member).birthdate 
'===== Keywords: MAXIS, member, array, dialog
dim enhanced_HH_member_array()
	call check_for_MAXIS(false)
	membs = 1
    'redim enhanced_HH_member_array(1)
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
    transmit
	EMREadScreen total_clients, 2, 2, 78
	total_clients = cint(replace(total_clients, " ", ""))
	redim enhanced_HH_member_array(total_clients, 5)
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		 'if only one MEMB screen, we don't need to display the dialog 
		
        EMReadScreen access_denied_check, 13, 24, 2
        'MsgBox access_denied_check
        If access_denied_check = "ACCESS DENIED" Then
            PF10
			EMWaitReady 0, 0
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
			ssn_last_4 = ""
			birthdate = ""
        Else
            EMReadscreen ref_nbr, 3, 4, 33
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
			EMReadScreen ssn, 11, 7, 42 
			EMReadScreen birthdate, 10, 8, 42
    		last_name = trim(replace(last_name, "_", "")) & " "
    		first_name = trim(replace(first_name, "_", "")) & " "
    		mid_initial = replace(mid_initial, "_", "")
			birthdate = replace(birthdate, " ", "/")
		End If
		client_string = last_name & first_name & mid_initial
		'Create an object for the member and add that members info, plus the checkbox defaults
        		
		enhanced_HH_member_array(membs, 0) = ref_nbr
		enhanced_HH_member_array(membs, 1) = client_string
		enhanced_HH_member_array(membs, 2) = replace(ssn, " ", "") 
		enhanced_HH_member_array(membs, 3) = birthdate
		enhanced_HH_member_array(membs, 4) = first_checkbox_default
		enhanced_HH_member_array(membs, 5) = second_checkbox_default

  		membs = membs + 1 'index the value up 1 for next member
		transmit
	    Emreadscreen edit_check, 7, 24, 2

	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	instruction_text_lines = (len(instruction_text) \ 80) + 1
	if total_clients > 8 Then instruction_text_lines = (len(instruction_text) \ 160) + 1
	If total_clients > 1 OR second_checkbox <> "" Then 'We only need the dialog if more than 1 client or multiple checkboxes to select
		'Generating the dialog
		split_number = 9
		If total_clients > 8 Then split_number = (total_clients \ 2) + 1
	    member_height = 15
	    If display_ssn = true Or display_birthdate = true  Then member_height = member_height + 15
	    If first_checkbox <> "" Then member_height = member_height + 15
	    If second_checkbox <> "" Then member_height = member_height + 15
	    
	    If total_clients < split_number Then 'Single column dialog
			dialog_width = 290
			dialog_height = (total_clients * 35) + (instruction_text_lines * 15) + 20
		Else
			dialog_width = 580
			dialog_height = (split_number * 35) + (instruction_text_lines * 15) + 20
		End If 
		dialog1 = ""

	    BEGINDIALOG dialog1, 0, 0, dialog_width, dialog_height, "HH Member Dialog"   
			y_pos = 5
	    	Text 10, y_pos, dialog_width - 20, 10 * instruction_text_lines, instruction_text
			y_pos = y_pos + (10 * instruction_text_lines) + 10

	    	FOR person = 1 to total_clients										
	    		'enhanced_HH_member_array(i).member_number
                x_pos = 10
				IF enhanced_HH_member_array(person, 0) <> "" THEN 
	    			if person > split_number THEN x_pos = 300
					display_string = enhanced_HH_member_array(person, 1) 'client name
					If display_birthdate = True Then display_string = display_string & " " & enhanced_HH_member_array(person, 3) 'birthdate
					If display_ssn = True Then display_string = display_string & "  XXX-XX-" & right(enhanced_HH_member_array(person, 2), 4) 'ssn
					Text x_pos, y_pos, 270, 10, enhanced_HH_member_array(person, 0) & " " & display_string   'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	    			If first_checkbox <> "" Then checkbox x_pos + 10, y_pos + 15, 125, 10, first_checkbox, enhanced_HH_member_array(person, 4) 'enhanced_HH_member_array(i).first_checkbox 
                    If second_checkbox <> "" Then checkbox x_pos + 140, y_pos + 15, 125, 10, second_checkbox, enhanced_HH_member_array(person, 5)   
	    			y_pos = y_pos + 30
					if person = split_number Then y_pos = 15 + (10 * instruction_text_lines) 'resets y value when moving to next column
                End If
	    	NEXT
	    	ButtonGroup ButtonPressed
	    	OkButton dialog_width - 115, dialog_height - 20, 50, 15
	    	CancelButton dialog_width - 60, dialog_height - 20, 50, 15 
	    ENDDIALOG
		'runs the dialog that has been dynamically created
                              
    	    
	    Dialog dialog1
	    Cancel_without_confirmation
	End If 
	'This section puts each person's info into objects in teh HH_member_array
	redim hh_member_array(total_clients)
	For memb = 1 to total_clients
		set HH_member_array(memb) = new member_data
		with HH_member_array(memb)
			.member_number = enhanced_HH_member_array(memb, 0)
			.name = enhanced_HH_member_array(memb, 1)
			.ssn = enhanced_HH_member_array(memb, 2)
			.birthdate = enhanced_HH_member_array(memb, 3)
			.first_checkbox = enhanced_HH_member_array(memb, 4)
			.second_checkbox = enhanced_HH_member_array(memb, 5)
		end with
	next
end function

Call MAXIS_case_number_finder(MAXIS_CASE_NUMBER)

call HH_member_enhanced_dialog(HH_member_array, "Select the HH Members that are potentially migrating to METS below. Do not select members that do not have a potential migration reason.", true, true, "No longer has a MAXIS basis.", 1, "Continues to meet a Maxis Basis.", 0)
 

For memb = 1 to ubound(hh_member_array)
	with hh_member_array(memb)
  	  If .first_checkbox = checked Then msgbox .member_number & " " & .name & " " & .ssn
	end with
Next

call HH_member_enhanced_dialog(HH_member_array, "Select household members that are in the assistance unit or deemers.", true, false, "Include in budget.", 1, "", 0)

for i = 1 to ubound(hh_member_array)
	If hh_member_array(i).first_checkbox = checked Then msgbox hh_member_array(i).member_number
Next
stopscript

