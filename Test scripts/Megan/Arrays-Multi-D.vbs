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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

EMConnect ""


'~~~~~~~~~~MULTI DIMENSIONAL ARRAYS~~~~~~~~~~ 
        'Errors/Facts
            'You cannot reassing a constant otherwise you will receive an error: Script Error at Line X, Column X: Illegal assignment: "CONST_NAME"
                'constant would look like: const cash_1_stat_cont = 4
                'reassigning would look like: case_1_stat_const = 4


    
'Step 1: Define constants for array. Think of constants as rows in excel. This is how you store information in an array. 
        'Use "const" at the end of const name to keep it unique from say a variable with the same name (i.e. case_nbr_const)
        'Add an extra constant at the end of your constant list. You can even name it "the_last_const" 
            'This helps keep the integrity of your coding following the array so you don't have to make multiple updates if you need to add to your list of constants. 
            'If you need to add a constant, you add it before the_last_const and update the numbers respectively 
const Case_Nbr_const		= 0
const Name_const			= 1
const Revw_Date_const		= 2
const Cash_1_prog_const		= 3
const Cash_1_stat_const		= 4
const Cash_2_prog_const		= 5
const Cash_2_stat_const		= 6
const FS_const				= 7
const HC_const				= 8
const EA_const				= 9
const GR_const				= 10
const IVE_const				= 11
const FIAT_const			= 12
const CC_const				= 13
const the_last_const		= 14


'Step 2: Dim and ReDim array. 
    'Dim Array_Name ()
    'ReDim Array_Name (last_const, #)
        'ReDim tells the system we want to make changes to the array
        'First field in the () is always the last constant
DIM REPT_ACVT_ARRAY()
ReDim REPT_ACVT_ARRAY(the_last_const, 0)


'Step 3: Navigating to the correct screen
Call back_to_SELF
Call navigate_to_MAXIS_screen("REPT", "ACTV")
Call write_value_and_transmit("X127MG2", 21, 13)		'Was using X127EZ2


'Step 4: Array counters and row defining. 
        'Need an array counter which should start at 0 (otherwise for loops get messy)
        'Defining the row ahead of time helps for dynamic coding. 
row = 7
array_counter = 0


'Step 5: Create a Do-Loop with read screen and array storage. 
    ' Arrays are used to store information. Typically we know what fields of information we want to collect, but not how many instances we need to collect which is why we use a Do-Loop to cycle through all of the instances.
        'i.e. Know we want to collect dates/names/cases/etc from a specific panel but we dont know how many cases are listed in said panel (case# is the unique identifier).
    'In this example, we are using a do-loop to page through all of the REPT/ACTV panels to read all of the case information listed.
    'If you are cycling through information over and over again using an array, it can be helpful to use "temp" to identify what info it is reading over and over again for each line/panel.
    'Array Counter: 
        'Only want to increment array counted if there is another entry. If there are no more rows, you do not want it to continue.
        'Array counters should be the last thing listed before the End If statement 
   
    
Do
	EMReadScreen case_numb_temp, 7, row, 13                                 'Readscreen first, leveraging "row" istead of listing a specific #, this allows us to use Row + 1 to move down a row with each loop cycle.
	case_numb_temp = trim(case_numb_temp)                                   'Always good to clean up your coding by eliminating extra spaces since not all case numbers are the same in length.

	if case_numb_temp <> "" Then                                            'Need If statement to determine if there is something in the case number field, if it's blank there is no information to capture. 
		ReDim Preserve REPT_ACVT_ARRAY(the_last_const, array_counter)       'CRITICAL: Need to tell our array how big we want it to be evertime! Preserve is critical to ensure the information is not written over with each cycle.

		REPT_ACVT_ARRAY(Case_Nbr_const, array_counter) = case_numb_temp     'TO DO: Is this correct? Array_Counter is telling it to save the info it read ? If we have rows of our constants is the array_counter our columns of info?

		EMReadScreen name_temp, 	20, row, 21                             'Reading the screen for every desired field. Row is used so that it updated with each loop. 
		EMReadScreen revw_date_temp, 8, row, 42
		EMReadScreen cash_1p_temp, 	 2, row, 51
		EMReadScreen cash_1s_temp, 	 1, row, 54
		EMReadScreen cash_2p_temp, 	 2, row, 56
		EMReadScreen cash_2s_temp, 	 1, row, 59
		EMReadScreen fs_temp, 		 1, row, 61
		EMReadScreen hc_temp, 		 1, row, 64
		EMReadScreen ea_temp, 		 1, row, 67
		EMReadScreen grh_temp,		 1, row, 70
		EMReadScreen ive_temp, 		 1, row, 73
		EMReadScreen fiat_temp, 	 1, row, 77
		EMReadScreen cc_temp, 		 1, row, 80


                                                                            'TODO: Clarify this chunck of code. Storing the read information in the array. Use array_counter so that it move over a column with each line of information read. Is this correct? 
                                                                            'This is a good time to do format cleanup
		REPT_ACVT_ARRAY(Name_const, array_counter) = trim(name_temp)            
		REPT_ACVT_ARRAY(Revw_Date_const, array_counter) = replace(revw_date_temp, " ", "/")                                         'This line of code is listed twice so we can do some formatting
		REPT_ACVT_ARRAY(Revw_Date_const, array_counter) = DateAdd("d", 0, REPT_ACVT_ARRAY(Revw_Date_const, array_counter))          'This whole thing is now the variable which is reading the date, adding slashes, and telling vbscript it's a date: REPT_ACVT_ARRAY(Revw_Date_const, array_counter)
		REPT_ACVT_ARRAY(Cash_1_prog_const, array_counter) = cash_1p_temp
		REPT_ACVT_ARRAY(Cash_1_stat_const, array_counter) = cash_1s_temp
		REPT_ACVT_ARRAY(Cash_2_prog_const, array_counter) = cash_2p_temp
		REPT_ACVT_ARRAY(Cash_2_stat_const, array_counter) = cash_2s_temp
		REPT_ACVT_ARRAY(FS_const		, array_counter) = fs_temp
		REPT_ACVT_ARRAY(HC_const		, array_counter) = hc_temp
		REPT_ACVT_ARRAY(EA_const		, array_counter) = ea_temp
		REPT_ACVT_ARRAY(GR_const		, array_counter) = grh_temp
		REPT_ACVT_ARRAY(IVE_const		, array_counter) = ive_temp
		REPT_ACVT_ARRAY(FIAT_const		, array_counter) = fiat_temp
		REPT_ACVT_ARRAY(CC_const		, array_counter) = cc_temp

		array_counter = array_counter + 1
	end if

	row = row + 1                                                               'This navigates us to the next page if you hit row 19, then redefine row as row =7 so it starts reading from the top again.
	If row = 19 Then
		PF8
		row = 7
	End if
Loop until case_numb_temp = ""                                                  'This is telling us to Loop Until there are no more case #s

' MsgBox "array_counter - " & array_counter
' array_counter = array_counter-1


'Step 6: For-Next defines the array and allows you to display information. In this exercise we are msgboxing what we've read and stored in the array for simplicity. 
    'For Next Loop Structure Below -  Note for i, we used each_thing in our example, it could also be array_counter but you have to be consistent with use. This allows you to move through the array correctly from top to bottom.
        'For i = 0 to Ubound(array_name)
            'Msgbox array_name (i)
        'Next

    'In projects, we won't likely msgbox, instead we will pull this info into a dialog box, store it in a spreadsheet, etc. 

for each_thing = 0 to UBound(REPT_ACVT_ARRAY, 2)                                    'Defaults to the first parameter unless you state what parameter, in this case we indicate 2                            
	'MsgBox "each_thing - " & each_thing												'Msgbox for each item found
	If REPT_ACVT_ARRAY(FS_const, each_thing) = "A" Then                       'Picking criteria to search within our array for CHANGED from Cash_1_prog_const to FS_const for testing
		MsgBox REPT_ACVT_ARRAY(Case_Nbr_const, each_thing) & " is on " & REPT_ACVT_ARRAY(FS_const, each_thing)     'This will msgbox each case	CHANGED from Cash_1_prog_const to FS_const for testing
	End If

Next

'MsgBox "UBound 1 - " & UBound(REPT_ACVT_ARRAY) & vbCr & "UBound 2 - " & UBound(REPT_ACVT_ARRAY, 2) 