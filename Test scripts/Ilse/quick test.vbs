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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'next_revw_date = "01/01/19"
'
'last_day_of_revw = dateadd("d", -1, next_revw_date) & "" 	'blank space added to make vorianble to make a string
'
'revw_start_date = dateadd("M", - 6, next_revw_date)	'blank space added to make vorianble to make a string
''revw_start_date = right("0" & DatePart("YYYY", next_revw_date), 2)
'


'Function for sorting numeric array descending (biggest to smallest)
FUNCTION find_smallest_number(values_array, separate_character, output_variable)
	'trimming and splitting the array
	values_array = trim(values_array)
	values_array = split(values_array, separate_character)
	num_of_values = ubound(values_array)
	
	array_position = 0

	DIM placeholder_array()
	REDIM placeholder_array(num_of_values, 1) ' position 0 is the number, position 1 is if the number has been put in the output array
	array_position = 0 'assigning the number values to the multi-dimensional placeholder array AND whether the specific value has been used for comparison yet (position 1)
	FOR EACH num_char IN values_array
		IF num_char <> "" THEN
			num_char = cdbl(num_char)
			placeholder_array(array_position, 0) = num_char
			placeholder_array(array_position, 1) = FALSE
			array_position = array_position + 1
		END IF
	NEXT
	
	array_position = 0 'reseting array_position for the generation of the output array
	i = 0
	all_sorted = FALSE
	DO
		lowest_value = FALSE 'stating that the number has not yet been put into the sorted array
		value_to_watch = placeholder_array(i, 0)
		msgbox value_to_watch
		IF placeholder_array(i, 1) = FALSE THEN
			FOR item = 0 TO num_of_values 'If the value is not blank AND if we still have not assigned this value to the output array. We need
				msgbox "item: " & item
				' to avoid a list of only the lowest values, which is what happens what you remove the placeholder_array(j, 1) bit
				IF placeholder_array(item, 0) <> ""  AND placeholder_array(item, 1) = FALSE THEN
					IF value_to_watch <= placeholder_array(item, 0) THEN
						lowest_value = TRUE 'If we confirm that this is the lowest value...
						msgbox "lowest_value " & lowest_value
						'...then we assign position 1 as TRUE (so we will not use this value for comparison in the future)
						placeholder_array(i, 1) = TRUE
						'...we assign it to the output array...
						'output_array = output_array & value_to_watch & ","
						output_variable = value_to_watch
						'...and we move on to the next position in the array...
						array_position = array_position + 1
						'...until we find that we have hit the ubound for the original array. Then we stop assigning.
						IF array_position = num_of_values THEN all_sorted = TRUE
					ELSE
						EXIT FOR  'If the function finds a value LOWER than the current one, it stops comparison and exits the FOR NEXT
					END IF
				END IF
			NEXT
		END IF
		'If we get through this specific number and find that it does not go next on the sorted list,
		' we need to get to the next number. If we find that we have got through all the numbers and the list
		' is not complete, we need to reset this value, and start back at the beginning of the original list.
		' This way, we avoid skipping numbers that should be showing up on the list.
		i = i + 1
		IF i = num_of_values AND all_sorted = FALSE THEN i = 0
	LOOP UNTIL all_sorted = TRUE

	'output_array = trim(output_array)
	'output_array = split(output_array, ",")
END FUNCTION

test_array = "16|2|23"

Call find_smallest_number(test_array, "|", testing_variable)
'msgbox "all done " & join(testing_array, ",")
msgbox "testing_variable: " & testing_variable
stopscript
 
'start_date = "02/01/2018"   'start and end service agreement dates
'end_date = "06/18/2018"
'total_units = datediff("D", start_date, end_date)
'msgbox total_units

'MsgBox(client_age("08/18/1963"))
'Function client_age(client_DOB)
'    Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
'    CurrentDate = CDate(client_DOB)
'    Years = DateDiff("yyyy", CurrentDate, Date)
'    ThisYear = DateAdd("yyyy", Years, CurrentDate)
'    Months = DateDiff("m", ThisYear, Date)
'    ThisMonth = DateAdd("m", Months, ThisYear)
'    Days = DateDiff("d", ThisMonth, Date)
'
'    Do While (Days < 0) Or (Months < 0)
'        If Days < 0 Then
'            Months = Months - 1
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'        If Months < 0 Then
'            Years = Years - 1
'            ThisYear = DateAdd("yyyy", Years, CurrentDate)
'            Months = DateDiff("m", ThisYear, Date)
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'    Loop
'    client_age = Years & "y/" & Months & "m/" & Days
'End Function


stopscript