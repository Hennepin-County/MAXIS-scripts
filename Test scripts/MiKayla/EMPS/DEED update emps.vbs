'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - EMPS PROVIDER UPDATE.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "150"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
'END OF stats block==============================================================================================

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("01/17/2018", "Updated for other provider handling.", "MiKayla Handley, Hennepin County")
call changelog_update("12/12/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------DIALOGS
BeginDialog EMPS_dialog, 0, 0, 266, 95, "EMPS provider update"
ButtonGroup ButtonPressed
	PushButton 200, 25, 50, 15, "Browse...", select_a_file_button
	OkButton 145, 75, 50, 15
CancelButton 200, 75, 50, 15
EditBox 15, 25, 180, 15, file_selection_path
GroupBox 10, 5, 250, 65, "EMPS provider update "
Text 20, 45, 170, 20, "Select the Excel file that contains the information by selecting the 'Browse' button, and finding the file."
EndDialog

'----------------------------------------------------------------------------------------------------THE SCRIPT
'Connects to BlueZone
EMConnect ""

'Shows dialog
Do
	Do
		err_msg = ""
		Dialog DEED_dialog
		cancel_confirmation
		If ButtonPressed = select_a_file_button THEN
			If file_selection_path <> "" THEN 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" THEN err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" THEN call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" THEN MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'Now the script goes back into MFCM and grabs the member # and client name, then checks the potentially exempt members
excel_row = 2           're-establishing the row to start checking the members for

Do

        MAXIS_case_number   = objExcel.cells(excel_row, 1).Value	're-establishing the case number to use for the case
        last_name           = objExcel.cells(excel_row, 3).Value
        first_name          = objExcel.cells(excel_row, 2).Value
      	'member_number       = objExcel.cells(excel_row, 4).Value
        DEED_agency         = objExcel.cells(excel_row, 5).Value	're-establishing the agency
        MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do						'exits do if the case number is ""
        client_name = last_name & ", " & first_name

		EMWriteScreen MAXIS_case_number, 18, 43				'enters member number

		CALL navigate_to_MAXIS_screen("STAT", "EMPS")

		EMReadScreen EMPS_panel_check, 4, 2, 55
      	If EMPS_panel_check <> "EMPS" then ObjExcel.Cells(excel_row, 7).Value = ""

        PF9	    'putting EMPS panel into edit mode
        EMReadScreen err_msg, 18, 24, 02
        If err_msg <> "" THEN
          error_reason = msgbox "UNABLE TO UPDATE EMPS"
          xcel_row = excel_row + 1
        END IF

        Call write_value_and_transmit("x", 19, 25)	'opening 'other provider information pop up box
        EMReadScreen other_box, 5, 4, 30
        IF other_box <> "Other"	THEN
        	error_reason = "Unable to get into Provider information"
        End if

        CALL clear_line_of_text(6, 37)
        CALL clear_line_of_text(7, 37)
        CALL clear_line_of_text(8, 37)
        CALL clear_line_of_text(9, 37)
        CALL clear_line_of_text(10, 47)
        CALL clear_line_of_text(12, 37)
        CALL clear_line_of_text(13, 39)
        CALL clear_line_of_text(13, 45)
        CALL clear_line_of_text(13, 49)

		IF DEED_agency = "AIOIC" THEN
				EMWriteScreen "HSPH.ESP.61AIO", 6, 37
				EMWriteScreen "AIOIC", 7, 37
				EMWriteScreen "1845 EAST FRANKLIN AVENUE", 8, 37
				EMWriteScreen "MINNEAPOLIS", 9, 37
				EMWriteScreen "MN", 10, 37
				EMWriteScreen "55404", 10, 47
				EMWriteScreen "612", 12, 39
				EMWriteScreen "341", 12, 45
				EMWriteScreen "3358", 12, 49
				start_a_blank_CASE_NOTE
					CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO AIOIC AS PART OF THE HC ADJUSTMENT OF CASELOADS")
					CALL write_variable_in_CASE_NOTE("HSPH.ESP.61AIO IS NEW ESP OFFICE")
					CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
				PF3 'saving the case note
				error_reason = "EMPS updated"

        Elseif DEED_agency = "AVIVO BROOKLYN CENTER" THEN
            EMWriteScreen "HSPH.ESP.20268", 6, 37
            EMWriteScreen "AVIVO BROOKLYN CENTER", 7, 37
            EMWriteScreen "5701 SHINGLE CREEK PARKWAY", 8, 37
            EMWriteScreen "BROOKLYN CENTER", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55430", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "752", 12, 45
            EMWriteScreen "8900", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE ASSIGNED TO AVIVO BROOKLYN CENTER")
            CALL write_variable_in_CASE_NOTE("20268 IS AGENCY RETAINING THE CASE")
            CALL write_variable_in_CASE_NOTE("HUB MODEL ENDED 12/31/17")
            PF3 'saving the case note
            error_reason = "EMPS updated"

        Elseif DEED_agency = "AVIVO BLOOMINGTON" THEN
            EMWriteScreen "HSPH.ESP.26AVO", 6, 37
            EMWriteScreen "AVIVO BLOOMINGTON", 7, 37
            EMWriteScreen "2626 EAST 82ND ST #370", 8, 37
            EMWriteScreen "BLOOMINGTON", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55425", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "752", 12, 45
            EMWriteScreen "8940", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO AVIVO BLOOMINGTON")
            CALL write_variable_in_CASE_NOTE("26AVO IS AGENCY RETAINING THE CASE")
            CALL write_variable_in_CASE_NOTE("HUB MODEL ENDED 12/31/17")
            PF3 'saving the case note
            error_reason = "EMPS updated"

        Elseif DEED_agency = "AVIVO NORTH" THEN
                EMWriteScreen "HSPH.ESP.27WIN", 6, 37
                EMWriteScreen "AVIVO NORTH MINNEAPOLIS", 7, 37
                EMWriteScreen "2143 LOWRY AVE NORTH", 8, 37
                EMWriteScreen "MINNEAPOLIS", 9, 37
                EMWriteScreen "MN", 10, 37
                EMWriteScreen "55411", 10, 47
                EMWriteScreen "612", 12, 39
                EMWriteScreen "752", 12, 45
                EMWriteScreen "8500", 12, 49
                start_a_blank_CASE_NOTE
                CALL write_variable_in_CASE_NOTE("ESP CASE ASSIGNED TO AVIVO NORTH MINNEAPOLIS")
                CALL write_variable_in_CASE_NOTE("27WIN IS AGENCY RETAINING THE CASE")
                CALL write_variable_in_CASE_NOTE("HUB MODEL ENDED 12/31/17")
                PF3 'saving the case note
                error_reason = "EMPS updated"

        Elseif DEED_agency = "CAPI BROOKLYN CENTER" THEN
            EMWriteScreen "HSPH.ESP.20297", 6, 37
            EMWriteScreen "CAPI BROOKLYN CENTER", 7, 37
            EMWriteScreen "5930 BROOKLYN BLVD", 8, 37
            EMWriteScreen "BROOKLYN CENTER", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55429", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "588", 12, 45
            EMWriteScreen "3592", 12, 49
			start_a_blank_CASE_NOTE
				CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO CAPI BROOKLYN CENTER AS PART OF THE HC ADJUSTMENT OF CASELOADS")
				CALL write_variable_in_CASE_NOTE("HSPH.ESP.20297 IS NEW ESP OFFICE")
				CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
			PF3 'saving the case note
            error_reason = "EMPS updated"

        Elseif DEED_agency = "CAPI SOUTH" THEN
            EMWriteScreen "HSPH.ESP.1CP50", 6, 37
            EMWriteScreen "CAPI SOUTH", 7, 37
            EMWriteScreen "3702 EAST LAKE ST", 8, 37
            EMWriteScreen "MINNEAPOLIS", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55406", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "721", 12, 45
            EMWriteScreen "0122", 12, 49
            start_a_blank_CASE_NOTE
            	CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO CAPI SOUTH AS PART OF THE HC ADJUSTMENT OF CASELOADS")
				CALL write_variable_in_CASE_NOTE("HSPH.ESP.1CP50 IS NEW ESP OFFICE")
				CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
			PF3 'saving the case note
            error_reason = "EMPS updated"

		Elseif DEED_agency = "EASTSIDE NEIGHBORHOOD" THEN
				EMWriteScreen "HSPH.ESP.55ESN", 6, 37
				EMWriteScreen "EASTSIDE NEIGHBORHOOD", 7, 37
				EMWriteScreen "1700 NE 2ND STREET", 8, 37
				EMWriteScreen "MINNEAPOLIS", 9, 37
				EMWriteScreen "MN", 10, 37
				EMWriteScreen "55413", 10, 47
				EMWriteScreen "612", 12, 39
				EMWriteScreen "781", 12, 45
				EMWriteScreen "6911", 12, 49
				start_a_blank_CASE_NOTE
				CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO EASTSIDE NEIGHBORHOOD AS PART OF HC ADJUSTMENT OF CASELOADS")
	            CALL write_variable_in_CASE_NOTE("HSPH.ESP.55ESN IS NEW ESP OFFICE")
				CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
				PF3 'saving the case note
				error_reason = "EMPS updated"

        Elseif DEED_agency = "EMERGE NORTH" THEN
            EMWriteScreen "HSPH.ESP.79UNI", 6, 37
            EMWriteScreen "EMERGE NORTH", 7, 37
            EMWriteScreen "1834 Emerson Ave North", 8, 37
            EMWriteScreen "Minneapolis", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55411", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "529", 12, 45
            EMWriteScreen "9267", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO EMERGE NORTH ")
            CALL write_variable_in_CASE_NOTE("DEED IS NO LONGER AN MFIP ESP 1/1/2018")
            CALL write_variable_in_CASE_NOTE(" 79UNI IS NEW ESP MFIP COORDINATION OFFICE")
            PF3 'saving the case note
            error_reason = "EMPS updated"

        'Elseif DEED_agency = "GOODWILL EASTER SEALS" THEN
        '    EMWriteScreen "HSPH.ESP.24GES", 6, 37
        '    EMWriteScreen "GOODWILL EASTER SEALS", 7, 37
        '    EMWriteScreen "1455 WEST LAKE ST", 8, 37
        '    EMWriteScreen "MINNEAPOLIS", 9, 37
        '    EMWriteScreen "MN", 10, 37
        '    EMWriteScreen "55408", 10, 47
        '    EMWriteScreen "612", 12, 39
        '    EMWriteScreen "721", 12, 45
        '    EMWriteScreen "8470", 12, 49
        '    start_a_blank_CASE_NOTE
        '    CALL write_variable_in_CASE_NOTE("ESP CASE ASSIGNED TO GOODWILL EASTER SEALS")
        '    CALL write_variable_in_CASE_NOTE("24GES IS AGENCY RETAINING THE CASE")
        '    CALL write_variable_in_CASE_NOTE("HUB MODEL ENDED 12/31/17")
        '    PF3 'saving the case note
        '    error_reason = "EMPS updated"

		Elseif DEED_agency = "GOODWILL EASTER SEALS" THEN
			EMWriteScreen "HSPH.ESP.24GES", 6, 37
			EMWriteScreen "GOODWILL EASTER SEALS", 7, 37
			EMWriteScreen "2801 21ST AVENUE SOUTH", 8, 37
			EMWriteScreen "MINNEAPOLIS", 9, 37
			EMWriteScreen "MN", 10, 37
			EMWriteScreen "55407", 10, 47
			EMWriteScreen "612", 12, 39
			EMWriteScreen "724", 12, 45
			EMWriteScreen "0128", 12, 49
			start_a_blank_CASE_NOTE
			CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO GOODWILL EASTER SEALS AS PART OF HC ADJUSTMENT OF CASELOADS")
            CALL write_variable_in_CASE_NOTE("HSPH.ESP.24GES IS NEW ESP OFFICE")
			CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
			PF3 'saving the case note
			error_reason = "EMPS updated"


        Elseif DEED_agency = "HIRED BLOOMINGTON" THEN
            EMWriteScreen "HSPH.ESP.17HIR", 6, 37
            EMWriteScreen "HIRED", 7, 37
            EMWriteScreen "1701 EAST 79TH ST", 8, 37
            EMWriteScreen "BLOOMINGTON,", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55425", 10, 47
            EMWriteScreen "952", 12, 39
            EMWriteScreen "853", 12, 45
            EMWriteScreen "9100", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO HIRED EAST BLOOMINGTON")
            CALL write_variable_in_CASE_NOTE("DEED IS NO LONGER AN MFIP ESP 1/1/2018")
            CALL write_variable_in_CASE_NOTE("17HIR IS NEW ESP MFIP COORDINATION OFFICE")
            PF3 'saving the case note
            error_reason = "EMPS updated"

        Elseif DEED_agency = "HIRED HENNEPIN NORTH" THEN
            EMWriteScreen "HSPH.ESP.1HD10", 6, 37
            EMWriteScreen "HIRED HENNEPIN NORTH", 7, 37
            EMWriteScreen "7225 NORTHLAND DRIVE", 8, 37
            EMWriteScreen "BROOKLYN PARK", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55428", 10, 47
            EMWriteScreen "763", 12, 39
            EMWriteScreen "210", 12, 45
            EMWriteScreen "6200", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO HIRED HENNEPIN NORTH")
            CALL write_variable_in_CASE_NOTE("DEED IS NO LONGER AN MFIP ESP 1/1/2018")
            CALL write_variable_in_CASE_NOTE("1HD10  IS NEW ESP MFIP COORDINATION OFFICE")
            PF3 'saving the case note
            error_reason = "EMPS updated"

		Elseif name_of_EMPS = "LSS" THEN
			EMPS_array(EMPS_update, item) = True
			EMWriteScreen "HSPH.ESP.42LSS", 6, 37
			EMWriteScreen "LUTHERAN SOCIAL SERVICE", 6, 37
			EMWriteScreen "2400 PARK AVENUE", 8, 37
			EMWriteScreen "MINNEAPOLIS", 9, 37
			EMWriteScreen "MN", 10, 37
			EMWriteScreen "55404", 10, 47
			EMWriteScreen "612", 12, 39
			EMWriteScreen "879", 12, 45
			EMWriteScreen "5372", 12, 49
			start_a_blank_CASE_NOTE
            	CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO LSS AS PART OF HC ADJUSTMENT OF CASELOADS")
            	CALL write_variable_in_CASE_NOTE("HSPH.ESP.42LSS IS NEW ESP OFFICE")
				CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
            PF3 'saving the case note
            error_reason = "EMPS updated"

		Elseif name_of_EMPS = "LIFETRACK: 1LT10" THEN
			EMPS_array(EMPS_update, item) = True
			EMWriteScreen "HSPH.ESP.1LT10", 6, 37
			EMWriteScreen "LIFETRACK RESOURCES", 6, 37
			EMWriteScreen "3433 BROADWAY STREET N.E.", 8, 37
			EMWriteScreen "MINNEAPOLIS", 9, 37
			EMWriteScreen "MN", 10, 37
			EMWriteScreen "55413", 10, 47
			EMWriteScreen "612", 12, 39
			EMWriteScreen "788", 12, 45
			EMWriteScreen "8855", 12, 49
			start_a_blank_CASE_NOTE
	           CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO LIFETRACk AS PART OF HC ADJUSTMENT OF CASELOADS")
	           CALL write_variable_in_CASE_NOTE("HSPH.ESP.1LT10 IS NEW ESP OFFICE")
				CALL write_variable_in_CASE_NOTE("TRANSFERRED VIA BULK SCRIPT")
	        PF3 'saving the case note
	        error_reason = "EMPS updated"

        Elseif DEED_agency = "NORTHPOINT" THEN
            EMWriteScreen "HSPH.ESP.NP027", 6, 37
            EMWriteScreen "NORTHPOINT HEALTH & WELLNESS", 7, 37
            EMWriteScreen "1315 PENN AVE NORTH", 8, 37
            EMWriteScreen "MINNEAPOLIS", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55411", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "767", 12, 45
            EMWriteScreen "0321", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CASE TRANSFER TO NORTHPOINT HEALTH & WELLNESS")
            CALL write_variable_in_CASE_NOTE("DEED IS NO LONGER AN MFIP ESP 1/1/2018")
            CALL write_variable_in_CASE_NOTE(" NP027 IS NEW ESP MFIP COORDINATION OFFICE")
            PF3 'saving the case note
            error_reason = "EMPS updated"

        Elseif DEED_agency = "RISE, INC SOUTH" THEN
            EMWriteScreen "HSPH.ESP.1RI50", 6, 37
            EMWriteScreen "RISE, INC SOUTH", 7, 37
            EMWriteScreen "3708 NICOLLET AVE SOUTH", 8, 37
            EMWriteScreen "MINNEAPOLIS", 9, 37
            EMWriteScreen "MN", 10, 37
            EMWriteScreen "55409", 10, 47
            EMWriteScreen "612", 12, 39
            EMWriteScreen "872", 12, 45
            EMWriteScreen "7720", 12, 49
            start_a_blank_CASE_NOTE
            CALL write_variable_in_CASE_NOTE("ESP CSE ASSIGNED TO RISE, INC SOUTH")
            CALL write_variable_in_CASE_NOTE("1RI50 IS AGENCY RETAINING THE CASE")
            CALL write_variable_in_CASE_NOTE("HUB MODEL ENDED 12/31/17")
            PF3 'saving the case note
            error_reason = "EMPS updated"
        END IF

        ObjExcel.Cells(excel_row,  7).Value = trim(error_reason)
        excel_row = excel_row + 1
        STATS_counter = STATS_counter + 1
        back_to_SELF

LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

STATS_counter = STATS_counter - 1
script_end_procedure("Success! Please review the list generated.")
