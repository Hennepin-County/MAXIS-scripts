'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-ACTV LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
call changelog_update("8/6/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function get_user_information(maxis_row, excel_row)
    EMReadScreen worker_x_number, 7, maxis_row, 5
    ObjExcel.Cells(excel_row, 1).Value = worker_x_number

    If all_workers_check = unchecked Then
        EMReadScreen worker_permissions, 29, maxis_row, 38
        ObjExcel.Cells(excel_row, 14).Value = trim(worker_permissions)
    End If

    EMWriteScreen "X", maxis_row, 3
    transmit

    EMReadScreen worker_full_name, 42, 7, 27
    worker_full_name = trim(worker_full_name)

    ObjExcel.Cells(excel_row, 4).Value = worker_full_name

    comma_place = InStr(worker_full_name, ",")
    worker_last_name = left(worker_full_name, comma_place - 1)
    If right(worker_full_name, 1) = "." Then worker_full_name = left(worker_full_name, len(worker_full_name) - 3)
    worker_full_name = trim(worker_full_name)
    worker_first_name = right(worker_full_name, len(worker_full_name) - comma_place)

    ObjExcel.Cells(excel_row, 2).Value = worker_first_name
    ObjExcel.Cells(excel_row, 3).Value = worker_last_name

    EMReadScreen worker_alias, 42, 8, 27
    EMReadScreen address_one, 42, 9, 27
    EMReadScreen address_two, 42, 10, 27
    EMReadScreen address_three, 42, 11, 27
    EMReadScreen address_city, 20, 12, 27
    EMReadScreen address_zip, 10, 12, 50
    EMReadScreen address_phone, 14, 13, 27
    EMReadScreen address_fax, 14, 14, 27
    EMReadScreen supr_id, 7, 14, 61

    ObjExcel.Cells(excel_row, 5).Value = trim(worker_alias)
    ObjExcel.Cells(excel_row, 6).Value = trim(address_one)
    ObjExcel.Cells(excel_row, 7).Value = trim(address_two)
    ObjExcel.Cells(excel_row, 8).Value = trim(address_three)
    ObjExcel.Cells(excel_row, 9).Value = trim(address_city)
    ObjExcel.Cells(excel_row, 10).Value = trim(address_zip)
    address_phone = replace(address_phone, " ", "")
    If address_phone = "()" Then address_phone = ""
    ObjExcel.Cells(excel_row, 11).Value = address_phone
    address_fax = replace(address_fax, " ", "")
    If address_fax = "()" Then address_fax = ""
    ObjExcel.Cells(excel_row, 12).Value = address_fax
    ObjExcel.Cells(excel_row, 13).Value = supr_id

    transmit
end function

'DIALOGS----------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 240, 135, "Pull REPT data into Excel dialog"
  EditBox 90, 25, 145, 15, supervisor_array
  CheckBox 10, 70, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 130, 115, 50, 15
    CancelButton 185, 115, 50, 15
  Text 50, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 90, 45, 145, 20, "Enter 7 digits of supervisors' x1 numbers (ex: x######), separated by a comma."
  Text 10, 30, 75, 10, "Supervisors to Check:"
  Text 10, 90, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

Do
    'Shows dialog
    Dialog pull_rept_data_into_Excel_dialog
    cancel_without_confirmation
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "FIRST NAME"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "LAST NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "FULL NAME"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "ALIAS"
objExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "ADDR TITLE"
objExcel.Cells(1, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "ADDR 1"
objExcel.Cells(1, 7).Font.Bold = TRUE
ObjExcel.Cells(1, 8).Value = "ADDR 2"
objExcel.Cells(1, 8).Font.Bold = TRUE
ObjExcel.Cells(1, 9).Value = "CITY"
objExcel.Cells(1, 9).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "ZIP"
objExcel.Cells(1, 10).Font.Bold = TRUE
ObjExcel.Cells(1, 11).Value = "PHONE"
objExcel.Cells(1, 11).Font.Bold = TRUE
ObjExcel.Cells(1, 12).Value = "FAX"
objExcel.Cells(1, 12).Font.Bold = TRUE
ObjExcel.Cells(1, 13).Value = "SUPERVISOR"
objExcel.Cells(1, 13).Font.Bold = TRUE
ObjExcel.Cells(1, 14).Value = "PERMISSIONS"
objExcel.Cells(1, 14).Font.Bold = TRUE

XL_row = 2
MX_row = 7

If all_workers_check = checked then

    'Getting to REPT/USER
    call navigate_to_MAXIS_screen("rept", "user")

    'Hitting PF5 to force sorting, which allows directly selecting a county
    PF5

    'Inserting county
    EMWriteScreen county_code, 21, 6
    transmit

    Do
        EmReadScreen check_for_user, 1, MX_row, 7

        If check_for_user <> " " Then
            Call get_user_information(MX_row, XL_row)
            XL_row = XL_row + 1
            MX_row = MX_row + 1
        ElseIf check_for_user = " " Then
            Exit Do
        End If

        If MX_row = 19 Then
            EMReadScreen last_page_check, 9, 19, 3

            If last_page_check <> "More:   -" Then
                PF8
                MX_row = 7
            End If
        End If

    Loop until last_page_check = "More:   -"

    end_of_list = XL_row

    XL_row = 2
    Call back_to_SELF
    Call navigate_to_MAXIS_screen("REPT", "USER")

    PF5
    PF5

    Do
        sup_x_numb = ObjExcel.Cells(XL_row, 13).Value
        wrk_x_numb = ObjExcel.Cells(XL_row, 1).Value

        EMWriteScreen sup_x_numb, 21, 12
        transmit

        EMReadScreen sup_found_check, 14, 24, 2
        If sup_found_check = "NO INFORMATION" OR trim(sup_x_numb) = "" Then
            XL_row = XL_row + 1
        Else

            MX_row = 7
            Do
                EMReadScreen mx_x_numb, 7, MX_row, 5

                If mx_x_numb = wrk_x_numb Then
                    EMReadScreen worker_permissions, 29, MX_row, 38
                    ObjExcel.Cells(XL_row, 14).Value = trim(worker_permissions)

                    Exit Do
                End If

                MX_row = MX_row + 1

                If MX_row = 19 Then
                    EMReadScreen last_page_check, 9, 19, 3

                    If last_page_check <> "More:   -" Then
                        PF8
                        EMReadScreen too_many_check, 14, 24, 2
                        If too_many_check = "MAXIMUM NUMBER" Then Exit Do
                        MX_row = 7
                    End If
                End If

            Loop
        End If

        XL_row = XL_row + 1
    Loop until XL_row = end_of_list

Else

    'Getting to REPT/USER
    CALL navigate_to_MAXIS_screen("REPT", "USER")


    'Sorting by supervisor
    PF5
    PF5

    'Splitting the list of inputted supervisors...
    supervisor_array = replace(supervisor_array, " ", "")
    supervisor_array = split(supervisor_array, ",")
    FOR EACH unit_supervisor IN supervisor_array
        IF unit_supervisor <> "" THEN
            'Entering the supervisor number and sending a transmit
            CALL write_value_and_transmit(unit_supervisor, 21, 12)

            MX_row = 7

            Do
                EmReadScreen check_for_user, 1, MX_row, 7

                If check_for_user <> " " Then
                    Call get_user_information(MX_row, XL_row)
                    XL_row = XL_row + 1
                    MX_row = MX_row + 1
                ElseIf check_for_user = " " Then
                    Exit Do
                End If

                If MX_row = 19 Then
                    EMReadScreen last_page_check, 9, 19, 3

                    If last_page_check <> "More:   -" Then
                        PF8
                        MX_row = 7
                    End If
                End If

            Loop until last_page_check = "More:   -"

        End If
    NEXT

End If

'Query date/time/runtime info
objExcel.Cells(1, 16).Font.Bold = TRUE
objExcel.Cells(2, 16).Font.Bold = TRUE
ObjExcel.Cells(1, 16).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 17).Value = now
ObjExcel.Cells(2, 16).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 17).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to 17
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your REPT/USER list has been created.")
