'Required for statistical purposes==========================================================================================
name_of_script = "FIND eDRS Match.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at one
STATS_manualtime = 45                	'manual run time in seconds
STATS_denomination = "M"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
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
db_provider = "SQLOLEDB.1"
db_data_source = "hssqlpw202"
db_catalog = "BlueZone_Statistics"
db_security = "SSPI"
db_translate = "False"

db_full_string = "Provider = " & db_provider & ";Data Source= " & db_data_source & ";Initial Catalog= " & db_catalog & "; Integrated Security=" & db_security & ";Auto Translate=" & db_translate & ";"

const last_name_const = 0
const first_name_const = 1
const mid_initial_const = 2
const dob_const = 3
const ssn_const = 4

'Now determines name of file
output_file_path = user_myDocs_folder & "found-edrs-matches.txt"

set FSOobj = CreateObject("Scripting.FileSystemObject")
'Creating an object for the stream of text which we'll use frequently
Dim objTextStream
If FSOobj.FileExists(output_file_path) = True then
	FSOobj.DeleteFile(output_file_path)
End If

Set objTextStream = FSOobj.OpenTextFile(output_file_path, ForWriting, true)




'declare the SQL statement that will query the database
objSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnap"

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objConnection.Open db_full_string
objRecordSet.Open objSQL, objConnection

'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
todays_cases_list = "*"
name_match_count = 0      'Setting an incrementor for the array to be filled
Dim CASE_PEOPLE()

Do While NOT objRecordSet.Eof
	memb_count = 0
	ReDim CASE_PEOPLE(ssn_const, 0)
	MAXIS_case_number = objRecordSet("CaseNumber")

	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
    EMReadScreen memb_check, 4, 2, 48
    If NOT is_this_priv and memb_check = "MEMB" Then
        EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
        transmit

        DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
            EMReadscreen ref_nbr, 3, 4, 33
            EMReadScreen access_denied_check, 13, 24, 2
            If access_denied_check <> "ACCESS DENIED" Then
                EMReadscreen last_name, 25, 6, 30
                EMReadscreen first_name, 12, 6, 63
                EMReadscreen mid_initial, 1, 6, 79
                EMReadscreen ssn, 11, 7, 42
                EMReadScreen dob, 10, 8, 42
                last_name = trim(replace(last_name, "_", "")) & " "
                first_name = trim(replace(first_name, "_", "")) & " "
                mid_initial = replace(mid_initial, "_", "")
                Redim Preserve CASE_PEOPLE(ssn_const, memb_count)

                CASE_PEOPLE(last_name_const, memb_count) = last_name
                CASE_PEOPLE(first_name_const, memb_count) = first_name
                CASE_PEOPLE(mid_initial_const, memb_count) = mid_initial
                CASE_PEOPLE(dob_const, memb_count) = replace(dob, " ", "-")
                CASE_PEOPLE(ssn_const, memb_count) = ssn
                memb_count = memb_count + 1
                STATS_counter = STATS_counter + 1
            End If
            transmit
            Emreadscreen edit_check, 7, 24, 2
        LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

        Call back_to_SELF
        CALL navigate_to_MAXIS_screen("INFC", "EDRS")

        For the_memb = 0 to UBound(CASE_PEOPLE, 2)
            'Write in SSN number into EDRS
            EMwritescreen replace(CASE_PEOPLE(ssn_const, the_memb), " ", ""), 2, 7
            transmit
            Emreadscreen SSN_output, 7, 24, 2

            'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
            IF SSN_output = "NO DISQ" THEN
                EMWritescreen CASE_PEOPLE(last_name_const, the_memb), 2, 24
                EMWritescreen CASE_PEOPLE(first_name_const, the_memb), 2, 58
                EMWritescreen CASE_PEOPLE(mid_initial_const, the_memb), 2, 76
                transmit
                EMreadscreen NAME_output, 7, 24, 2
                IF NAME_output <> "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name.
                    edrs_row = 5
                    entry_header = "  -  "
                    Do
                        EMReadScreen edrs_match_dob, 10, edrs_row, 58
                        If edrs_match_dob = CASE_PEOPLE(dob_const, the_memb) Then
                            entry_header = "* "
                            name_match_count = name_match_count + 1
                            Exit Do
                        End If
                        If trim(edrs_match_dob) = "" Then Exit Do
                        edrs_row = edrs_row + 1
                        If edrs_row = 20 Then
                            PF8
                            EMReadScreen end_of_list, 9, 24, 14
                            If end_of_list = "LAST PAGE" Then Exit Do
                            edrs_row = 5
                        End If
                        If edrs_row = 16 and entry_header <> "* " Then entry_header = "  ?? "
                    Loop until edrs_row > 20
                    objTextStream.WriteLine entry_header & CASE_PEOPLE(last_name_const, the_memb) & ", " & CASE_PEOPLE(first_name_const, the_memb) & " " & CASE_PEOPLE(mid_initial_const, the_memb) & ".  --  " & CASE_PEOPLE(dob_const, the_memb) & "  --  " & MAXIS_case_number
                END IF
            ELSE
                objTextStream.WriteLine "* " & CASE_PEOPLE(ssn_const, the_memb) & "  --  " & MAXIS_case_number
            END IF
        Next
        Call back_to_SELF
    End If

    ' If name_match_count > 5 Then Exit Do
	' If STATS_counter > 300 Then Exit Do
	objRecordSet.MoveNext

Loop

'close the connection and recordset objects to free up resources
objRecordSet.Close
objConnection.Close
Set objRecordSet=nothing
Set objConnection=nothing

'Close the object so it can be opened again shortly
objTextStream.Close

MAXIS_case_number = ""

call script_end_procedure("DONE")
