'Required for statistical purposes==========================================================================================
name_of_script = "CA - XML FILE CLEANUP.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "I"                   'I is for Item - based on search criteria
'END OF stats block=========================================================================================================

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
changelog = array()
call changelog_update("02/24/2023", "Initial version.", "Dave Courtright, Hennepin County")
changelog_display

'END CHANGELOG BLOCK =======================================================================================================
dim archive_date, delete_date
file_path = "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\"
archive_file_path = "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\Archive\"
archive_date = dateadd("m", -1, date)
delete_date = dateadd("m", -3, date)
archive_date = cstr(archive_date)
delete_date = cstr(delete_date)
deleted_count = 0
archived_count = 0


If user_ID_for_validation = "HOAB001" OR user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "MEGE001" OR user_ID_for_validation = "MARI001" OR user_ID_for_validation = "DACO003" Then

	Dialog1 = ""												
	BeginDialog Dialog1, 0, 0, 206, 105, "XML Cleanup"
	  EditBox 140, 30, 60, 15, archive_date
	  EditBox 140, 55, 60, 15, delete_date
	  ButtonGroup ButtonPressed
	    OkButton 90, 85, 50, 15
	    CancelButton 150, 85, 50, 15
	  Text 5, 0, 195, 20, "Use this script to remove old xml files for MNBenefits applications from the shared folder."
	  Text 5, 30, 115, 20, "Archive all files created before this date:"
	  Text 5, 55, 115, 20, "Permanently delete all files created before this date:"
	EndDialog


	DO
				err_msg = ""
				Dialog Dialog1
				If ButtonPressed = 0 then Stopscript
				If isdate(archive_date) = false then err_msg = err_msg & vbCr & "Please enter a valid archive date."
				If isdate(delete_date) = false then err_msg = err_msg & vbCr & "Please enter a valid file delete date."
				If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr
	LOOP UNTIL err_msg = ""

	'Creating object, grabbing the folder content)
	''move to archive folder after a certain time frame
	For each file in main_folder.files
	 	If datediff("d", file.DateCreated, archive_date) > 0 and file.Type = "XML Source File" Then
		    archived_files = archived_files & vbCr & file.name 'add file name to list for log file
			archived_count = archived_count + 1
			file.Move archive_file_path 'move to archive folder
		End If 
	Next
	
	'Delete older than 3 months
	Set archive = fso.GetFolder(archive_file_path)
	For each file in archive.files
		If datediff("d", file.DateCreated, delete_date) > 0 Then'if older than delete date
			deleted_files = deleted_files & vbCr & file.name 'add file name to list for log file
			deleted_count = deleted_count + 1
			file.delete 'deleting the file
		End if 'delete
	Next


	'THIS creates a log file of all cases deleted
	log_usage = true
	If log_usage = true then
		txt_file_name = "xml_cleanup_report_" & replace(date, "/", "_") & ".txt"
		script_test_info_file_path = t_drive &"\Eligibility Support\Assignments\Script Testing Logs\"  & txt_file_name

		Call find_user_name(script_run_worker)

		'CREATING THE TESTING REPORT - perhaps just write a list of deleted files
		With (CreateObject("Scripting.FileSystemObject"))
			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream

			Set objTextStream = .OpenTextFile(script_test_info_file_path, ForWriting, true)

			objTextStream.WriteLine "SCRIPT Run Date and Time: " & now
			objTextStream.WriteLine "Script run be: " & script_run_worker
			objTextStream.WriteLine "Archived Files prior to : " & archive_date & " total: " & archived_count
			objTextStream.WriteLine "-------------------------------------------------"
			objTextStream.WriteLine archived_files
			objTextStream.WriteLine "Deleted Files prior to: " &  delete_date & " total: " & deleted_count
			objTextStream.WriteLine "-------------------------------------------------"

			objTextStream.WriteLine deleted_files
			objTextStream.Close
		End With
	End If
	'END OF TESTING REPORT

	script_end_procedure("Success! Archived " & archived_count & " files and deleted " & deleted_count & " files.")
ELSE
	MSgbox "You are not an authorized user for this script. The script will now stop."
	Stopscript
End If