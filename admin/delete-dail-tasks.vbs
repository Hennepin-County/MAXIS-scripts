'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DELETE DAIL TASKS.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 600
STATS_denomination = "I"       			'I is for each Item
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
call changelog_update("04/26/2021", "Removed emailing Todd Bennington per request.", "Ilse Ferris, Hennepin County")
call changelog_update("02/10/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

'Gathering windows user information for transpancy purposes. 
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
Call find_user_name(the_person_running_the_script)

BeginDialog Dialog1, 0, 0, 281, 140, "Delete DAIL Tasks Dialog"
  DropListBox 225, 100, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", deletion_choice
  ButtonGroup ButtonPressed
    OkButton 190, 120, 40, 15
    CancelButton 235, 120, 40, 15
  Text 10, 35, 255, 25, "This script will delete ALL actionable DAIL messages captured to be assigned in task-based processing from the SQL Database which feeds the Big Scoop Report."
  GroupBox 5, 20, 270, 75, "Using This Script:"
  Text 10, 70, 255, 20, "By answering with the question below with a YES, you will be confirming your actions."
  Text 65, 5, 75, 10, "---Delete DAIL Tasks---"
  Text 35, 105, 185, 10, "Do you wish to delete the task-based DAILs in database?"
EndDialog

Do 
    err_msg = ""
    dialog Dialog1
    cancel_without_confirmation 
    If deletion_choice = "Select one..." then err_msg = err_msg & vbcr & "* Select a deletion option."
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
Loop Until err_msg = ""  

If deletion_choice = "No" then script_end_procedure("The database will not be deleted. The script will now end.")
    
'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

''Creating objects for Database
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'How to connect to the database
'Provider: the type of connection you are establishing, in this case SQL Server.
'Data Source: The server you are connecting to.
'Initial Catalog: The name of the database.
'user id: your username.
'password: um, your password. ;)

objConnection.Open "Provider = SQLOLEDB.1;Data Source= HSSQLPW017;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'Deleting ALL data fom DAIL table prior to loading new DAIL messages.
objRecordSet.Open "DELETE FROM EWS.DAILDecimator",objConnection, adOpenStatic, adLockOptimistic

'Closing the connection
objConnection.Close

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("Faughn.Ramisch-Church@hennepin.us", "Ilse.Ferris@hennepin.us", "Task-Based Assignment DAIL Messages deleted in SQL Table by " & windows_user_ID & ": " & the_person_running_the_script & ". EOM.", "", "", True)

script_end_procedure("Success! The database has been deleted.")