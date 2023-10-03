'LOADING GLOBAL VARIABLES
Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
If user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "MEGE001" OR user_ID_for_validation = "MARI001" OR user_ID_for_validation = "DACO003" Then
	Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-Scripts\locally-installed-files\SETTINGS - GLOBAL VARIABLES.vbs")
Else
	Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
End If
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING SCRIPT
script_url = script_repository & "/utilities/hot-topics.vbs"
IF run_locally = False THEN
    SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_URL
    req.open "GET", script_URL, FALSE									'Attempts to open the script_URL
    req.send													'Sends request
    IF req.Status = 200 THEN									'200 means great success
    	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
    	Execute req.responseText								'Executes the script code
    ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
    	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
    	vbCr & _
    	"Before contacting the BlueZone script team at HSPH.EWS.BlueZoneScripts@Hennepin.us, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
    	vbCr & _
    	"If you can reach GitHub.com, but this script still does not work, contact the BlueZone script team at HSPH.EWS.BlueZoneScripts@Hennepin.us and provide the following information:" & vbCr &_
    	vbTab & "- The name of the script you are running." & vbCr &_
    	vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
    	vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
    	vbCr & _
    	"We will work with your IT department to try and solve this issue, if needed." & vbCr &_
    	vbCr &_
    	"URL: " & url
    	StopScript
    END IF
ELSE
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_url)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
END IF
