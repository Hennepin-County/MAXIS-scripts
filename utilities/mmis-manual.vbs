'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - HOT TOPICS.vbs"
start_time = timer

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
call changelog_update("06/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Call open_URL_in_browser("https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=user_manual")
' run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=user_manual"

' Option Explicit

' Public Sub Sample()
'   Dim pid As Long
'   Dim pno As Long: pno = 17556
'
'   pid = StartEdgeDriver(PortNo:=pno)
'   If pid = 0 Then Exit Sub
'   With CreateObject("Selenium.WebDriver")
'     .StartRemotely "http://localhost:" & pno & "/", "MicrosoftEdge"
'     .Get "https://www.bing.com/"
'     .FindElementById("sb_form_q").SendKeys "abcdefg"
'     MsgBox "pause", vbInformation + vbSystemModal
'     .Quit
'   End With
'   TerminateEdgeDriver pid
'   MsgBox "done.", vbInformation + vbSystemModal
' End Sub
'
' Private Function StartEdgeDriver( _
'   Optional ByVal DriverPath As String = "C:\Windows\System32\MicrosoftWebDriver.exe", _
'   Optional ByVal PortNo As Long = 17556) As Long
'
'   Dim DriverFolderPath As String
'   Dim DriverName As String
'   Dim Options As String
'   Dim itm As Object, itms As Object
'   Dim pid As Long: pid = 0
'
'   With CreateObject("Scripting.FileSystemObject")
'     If .FileExists(DriverPath) = False Then GoTo Fin
'     DriverFolderPath = .GetParentFolderName(DriverPath)
'     DriverName = .GetFileName(DriverPath)
'   End With
'
'   'check already running process
'   Set itms = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery _
'              ("Select * From Win32_Process Where Name = '" & DriverName & "'")
'   If itms.Count > 0 Then
'     For Each itm In itms
'       pid = itm.ProcessId: GoTo Fin
'     Next
'   End If
'
'   'execute WebDriver
'   Options = " --host=localhost --jwp --port=" & PortNo
'   With CreateObject("WbemScripting.SWbemLocator").ConnectServer.Get("Win32_Process")
'     .Create DriverPath & Options, DriverFolderPath, Null, pid
'   End With
'
' Fin:
'   StartEdgeDriver = pid
' End Function
'
' Private Sub TerminateEdgeDriver(ByVal ProcessId As Long)
'   Dim itm As Object, itms As Object
'
'   Set itms = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery _
'              ("Select * From Win32_Process Where ProcessId = " & ProcessId & "")
'   If itms.Count > 0 Then
'     For Each itm In itms
'       itm.Terminate: Exit For
'     Next
'   End If
' End Sub


'Script ends
script_end_procedure("")
