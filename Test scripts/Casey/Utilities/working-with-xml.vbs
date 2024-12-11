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



xmlPath = user_myDocs_folder & "test_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".xml"
MsgBox "xmlPath - " & vbCr & vbCr & xmlPath
'Grabbing some information from the xml file for testing report
' file_name = replace(xmlPath, "T:\Eligibility Support\EA_ADAD\EA_ADAD_Common\CASE ASSIGNMENT\MNB_XML_files\", "")
running_error = ""

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.async = False

' Load the XML file
Set root = xmlDoc.createElement("form")
xmlDoc.appendChild root
Set element = xmlDoc.createElement("DHSNumber")
root.appendChild element
Set info = xmlDoc.createTextNode("5223")
element.appendChild info
' xmlDoc.getElementsByTagName("form").insertBrefore(xmlDoc.getElementsByTagName("DHSNumber"), xmlDoc.getElementsByTagName("form").childNodes.item(0))
' NewInput = element.insertBrefore(info, element.childNodes.item(0))
' element = xmlDoc.createElement("Name")
' insertElemnet = root.insertBrefore(element, root.childNodes.item(1))
' info =xmlDoc.createTextNode("CAF")
' NewInput = element.insertBrefore(info, element.childNodes.item(0))

xmlDoc.save(xmlPath)
' xmlDoc.createElement("form")
' xmlDoc.createElement("DHSNumber")
' xmlDoc.save(xmlPath)
MsgBox "PAUSE"
stopscript