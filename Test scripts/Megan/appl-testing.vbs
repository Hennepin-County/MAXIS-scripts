'STATS GATHERING=============================================================================================================
name_of_script = "APPL-Testing.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 10            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'END CHANGELOG BLOCK =======================================================================================================

CALL file_selection_system_dialog(xmlPath, ".xml")

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.async = False

' Load the XML file
xmlDoc.load(xmlPath)

MAXIS_case_number = "________"

'CONNECTING TO MAXIS, STOPPING THE CASE NUMBER FROM CARRYING THROUGH
EMConnect ""

'NAVIGATING TO THE SCREEN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'checking for active MAXIS session
call check_for_MAXIS(True)

' Constants for categories
Const global_child_const = 0
Const global_other_family_const = 1
Const global_unknown_const = 2
Const global_other_non_family_const = 3

' Array to store the last used reference number for each category
Dim global_last_used_reference_number(3)
global_last_used_reference_number(global_child_const) = 2
global_last_used_reference_number(global_other_family_const) = 15
global_last_used_reference_number(global_unknown_const) = 26
global_last_used_reference_number(global_other_non_family_const) = 29

' Check if the XML file is loaded successfully and set objects
If xmlDoc.parseError.errorCode <> 0 Then
    MsgBox "Error in XML: " & xmlDoc.parseError.reason
Else
    Dim formattedDate, objApplicationDate, appl_date, appl_month, appl_day, appl_year
    ' Application Date - May need to add logic if Application Date falls on a weekend or holiday, but we can add that later.
    Set objApplicationDate = xmlDoc.selectSingleNode("//CONTENT/ap:Bytes/io4:SubmitDate")
    If Not objApplicationDate Is Nothing Then
        appl_date = objApplicationDate.Text
        appl_month = Mid(appl_date, 6, 2)
        appl_day = Mid(appl_date, 9, 2)
        appl_year = Mid(appl_date, 1, 4)
    Else ' Use the current date if the application date is not available
        Dim currentDate
        currentDate = Date
        appl_month = Right("0" & Month(currentDate), 2)
        appl_day = Right("0" & Day(currentDate), 2)
        appl_year = Year(currentDate)
    End If

    formattedDate = appl_month & "/" & appl_day & "/" & appl_year
    MAXIS_footer_month = appl_month
    MAXIS_footer_year = Mid(appl_year, 3, 2)
    'MsgBox "Application Date: " & formattedDate
    
    'Appl screen
    call navigate_to_MAXIS_screen("Appl", "____")

    ' Find ContactInfo occurring once
    Set contactInfo = xmlDoc.selectSingleNode("//ns4:ContactInfo")
    
    ' Find MiscellaneousInfo occurring once
    Set MiscellaneousInfo = xmlDoc.selectSingleNode("//ns4:MiscellaneousInfo")
    
    If Not contactInfo Is Nothing And Not contactInfo Is Nothing Then
        'MsgBox contactInfo.getElementsByTagName("ns4:Person").Item(0).Text

        ' Access individual contact info data
        Con_LastName = contactInfo.getElementsByTagName("ns4:Person").Item(0).getElementsByTagName("ns4:LastName").Item(0).Text
        'MsgBox "LastName " & Con_LastName
        Con_FirstName = contactInfo.getElementsByTagName("ns4:Person").Item(0).getElementsByTagName("ns4:FirstName").Item(0).Text
        'MsgBox "Con firstName " & Con_FirstName
        
        Con_Line = contactInfo.getElementsByTagName("ns4:Address").Item(0).getElementsByTagName("ns4:Line").Item(0).Text
        'MsgBox "Con Line " & Con_Line
        Con_City = contactInfo.getElementsByTagName("ns4:Address").Item(0).getElementsByTagName("ns4:City").Item(0).Text
        'MsgBox "Con City " & Con_City
        Con_State_code = contactInfo.getElementsByTagName("ns4:Address").item(0).getElementsByTagName("ns4:StateCode").Item(0).Text
        'MsgBox "Con State code " & Con_State_code
        Con_Zip5 = contactInfo.getElementsByTagName("ns4:Address").item(0).getElementsByTagName("ns4:Zip5").Item(0).Text
        'MsgBox "Con Zip5 " & Con_Zip5
        
        Mail_Line = contactInfo.getElementsByTagName("ns4:MailingAddress").Item(0).getElementsByTagName("ns4:Line").Item(0).Text
        'MsgBox "Mail Line " & Mail_Line
        Mail_City = contactInfo.getElementsByTagName("ns4:MailingAddress").Item(0).getElementsByTagName("ns4:City").Item(0).Text
        'MsgBox "Mail City " & Mail_City
        Mail_State_code = contactInfo.getElementsByTagName("ns4:MailingAddress").item(0).getElementsByTagName("ns4:StateCode").Item(0).Text
        'MsgBox "Mail State code " & Mail_State_code
        Mail_Zip5 = contactInfo.getElementsByTagName("ns4:MailingAddress").item(0).getElementsByTagName("ns4:Zip5").Item(0).Text
        'MsgBox "Mail Zip5 " & Mail_Zip5
        
        Dim countyResidenceNodes, corNodes
        Set countyResidenceNodes = contactInfo.getElementsByTagName("ns4:CountyResidence")

        If countyResidenceNodes.Length > 0 Then
            Set corNodes = countyResidenceNodes.Item(0).getElementsByTagName("ns4:COR")
            If corNodes.Length > 0 Then
                COR = corNodes.Item(0).Text
            Else
                COR = "Dakota" ' or some default value
            End If
        Else
            COR = "Dakota" ' or some default value
        End If
        'MsgBox "COR " & COR

        Con_phone = contactInfo.getElementsByTagName("ns4:Phone").Item(0).getElementsByTagName("ns4:PhoneNumber").Item(0).Text
        'MsgBox "Contact phone " & Con_phone
        EMWriteScreen left(formattedDate, 3), 4, 63
        EMWriteScreen mid(formattedDate, 4, 3), 4, 66
        EMWriteScreen right(formattedDate, 2), 4, 69
        'MsgBox "pause"
        EMWritescreen Con_LastName, 7, 30
        EMWritescreen Con_FirstName, 7, 63
        'MsgBox "pause"
        Transmit
        EMReadScreen scrLine, 6, 2, 44
        If scrLine = "(APPL)" Then
            Call Error_Process
            StopScript
        End If
        'MsgBox "pause"
    
        If Not MiscellaneousInfo Is Nothing Then
            ' Access Miscellaneous info data
            PrimarySpokenLanguage = MiscellaneousInfo.getElementsByTagName("ns4:PrimarySpokenLanguage").Item(0).Text
            'MsgBox "Spoken lang " & Spoken_Language
            PreferredWrittenLanguage = MiscellaneousInfo.getElementsByTagName("ns4:PreferredWrittenLanguage").Item(0).Text
            'MsgBox "Written lang " & Written_Lang
            NeedInterpreterInd = MiscellaneousInfo.getElementsByTagName("ns4:NeedInterpreterInd").Item(0).Text
            'MsgBox "Written lang " & NeedInterpreterInd
        End If

        ' Find all household members
        Set householdMembers = xmlDoc.selectNodes("//ns4:HouseholdMember")
        
        Dim memberCounter
        memberCounter = 0
        Dim loopCompleted
        loopCompleted = False
        Dim firstNameEmpty
        firstNameEmpty = False
        Dim SegmentEmpty
        SegmentEmpty = True

        'MsgBox "HouseholdMembers " & householdMembers.length
        appl_membs = 0 'Incrementor for array of members

        For Each member In householdMembers
            ' Check if ns4:FirstName node exists
            'Set firstNameNode = member.selectSingleNode(".//ns4:FirstName")
            'If firstNameNode Is Nothing Then
            ' Check if ns4:FirstName node has no text
            '   If Trim(firstNameNode.text) = "" Then
            '       firstNameEmpty = True
            '       loopCompleted = True
            '       Exit For ' Exit the loop
            '   End If
            'End If
            SegmentEmpty = True
            
            'If appl_membs = 0 Then
                Set personalInfo = member.selectSingleNode("ns4:PersonalInfo")
            'Else
            '    Set personalInfo = member.selectSingleNode("ns4:PersonalInfo/ns4:Person")
            'End If
            
            ' Check if LastName node contains text
            Set lastNameNode = personalInfo.selectSingleNode("ns4:Person/ns4:LastName")
            If Not lastNameNode Is Nothing Then
                If Trim(lastNameNode.text) <> "" Then
                    SegmentEmpty = False
                End If
            End If
            
            ' Check if FirstName node contains text
            Set firstNameNode = personalInfo.selectSingleNode("ns4:Person/ns4:FirstName")
            If Not firstNameNode Is Nothing Then
                If Trim(firstNameNode.text) <> "" Then
                    SegmentEmpty = False
                End If
            End If
            
            ' Check if DOB node contains text
            Set dobNode = personalInfo.selectSingleNode("ns4:DOB")
            If Not dobNode Is Nothing Then
                If Trim(dobNode.text) <> "" Then
                    SegmentEmpty = False
                End If
            End If
            
            ' Check if SSN node contains text
            Set ssnNode = member.selectSingleNode("ns4:CitizenshipInfo/ns4:SSNInfo/ns4:SSN")
            If Not ssnNode Is Nothing Then
                If Trim(ssnNode.text) <> "" Then
                    SegmentEmpty = False
                End If
            End If

            ' If all relevant nodes are empty, consider the HouseholdMember occurrence as empty
            'MsgBox "Empty " & SegmentEmpty
            If SegmentEmpty Then
                loopCompleted = True
                Exit For ' Exit the loop
                MsgBox "HouseholdMember occurrence is empty."
            End If
            
            ' To Do: Add a routine to check the MEMB screen for if it is Mode:A or Mode:D
            ' If it is Mode:D, then we need to perform a PF9 to get it in edit mode so we can rewrite the values.
            ' This will most likely happen when the script has to be run again because of it matched an existing person and requires a manual match.
            
            'Set personalInfo = member.selectSingleNode(".//ns4:PersonalInfo")
            Set CitizenshipInfo = member.selectSingleNode("ns4:CitizenshipInfo")
            ' Extract data from PersonalInfo
            
            EMWriteScreen left(formattedDate, 3), 5, 70
            EMWriteScreen mid(formattedDate, 4, 3), 5, 73
            EMWriteScreen right(formattedDate, 4), 5, 76

            appl_relationship_text = personalInfo.selectSingleNode("ns4:Relationship").text
            ' If rel_code is empty we want to pop up a dialog box to supply the relationship code manually.
            If appl_relationship_text = "" Then
                ' This will work temporarily, but we will want a dialog box with all of the values in a dropdown.
                appl_relationship_text = InputBox("Please enter the Relationship Code for " & firstName & " " & lastName & ".", "Relationship Code")
            End If
            appl_relationship_text = LCase(appl_relationship_text)

            appl_gender_text = personalInfo.selectSingleNode("ns4:Gender").text
            appl_gender_text = LCase(appl_gender_text)

            'MsgBox "Rel " & Rel_code
            maxis_relationship_code = Get_MAXIS_Relationship_Code(appl_relationship_text, appl_gender_text)
            
            fmt_Ref_number = Get_MAXIS_Reference_Number(maxis_relationship_code)
            If Len(fmt_Ref_number) > 0 Then
                EMWriteScreen fmt_Ref_number, 4, 33
            End If
            
            lastName = lastNameNode.text 'personalInfo.selectSingleNode("ns4:LastName").text
            'lastName = member.getElementsByTagName("ns4:PersonalInfo").Item(0).getElementsByTagName("ns4:Person").Item(0).getElementsByTagName("ns4:LastName").Item(0).Text
            EMWritescreen lastName, 6, 30
            'MsgBox "last name" & lastName
            
            firstName = firstNameNode.text 'personalInfo.selectSingleNode("ns4:FirstName").text
            EMWritescreen firstName, 6, 63
            'MsgBox "first name" & firstName

            on error resume next
            If err.number <> 0 then
                MsgBox "error1 " & err.description
            End If
            
            SSN = ssnNode.text 'CitizenshipInfo.selectSingleNode("ns4:SSNInfo/ns4:SSN").text
			If Len(SSN) > 0 Then
            	EMWriteScreen left(SSN, 3), 7, 42
            	EMWriteScreen mid(SSN, 5, 2), 7, 46
            	EMWriteScreen right(SSN, 4), 7, 49

				EMWriteScreen "P", 7, 68
			Else
				EMWriteScreen "N", 7, 68
			End If	
            'MsgBox "SSN " & SSN			
            
            DOB = dobNode.text 'personalInfo.selectSingleNode("ns4:DOB").text
            EMWriteScreen left(DOB, 2), 8, 42
            EMWriteScreen mid(DOB, 4, 2), 8, 45
            EMWriteScreen right(DOB, 4), 8, 48
            'MsgBox "DOB " & DOB
            
            EMWriteScreen "NO", 8, 68
            EMWriteScreen "NO", 9, 68
            
            If appl_gender_text = "male" THEN
                EMWritescreen "M", 9, 42
            End If
            If appl_gender_text = "female" THEN
                EMWritescreen "F", 9, 42
            End If
            'MsgBox "Gender " & Gender
            
            marital_status = LCase(personalInfo.selectSingleNode("ns4:MaritalStatus").text)
            If marital_status = "never married" Then
                marital_status_code = "N"
            ElseIf marital_status = "married" Then
                marital_status_code = "M"
            ElseIf marital_status = "married living apart" Or marital_status = "separated" Then
                marital_status_code = "S"
            ElseIf marital_status = "legally sep" Then
                marital_status_code = "L"
            ElseIf marital_status = "divorced" Then
                marital_status_code = "D"
            ElseIf marital_status = "widowed" Then
                marital_status_code = "W"
            End If
            'MsgBox "Marital Status " & marital_status
            
            If Len(maxis_relationship_code) > 0 Then
                EMWriteScreen maxis_relationship_code, 10, 42    
            End If
                 
            Call Populate_Language_Code
            Call PopulatePreferredWrittenLanguage
            If NeedInterpreterInd = "True" THEN
                EMWritescreen "Y", 14, 68
            ELSE
                EMWritescreen "N", 14, 68
            End If

            'MsgBox "pause 2"
            Transmit
            
            EMReadScreen scrLine, 6, 2, 50
            If scrLine <> "(MTCH)" Then
                Call Error_Process
                StopScript
            End If
            
            PF8
            PF8
            PF5
            'MsgBox "pause"
            EMWriteScreen "Y", 6, 67
            EMReadScreen scrLine, 1, 6, 67
            If scrLine <> "Y" Then
                MsgBox "Member already exists in MAXIS. THE SCRIPT HAS STOPPED."
                StopScript
            End If
            'MsgBox "pause 3"
            Transmit
            ' Let's pull marital status above and populate it here.
            If Len(marital_status_code) > 0 Then
                EMWriteScreen marital_status_code, 7, 40
            End If
            Transmit
            'MsgBox "pause"
            ' Increment Reference number
            Ref_number = Ref_number + 1
            fmt_Ref_number = Right("00" & Ref_number, 2)
            
            ' Increment counter
            memberCounter = memberCounter + 1
            
            ' Check if the loop has completed
            If memberCounter = householdMembers.length Then
                loopCompleted = True
                'MsgBox "loop completed"
                Exit For ' Exit the loop
            End If

            appl_membs = appl_membs + 1
        Next
        If loopCompleted = True Then
            Call Address_Screen
        End If
    End If
End If

Function Populate_Language_Code
    If PrimarySpokenLanguage = "Amharic" Then
        EMWriteScreen "09", 12, 42
    End If

    If PrimarySpokenLanguage = "ASL" Then
        EMWriteScreen "08", 12, 42
    End If

    If PrimarySpokenLanguage = "Arabic" Then
        EMWriteScreen "10", 12, 42
    End If

    If PrimarySpokenLanguage = "Burmese" Then
        EMWriteScreen "14", 12, 42
    End If

    If PrimarySpokenLanguage = "Cantonese" Then
        EMWriteScreen "15", 12, 42
    End If

    If PrimarySpokenLanguage = "English" Then
        EMWriteScreen "99", 12, 42
    End If

    If PrimarySpokenLanguage = "French" Then
        EMWriteScreen "16", 12, 42
    End If

    If PrimarySpokenLanguage = "Hmong" Then
        EMWriteScreen "02", 12, 42
    End If

    If PrimarySpokenLanguage = "Khmer" Then
        EMWriteScreen "04", 12, 42
    End If

    If PrimarySpokenLanguage = "Korean" Then
        EMWriteScreen "20", 12, 42
    End If

    If PrimarySpokenLanguage = "Karen" Then
        EMWriteScreen "21", 12, 42
    End If

    If PrimarySpokenLanguage = "Laotian" Then
        EMWriteScreen "05", 12, 42
    End If

    If PrimarySpokenLanguage = "Mandarin" Then
        EMWriteScreen "17", 12, 42
    End If

    If PrimarySpokenLanguage = "Oromo" Then
        EMWriteScreen "12", 12, 42
    End If

    If PrimarySpokenLanguage = "Russian" Then
        EMWriteScreen "06", 12, 42
    End If

    If PrimarySpokenLanguage = "Serbo-Croatian" Then
        EMWriteScreen "11", 12, 42
    End If

	If PrimarySpokenLanguage = "Somali"  Then
	   EMWriteScreen "07", 12, 42
	End If

	If PrimarySpokenLanguage = "Spanish"  Then
	   EMWriteScreen "01", 12, 42
	End If

	If PrimarySpokenLanguage = "Swahili"  Then
	   EMWriteScreen "18", 12, 42
	End If

	If PrimarySpokenLanguage = "Tigrinya"  Then
	   EMWriteScreen "13", 12, 42
	End If

	If PrimarySpokenLanguage = "Vietnamese"  Then
	   EMWriteScreen "03", 12, 42
	End If

	If PrimarySpokenLanguage = "Yoruba"  Then
	   EMWriteScreen "19", 12, 42
	End If

	If PrimarySpokenLanguage = "Unknown"  Then
	   EMWriteScreen "97", 12, 42
	End If

	If PrimarySpokenLanguage = "Other"  Then
	   EMWriteScreen "98", 12, 42
	End If   
End Function  

Function PopulatePreferredWrittenLanguage
	If PreferredWrittenLanguage = "Amharic" Then
       EMWriteScreen "09", 13, 42
    End If

	If PreferredWrittenLanguage = "Arabic"  Then
	   EMWriteScreen "10", 13, 42
	End If

	If PreferredWrittenLanguage = "Burmese"  Then
	   EMWriteScreen "14", 13, 42
	End If

	If PreferredWrittenLanguage = "Cantonese"  Then
	   EMWriteScreen "15", 13, 42
	End If

	If PreferredWrittenLanguage = "English"  Then
	   EMWriteScreen "99", 13, 42
	End If

	If PreferredWrittenLanguage = "French"  Then
	   EMWriteScreen "16", 13, 42
	End If

	If PreferredWrittenLanguage = "Hmong"  Then
	   EMWriteScreen "02", 13, 42
	End If

	If PreferredWrittenLanguage = "Khmer"  Then
	   EMWriteScreen "04", 13, 42
	End If

	If PreferredWrittenLanguage = "Korean"  Then
	   EMWriteScreen "20", 13, 42
	End If

	If PreferredWrittenLanguage = "Karen"  Then
	   EMWriteScreen "21", 13, 42
	End If

	If PreferredWrittenLanguage = "Laotian"  Then
	   EMWriteScreen "05", 13, 42
	End If

	If PreferredWrittenLanguage = "Mandarin"  Then
	   EMWriteScreen "17", 13, 42
	End If

	If PreferredWrittenLanguage = "Oromo"  Then
	   EMWriteScreen "12", 13, 42
	End If

	If PreferredWrittenLanguage = "Russian"  Then
	   EMWriteScreen "06", 13, 42
	End If

	If PreferredWrittenLanguage = "Serbo-Croatian"  Then
	   EMWriteScreen "11", 13, 42
	End If

	If PreferredWrittenLanguage = "Somali"  Then
	   EMWriteScreen "07", 13, 42
	End If

	If PreferredWrittenLanguage = "Spanish"  Then
	   EMWriteScreen "01", 13, 42
	End If

	If PreferredWrittenLanguage = "Swahili"  Then
	   EMWriteScreen "18", 13, 42
	End If

	If PreferredWrittenLanguage = "Tigrinya"  Then
	   EMWriteScreen "13", 13, 42
	End If

	If PreferredWrittenLanguage = "Vietnamese"  Then
	   EMWriteScreen "03", 13, 42
	End If

	If PreferredWrittenLanguage = "Yoruba"  Then
	   EMWriteScreen "19", 13, 42
	End If

	If PreferredWrittenLanguage = "Unknown"  Then
	   EMWriteScreen "97", 13, 42
	End If

	If PreferredWrittenLanguage = "Other"  Then
	   EMWriteScreen "98", 13, 42
	End If   
End Function

Function Populate_County_Code

	If COR = "Aitkin" Then
       EMWriteScreen "01", 9, 66
    End If
	If COR = "Anoka" Then
       EMWriteScreen "02", 9, 66
    End If
	If COR = "Becker" Then
       EMWriteScreen "03", 9, 66
    End If
	If COR = "Beltrami" Then
       EMWriteScreen "04", 9, 66
    End If
	If COR = "Benton" Then
       EMWriteScreen "05", 9, 66
    End If
	If COR = "Big Stone" Then
       EMWriteScreen "06", 9, 66
    End If
	If COR = "Blue Earth" Then
       EMWriteScreen "07", 9, 66
    End If
	If COR = "Brown" Then
       EMWriteScreen "08", 9, 66
    End If
	If COR = "Carlton" Then
       EMWriteScreen "09", 9, 66
    End If
	If COR = "Carver" Then
       EMWriteScreen "10", 9, 66
    End If
	If COR = "Cass" Then
       EMWriteScreen "11", 9, 66
    End If
	If COR = "Chippewa" Then
       EMWriteScreen "12", 9, 66
    End If
	If COR = "Chisago" Then
       EMWriteScreen "13", 9, 66
    End If
	If COR = "Clay" Then
       EMWriteScreen "14", 9, 66
    End If
	If COR = "Clearwater" Then
       EMWriteScreen "15", 9, 66
    End If
	If COR = "Cook" Then
       EMWriteScreen "16", 9, 66
    End If
	If COR = "Cottonwood" Then
       EMWriteScreen "17", 9, 66
    End If
	If COR = "Crow Wing" Then
       EMWriteScreen "18", 9, 66
    End If
	If COR = "Dakota" Then
       EMWriteScreen "19", 9, 66
    End If
	If COR = "Dodge" Then
       EMWriteScreen "20", 9, 66
    End If
	If COR = "Douglas" Then
       EMWriteScreen "21", 9, 66
    End If
	If COR = "Fairbault" Then
       EMWriteScreen "22", 9, 66
    End If
	If COR = "Fillmore" Then
       EMWriteScreen "23", 9, 66
    End If
	If COR = "Freeborn" Then
       EMWriteScreen "24", 9, 66
    End If
	If COR = "Goodhue" Then
       EMWriteScreen "25", 9, 66
    End If
	If COR = "Grant" Then
       EMWriteScreen "26", 9, 66
    End If
	If COR = "Hennepin" Then
       EMWriteScreen "27", 9, 66
    End If
	If COR = "Houston" Then
       EMWriteScreen "28", 9, 66
    End If
	If COR = "Hubbard" Then
       EMWriteScreen "29", 9, 66
    End If
	If COR = "Isanti" Then
       EMWriteScreen "30", 9, 66
    End If
	If COR = "Itasca" Then
       EMWriteScreen "31", 9, 66
    End If
	If COR = "Jackson" Then
       EMWriteScreen "32", 9, 66
    End If
	If COR = "Kanabec" Then
       EMWriteScreen "33", 9, 66
    End If
	If COR = "Kandiyohi" Then
       EMWriteScreen "34", 9, 66
    End If
	If COR = "Kittson" Then
       EMWriteScreen "35", 9, 66
    End If
	If COR = "Koochiching" Then
       EMWriteScreen "36", 9, 66
    End If
	If COR = "Lac Qui Parle" Then
       EMWriteScreen "37", 9, 66
    End If
	If COR = "Lake" Then
       EMWriteScreen "38", 9, 66
    End If
	If COR = "Lake Of Woods" Then
       EMWriteScreen "39", 9, 66
    End If
	If COR = "Le Sueur" Then
       EMWriteScreen "40", 9, 66
    End If
	If COR = "Lincoln" Then
       EMWriteScreen "41", 9, 66
    End If
	If COR = "Lyon" Then
       EMWriteScreen "42", 9, 66
    End If
	If COR = "Mcleod" Then
       EMWriteScreen "43", 9, 66
    End If
	If COR = "Mahnomen" Then
       EMWriteScreen "44", 9, 66
    End If
	If COR = "Marshall" Then
       EMWriteScreen "45", 9, 66
    End If
	If COR = "Martin" Then
       EMWriteScreen "46", 9, 66
    End If
	If COR = "Meeker" Then
       EMWriteScreen "47", 9, 66
    End If
	If COR = "Mille Lacs" Then
       EMWriteScreen "48", 9, 66
    End If
	If COR = "Morrison" Then
       EMWriteScreen "49", 9, 66
    End If
	If COR = "Mower" Then
       EMWriteScreen "50", 9, 66
    End If
	If COR = "Murray" Then
       EMWriteScreen "51", 9, 66
    End If
	If COR = "Nicollet" Then
       EMWriteScreen "52", 9, 66
    End If
	If COR = "Nobles" Then
       EMWriteScreen "53", 9, 66
    End If
	If COR = "Norman" Then
       EMWriteScreen "54", 9, 66
    End If
	If COR = "Olmsted" Then
       EMWriteScreen "55", 9, 66
    End If
	If COR = "Otter Tail" Then
       EMWriteScreen "56", 9, 66
    End If
	If COR = "Pennington" Then
       EMWriteScreen "57", 9, 66
    End If
	If COR = "Pine" Then
       EMWriteScreen "58", 9, 66
    End If
	If COR = "Pipestone" Then
       EMWriteScreen "59", 9, 66
    End If
	If COR = "Polk" Then
       EMWriteScreen "60", 9, 66
    End If
	If COR = "Pope" Then
       EMWriteScreen "61", 9, 66
    End If
	If COR = "Ramsey" Then
       EMWriteScreen "62", 9, 66
    End If
	If COR = "Red Lake" Then
       EMWriteScreen "63", 9, 66
    End If
	If COR = "Redwood" Then
       EMWriteScreen "64", 9, 66
    End If
	If COR = "Renville" Then
       EMWriteScreen "65", 9, 66
    End If
	If COR = "Rice" Then
       EMWriteScreen "66", 9, 66
    End If
	If COR = "Rock" Then
       EMWriteScreen "67", 9, 66
    End If
	If COR = "Roseau" Then
       EMWriteScreen "68", 9, 66
    End If
	If COR = "St. Louis" Then
       EMWriteScreen "69", 9, 66
    End If
	If COR = "Scott" Then
       EMWriteScreen "70", 9, 66
    End If
	If COR = "Sherburne" Then
       EMWriteScreen "71", 9, 66
    End If
	If COR = "Sibley" Then
       EMWriteScreen "72", 9, 66
    End If
	If COR = "Stearns" Then
       EMWriteScreen "73", 9, 66
    End If
	If COR = "Steele" Then
       EMWriteScreen "74", 9, 66
    End If
	If COR = "Stevens" Then
       EMWriteScreen "75", 9, 66
    End If
	If COR = "Swift" Then
       EMWriteScreen "76", 9, 66
    End If
	If COR = "Todd" Then
       EMWriteScreen "77", 9, 66
    End If
	If COR = "Traverse" Then
       EMWriteScreen "78", 9, 66
    End If
	If COR = "Wabasha" Then
       EMWriteScreen "79", 9, 66
    End If
	If COR = "Wadena" Then
       EMWriteScreen "80", 9, 66
    End If
	If COR = "Waseca" Then
       EMWriteScreen "81", 9, 66
    End If
	If COR = "Washington" Then
       EMWriteScreen "82", 9, 66
    End If
	If COR = "Watonwan" Then
       EMWriteScreen "83", 9, 66
    End If
	If COR = "Wilkin" Then
       EMWriteScreen "84", 9, 66
    End If
	If COR = "Winona" Then
       EMWriteScreen "85", 9, 66
    End If
	If COR = "Wright" Then
       EMWriteScreen "86", 9, 66
    End If
	If COR = "Yellow Medicine" Then
       EMWriteScreen "87", 9, 66
    End If
	If COR = "Out-of-State" Then
       EMWriteScreen "89", 9, 66
    End If
		   
End Function  

Function Get_MAXIS_Relationship_Code(rel_code, gender)
    Dim return_code
    Select Case rel_code
        Case "applicant", "self"
            return_code = "01"     
        Case "spouse"
            return_code = "02"
        Case "child"
            return_code = "03"
        Case "parent"
            return_code = "04"
        Case "sibling", "brother or sister"
            return_code = "05"
        Case "step sibling"
            return_code = "06"
        Case "step child", "step-child"
            return_code = "08"
        Case "step parent"
            return_code = "09"
        Case "aunt"
            return_code = "10"
        Case "uncle"
            return_code = "11"
        Case "niece", "niece or nephew", "nephew or niece"
            return_code = "12"
            If gender = "male" Then
                return_code = "13"
            End If
        Case "nephew"
            return_code = "13"
        Case "cousin"
            return_code = "14"
        Case "grandparent"
            return_code = "15"
        Case "grandchild"
            return_code = "16"
        Case "other relative"
            return_code = "17"
        Case "legal guardian", "parent or guardian"
            return_code = "18"
        Case "live-in attendent"
            return_code = "25"
        Case "unknown on caf i", "other"
            return_code = "27"
    End Select

	Get_MAXIS_Relationship_Code = return_code
		
End Function

Function Address_Screen
    
	EMWritescreen "ADDR", 20, 71
	'MsgBox "pause"
	Transmit
    Transmit  	
	'MsgBox "pause"
	EMWriteScreen left(formattedDate, 3), 4, 43
	EMWriteScreen mid(formattedDate, 4, 3), 4, 46
	EMWriteScreen right(formattedDate, 2), 4, 49
	EMWritescreen Con_Line, 6, 43
	EMWritescreen Con_City, 8, 43
	EMWritescreen Con_State_code, 8, 66
	EMWritescreen left(Con_Zip5, 5), 9, 43
		
	EMWritescreen Mail_Line, 12, 49
	EMWritescreen Mail_City, 14, 49
	EMWritescreen Mail_State_code, 15, 49
	EMWritescreen left(Mail_Zip5, 5), 15, 58
		
	Call Populate_County_Code
	EMWritescreen "10", 11, 43
	EMWriteScreen "NO", 9, 74
	'MsgBox "Pause"
	Transmit
	EMReadScreen scrLine, 33, 3, 6
	If scrLine = "RESIDENCE ADDRESS IS STANDARDIZED" Then
	   PF3
	   PF3
	   MsgBox "Record saved"
	   StopScript
	End If   
	EMReadScreen scrLine, 33, 24, 02
	If scrLine <> "RESIDENCE ADDRESS IS STANDARDIZED" and scrLine <> "" Then
	   Call Error_Process 
	   StopScript
	End If
	PF3
	MsgBox "Record saved"	
End Function

Function Error_Process
    EMReadScreen scrLine, 80, 24, 1
	If scrLine <> "" Then
		MsgBox "ERROR: '" & trim(scrLine) & "'" & vbCrlf & vbCrlf & "   THE SCRIPT HAS STOPPED."
		StopScript
	End If   
End Function

' Function to get the MAXIS reference number based on the relationship code
Function Get_MAXIS_Reference_Number(maxis_relationship_code)
    Dim reference_number
    
    Select Case maxis_relationship_code
        Case "01" ' Self
            reference_number = "01"
        Case "02" ' Spouse
            reference_number = "02"
        Case "03", "08" ' Child, Step Child
            reference_number = Get_Reference_Number_Recursive(global_child_const, 15, global_other_family_const)
        Case "04", "05", "06", "09", "10", "11", "12", "13", "14", "15", "16", "17" ' Other family
            reference_number = Get_Reference_Number_Recursive(global_other_family_const, 26, global_unknown_const)
        Case "18", "27" ' Unknown
            reference_number = Get_Reference_Number_Recursive(global_unknown_const, 29, global_other_family_const)
        Case "25" ' Other non-family
            global_last_used_reference_number(global_other_non_family_const) = global_last_used_reference_number(global_other_non_family_const) + 1
            reference_number = Right("0" & global_last_used_reference_number(global_other_non_family_const), 2)
        Case Else
            reference_number = Get_Reference_Number_Recursive(global_unknown_const, 29, global_other_family_const) ' Default case if the relationship code is not recognized
    End Select

    Get_MAXIS_Reference_Number = reference_number
End Function

' Recursive helper function to handle overflow logic
Function Get_Reference_Number_Recursive(category, max_value, overflow_category)
    global_last_used_reference_number(category) = global_last_used_reference_number(category) + 1
    
    If category = global_other_family_const And global_last_used_reference_number(category) = 20 Then
        global_last_used_reference_number(category) = 21
    End If
    
    If global_last_used_reference_number(category) > max_value Then
        If category = global_child_const Then
            Get_Reference_Number_Recursive = Get_Reference_Number_Recursive(global_other_family_const, 26, global_unknown_const)
        ElseIf category = global_other_family_const Then
            Get_Reference_Number_Recursive = Get_Reference_Number_Recursive(global_unknown_const, 29, global_other_family_const)
        ElseIf category = global_unknown_const Then
            Get_Reference_Number_Recursive = Get_Reference_Number_Recursive(global_other_family_const, 26, global_unknown_const)
        End If
    Else
        Get_Reference_Number_Recursive = Right("0" & global_last_used_reference_number(category), 2)
    End If
End Function

Set xmlDoc = Nothing

script_end_procedure("")   
		

   


	
	