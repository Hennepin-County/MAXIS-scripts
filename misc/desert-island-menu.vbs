'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DESERT ISLAND MAIN MENU.vbs"
start_time = timer

desert_island_repository = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\"
on_the_desert_island = TRUE

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County"
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("04/01/2023", "Added ELIGIBILITY SUMMARY and HEALTH CARE DETERMINATION scripts. Removed other elgibilty noting and health care scripts.", "Ilse Ferris, Hennepin County")
call changelog_update("04/28/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

If git_hub_issue_known = FALSE Then
    email_body = "I accessed the Desert Island Scripts and it does not appear that you are aware the repository is unreachable." & vbCr & vbCr & "Today the script redirect sent me to the Desert Island Menu. There appeared to be a problem with GitHub." & vbCR & "https://www.githubstatus.com/" & vbCr & vbCr & "EMAIL sent from Desert Island Menu." & vbCr & vbCr & worker_signature
    Call create_outlook_email("", "HSPH.EWS.BlueZoneScripts@hennepin.us", "", "", "URGENT! - Reporting a Possible GitHub Issue", 1, False, "", "", False, "", email_body, False, "", True)
End If

Do
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 585, 350, "You are on the Desert Island - Essential Scripts during Outage"
      ButtonGroup ButtonPressed
	  	GroupBox 5, 5, 575, 45, "Script Repository Issue Occurrence - Also Known As The Desert Island"
      	IF git_hub_issue_known = TRUE Then
			Text 15, 20, 150, 10, "Oh no, you've ended up on the Desert Island!"
			Text 280, 20, 205, 10, "These are essential scripts - what you would need if deserted."
			Text 25, 30, 295, 20, "We are experiencing an outage of the repository system that stores and runs our script files. We are aware of this outage and keeping up with GitHub and the status of the repository."
            PushButton 165, 17, 110, 13, "More about the Desert Island", desert_island_info_btn
            PushButton 490, 15, 85, 13, "Check GitHub", check_git_hub_btn
            PushButton 315, 32, 95, 13, "Why does this happen?", why_btn
            PushButton 450, 32, 125, 13, "E-mail the BlueZone Script Team", email_script_team_btn
      	ELSE
			Text 15, 20, 150, 10, "Oh no, you've ended up on the Desert Island!"
			Text 280, 20, 205, 10, "These are essential scripts - what you would need if deserted."
			Text 45, 35, 410, 10, "The Power Pad had difficulty accessing the script files stored online in our script repository on GitHub.This is likely very temporary."
            PushButton 165, 17, 110, 13, "More about the Desert Island", desert_island_info_btn
            PushButton 490, 15, 85, 13, "Check GitHub", check_git_hub_btn
            PushButton 460, 32, 95, 13, "Why does this happen?", why_btn
      	END IF
      	  ButtonGroup ButtonPressed
            PushButton 460, 50, 115, 15, "NOTICES - Add WCOM", add_wcom_btn
            PushButton 10, 60, 115, 15, "Check EDRs", check_edrs_btn
            Text 135, 60, 300, 10, "Checks EDRS for HH members with disqualifications on a case."
            Text 10, 50, 440, 10, "----------------------------------------------------------------------------------------------------- ACTIONS -----------------------------------------------------------------------------------------------------"
            ButtonGroup ButtonPressed
              PushButton 10, 75, 115, 15, "Transfer Case", transfer_case_btn
            Text 135, 75, 445, 10, "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
            Text 10, 95, 500, 10, "------------------------------------------------------------------------------------------------------- NOTES -------------------------------------------------------------------------------------------------------"
            ButtonGroup ButtonPressed
              PushButton 10, 110, 115, 15, "Application Received", application_received_btn
            Text 135, 115, 445, 10, "Case notes an application, screens for expedited SNAP, sends the appointment letter and transfers case (if applicable)."
            ButtonGroup ButtonPressed
              PushButton 10, 125, 115, 15, "CAF", caf_btn
            Text 135, 130, 445, 10, "Document actions and processing when an interview has been completed on a CAF and STAT panels are updated."
            ButtonGroup ButtonPressed
              PushButton 10, 140, 115, 15, "Client Contact", client_contact_btn
            Text 135, 145, 445, 10, "Template for documenting client contact, either from or to a client."
            ButtonGroup ButtonPressed
              PushButton 10, 155, 115, 15, "CSR", csr_btn
            Text 135, 160, 445, 10, "Template for the Combined Six-month Report (CSR)."
            ButtonGroup ButtonPressed
              PushButton 10, 170, 115, 15, "Documents Received", documents_received_btn
            Text 135, 175, 445, 10, "Template for case noting information about documents received."
            ButtonGroup ButtonPressed
              PushButton 10, 185, 115, 15, "Eligibility Summary", elig_summary_btn
            Text 135, 190, 445, 10, "All-in-one case noting for approved, denied and/or closed programs."
            ButtonGroup ButtonPressed
              PushButton 10, 200, 115, 15, "Emergency", emergency_btn
            Text 135, 205, 445, 10, "Template for EA/EGA applications."
            ButtonGroup ButtonPressed
              PushButton 10, 215, 115, 15, "Expedited Determination", expedited_determination_btn
            Text 135, 220, 445, 10, "Work flow for assessing if a case meets Expedited SNAP Criteria"
            ButtonGroup ButtonPressed
              PushButton 10, 230, 115, 15, "Health Care Evaluation", health_care_btn
            Text 135, 235, 445, 10, "Template for Health Care applications and/or renewals."
            ButtonGroup ButtonPressed
              PushButton 10, 245, 115, 15, "HRF", hrf_btn
            Text 135, 250, 445, 10, "Template for HRFs (for GRH, use the ''GRH - HRF'' script)."
            ButtonGroup ButtonPressed
              PushButton 10, 260, 115, 15, "Interview", interview_btn
            Text 135, 265, 445, 10, "Workflow for a quality interview."
            ButtonGroup ButtonPressed
              PushButton 10, 275, 115, 15, "Verifications Needed", verifications_needed_btn
            Text 135, 280, 445, 10, "Template for when verifications are needed (enters each verification clearly)."
            ButtonGroup ButtonPressed
            CancelButton 530, 325, 50, 15
    EndDialog

    dialog Dialog1

    Select Case ButtonPressed
        Case 0
            stopscript
        Case check_git_hub_btn
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.githubstatus.com/"
        Case email_script_team_btn
            email_body = "I accessed the Desert Island Scripts." & vbCr & vbCr & "Today the script redirect sent me to the Desert Island Menu. There appeared to be a problem with GitHub." & vbCr & vbCr & worker_signature
            Call create_outlook_email("", "HSPH.EWS.BlueZoneScripts@hennepin.us", "", "", "Reporting a Possible GitHub Issue", 1, False, "", "", False, "", email_body, False, "", True)
        Case why_btn
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 446, 295, "Why did I end up on the Desert Island"
              ButtonGroup ButtonPressed
                OkButton 390, 275, 50, 15
              GroupBox 10, 10, 430, 95, "How does this happen?"
              Text 25, 25, 210, 10, "We want the script project to be quick, adaptive, and updatable."
              Text 25, 40, 370, 20, "The best way to make sure this happens is for us to keep our scripts in an online repository (or storage location). We use the repository on GitHub."
              Text 25, 65, 340, 20, "Since the repository is online, the scripts have to actually read the functionality from the online source. As we all know, sometimes accessing online sites doesn't work as well as we would always like."
              Text 25, 90, 305, 10, "You reach the Desert Island if the request to read the information times out, or takes too long."
              GroupBox 10, 110, 430, 90, "Why does this happen?"
              Text 25, 125, 240, 10, "This typically happens due to a slow-down somewhere on the internet."
              Text 25, 140, 290, 10, "Slow downs on the internet can happen anywhere or for any reason, some examples are:"
              Text 40, 150, 185, 10, "Too much traffic out from Hennepin County's network."
              Text 40, 160, 185, 10, "Too much traffic at the repository servers."
              Text 40, 170, 370, 10, "High bandwidth use from your location. (Someone streaming videos or playing video games are bandwidth hogs.)"
              Text 25, 185, 415, 10, "Sometimes GitHub has an outage that prevents us from accessing our scripts. When this happens the outage is much longer."
              GroupBox 10, 205, 430, 45, "How long will this last?"
              If git_hub_issue_known = FALSE Then
                  Text 25, 220, 420, 10, "Most of the time these outages are short (less than a few minutes) and limited to only a few users."
                  Text 25, 235, 210, 10, "If this is a GitHub Outage, it is typically resolved within an hour."
                  Text 10, 260, 435, 10, "The BlueZone Script Team has been notified via Email that you ended up on the Desert Island. We will keep an eye on the situation."
              Else
                  Text 25, 220, 420, 10, "This is a confirmed interruption of GitHub service. We cannot say for certain how long this will be."
                  Text 25, 235, 210, 10, "Typically GitHub outages are resolved within an hour."
                  Text 10, 260, 435, 10, "The BlueZone Script Team is aware of the outage and we are monitoring the outage."
              End If
            EndDialog

            dialog Dialog1
        Case desert_island_info_btn
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 356, 75, "What is the Desert Island"
              ButtonGroup ButtonPressed
                OkButton 300, 55, 50, 15
              Text 10, 10, 285, 10, "The Desert Island Menu is our safety net of essential scripts in the case of an outage."
              Text 10, 25, 355, 10, "It is the Desert Island because if your were stranded on a Desert Island, you would want your essentials."
              Text 10, 35, 325, 10, "Now you're stranded - temporarily - on a Script Desert Island. You only have your essential scripts. "
            EndDialog

            dialog Dialog1
        Case check_edrs_btn
            call run_another_script(desert_island_repository & "check-edrs.vbs")
        Case transfer_case_btn
            call run_another_script(desert_island_repository & "transfer-case.vbs")
        Case application_received_btn
            call run_another_script(desert_island_repository & "application-received.vbs")
        Case caf_btn
            call run_another_script(desert_island_repository & "caf.vbs")
        Case client_contact_btn
            call run_another_script(desert_island_repository & "client-contact.vbs")
        Case csr_btn
            call run_another_script(desert_island_repository & "csr.vbs")
        Case documents_received_btn
            call run_another_script(desert_island_repository & "documents-received.vbs")
        Case elig_summary_btn
            call run_another_script(desert_island_repository & "eligibility-summary.vbs")
        Case emergency_btn
			call run_another_script(desert_island_repository & "emergency.vbs")
		Case expedited_determination_btn
            call run_another_script(desert_island_repository & "expedited-determination.vbs")
        Case health_care_btn
            call run_another_script(desert_island_repository & "health-care-evaluation.vbs")
        Case hrf_btn
            call run_another_script(desert_island_repository & "hrf.vbs")
		Case interview_btn
            call run_another_script(desert_island_repository & "interview.vbs")  
        Case verifications_needed_btn
            call run_another_script(desert_island_repository & "verifications-needed.vbs")
        Case add_wcom_btn
            call run_another_script(desert_island_repository & "add-wcom.vbs")
    End Select
Loop