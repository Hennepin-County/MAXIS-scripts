'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DESERT ISLAND MAIN MENU.vbs"
start_time = timer
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
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
call changelog_update("04/28/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

desert_island_respository = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\"
on_the_desert_island = TRUE

If git_hub_issue_known = FALSE Then
    email_body = "I accessed the Desert Island Scripts and it does not appear that you are aware the respository is unreachable." & vbCr & vbCr & "Today the script redirect sent me to the Desert Island Menu. There appeared to be a problem with GitHub." & vbCR & "https://www.githubstatus.com/" & vbCr & vbCr & "EMAIL sent from Desert Island Menu." & vbCr & vbCr & worker_signature
    Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", "URGENT! - Reporting a Possible GitHub Issue", email_body, "", TRUE)
End If

Do

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 585, 375, "You are on the Desert Island - Essential Scripts during Outage"
      ButtonGroup ButtonPressed
	  	GroupBox 5, 5, 575, 45, "Script Repository Issue Occurance - Also Known As The Desert Island"
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
			Text 45, 35, 410, 10, "The Power Pad had difficulty accessing the script files stored online in our script reporistory on GitHub.This is likely very temporary."
            PushButton 165, 17, 110, 13, "More about the Desert Island", desert_island_info_btn
            PushButton 490, 15, 85, 13, "Check GitHub", check_git_hub_btn
            PushButton 460, 32, 95, 13, "Why does this happen?", why_btn
      	END IF
      	Text 10, 50, 500, 10, "----------------------------------------------------------------------------------------------------- ACTIONS -----------------------------------------------------------------------------------------------------"
        PushButton 10, 60, 115, 13, "Check EDRs", check_edrs_btn
      	Text 135, 63, 445, 10, "Checks EDRS for HH members with disqualifications on a case."
		PushButton 460, 53, 115, 15, "NOTICES - Add WCOM", add_wcom_btn
        PushButton 10, 75, 115, 13, "Transfer Case", transfer_case_btn
      	Text 135, 78, 445, 10, "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
      	Text 10, 95, 500, 10, "------------------------------------------------------------------------------------------------------- NOTES -------------------------------------------------------------------------------------------------------"
        PushButton 10, 105, 115, 13, "Application Check", application_check_btn
      	Text 135, 108, 445, 10, "Template for documenting details and tracking pending cases."
        PushButton 10, 120, 115, 13, "Approved Programs", approved_programs_btn
      	Text 135, 123, 445, 10, "Template for when you approve a client's programs."
        PushButton 10, 135, 115, 13, "CAF", caf_btn
      	Text 135, 138, 445, 10, "Document actions and processing when an interview has been completed on a CAF and STAT panels are updated."
        PushButton 10, 150, 115, 13, "Client Contact", client_contact_btn
      	Text 135, 153, 445, 10, "Template for documenting client contact, either from or to a client."
        PushButton 10, 165, 115, 13, "Closed Programs", closed_programs_btn
      	Text 135, 168, 445, 10, "Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
        PushButton 10, 180, 115, 13, "CSR", csr_btn
      	Text 135, 183, 445, 10, "Template for the Combined Six-month Report (CSR)."
        PushButton 10, 195, 115, 13, "Denied Programs", denied_programs_btn
      	Text 135, 198, 445, 10, "Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
        PushButton 10, 210, 115, 13, "Documents Received", documents_received_btn
      	Text 135, 213, 445, 10, "Template for case noting information about documents received."
        PushButton 10, 225, 115, 13, "Emergency", emergency_btn
      	Text 135, 228, 445, 10, "Template for EA/EGA applications."
		PushButton 10, 240, 115, 13, "Expedtied Determination", expedited_determination_btn
      	Text 135, 243, 445, 10, "Workflow for assessing if a case meets Expedited SNAP Criteria"
        PushButton 10, 255, 115, 13, "HC Renewal", hc_renewal_btn
      	Text 135, 258, 445, 10, "Template for HC renewals."
        PushButton 10, 270, 115, 13, "HCAPP", hcapp_btn
      	Text 135, 273, 445, 10, "Template for HCAPPs."
        PushButton 10, 285, 115, 13, "HRF", hrf_btn
      	Text 135, 288, 445, 10, "Template for HRFs (for GRH, use the ''GRH - HRF'' script)."
        PushButton 10, 300, 115, 13, "Interview", interview_btn
      	Text 135, 303, 445, 10, "Workflow for a quality interview."
        PushButton 10, 315, 115, 13, "LTC - Renewal", ltc_renewal_btn
      	Text 135, 318, 445, 10, "Template for LTC renewals."
        PushButton 10, 330, 115, 13, "Verifications Needed", verifications_needed_btn
		Text 135, 333, 445, 10, "Template for when verifications are needed (enters each verification clearly)."
		Text 15, 350, 500, 10, "--------------------------------------------------------------------------------------------------------- CASE ASSIGNMENT ---------------------------------------------------------------------------------------------------------"
        PushButton 10, 360, 115, 13, "Application Received", application_received_btn
      	Text 135, 363, 445, 10, "Case notes an application, screens for expedited SNAP, sends the appointment letter and transfers case (if applicable)."
        CancelButton 530, 357, 50, 15
    EndDialog



    dialog Dialog1

    Select Case ButtonPressed
        Case 0
            stopscript
        Case check_git_hub_btn
            run "C:\Program Files\Internet Explorer\iexplore.exe https://www.githubstatus.com/"		'Goes to SIR if button is pressed
        Case email_script_team_btn
            email_body = "I accessed the Desert Island Scripts." & vbCr & vbCr & "Today the script redirect sent me to the Desert Island Menu. There appeared to be a problem with GitHub." & vbCr & vbCr & worker_signature
            Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", "Reporting a Possible GitHub Issue", email_body, "", TRUE)
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
            call run_another_script(desert_island_respository & "check-edrs.vbs")
        Case transfer_case_btn
            call run_another_script(desert_island_respository & "transfer-case.vbs")
        Case application_check_btn
            call run_another_script(desert_island_respository & "application-check.vbs")
        Case approved_programs_btn
            call run_another_script(desert_island_respository & "approved-programs.vbs")
        Case caf_btn
            call run_another_script(desert_island_respository & "caf.vbs")
        Case client_contact_btn
            call run_another_script(desert_island_respository & "client-contact.vbs")
        Case closed_programs_btn
            call run_another_script(desert_island_respository & "closed-programs.vbs")
        Case csr_btn
            call run_another_script(desert_island_respository & "csr.vbs")
        Case denied_programs_btn
            call run_another_script(desert_island_respository & "denied-programs.vbs")
        Case documents_received_btn
            call run_another_script(desert_island_respository & "documents-received.vbs")
        Case emergency_btn
			call run_another_script(desert_island_respository & "emergency.vbs")
		Case expedited_determination_btn
            call run_another_script(desert_island_respository & "expedited-determination.vbs")
        Case hc_renewal_btn
            call run_another_script(desert_island_respository & "hc-renewal.vbs")
        Case hcapp_btn
            call run_another_script(desert_island_respository & "hcapp.vbs")
        Case hrf_btn
            call run_another_script(desert_island_respository & "hrf.vbs")
		Case interview_btn
            call run_another_script(desert_island_respository & "interview.vbs")
        Case ltc_renewal_btn
            call run_another_script(desert_island_respository & "ltc-renewal.vbs")
        Case verifications_needed_btn
            call run_another_script(desert_island_respository & "verifications-needed.vbs")
        Case application_received_btn
            call run_another_script(desert_island_respository & "application-received.vbs")
        Case add_wcom_btn
            call run_another_script(desert_island_respository & "add-wcom.vbs")
    End Select
Loop
