'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
Call changelog_update("09/27/2024", "Fixed an isssue with identifying case status when a second cash program is pending. New functionality will be more reliable in these situations.##~##", "Casey Love, Hennepin County.")
Call changelog_update("08/08/2024", "Update to the CA Transfer process to transfer GRH/HS cases less often to maintain the caseload structure the GRH team uses. Additionally adds a separation of adult vs family GRH cases.", "Casey Love, Hennepin County.")
Call changelog_update("05/23/2024", "Added contracted caseload selection for HCMC and North Memorial.", "Casey Love, Hennepin County.")
Call changelog_update("04/29/2024", "Enhanced SPEC/XFER reminder when ransferring cases in ECF Next prior to transferring the case in MAXIS.", "Ilse Ferris, Hennepin County.")
Call changelog_update("04/27/2024", "Added reminder option prior to SPEC/XFER about transferring cases in ECF Next prior to transferring the case in MAXIS.", "Ilse Ferris, Hennepin County.")
call changelog_update("03/25/2024", "Update to alight with Caseload Assignment and Transfer Process updates. This functionality supports a large number of Caseloads and reduces the transfering of cases within the county. There is also support to reduce the number of caseloads with PND2 display limits.", "Casey Love, Hennepin County")
call changelog_update("01/22/2023", "BUG FIX - the script would error anytime an apostrophe (') was in the case name when loading into a data table at the very end. This fix will resolve this error by substituting a dash in place of the apostrophe.", "Casey Love, Hennepin County")
call changelog_update("09/22/2023", "Updated format of appointment notice and digital experience in SPEC/MEMO", "Megan Geissler, Hennepin County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
CALL changelog_update("04/24/2023", "Changed the CASE/NOTE for the Expedited Screening to a standard number format and align for easier viewing.", "Casey Love, Hennepin County")
call changelog_update("03/22/2023", "Updated form names and simplified selections for how an application is received by the Case Assignment team. Updated email verbiage on a response to the 'Request for APPL' form. These updates are meant to align the script to official language and information.", "Casey Love, Hennepin County")
call changelog_update("03/21/2023", "Removed the functionality to e-mail the CCAP team if CCAP was requested with other programs on MNbenefits. This process is now supported in ECF Next and the manual e-mail process is no longer required.", "Casey Love, Hennepin County")
call changelog_update("02/23/2023", "BUG FIX for cases with a second application to better determine which application is for HC and which is for CAF Based Programs.", "Casey Love, Hennepin County")
call changelog_update("02/21/2023", "BUG FIX for cases with a second application that is for HC. These cases are not subsequent applications, and need to be handled within this script and not duplicate MEMOs and Screenings.", "Casey Love, Hennepin County")
call changelog_update("01/30/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("01/30/2023", "Script will now redirect to new script (NOTES - Subsequent Application) for cases that are already pending with a new application form received. This new script supports case actions needed for subsequent applications received.", "Casey Love, Hennepin County")
CALL changelog_update("09/12/2022", "Updated EBT card availibilty in the office direction for expedited cases.", "Ilse Ferris, Hennepin County")
call changelog_update("05/24/2022", "CASE/NOTE format updated to exclude the 'How App Received' detail. This information is important for the script operation, but is not necessary to be included in the CASE/NOTE", "Casey Love, Hennepin County")   '#799
call changelog_update("05/01/2022", "Updated the Appointment Notice to have information for residents about in person support.", "Casey Love, Hennepin County")
call changelog_update("04/27/2022", "The Application Received script is updated to check cases to find if the ADDR panel is missing or has an error. The script will stop if it discovers a possible issue with an ADDR panel as that is a mandatory panel for all cases.", "Casey Love, Hennepin County")
call changelog_update("03/29/2022", "Removed APPLYMN as application option.", "Ilse Ferris, Hennepin County")
call changelog_update("03/11/2022", "Added randomizer functionality for Adults appplications that appear expedited. Caseloads suggested will be either EX1 or EX2", "Ilse Ferris, Hennepin County")
call changelog_update("03/07/2022", "Updated METS retro contact from Team 601 to Team 603.", "Ilse Ferris, Hennepin County")
call changelog_update("1/6/2022", "The script no longer allows you to change the Appointment Notice date if one is required based on the pending programs. This change is to ensure compliance with notification requirements of the On Demand Waiver.", "Casey Love, Hennepin County")
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
call changelog_update("09/29/2021", "Added functionality to determine HEST utility allowances based on application date. ", "Ilse Ferris, Hennepin County")
call changelog_update("09/17/2021", "Removed the field for 'Requested by X#' in the 'Request to APPL' option as this information will no longer be in the CASE/NOTE as this information is not pertinent to case actions/decisions.##~##", "Casey Love, Hennepin County")
call changelog_update("09/10/2021", "This is a very large update for the script.##~## ##~##We have reordered the functionality and consolidated the dialogs to have fewer interruptions in the process and to support the natural order of completing a pending update.##~##", "Casey Love, Hennepin County")
call changelog_update("08/03/2021", "GitHub Issue #547, added Mail as an option for how an application can be received.", "MiKayla Handley, Hennepin County")
call changelog_update("08/01/2021", "Changed the notices sent in 2 ways:##~## ##~## - Updated verbiage on how to submit documents to Hennepin.##~## ##~## - Appointment Notices will now be sent with a date of 5 days from the date of application.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("01/29/2021", "Updated Request for APPL handling per case assignement request. Issue #322", "MiKayla Handley, Hennepin County")
call changelog_update("01/07/2021", "Updated worker signature as a mandatory field.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Temporarily removed in person option for how applications are received.", "MiKayla Handley, Hennepin County")
call changelog_update("12/18/2020", "Update to make confirmation number mandatory for MN Benefit application.", "MiKayla Handley, Hennepin County")
call changelog_update("11/15/2020", "Updated droplist to add virtual drop box option to how the application was received.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/13/2020", "Enhanced date evaluation functionality when which determining HEST standards to use.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/24/2020", "Added Mnbenefits application and removed SHIBA and apply MN options.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/01/2020", "Updated Standard Utility Allowances for 10/2020.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/24/2020", "Added SHIBA application and combined CA and NOTES scripts.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/10/2020", "Email functionality removed for Triagers.", "MiKayla Handley, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2020", "Update to the notice wording. Information and direction for in-person interview option removed. County offices are not currently open due to the COVID-19 Peacetime Emergency.", "Casey Love, Hennepin County")
call changelog_update("03/09/2020", "Per project request- Updated checkbox for the METS Retro Request to team 601.", "MiKayla Handley, Hennepin County")
call changelog_update("01/13/2020", "Updated requesting worker for the Request to APPL form process.", "MiKayla Handley, Hennepin County")
call changelog_update("11/04/2019", "New version pulled to support the request for APPL process.", "MiKayla Handley, Hennepin County")
call changelog_update("10/01/2019", "Updated the utility standards for SNAP.", "Casey Love, Hennepin County")
call changelog_update("08/27/2019", "Added handling to push the case into background to ensure pending programs are read.", "MiKayla Handley, Hennepin County")
call changelog_update("08/27/2019", "Added GRH to appointment letter handling for future enhancements.", "MiKayla Handley, Hennepin County")
call changelog_update("08/20/2019", "Bug on the script when a large PND2 list is accessed.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/26/2019", "Reverted the script to not email Team 603 for METS cases. CA workers will need to manually complete the email to: field.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/24/2019", "Removed Mail & Fax option and added MDQ per request.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/22/2019", "Updated the script to automatically email Team 603 for METS cases.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
CALL changelog_update("02/05/2019", "Updated case correction handling.", "Casey Love, Hennepin County")
CALL changelog_update("11/15/2018", "Enhanced functionality for SameDay interview cases.", "Casey Love, Hennepin County")
CALL changelog_update("11/06/2018", "Updated handling for HC only applications.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/25/2018", "Updated script to add handling for case correction.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/17/2018", "Updated appointment letter to address EGA programs.", "MiKayla Handley, Hennepin County")
CALL changelog_update("09/01/2018", "Updated Utility standards that go into effect for 10/01/2018.", "Ilse Ferris, Hennepin County")
CALL changelog_update("07/20/2018", "Changed wording of the Appointment Notice and changed default interview date to 10 days from application for non-expedidted cases.", "Casey Love, Hennepin County")
CALL changelog_update("07/16/2018", "Bug Fix that was preventing notices from being sent.", "Casey Love, Hennepin County")
CALL changelog_update("03/28/2018", "Updated appt letter case note for bulk script process.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/21/2018", "Added on demand waiver handling.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/16/2018", "Added case transfer confirmation coding.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/03/2017", "Email functionality - only expedited emails will be sent to Triagers.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/25/2017", "Email functionality - will create email, and send for all CASH and FS applications.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/12/2017", "Email functionality will create email, but not send it. Staff will need to send email after reviewing email.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
random_team_needed = ""
set caseload_info = CreateObject("Scripting.Dictionary")

caseload_info.add "X127FA5", "YET"
' caseload_info.add "X127F3P", "Families - General"		- this is MAEPD
'Team 1 Clifton 
caseload_info.add "X127EK8", "Adults - Pending 1"
caseload_info.add "X127EH1", "Adults - Pending 1"
caseload_info.add "X127EP1", "Adults - Pending 1"
caseload_info.add "X127EZ6", "Families - Pending 1"
caseload_info.add "X127EZ8", "Families - Pending 1"
'Active casebanks for Clifton
caseload_info.add "X127EQ5", "Adults Active 1" 
caseload_info.add "X127EQ6", "Adults Active 1"
caseload_info.add "X127EQ7", "Adults Active 1"
caseload_info.add "X127EQ8", "Adults Active 1"
caseload_info.add "X127EX1", "Adults Active 1"
caseload_info.add "X127EX2", "Adults Active 1"
caseload_info.add "X127EX3", "Adults Active 1"
caseload_info.add "X127EX4", "Adults Active 1"
caseload_info.add "X127EX5", "Adults Active 1"
caseload_info.add "X127EX7", "Adults Active 1"
'caseload_info.add "X127F3H", "Adults Active 1"
caseload_info.add "X127ET5", "Families Active 1"
caseload_info.add "X127ET6", "Families Active 1"
caseload_info.add "X127ET7", "Families Active 1"
caseload_info.add "X127ET8", "Families Active 1"
caseload_info.add "X127ET9", "Families Active 1"
caseload_info.add "X127EZ1", "Families Active 1"
'Team 2 Coenen
caseload_info.add "X127EP2", "Adults - Pending 2"
caseload_info.add "X127EH8", "Adults - Pending 2"
caseload_info.add "X127EP6", "Adults - Pending 2"
caseload_info.add "X127EZ9", "Families - Pending 2"
caseload_info.add "X127EH4", "Families - Pending 2"
'Active casebanks for Coenen
caseload_info.add "X127EL7", "Adults Active 2"
caseload_info.add "X127EL8", "Adults Active 2"
caseload_info.add "X127EL9", "Adults Active 2"
caseload_info.add "X127EN1", "Adults Active 2"
caseload_info.add "X127EN2", "Adults Active 2"
caseload_info.add "X127EN3", "Adults Active 2"
caseload_info.add "X127EN5", "Adults Active 2"
caseload_info.add "X127EN4", "Adults Active 2"
caseload_info.add "X127EN7", "Adults Active 2"
caseload_info.add "X127ES1", "Families Active 2"
caseload_info.add "X127ES2", "Families Active 2"
caseload_info.add "X127ET1", "Families Active 2"
'caseload_info.add "X127F4E", "Families Active 2"
caseload_info.add "X127EZ7", "Families Active 2"
caseload_info.add "X127FB7", "Families Active 2"
'Team 3 Garrett
caseload_info.add "X127EP7", "Adults - Pending 3"
caseload_info.add "X127EP8", "Adults - Pending 3"
caseload_info.add "X127EP3", "Adults - Pending 3"
caseload_info.add "X127EH5", "Families - Pending 3"
caseload_info.add "X127EH6", "Families - Pending 3"
'Active casebanks for Garrett
caseload_info.add "X127EN8", "Adults Active 3"
caseload_info.add "X127EN9", "Adults Active 3"
caseload_info.add "X127EQ1", "Adults Active 3"
caseload_info.add "X127EQ2", "Adults Active 3"
caseload_info.add "X127EQ3", "Adults Active 3"
caseload_info.add "X127EQ4", "Adults Active 3"
caseload_info.add "X127EX8", "Adults Active 3"
caseload_info.add "X127EX9", "Adults Active 3"
caseload_info.add "X127EG4", "Adults Active 3"
caseload_info.add "X127ET2", "Families Active 3"
caseload_info.add "X127ET3", "Families Active 3"
caseload_info.add "X127ET4", "Families Active 3"
'Team 4 Groves
caseload_info.add "X127EH7", "Adults - Pending 4"
caseload_info.add "X127EK3", "Adults - Pending 4"
caseload_info.add "X127EK7", "Adults - Pending 4"
caseload_info.add "X127EZ3", "Families - Pending 4"
caseload_info.add "X127EZ4", "Families - Pending 4"
'Active casebanks for Groves, not currently utilized for assignment

caseload_info.add "X127EE1", "Adults Active 4"
caseload_info.add "X127EE2", "Adults Active 4"
caseload_info.add "X127EE3", "Adults Active 4"
caseload_info.add "X127EE4", "Adults Active 4"
caseload_info.add "X127EE5", "Adults Active 4"
caseload_info.add "X127EE6", "Adults Active 4"
caseload_info.add "X127EE7", "Adults Active 4"
caseload_info.add "X127EL1", "Adults Active 4"
caseload_info.add "X127EL2", "Adults Active 4"
caseload_info.add "X127EL3", "Adults Active 4"
caseload_info.add "X127EL4", "Adults Active 4"
caseload_info.add "X127EL5", "Adults Active 4"
caseload_info.add "X127EL6", "Adults Active 4"
caseload_info.add "X127ES3", "Families Active 4"
caseload_info.add "X127ES4", "Families Active 4"
caseload_info.add "X127ES5", "Families Active 4"
caseload_info.add "X127ES6", "Families Active 4"
caseload_info.add "X127ES7", "Families Active 4"
caseload_info.add "X127ES8", "Families Active 4"
caseload_info.add "X127ES9", "Families Active 4"
'Healthcare Pending caseloads
caseload_info.add "X127ED8", "Healthcare - Pending"
caseload_info.add "X127ER1", "Healthcare - Pending"
caseload_info.add "X127ER2", "Healthcare Only Active"
caseload_info.add "X127ER3", "Healthcare Only Active"
caseload_info.add "X127ER4", "Healthcare Mixed Active"
caseload_info.add "X127ER5", "Healthcare Mixed Active"
caseload_info.add "X127ER6", "Healthcare Mixed Active"
'This is the casebank for DWP team
caseload_info.add "X127EY9", "Families - Cash"
' caseload_info.add "X127EY8", "Families - Cash"		removed from assignment selection until additional process clarification can be identified. There are concerns with all cases being entered into a single basket with pending status.
caseload_info.add "X127EN6", "TEFRA"
caseload_info.add "X127FG1", "Foster Care / IV-E"
caseload_info.add "X127EW6", "Foster Care / IV-E"
caseload_info.add "X1274EC", "Foster Care / IV-E"
caseload_info.add "X127FG2", "Foster Care / IV-E"
caseload_info.add "X127EW4", "Foster Care / IV-E"

caseload_info.add "X127EM8", "GRH / HS - Adults Pending"
caseload_info.add "X127FE6", "GRH / HS - Adults Pending"
caseload_info.add "X127EZ2", "GRH / HS - Families Pending"
caseload_info.add "X127EM2", "GRH / HS - Maintenance"
caseload_info.add "X127EH9", "GRH / HS - Maintenance"
'caseload_info.add "X127FE6", "GRH / HS - Maintenance" This bank is being changed to accept pending, but cases already here for maintenance will not transfer.
caseload_info.add "X127EJ4", "GRH / HS - Maintenance"
caseload_info.add "X127EH2", "GRH / HS - Maintenance"
caseload_info.add "X127EP4", "GRH / HS - Maintenance"
caseload_info.add "X127EK5", "GRH / HS - Maintenance"
caseload_info.add "X127EG5", "GRH / HS - Maintenance"
'caseload_info.add "X127EG4", "MIPPA"
caseload_info.add "X127F3D", "MA - BC"
caseload_info.add "X127EK4", "LTC+ - General"
caseload_info.add "X127EK9", "LTC+ - General"
caseload_info.add "X127EF8", "1800 - Team 160"
caseload_info.add "X127EF9", "1800 - Team 160"
caseload_info.add "X1275H5", "Privileged Cases"
caseload_info.add "X127FAT", "Privileged Cases"
caseload_info.add "X127F3H", "Privileged Cases"
caseload_info.add "X127FG7", "Contracted - Monarch Facilities Contract"
caseload_info.add "X127EM4", "Contracted - A Villa Facilities Contract"
caseload_info.add "X127EW8", "Contracted - Ebenezer Care Center/ Martin Luther Care Center"

caseload_info.add "X127FF8", "Contracted - North Memorial"
caseload_info.add "X127FF6", "Contracted - HCMC"
caseload_info.add "X127FF7", "Contracted - HCMC"

caseload_info.add "X127FI1", "METS Retro Request"

' MsgBox "The caseload type of Families - General is " & join(caseload_info.item("Families - General"), ", ")
' MsgBox "The caseload type of Families - General is ~" & caseload_info.item("Families - General") & "~"
function select_random_index(ubound_of_array, index_selection)
	If ubound_of_array = 0 Then
		index_selection = 0
	Else
		options = ubound_of_array + 1
		Randomize       'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
		rnd_nbr = rnd
		size_up = rnd_nbr * options
		index_selection = int(size_up)
	End If
end function

' For i = 1 to 10
' 	call select_random_index(9, select_out_of_10)
' 	call select_random_index(4, select_out_of_5)
' 	call select_random_index(19, select_out_of_20)
' 	call select_random_index(99, select_out_of_100)
' 	call select_random_index(1, select_out_of_2)
' 	call select_random_index(34, select_out_of_35)
' 	call select_random_index(16, select_out_of_17)
' 	MsgBox "Select out of 10:  " & select_out_of_10 & vbCr &_
' 			"Select out of 5: " & select_out_of_5 & vbCr &_
' 			"Select out of 20: " & select_out_of_20 & vbCr &_
' 			"Select out of 100: " & select_out_of_100 & vbCr &_
' 			"Select out of 2: " & select_out_of_2 & vbCr &_
' 			"Select out of 35: " & select_out_of_35 & vbCr &_
' 			"Select out of 17: " & select_out_of_17
' Next

' MsgBox "That's all"

' MsgBox "The caseload type of X127EX9 is " & caseload_info.item("X127EX9")

' Call get_caseload_array_by_type("Families - General", test_array)

' MsgBox join(test_array, ", ")

function get_caseload_array_by_type(caseload_type, available_caseload_array)
	all = caseload_info.items
	things = caseload_info.keys

	Dim temp_array()
	ReDim temp_array(0)
	counter = 0

	for i = 0 to UBound(all) - 1
        If right(caseload_type, 1) = "?" Then random_team_needed = True 'failsafe to ensure that the random team is selected when something slips through the cracks
        If random_team_needed = TRUE Then  'This will be used to randomly select a PM team for the case to be transferred to for cash/snap pending caseload
            team_to_check = all(i) & ""
            If left(team_to_check, len(team_to_check)-2) = left(caseload_type, len(caseload_type) - 2) Then
			    ReDim preserve temp_array(counter)
			    temp_array(counter) = things(i)
			    counter = counter + 1
		    End If
        Else
		    If all(i) = caseload_type Then
		    	ReDim preserve temp_array(counter)
		    	temp_array(counter) = things(i)
		    	counter = counter + 1
		    End If
        End If 
	Next
	available_caseload_array = temp_array
end function

function gather_current_caseload(current_caseload, secondary_caseload, find_previous_pw, previous_pw)
	call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen current_caseload, 7, 21, 14
	EMReadScreen secondary_caseload, 7, 21, 26
	current_caseload = trim(current_caseload)
	secondary_caseload = trim(secondary_caseload)
	If find_previous_pw = True Then
		call navigate_to_MAXIS_screen("SPEC", "XFER")
		call write_value_and_transmit("X", 5, 16)
		EMReadScreen previous_pw, 7, 18, 28
		PF3
		call navigate_to_MAXIS_screen("CASE", "CURR")
	End If
end function

function find_correct_caseload(current_caseload, secondary_caseload, user_x_number, previous_pw, transfer_needed, correct_caseload_type, new_caseload, application_form, appears_ltc_checkbox, METS_retro_checkbox, caseload_contract_facility, case_has_child_under_19, case_has_guardian, age_of_memb_01, case_has_child_under_22, preg_person_on_case, addr_on_1800_faci_list, case_name, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type)
	transfer_needed = True

	current_caseload_type = caseload_info.item(current_caseload)		'this could be blank if the basket is not identified in the dictionary

	'TODO - handle for cases that were automatically transferred by MAXIS when APPLd to see if they should be transferred back to wthe INAC basket
	'QUESTION - do we care how long ago that was?
	pended_from_inactive = False
	If current_caseload = user_x_number Then pended_from_inactive = True

	If current_caseload_type = "Privileged Cases" or current_caseload_type = "Foster Care / IV-E" or current_caseload_type = "1800 - Team 160" or left(current_caseload_type, 10) = "Contracted" Then
		transfer_needed = False
		correct_caseload_type = current_caseload_type
		new_caseload = current_caseload
	End If

	'   DropListBox 85, 80, 170, 45, "Select One:"+chr(9)+"CAF - 5223"+chr(9)+"MNbenefits CAF - 5223"+chr(9)+
	'   "SNAP App for Seniors - 5223F"+chr(9)+
	'   "MNsure App for HC - 6696"+chr(9)+
	'   "MHCP App for Certain Populations - 3876"+chr(9)+
	'   "App for MA for LTC - 3531"+chr(9)+
	'   "MHCP App for B/C Cancer - 3523"+chr(9)+
	'   "EA/EGA Application"+chr(9)+
	'   "No Application Required", application_type


	alpha_split_one_a_l = "ABCDEFGHIJKL"
	alpha_split_two_m_z = "MNOPQRSTUVWXYZ"


	If transfer_needed = True Then
		If application_form = "MHCP App for B/C Cancer - 3523" Then
			correct_caseload_type = "MA - BC"
			If correct_caseload_type = current_caseload_type Then transfer_needed = False
		End If
		If application_form = "No Application Required" and METS_retro_checkbox = checked Then
			correct_caseload_type = "METS Retro Request"
			If correct_caseload_type = current_caseload_type Then transfer_needed = False
		End If
	End If

	If correct_caseload_type = "" Then
		If caseload_contract_facility <> "None of the facilities list on Application" Then
			correct_caseload_type = "Contracted - " & caseload_contract_facility
		End If
	End If

	If correct_caseload_type = "" Then
		If unknown_hc_pending = True or ma_status <> "INACTIVE" or msp_status <> "INACTIVE" Then
			If appears_ltc_checkbox = checked Then
				correct_caseload_type = "LTC+ - General"
				'MsgBox left(case_name, 1) & vbCr & InStr(alpha_split_two_m_z, left(case_name, 1))
				If InStr(alpha_split_one_a_l, left(case_name, 1)) <> 0 Then new_caseload = "X127EK4"
				If InStr(alpha_split_two_m_z, left(case_name, 1)) <> 0 Then new_caseload = "X127EK9"
				If current_caseload = new_caseload Then transfer_needed = False
			End If
		End If
	End If

	population = ""
	If correct_caseload_type = "" Then
		If grh_status = "ACTIVE" or grh_status = "PENDING" or grh_status = "REIN" Then
			correct_caseload_type = "GRH / HS"
			If correct_caseload_type = left(current_caseload_type, 8) Then transfer_needed = False
			population = "Adults"
			If addr_on_1800_faci_list = True Then
				correct_caseload_type = "1800 - Team 160"
				If InStr(alpha_split_one_a_l, left(case_name, 1)) <> 0 Then new_caseload = "X127EF8"
				If InStr(alpha_split_two_m_z, left(case_name, 1)) <> 0 Then new_caseload = "X127EF9"
				If new_caseload <> "" and current_caseload <> new_caseload Then transfer_needed = True
			ElseIf transfer_needed = True Then
				correct_caseload_type = "GRH / HS - Adults Pending"
				If case_has_child_under_19 = True or preg_person_on_case = True Then
					correct_caseload_type = "GRH / HS - Families Pending"
					population = "Families"
				End If
			End If
		End If
	End If

	If correct_caseload_type = "" Then
		If application_form = "MHCP App for Certain Populations - 3876" or application_form = "MNsure App for HC - 6696" or application_form = "No Application Required" Then
			If age_of_memb_01 < 18 Then
				correct_caseload_type = "TEFRA"
				If correct_caseload_type = current_caseload_type Then transfer_needed = False
			End If
		End If
	End If

	If correct_caseload_type = "" Then
		If dwp_status = "PENDING" or mfip_status = "PENDING" Then
			correct_caseload_type = "Families - Cash"
			If age_of_memb_01 < 20 Then correct_caseload_type = "YET"
			population = "Families"
		ElseIf unknown_cash_pending = True or ga_status = "PENDING" or msa_status = "PENDING" Then
			If age_of_memb_01 < 20 Then
				correct_caseload_type = "YET"
				population = "Families"
			ElseIf case_has_child_under_19 = True or preg_person_on_case = True Then
				correct_caseload_type = "Families - Cash"
				population = "Families"
			Else
				correct_caseload_type = "Adults - Pending"
				population = "Adults"
			End If
		ElseIf emer_status = "PENDING" Then
			If age_of_memb_01 < 20 Then
				correct_caseload_type = "YET"
				population = "Families"
			ElseIf emer_type = "EGA" Then
				correct_caseload_type = "Adults - Pending"
				population = "Adults"
			ElseIf emer_type = "EA" Then
				correct_caseload_type = "Families - Pending"
				population = "Families"
			End If
		End If
	End If

	If correct_caseload_type = "" Then
		If age_of_memb_01 < 20 Then
			correct_caseload_type = "YET"
			population = "Families"
		ElseIf case_has_child_under_19 = True Then
			correct_caseload_type = "Families - Pending"
			population = "Families"
		ElseIf (case_has_child_under_22 = False or case_has_guardian = False) and preg_person_on_case = False Then
			correct_caseload_type = "Adults - Pending"
			population = "Adults"
		Else
			correct_caseload_type = "Families - Pending"
			population = "Families"
		End If
	End If
    
    'HC Pending cases 'cases that are HC pending or have active HC programs with pending SNAP/CASH not to a specific team go to pending HC caseload
    If correct_caseload_type = "" OR correct_caseload_type = "Adults - Pending" OR correct_caseload_type = "Families - Pending" Then
        If unknown_hc_pending = True OR ma_status <> "INACTIVE"  OR msp_status <> "INACTIVE" Then 
            correct_caseload_type = "Healthcare - Pending"
        End If 
    End If

    'Adjust correct_caseload_type for correct Team
    If (correct_caseload_type = "Adults - Pending" OR correct_caseload_type = "Families - Pending") AND (case_active <> TRUE OR isnumeric(right(current_caseload_type, 1)) = False) Then random_team_needed = TRUE

    If correct_caseload_type = "Adults - Pending" Then 'Grabs the current team for the caseload type for cases already active on a program
        If isnumeric(right(current_caseload_type, 1)) = True Then
            correct_caseload_type = "Adults - Pending " & right(current_caseload_type, 1)
        Else
            correct_caseload_type = "Adults - Pending ?" 'making correct length if current_caseload_type = ""
        End If 
    ElseIf correct_caseload_type = "Families - Pending" Then 
        If isnumeric(right(current_caseload_type, 1)) = True Then
            correct_caseload_type = "Families - Pending " & right(current_caseload_type, 1)
        Else
            correct_caseload_type = "Families - Pending ?" 'making correct length if current_caseload_type = ""
        End If
    End If 

	If correct_caseload_type = current_caseload_type Then transfer_needed = False
	' MsgBox "correct_caseload_type - " & correct_caseload_type & vbCr & "current_caseload_type - " & current_caseload_type & vbCr & "transfer_needed - " & transfer_needed
	' MsgBox "current_caseload_type - " & current_caseload_type & vbCr & "current_caseload - " & current_caseload & vbCr & "correct_caseload_type - " & correct_caseload_type & vbCr & "transfer_needed - " & transfer_needed


	' case_has_child_under_19
	' case_has_guardian
	' age_of_memb_01
	' case_has_child_under_22
	' addr_on_1800_faci_list
	' case_name

	' unknown_cash_pending
	' unknown_hc_pending
	' ga_status
	' msa_status
	' mfip_status
	' dwp_status
	' grh_status
	' snap_status
	' ma_status
	' msp_status
	' msp_type
	' emer_status
	' emer_type
	If transfer_needed = True and new_caseload = "" Then
        
        Call get_caseload_array_by_type(correct_caseload_type, available_caseload_array)
		number_of_options = UBound(available_caseload_array)
		Do
			pnd2_limit_issue = False
			call select_random_index(number_of_options, caseload_index)
			new_caseload = available_caseload_array(caseload_index)
			If number_of_options = 0 Then Exit Do
			call navigate_to_MAXIS_screen("REPT", "PND2")
			Call write_value_and_transmit(new_caseload, 21, 13)
            EMReadScreen pnd2_disp_limit, 13, 6, 35
            If pnd2_disp_limit = "Display Limit" Then
				pnd2_limit_issue = True
			Else
				EMReadScreen total_pages, 2, 3, 79
				total_pages = trim(total_pages)
				total_pages = total_pages * 1
				if total_pages > 34 Then pnd2_limit_issue = True
			End If


		Loop until pnd2_limit_issue = False
	End If
end function

'---------------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMReadScreen user_x_number, 7, 22, 8
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

'Initial Dialog - Case number
Dialog1 = ""                                        'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 191, 150, "Application Received"
  EditBox 60, 35, 45, 15, MAXIS_case_number
  If running_from_ca_menu = True Then CheckBox 60, 55, 105, 10, "Check here if case is PRIV", priv_case_checkbox
  ButtonGroup ButtonPressed
    PushButton 90, 110, 95, 15, "Script Instructions", script_instructions_btn
    OkButton 80, 130, 50, 15
    CancelButton 135, 130, 50, 15
  Text 5, 10, 185, 20, "Multiple CASE:NOTEs will be entered with this script run to document the actions for pending new applications."
  Text 5, 40, 50, 10, "Case Number:"
  Text 5, 70, 185, 10, "This case should be in PND2 status for this script to run."
  Text 5, 80, 185, 30, "If the programs requested on the application are not yet pending in MAXIS, cancel this script run, pend the case to PND2 status and run the script again."
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
        If ButtonPressed = script_instructions_btn Then             'Pulling up the instructions if the instruction button was pressed.
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx"
            err_msg = "LOOP"
        Else                                                        'If the instructions button was NOT pressed, we want to display the error message if it exists.
		    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        End If
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Checking for PRIV cases.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
IF is_this_priv = True THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
MAXIS_background_check      'Making sure we are out of background.
EMReadScreen initial_pw_for_data_table, 7, 21, 17
EMReadScreen case_name_for_data_table, 20, 21, 46
case_name_for_data_table = replace(case_name_for_data_table, "'", "-")

'Grabbing case and program status information from MAXIS.
'For tis script to work correctly, these must be correct BEFORE running the script.
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, active_programs, programs_applied_for)
EMReadScreen pnd2_appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
Call gather_current_caseload(current_caseload, secondary_caseload, True, previous_pw)

Call navigate_to_MAXIS_screen("CASE", "PERS")               'Getting client eligibility of HC from CASE PERS
pers_row = 10                                               'This is where client information starts on CASE PERS
clt_hc_is_pending = False                                   'defining this at the beginning of each row of CASE PERS
HH_members_pending = ""
Do
	EMReadScreen clt_hc_ref_numb, 2, pers_row, 3     'this reads for the end of the list
	EMReadScreen clt_hc_status, 1, pers_row, 61             'reading the HC status of each client
	'MsgBox clt_hc_status
	If clt_hc_status = "P" Then
		clt_hc_is_pending = True                             'if HC is active then we will add this client to the array to find additional information
		HH_members_pending = HH_members_pending & ", MEMB " & clt_hc_ref_numb
	End If

	pers_row = pers_row + 3         'next client information is 3 rows down
	If pers_row = 19 Then           'this is the end of the list of client on each list
		PF8                         'going to the next page of client information
		on_page = on_page + 1       'saving that we have gone to a new page
		pers_row = 10               'resetting the row to read at the top of the next page
		EMReadScreen end_of_list, 9, 24, 14
		If end_of_list = "LAST PAGE" Then Exit Do
	End If
	EMReadScreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list
	' MsgBox "next_pers_ref_numb - " & next_pers_ref_numb & vbCr & "clt_hc_status - " & clt_hc_status
Loop until next_pers_ref_numb = "  "
If left(HH_members_pending, 1) = "," Then HH_members_pending = right(HH_members_pending, len(HH_members_pending)-1)
HH_members_pending = trim(HH_members_pending)
PF3
If clt_hc_is_pending = True and InStr(programs_applied_for, "HC") = 0 Then
	If programs_applied_for <> "" Then programs_applied_for = programs_applied_for & ", HC"
	If programs_applied_for = "" Then programs_applied_for = "HC"
End If

call back_to_SELF           'resetting
EMReadScreen mx_region, 10, 22, 48
mx_region = trim(mx_region)
If mx_region = "INQUIRY DB" Then
    ' continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
    ' If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
End If

case_status = trim(case_status)     'cutting off any excess space from the case_status read from CASE/CURR above
script_run_lowdown = "CASE STATUS - " & case_status & vbCr & "CASE IS PENDING - " & case_pending        'Adding details about CASE/CURR information to a script report out to BZST
If case_status = "CAF1 PENDING" Then                    'The case MUST be pending and NOT in PND1 to continue.
    call script_end_procedure_with_error_report("This case is not in PND2 status. Current case status in MAXIS is " & case_status & ". Update MAXIS to put this case in PND2 status and then run the script again.")
Else
	multiple_app_dates = False                          'defaulting the boolean about multiple application dates to FALSE
	EMWriteScreen MAXIS_case_number, 18, 43             'now we are going to try to get to REPT/PND2 for the case to read the application date.
	Call navigate_to_MAXIS_screen("REPT", "PND2")
	EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
	If pnd2_disp_limit = "Display Limit" Then transmit
	row = 1                                             'searching for the CASE NUMBER to read from the right row
	col = 1
	EMSearch MAXIS_case_number, row, col
	If row <> 24 and row <> 0 Then pnd2_row = row
	EMReadScreen pdn2_error_info, 40, 24, 2
	pdn2_error_info = ucase(pdn2_error_info)
	If InStr(pdn2_error_info, "IS NOT PENDING") Then call script_end_procedure_with_error_report("This case is not in PND2 status. Current case status in MAXIS is " & case_status & ". Update MAXIS to put this case in PND2 status and then run the script again.")
	If pnd2_row = "" Then
		EMReadScreen too_big_basket, 7, 21, 13
		Call script_end_procedure_with_error_report("This script - Application Received - cannot read this case on REPT/PND2. This is likely because PND2 for the caseload " & too_big_basket & " has reached the MAXIS display limit and the case is not displayed on the page. Transfer this case to a caseload with fewer cases in PND2 status for this script to operate.")
	End If
End If

EMReadScreen application_date, 8, pnd2_row, 38                                  'reading and formatting the application date
application_date = replace(application_date, " ", "/")
oldest_app_date = application_date
EMReadScreen CA_1_code, 1, pnd2_row, 54                                         'reading the pending codes by program for the application date line.
EMReadScreen FS_1_code, 1, pnd2_row, 62
EMReadScreen HC_1_code, 1, pnd2_row, 65
EMReadScreen EA_1_code, 1, pnd2_row, 68
EMReadScreen GR_1_code, 1, pnd2_row, 72

'This section checks to see if the case has multiple application dates
'it will ignore any dates that are for CCAP only as those are not pertinent to our work.
EMReadScreen additional_application_check_one, 14, pnd2_row + 1, 17                 'looking to see if this case has a secondary application date entered
EMReadScreen additional_app_one_hc, 1, pnd2_row + 1, 65
EMReadScreen additional_app_one_ccap, 27, pnd2_row + 1, 54
If additional_application_check_one = "ADDITIONAL APP" Then
	EMReadScreen additional_application_check_two, 14, pnd2_row + 2, 17                 'looking to see if this case has a third application date entered
	EMReadScreen additional_app_two_hc, 1, pnd2_row + 2, 65
	EMReadScreen additional_app_two_ccap, 27, pnd2_row + 2, 54
End If

'Once we have read the lines of REPT/PND2, we need to determine if the additional application should be considered
additional_es_application = False
additional_app_row = ""
If additional_application_check_one = "ADDITIONAL APP" Then
	If additional_app_one_hc = "P" Then											'secondary application is for HC - we count it
		additional_es_application = True
		additional_app_row = pnd2_row + 1										'set the row to read for secondary application date
	ElseIf additional_app_one_ccap <> "_       _     _   _       P" Then		'secondary application is for CCAP only - we do NOT count it
		additional_es_application = True
		additional_app_row = pnd2_row + 1										'set the row to read for secondary application date
	End If
	If additional_application_check_two = "ADDITIONAL APP" Then
		If additional_app_two_hc = "P" Then										'third application is for HC - we count it
			additional_es_application = True
			additional_app_row = pnd2_row + 2									'set the row to read for secondary application date
		ElseIf additional_app_two_ccap <> "_       _     _   _       P" Then	'third application is for CCAP only - we do NOT count it
			additional_es_application = True
			additional_app_row = pnd2_row + 2									'set the row to read for secondary application date
		End If
	End If
End If

If additional_es_application = True THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.


' EMReadScreen additional_application_check, 14, pnd2_row + 1, 17                 'looking to see if this case has a secondary application date entered
' If additional_application_check = "ADDITIONAL APP" THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.
    multiple_app_dates = True           'identifying that this case has multiple application dates - this is not used specifically yet but is in place so we can output information for managment of case handling in the future.

    EMReadScreen additional_application_date, 8, additional_app_row, 38               'reading the app date from the other application line
    additional_application_date = replace(additional_application_date, " ", "/")
    newest_app_date = additional_application_date
    EMReadScreen CA_2_code, 1, pnd2_row, 54                                     'reading the pending codes by program for the second application date line.
    EMReadScreen FS_2_code, 1, pnd2_row, 62
    EMReadScreen HC_2_code, 1, pnd2_row, 65
    EMReadScreen EA_2_code, 1, pnd2_row, 68
    EMReadScreen GR_2_code, 1, pnd2_row, 72


    'There is a specific dialog that will display if there is more than one application date so we can select the right one for this script run
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 166, 160, "Application Received"
      DropListBox 15, 70, 100, 45, application_date+chr(9)+additional_application_date, app_date_to_use
      ButtonGroup ButtonPressed
        PushButton 65, 120, 95, 15, "Open CM 05.09.06", cm_05_09_06_btn
        OkButton 55, 140, 50, 15
        CancelButton 110, 140, 50, 15
      Text 5, 10, 135, 10, "This case has a second application date."
      Text 5, 25, 165, 25, "Per CM 0005.09.06 - if a case is pending and a new app is received you should use the original application date."
      Text 5, 55, 115, 10, "Select which date you need to use:"
      Text 5, 90, 145, 30, "Please contact Knowledge Now or your Supervisor if you have questions about dates to enter in MAXIS for applications."
    EndDialog

    Do
    	Do
    		Dialog Dialog1
    		cancel_without_confirmation

            'referncing the CM policy about application dates.
            If ButtonPressed = cm_05_09_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00050906"
    	Loop until ButtonPressed = -1
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    application_date = app_date_to_use                          'setting the application date selected to the application_date variable
End If

If (case_pending = False and clt_hc_is_pending = False) Then
	If CA_1_code <> "" and CA_1_code <> " " and CA_1_code <> "_" Then unknown_cash_pending = True
	If CA_2_code <> "" and CA_2_code <> " " and CA_2_code <> "_" Then unknown_cash_pending = True
	If FS_1_code <> "" and FS_1_code <> " " and FS_1_code <> "_" Then snap_status = "PENDING"
	If FS_2_code <> "" and FS_2_code <> " " and FS_2_code <> "_" Then snap_status = "PENDING"
	If HC_1_code <> "" and HC_1_code <> " " and HC_1_code <> "_" Then unknown_hc_pending = True
	If HC_2_code <> "" and HC_2_code <> " " and HC_2_code <> "_" Then unknown_hc_pending = True
	If EA_1_code <> "" and EA_1_code <> " " and EA_1_code <> "_" Then emer_status = "PENDING"
	If EA_2_code <> "" and EA_2_code <> " " and EA_2_code <> "_" Then emer_status = "PENDING"
	If GR_1_code <> "" and GR_1_code <> " " and GR_1_code <> "_" Then grh_status = "PENDING"
	If GR_2_code <> "" and GR_2_code <> " " and GR_2_code <> "_" Then grh_status = "PENDING"
End If

IF IsDate(application_date) = False THEN                   'If we could NOT find the application date - then it will use the PND2 application date.
    application_date = pnd2_appl_date
End if

app_recvd_note_found = False
Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
too_old_date = DateAdd("D", -1, oldest_app_date)              'We don't need to read notes from before the CAF date

note_row = 5
previously_pended_progs = ""
MEMO_NOTE_found = False
screening_found = False
Do
    EMReadScreen note_date, 8, note_row, 6                  'reading the note date

    EMReadScreen note_title, 55, note_row, 25               'reading the note header
    note_title = trim(note_title)

    If left(note_title, 22) = "~ Application Received" Then
        app_recvd_note_found = True
        Call write_value_and_transmit("X", note_row, 3)
        in_note_row = 4
        Do
            EMReadScreen note_line, 78, in_note_row, 3
            note_line = trim(note_line)

            If left(note_line, 25) = "* Application Requesting:" Then
                previously_pended_progs = right(note_line, len(note_line)-25)
                previously_pended_progs = trim(previously_pended_progs)
            End If

            If left(note_line, 18) = "* Case Population:" Then
                population_of_case = right(note_line, len(note_line)-18)
                population_of_case = trim(population_of_case)
            End If

            in_note_row = in_note_row + 1
            If in_note_row = 18 Then
                PF8
                in_note_row = 4
                EMReadScreen end_of_note, 9, 24, 14
                If end_of_note = "LAST PAGE" Then Exit Do
            End If
        Loop until note_line = ""
        PF3
    end If
    If left(note_title, 33) = "~ Appointment letter sent in MEMO" Then MEMO_NOTE_found = True       'MEMO case note
    If left(note_title, 31) = "~ Received Application for SNAP" Then screening_found = True         'Exp screening case note

    if note_date = "        " then Exit Do

    note_row = note_row + 1                         'Going to the next row of the CASE/NOTEs to read the next NOTE
    if note_row = 19 then
        note_row = 5
        PF8
        EMReadScreen check_for_last_page, 9, 24, 14
        If check_for_last_page = "LAST PAGE" Then Exit Do
    End If
    EMReadScreen next_note_date, 8, note_row, 6
    If next_note_date = "        " then Exit Do                         'if we are out of notes to read - leave the loop
Loop until DateDiff("d", too_old_date, next_note_date) <= 0             'once we are past the first application date, we stop reading notes

'If we have found the application received CASE/NOTE, we want to evaluate for if we need a subsequent application or app received run
If app_recvd_note_found = True Then
    skip_start_of_subsequent_apps = True                                'defaults
    hc_case = False
    If unknown_hc_pending = True Then hc_case = True                    'finding if the case has HC pending
    If ma_status = "PENDING" Then hc_case = True
    If msp_status = "PENDING" Then hc_case = True
	If clt_hc_is_pending = True Then hc_case = True

    'if HC is pending, we need to confirm that there are 2 different applications to process.
    If hc_case = True Then hc_request_on_second_app = MsgBox("It appears this case has already had the 'Application Received' script on this case. For CAF based programs, we should only run Application Received once since the application dates need to be aligned." & vbCr & vbCR &_
															 "Case currently has the following programs pending: " & programs_applied_for & vbCr & vbCR &_
															 "The following household members have Health Care pending: " & HH_members_pending & vbCr & vbCR &_
                                                             "Are there 2 seperate applications? One for Health Care and another for CAF based program(s)?", vbquestion + vbYesNo, "Type of Application Process")
    'If no HC or if answered 'No' we need to run Subsequent Application instead
	If hc_case = False or hc_request_on_second_app = vbNo Then  call run_from_GitHub(script_repository & "notes/subsequent-application.vbs")
    If hc_case = True and hc_request_on_second_app = vbYes Then     'if this is a HC application and CAF application, we need to determine which is which.
        If application_date = oldest_app_date Then                  'defaulting the program selection based on the application dates and programs
            not_processed_app_date = newest_app_date
            If HC_1_code = "P" Then processing_application_program = "Health Care Programs"
            If HC_1_code = "P" Then other_application_program = "CAF Based Programs"
            If HC_2_code = "P" Then other_application_program = "Health Care Programs"
            If HC_2_code = "P" Then processing_application_program = "CAF Based Programs"
        End If
        If application_date = newest_app_date Then
            not_processed_app_date = oldest_app_date
            If HC_2_code = "P" Then processing_application_program = "Health Care Programs"
            If HC_2_code = "P" Then other_application_program = "CAF Based Programs"
            If HC_1_code = "P" Then other_application_program = "Health Care Programs"
            If HC_1_code = "P" Then processing_application_program = "CAF Based Programs"
        End If

        'this dialog will allow the worker to assign the type of application to the correct application date so the rest of the secipt
        Do
            Do
                err_msg = ""

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 321, 135, "Application Date Information"
                  DropListBox 215, 60, 95, 45, "Select One..."+chr(9)+"Health Care Programs"+chr(9)+"CAF Based Programs", processing_application_program
                  DropListBox 215, 85, 95, 45, "Select One..."+chr(9)+"Health Care Programs"+chr(9)+"CAF Based Programs", other_application_program
                  ButtonGroup ButtonPressed
                    OkButton 210, 115, 50, 15
                    CancelButton 265, 115, 50, 15
                  Text 130, 10, 90, 10, "Multiple Application Dates"
                  Text 10, 30, 300, 20, "This case has Health Care pending and multiple application dates. We need to determine if this run of the script is for a seperate Health Care application or a CAF application."
                  GroupBox 5, 50, 310, 30, "THIS APPLICATION"
                  Text 10, 65, 135, 10, "Application we are currently processing:"
                  Text 165, 65, 40, 10, application_date
                  Text 70, 90, 75, 10, " Previous Application:"
                  Text 165, 90, 40, 10, not_processed_app_date
                  Text 10, 110, 150, 20, "*** CAF Based Programs mean Cash, SNAP,         Emergency, or Housing Support. "
                EndDialog


                dialog Dialog1
                cancel_confirmation

                If processing_application_program = "Select One..." Then err_msg = err_msg & vbCr & "* Please indicate what types of programs are requested on the application you are currently processing."
                If other_application_program = "Select One..." Then err_msg = err_msg & vbCr & "* Please indicate what types of programs are requested on the application that wass previously worked on."
                If processing_application_program = other_application_program Then err_msg = err_msg & vbCr & "* If both applications are for the same types of programs, there should not be seperate application dates. Review the answers and update if incorrect. If correct, cancel the script and call the TSS Help Desk to remove the second application date."

                If err_msg <> "" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine

            Loop until err_msg = ""
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

        script_run_lowdown = script_run_lowdown & vbCr & "processing_application_program - " & processing_application_program

        'reset the programs based on what was answered for the current application and force the processing in the right way.
        If processing_application_program = "Health Care Programs" Then
            unknown_cash_pending = False
            ga_status = ""
            msa_status = ""
            mfip_status = ""
            dwp_status = ""
            grh_status = ""
            snap_status = ""
            emer_status = ""
            emer_type = ""
            programs_applied_for = "HC"
        End If

        If processing_application_program = "CAF Based Programs" Then
            unknown_hc_pending = False
            ma_status = ""
            msp_status = ""
            msp_type = ""

            programs_applied_for = replace(programs_applied_for, "HC", "")
            programs_applied_for = replace(programs_applied_for, ", ,", "")
            programs_applied_for = trim(programs_applied_for)

            If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)
        End If
    End If
End If

Call navigate_to_MAXIS_screen("SPEC", "MEMO")                   'checking to make sure the ADDR panel exists. the MEMO functionality doesnt list 'client' if the ADDR is missing
PF5
EMReadScreen recipient_type, 6, 5, 15
PF3
If recipient_type <> "CLIENT" Then
    script_run_lowdown = script_run_lowdown & vbCr & "First MEMO Recipient was: " & recipient_type
    Call script_end_procedure_with_error_report("This case appears to have an issue with the ADDR panel. Proper case actions cannot occur if the ADDR panel is missing, blank, or has another error.")
End If

Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadscreen last_name, 25, 6, 30
EMReadscreen first_name, 12, 6, 63
EMReadScreen age_of_memb_01, 3, 8, 76
last_name = replace(last_name, "_", "")
first_name = replace(first_name, "_", "")
case_name = last_name & ", " & first_name
age_of_memb_01 = trim(age_of_memb_01)
If age_of_memb_01 = "" Then age_of_memb_01 = 0
age_of_memb_01 = age_of_memb_01*1

case_has_child_under_19 = False
case_has_child_under_22 = False
case_has_guardian = False
Do
	EMReadScreen memb_age, 3, 8, 76
	memb_age = trim(memb_age)
	If memb_age = "" Then memb_age = 0
	memb_age = memb_age*1

	If memb_age < 22 Then
		case_has_child_under_22 = True
		If memb_age < 19 Then case_has_child_under_19 = True
	End If

	EMReadScreen rel_to_applicant, 2, 10, 42
	If rel_to_applicant = "03" Then case_has_guardian = True
	If rel_to_applicant = "04" Then case_has_guardian = True
	If rel_to_applicant = "18" Then case_has_guardian = True
	If rel_to_applicant = "08" Then case_has_guardian = True
	If rel_to_applicant = "09" Then case_has_guardian = True
	If rel_to_applicant = "10" Then case_has_guardian = True
	If rel_to_applicant = "11" Then case_has_guardian = True
	If rel_to_applicant = "12" Then case_has_guardian = True
	If rel_to_applicant = "13" Then case_has_guardian = True
	If rel_to_applicant = "15" Then case_has_guardian = True
	If rel_to_applicant = "16" Then case_has_guardian = True

	transmit
	EMReadScreen end_of_MEMB, 7, 24, 2
Loop until end_of_MEMB = "ENTER A"

preg_person_on_case = False
Call navigate_to_MAXIS_screen("STAT", "PREG")
Do
	EMReadScreen preg_exists, 1, 2, 73
	If preg_exists = "1" Then preg_person_on_case = True
	transmit
	EMReadScreen end_of_MEMB, 7, 24, 2
Loop until end_of_MEMB = "ENTER A"

Call navigate_to_MAXIS_screen("STAT", "PARE")
EMReadScreen pare_exists, 1, 2, 73
If pare_exists = "1" Then case_has_guardian = True


child_under_19_question = "No"
If case_has_child_under_19 = True Then child_under_19_question = "Yes"
child_under_22_question = "No"
If case_has_child_under_22 = True Then child_under_22_question = "Yes"
guardian_question = "No"
If case_has_guardian = True Then guardian_question = "Yes"
pregnant_question = "No"
If preg_person_on_case = True Then pregnant_question = "Yes"
faci_1800_question = "No"

cash_hh_definition_applies = False
If emer_status = "PENDING" Then cash_hh_definition_applies = True
If dwp_status = "PENDING" Then cash_hh_definition_applies = True
If mfip_status = "PENDING" Then cash_hh_definition_applies = True
If msa_status = "PENDING" Then cash_hh_definition_applies = True
If ga_status = "PENDING" Then cash_hh_definition_applies = True
If unknown_cash_pending = True Then cash_hh_definition_applies = True

Call navigate_to_MAXIS_screen("STAT", "PROG")           'going here because this is a good background for the dialog to display against.

IF IsDate(application_date) = False THEN
    stop_early_msg = "This script cannot continue as the application date could not be found from MAXIS."
    stop_early_msg = stop_early_msg & vbCr & vbCr & "CASE: " & MAXIS_case_number
    stop_early_msg = stop_early_msg & vbCr & "Application Date: " & application_date
    stop_early_msg = stop_early_msg & vbCr & "Programs applied for: " & programs_applied_for
    stop_early_msg = stop_early_msg & vbCr & vbCr & "If you are unsure why this happened, screenshot this and send it to HSPH.EWS.BlueZoneScripts@hennepin.us"
    Call script_end_procedure_with_error_report(stop_early_msg)
End If

'Application form list for dialog
app_form_list = chr(9)+"CAF - 5223"
app_form_list = app_form_list+chr(9)+"MNbenefits CAF - 5223"
app_form_list = app_form_list+chr(9)+"SNAP App for Seniors - 5223F"
app_form_list = app_form_list+chr(9)+"MNsure App for HC - 6696"
app_form_list = app_form_list+chr(9)+"MHCP App for Certain Populations - 3876"
app_form_list = app_form_list+chr(9)+"App for MA for LTC - 3531"
app_form_list = app_form_list+chr(9)+"MHCP App for B/C Cancer - 3523"
app_form_list = app_form_list+chr(9)+"EA/EGA Application"
If running_from_ca_menu = True Then app_form_list = app_form_list+chr(9)+"Form other than Application"
app_form_list = app_form_list+chr(9)+"No Application Required"

app_facilities = "Monarch Facilities Contract"		'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Bloomington "			'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Chateau"					'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Excelsior"				'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at St. Louis Park"			'MONARCH
app_facilities = app_facilities+chr(9)+"A Villa Facilities Contract"		'VILLA
app_facilities = app_facilities+chr(9)+"Brookview"							'VILLA
app_facilities = app_facilities+chr(9)+"Park Health and Rehab"				'VILLA
app_facilities = app_facilities+chr(9)+"Richfield Villa Center/ Rehab"		'VILLA
app_facilities = app_facilities+chr(9)+"Robbinsdale Rehab and Care Center"	'VILLA
app_facilities = app_facilities+chr(9)+"Texas Terrace"						'VILLA
app_facilities = app_facilities+chr(9)+"Villa at Bryn Mawr"					'VILLA
app_facilities = app_facilities+chr(9)+"Villa at Osseo"						'VILLA
app_facilities = app_facilities+chr(9)+"Villa at St. Louis Park"			'VILLA
app_facilities = app_facilities+chr(9)+"Ebenezer Care Center/ Martin Luther Care Center"	'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Ebenezer Care Center"				'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Ebenezer Loren on Park"				'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Martin Luther Care Center"			'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Meadow Woods"						'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"North Memorial"
app_facilities = app_facilities+chr(9)+"HCMC"

'since this dialog has different displays for SNAP cases vs non-snap cases - there are differences in the dialog size
dlg_len = 190
If snap_status = "PENDING" Then dlg_len = 280

'This is the dialog with the application information.
Dialog1 = "" 'Blanking out previous dialog detail
'265 - previous dlg width
BeginDialog Dialog1, 0, 0, 515, dlg_len, "Application Received for: " & programs_applied_for & " on " & application_date
  GroupBox 5, 5, 255, 115, "Application Information"
  DropListBox 85, 30, 95, 45, "Select One:"+chr(9)+"ECF"+chr(9)+"Online"+chr(9)+"Request to APPL Form"+chr(9)+"In Person", how_application_rcvd
  DropListBox 85, 50, 170, 45, "Select One:"+app_form_list, application_type
  EditBox 85, 75, 105, 15, confirmation_number
  CheckBox 85, 95, 160, 10, "Check here if LTC or a Waiver is indicated.", appears_ltc_checkbox
  CheckBox 85, 110, 85, 10, "METS Retro Coverage", METS_retro_checkbox


  '   DropListBox 85, 60, 95, 45, "Select One:"+chr(9)+"Adults"+chr(9)+"Families"+chr(9)+"Specialty", population_of_case
  Text 15, 20, 65, 10, "Date of Application:"
  Text 85, 20, 60, 10, application_date

  y_pos = 15
  If unknown_cash_pending = True Then
    Text 275, y_pos, 50, 10, "Cash"
    y_pos = y_pos + 10
  End If
  If ga_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "GA"
    y_pos = y_pos + 10
  End If
  If msa_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "MSA"
    y_pos = y_pos + 10
  End If
  If mfip_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "MFIP"
    y_pos = y_pos + 10
  End If
  If dwp_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "DWP"
    y_pos = y_pos + 10
  End If
  If ive_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "IV-E"
    y_pos = y_pos + 10
  End If
  If grh_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "GRH"
    y_pos = y_pos + 10
  End If
  If snap_status = "PENDING" Then
    Text 275, y_pos, 50, 10, "SNAP"
    y_pos = y_pos + 10
  End If
  If emer_status = "PENDING" Then
    Text 275, y_pos, 50, 10, emer_type
    y_pos = y_pos + 10
  End If
  If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True OR clt_hc_is_pending = True Then
    Text 275, y_pos, 50, 10, "HC"
    y_pos = y_pos + 10
  End If

  GroupBox 265, 5, 70, y_pos, "Pending Programs:"

  Text 10, 35, 70, 10, "Application Received:"
'   Text 10, 65, 70, 10, "Population/Specialty"
  Text 15, 55, 65, 10, "Type of Application:"
  Text 85, 65, 50, 10, "Confirmation #:"

  y_pos = 5
  y_pos = y_pos + 10
  Text 345, y_pos, 125, 10, "Age of MEMB 01: " & age_of_memb_01	'"age_of_memb_01 - " & age_of_memb_01

  y_pos = y_pos + 10
  Text 345, y_pos, 200, 10, "Case Name: " & case_name	'"case_name - " & case_name

  If cash_hh_definition_applies = True Then
	y_pos = y_pos + 10
	If case_has_child_under_19 = False Then y_pos = y_pos + 5
	Text 345, y_pos, 100, 10, "Case has a child under 19: " 	'"case_has_child_under_19 - " & case_has_child_under_19
	If case_has_child_under_19 = False Then DropListBox 445, y_pos-5, 30, 15, "Yes"+chr(9)+"No", child_under_19_question
	If case_has_child_under_19 = True Then Text 445, y_pos, 30, 15, child_under_19_question
  Else
	y_pos = y_pos + 10
	If case_has_child_under_22 = False Then y_pos = y_pos + 5
	Text 345, y_pos, 100, 10, "Case has a member under 22: "	'"case_has_child_under_22 - " & case_has_child_under_22
	If case_has_child_under_22 = False Then DropListBox 445, y_pos-5, 30, 15, "Yes"+chr(9)+"No", child_under_22_question
	If case_has_child_under_22 = True Then Text 445, y_pos, 30, 15, child_under_22_question

	y_pos = y_pos + 10
	If case_has_guardian = False Then y_pos = y_pos + 5
	Text 345, y_pos, 100, 10, "Case has a guardian (parent): "	'"case_has_guardian - " & case_has_guardian
	If case_has_guardian = False Then DropListBox 445, y_pos-5, 30, 15, "Yes"+chr(9)+"No", guardian_question
	If case_has_guardian = True Then Text 445, y_pos, 30, 15, guardian_question
  End If

  y_pos = y_pos + 10
  If preg_person_on_case = False Then y_pos = y_pos + 5
  Text 345, y_pos, 100, 10, "Case has a pregnant person: "	'"case_has_guardian - " & case_has_guardian
  If preg_person_on_case = False Then DropListBox 445, y_pos-5, 30, 15, "Yes"+chr(9)+"No", pregnant_question
  If preg_person_on_case = True Then Text 445, y_pos, 30, 15, pregnant_question

  If grh_case = True Then
	y_pos = y_pos + 10
	Text 345, y_pos+5, 135, 20, "Is the ADDR on the 1800 FACI List: "	'"addr_on_1800_faci_list - " & addr_on_1800_faci_list
	DropListBox 465, y_pos, 30, 15, "Yes"+chr(9)+"No", faci_1800_question
	y_pos = y_pos + 5
  End If
  GroupBox 340, 5, 165, y_pos+10, "Population Questions"

  y_pos = y_pos + 15
  Text 275, y_pos, 250, 10, "Contracted Facilities (if written on the application, indicate which):"
  DropListBox 275, y_pos+10, 190, 45, "None of the facilities list on Application"+chr(9)+app_facilities, contracted_facility

  If dlg_len = 280 Then Text 300, y_pos+40, 200, 40, "The population questions are used to determine if a case needs to be transferred and which Caseload (X-Number) the case should be in. The script has attempted to identify and answer these questions from STAT inofmraiton, but additional details may only be on the application form."
  	' case_has_child_under_19
	' case_has_guardian
	' age_of_memb_01
	' case_has_child_under_22
	' addr_on_1800_faci_list
	' case_name

  y_pos = 130
  If snap_status = "PENDING" Then
      GroupBox 5, 120, 270, 95, "Expedited Screening"
      EditBox 130, 130, 50, 15, income
      EditBox 130, 150, 50, 15, assets
      EditBox 130, 170, 50, 15, rent
      CheckBox 200, 150, 55, 10, "Heat (or AC)", heat_AC_check
      CheckBox 200, 160, 45, 10, "Electricity", electric_check
      CheckBox 200, 170, 35, 10, "Phone", phone_check
      Text 35, 135, 90, 10, "Income received this month:"
      Text 35, 155, 90, 10, "Cash, checking, or savings:"
      Text 35, 175, 90, 10, "AMT paid for rent/mortgage:"
      GroupBox 190, 130, 75, 60, "Utilities claimed "
	  Text 195, 138, 60, 10, "(check below):"
      Text 25, 190, 100, 10, "**IMPORTANT**"'" The income, assets and shelter costs fields will default to $0 if left blank."
      Text 25, 200, 230, 10, "The income, assets and shelter costs fields will default to $0 if left blank. "
      y_pos = 220
  End If
  CheckBox 15, y_pos, 220, 10, "Check here if a HH Member is active on another MAXIS Case.", hh_memb_on_active_case_checkbox
'   y_pos = y_pos + 15
  CheckBox 250, y_pos, 220, 10, "Check here if only CAF1 is completed on the application.", only_caf1_recvd_checkbox
  y_pos = y_pos + 15
  EditBox 70, y_pos, 430, 15, other_notes
  Text 25, y_pos + 5, 45, 10, "Other Notes:"
  y_pos = y_pos + 20
  EditBox 70, y_pos, 190, 15, worker_signature
  Text 5, y_pos + 5, 60, 10, "Worker Signature:"
'   y_pos = y_pos + 20
  ButtonGroup ButtonPressed
    OkButton 400, y_pos, 50, 15
    CancelButton 455, y_pos, 50, 15
EndDialog

'Displaying the dialog
Do
	Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        If application_type = "MNbenefits CAF - 5223" AND how_application_rcvd = "Select One:" Then how_application_rcvd = "Online"
		If application_type = "No Application Required" AND how_application_rcvd <> "Request to APPL Form" Then err_msg = err_msg & vbNewLine & "* No Application cases can only be processed with a 'Request to APPL' form."
	    IF how_application_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
        'IF application_type = "N/A" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please enter in other notes what type of application was received to the agency."
	    IF application_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
        IF application_type = "MNbenefits CAF - 5223" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "* If a MNbenefits app was received, you must enter the confirmation number and time received."
        If population_of_case = "Select One:" then err_msg = err_msg & vbNewLine & "* Please indicate the population or specialty of the case."
        If snap_status = "PENDING" Then
            If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) THEN err_msg = err_msg & vbnewline & "* The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
        End If
	    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        If snap_status = "PENDING" Then
        End If
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

app_date_with_blanks = replace(application_date, "/", " ")                       'creating a variable formatted with spaces instead of '/' for reading on HCRE if needed later in the script

Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)      'We want to be acting in the application month generally

Call hest_standards(heat_AC_amt, electric_amt, phone_amt, application_date) 'function to determine the hest standards depending on the application date.

send_appt_ltr = FALSE                                           'Now we need to determine if this case needs an appointment letter based on the program(s) pending
If unknown_cash_pending = True Then send_appt_ltr = TRUE
If ga_status = "PENDING" Then send_appt_ltr = TRUE
If msa_status = "PENDING" Then send_appt_ltr = TRUE
If mfip_status = "PENDING" Then send_appt_ltr = TRUE
If dwp_status = "PENDING" Then send_appt_ltr = TRUE
If grh_status = "PENDING" Then send_appt_ltr = TRUE
If snap_status = "PENDING" Then send_appt_ltr = TRUE
If emer_status = "PENDING" and emer_type = "EGA" Then send_appt_ltr = TRUE
' If emer_status = "PENDING" and emer_type = "EA" Then send_appt_ltr = TRUE

If emer_status = "PENDING" and emer_type = "EGA" Then transfer_to_worker = "EP8"           'defaulting the transfer working for EGA cases as these are to be sent to this basket'

'Now we will use the entries in the Application information to determine if this case is screened as expedited
IF heat_AC_check = CHECKED THEN
    utilities = heat_AC_amt
ELSEIF electric_check = CHECKED and phone_check = CHECKED THEN
    utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
ELSEIF phone_check = CHECKED and electric_check = UNCHECKED THEN
    utilities = phone_amt
ELSEIF electric_check = CHECKED and phone_check = UNCHECKED THEN
    utilities = electric_amt
END IF

'in case no options are clicked, utilities are set to zero.
IF phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked THEN utilities = 0
'If nothing is written for income/assets/rent info, we set to zero.
IF income = "" THEN income = 0
IF assets = "" THEN assets = 0
IF rent   = "" THEN rent   = 0

If child_under_19_question = "No" Then case_has_child_under_19 = False
If child_under_19_question = "Yes" Then case_has_child_under_19 = True

If child_under_22_question = "No" Then case_has_child_under_22 = False
If child_under_22_question = "Yes" Then case_has_child_under_22 = True

If guardian_question = "No" Then case_has_guardian = False
If guardian_question = "Yes" Then case_has_guardian = True

If pregnant_question = "No" Then preg_person_on_case = False
If pregnant_question = "Yes" Then preg_person_on_case = True

If faci_1800_question = "No" Then addr_on_1800_faci_list = False
If faci_1800_question = "Yes" Then addr_on_1800_faci_list = True

caseload_contract_facility = contracted_facility
If caseload_contract_facility = "Estates at Bloomington " Then caseload_contract_facility = "Monarch Facilities Contract"
If caseload_contract_facility = "Estates at Chateau " Then caseload_contract_facility = "Monarch Facilities Contract"
If caseload_contract_facility = "Estates at Excelsior" Then caseload_contract_facility = "Monarch Facilities Contract"
If caseload_contract_facility = "Estates at St. Louis Park" Then caseload_contract_facility = "Monarch Facilities Contract"

If caseload_contract_facility = "Brookview" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Park Health and Rehab" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Richfield Villa Center/ Rehab" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Robbinsdale Rehab and Care Center" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Texas Terrace" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Villa at Bryn Mawr" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Villa at Osseo" Then caseload_contract_facility = "A Villa Facilities Contract"
If caseload_contract_facility = "Villa at St. Louis Park" Then caseload_contract_facility = "A Villa Facilities Contract"

If caseload_contract_facility = "Ebenezer Care Center" Then caseload_contract_facility = "Ebenezer Care Center/ Martin Luther Care Center"
If caseload_contract_facility = "Ebenezer Loren on Park" Then caseload_contract_facility = "Ebenezer Care Center/ Martin Luther Care Center"
If caseload_contract_facility = "Martin Luther Care Center" Then caseload_contract_facility = "Ebenezer Care Center/ Martin Luther Care Center"
If caseload_contract_facility = "Meadow Woods" Then caseload_contract_facility = "Ebenezer Care Center/ Martin Luther Care Center"

If application_type = "App for MA for LTC - 3531" Then appears_ltc_checkbox = checked
call find_correct_caseload(current_caseload, secondary_caseload, user_x_number, previous_pw, transfer_needed, correct_caseload_type, new_caseload, application_type, appears_ltc_checkbox, METS_retro_checkbox, caseload_contract_facility, case_has_child_under_19, case_has_guardian, age_of_memb_01, case_has_child_under_22, preg_person_on_case, addr_on_1800_faci_list, case_name, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type)

'Calculates expedited status based on above numbers - only for snap pending cases
If snap_status = "PENDING" Then
    IF (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) THEN
        ' If population_of_case = "Families" Then transfer_to_worker = "EZ1"      'cases that screen as expedited are defaulted to expedited specific baskets based on population
        ' If population_of_case = "Adults" Then
        '     'making sure that Adults EXP baskets are not at limit
        '     EX1_basket_available = True
        '     Call navigate_to_MAXIS_screen("REPT", "PND2")
        '     Call write_value_and_transmit("EX1", 21, 17)
        '     EMReadScreen pnd2_disp_limit, 13, 6, 35
        '     If pnd2_disp_limit = "Display Limit" Then EX1_basket_available = False

        '     EX2_basket_available = True
        '     Call navigate_to_MAXIS_screen("REPT", "PND2")
        '     Call write_value_and_transmit("EX2", 21, 17)
        '     EMReadScreen pnd2_disp_limit, 13, 6, 35
        '     If pnd2_disp_limit = "Display Limit" Then EX2_basket_available = False

        '     If (EX1_basket_available = True and EX2_basket_available = False) then
        '         transfer_to_worker = "EX1"
        '     ElseIf (EX1_basket_available = False and EX2_basket_available = True) then
        '         transfer_to_worker = "EX2"
        '     Else
        '     'Do all the randomization here
        '         Randomize       'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
        '         random_number = Int(100*Rnd) 'rnd function returns a value greater or equal 0 and less than 1.
        '         If random_number MOD 2 = 1 then transfer_to_worker = "EX1"		'odd Number
        '         If random_number MOD 2 = 0 then transfer_to_worker = "EX2"		'even Number
        '     End if
        ' End If
        expedited_status = "Client Appears Expedited"                           'setting a variable with expedited information
    End If
    IF (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) THEN expedited_status = "Client Does Not Appear Expedited"

    'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
    Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)
    Call navigate_to_MAXIS_screen("STAT", "DISQ")
    EMReadScreen DISQ_member_check, 34, 24, 2   'Reads the DISQ info for the case note.
    If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then
    	has_DISQ = False
    Else
    	has_DISQ = True
    End if

    'Reads MONY/DISB 'Head of Household" coding to see if a card has been issued. B, H, p and R codes mean that a resident has already received a card and cannot get another in office.
    'DHS webinar meeting 07/20/2022
    in_office_card = True   'Defaulting to true
    IF expedited_status = "client appears expedited" THEN
        Call navigate_to_MAXIS_screen("MONY", "DISB")
        EmReadscreen HoH_card_status, 1, 15, 27
        If HoH_card_status = "B" or _
           HoH_card_status = "H" or _
           HoH_card_status = "P" or _
           HoH_card_status = "R" then
           in_office_card = False
        End if
    End if
End If

'if the case is determined to need an appointment letter the script will default the interview date
IF send_appt_ltr = TRUE THEN
    interview_date = dateadd("d", 5, application_date)
    If interview_date <= date then interview_date = dateadd("d", 5, date)
    Call change_date_to_soonest_working_day(interview_date, "FORWARD")

    application_date = application_date & ""
    interview_date = interview_date & ""                                        'turns interview date into string for variable
End If

' If population_of_case = "Families" Then                                         'families cases that have cash pending need to to to these specific baskets
'     If unknown_cash_pending = True Then transfer_to_worker = "EY9"
'     If mfip_status = "PENDING" Then transfer_to_worker = "EY9"
'     If dwp_status = "PENDING" Then transfer_to_worker = "EY9"
' End if

' 'The familiy cash basket has a backup if it has hit the display limit.
' If transfer_to_worker = "EY9" Then
'     Call navigate_to_MAXIS_screen("REPT", "PND2")
'     EMWriteScreen "EY9", 21, 17
'     transmit
'     EMReadScreen pnd2_disp_limit, 13, 6, 35
'     If pnd2_disp_limit = "Display Limit" Then transfer_to_worker = "EY8"
' End If


dlg_len = 75                'this is another dynamic dialog that needs different sizes based on what it has to display.
IF send_appt_ltr = TRUE THEN dlg_len = dlg_len + 85
IF how_application_rcvd = "Request to APPL Form" THEN dlg_len = dlg_len + 80

back_to_self                                        'added to ensure we have the time to update and send the case in the background
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

If priv_case_checkbox = checked Then transfer_needed = False
If mx_region = "TRAINING" Then transfer_needed = False
transfer_to_worker = right(new_caseload, 3)
If transfer_needed = False Then
	no_transfer_checkbox = checked
	transfer_to_worker = ""
End If


'defining the actions dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 325, dlg_len, "Actions in MAXIS"
  Text 15, 15, 310, 10, "Case appears to be of the caseload type " & correct_caseload_type
  If mx_region = "TRAINING" Then
	Text 25, 30, 250, 10, "TRAINING Region case - will not transfer."
	Text 25, 40, 250, 10, "Correct caseload - " & new_caseload
  ElseIf transfer_needed = True Then
	Text 25, 30, 250, 10, "** Case will be transferred to " & new_caseload
    Text 35, 45, 135, 10, "Case has been transferred in ECF Next?"
    DropListBox 170, 40, 65, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", ECF_transfer_confirmation
  Else
	Text 25, 30, 250, 10, "** Case is currently in " & current_caseload & ", which is the correct caseload."
	Text 30, 40, 250, 10, "The case will not be transferred."
  End If

	'   "TESTING OVERRIDE - review caseload assignments if changing."
	'   EditBox 95, 15, 30, 15, transfer_to_worker
	'   CheckBox 20, 35, 185, 10, "Check here if this case does not require a transfer.", no_transfer_checkbox
	'   '   If expedited_status = "Client Appears Expedited" Then Text 130, 20, 130, 10, "This case screened as EXPEDITED."
	'   '   If expedited_status = "Client Does Not Appear Expedited" Then Text 130, 20, 130, 10, "Case screened as NOT EXPEDITED."
	'   Text 10, 20, 85, 10, "Transfer the case to x127"
  GroupBox 5, 5, 315, 50, "Transfer Information"
  y_pos = 55
  IF send_appt_ltr = TRUE THEN
      GroupBox 5, 55, 315, 80, "Appointment Notice"
      y_pos = y_pos + 15
      Text 15, y_pos, 35, 10, "CAF date:"
      Text 50, y_pos, 55, 15, application_date
      Text 120, y_pos, 60, 10, "Appointment date:"
      Text 185, y_pos, 55, 15, interview_date
      y_pos = y_pos + 15
      Text 50, y_pos, 195, 10, "The NOTICE cannot be cancelled or changed from this script."
      y_pos = y_pos + 10
      Text 50, y_pos, 250, 10, "An Eligibility Worker can make changes/cancellations to the notice in MAXIS."
      y_pos = y_pos + 10
      Text 50, y_pos, 200, 10, "This script follows the requirements for the On Demand Waiver."
      y_pos = y_pos + 10
      odw_btn_y_pos = y_pos
      y_pos = y_pos + 25
  End If
  IF how_application_rcvd = "Request to APPL Form" THEN
      GroupBox 5, y_pos, 315, 75, "Request to APPL Information"
      y_pos = y_pos + 10
      reset_y = y_pos
      EditBox 85, y_pos, 45, 15, request_date
      Text 15, y_pos + 5, 60, 10, "Submission Date:"
      y_pos = y_pos + 20
      ' EditBox 85, y_pos, 45, 15, request_worker_number
      ' Text 15, y_pos + 5, 60, 10, "Requested By X#:"
      ' y_pos = y_pos + 20
      EditBox 85, y_pos, 45, 15, METS_case_number
      Text 15, y_pos + 5, 55, 10, "METS Case #:"
      y_pos = reset_y
      CheckBox 150, y_pos, 55, 10, "MA Transition", MA_transition_request_checkbox
      y_pos = y_pos + 15
      CheckBox 150, y_pos, 60, 10, "Auto Newborn", Auto_Newborn_checkbox
      y_pos = y_pos + 15
      CheckBox 150, y_pos, 85, 10, "METS Retro Coverage", METS_retro_checkbox
      y_pos = y_pos + 15
      CheckBox 150, y_pos, 85, 10, "Team 603 will process", team_603_email_checkbox
      y_pos = y_pos + 25
  End If
  ButtonGroup ButtonPressed
    OkButton 215, y_pos, 50, 15
    CancelButton 270, y_pos, 50, 15
    IF send_appt_ltr = TRUE THEN PushButton 50, odw_btn_y_pos, 125, 13, "HSR Manual - On Demand Waiver", on_demand_waiver_button
EndDialog
'THIS COMMENTED OUT DIALOG IS THE DLG EDITOR FRIENDLY VERSION SINCE THERE IS LOGIC IN THE DIALOG
'-------------------------------------------------------------------------------------------------DIALOG
' BeginDialog Dialog1, 0, 0, 266, 220, "Request to Appl"
'   EditBox 95, 15, 30, 15, transfer_to_worker
'   CheckBox 20, 35, 185, 10, "Check here if this case does not require a transfer.", no_transfer_checkbox
'   GroupBox 5, 5, 255, 45, "Transfer Information"
'   Text 10, 20, 85, 10, "Transfer the case to x127"
'   GroupBox 5, 55, 255, 60, "Appointment Notice"
'   Text 50, 70, 55, 15, "application_date"
'   EditBox 185, 65, 55, 15, interview_date
'   Text 15, 70, 35, 10, "CAF date:"
'   Text 120, 70, 60, 10, "Appointment date:"
'   Text 50, 85, 185, 10, "If interview is being completed please use today's date."
'   Text 50, 95, 190, 20, "Enter a new appointment date only if it's a date county offices are not open."
'   GroupBox 5, 120, 255, 75, "Request to Appl Information"
'   EditBox 85, 130, 45, 15, request_date
'   EditBox 85, 150, 45, 15, request_worker_number
'   EditBox 85, 170, 45, 15, METS_case_number
'   CheckBox 150, 130, 55, 10, "MA Transition", MA_transition_request_checkbox
'   CheckBox 150, 145, 60, 10, "Auto Newborn", Auto_Newborn_checkbox
'   CheckBox 150, 160, 85, 10, "METS Retro Coverage", METS_retro_checkbox
'   CheckBox 150, 175, 85, 10, "Team 603 will process", team_603_email_checkbox
'   Text 15, 135, 60, 10, "Submission Date:"
'   Text 15, 155, 60, 10, "Requested By X#:"
'   Text 15, 175, 55, 10, "METS Case #:"
'   ButtonGroup ButtonPressed
'     OkButton 155, 200, 50, 15
'     CancelButton 210, 200, 50, 15
' EndDialog

'displaying the dialog
Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        ' IF no_transfer_checkbox = UNCHECKED AND transfer_to_worker = "" then err_msg = err_msg & vbNewLine & "* You must enter the basket number the case to be transferred by the script or check that no transfer is needed."
        ' IF no_transfer_checkbox = CHECKED and transfer_to_worker <> "" then err_msg = err_msg & vbNewLine & "* You have checked that no transfer is needed, please remove basket number from transfer field."
        ' IF no_transfer_checkbox = UNCHECKED AND len(transfer_to_worker) > 3 AND isnumeric(transfer_to_worker) = FALSE then err_msg = err_msg & vbNewLine & "* Please enter the last 3 digits of the worker number for transfer."
        IF send_appt_ltr = TRUE THEN
            If IsDate(interview_date) = False Then err_msg = err_msg & vbNewLine & "* The Interview Date needs to be entered as a valid date."
        End If
        IF how_application_rcvd = "Request to APPL Form" THEN
            IF request_date = "" THEN err_msg = err_msg & vbNewLine & "* If a request to APPL was received, you must enter the date the form was submitted."
            IF METS_retro_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg & vbNewLine & "* You have checked that this is a METS Retro Request, please enter a METS IC #."
            IF MA_transition_request_checkbox = CHECKED and METS_case_number = "" THEN err_msg = err_msg &  vbNewLine & "* You have checked that this is a METS Transition Request, please enter a METS IC #."
        End If
        If transfer_needed = True Then
            If ECF_transfer_confirmation = "Select one..." then err_msg = err_msg & vbNewLine & "* This case needs to be transferred in ECF Next first. Transfer the case in ECF Next now."
        End if
        If ButtonPressed = on_demand_waiver_button Then
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/On_Demand_Waiver.aspx"
            err_msg = "LOOP"
        Else
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        End If
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has     not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE

transfer_to_worker = trim(transfer_to_worker)               'formatting the information entered in the dialog
transfer_to_worker = Ucase(transfer_to_worker)
request_worker_number = trim(request_worker_number)
request_worker_number = Ucase(request_worker_number)
f = date

If how_application_rcvd = "Request to APPL Form" THEN                           'specific functionality if the application was pended from a request to APPL form
    If ma_status = "PENDING" OR msp_status = "PENDING" OR unknown_hc_pending = True Then        'HC cases - we need to add the persons pending HC to the CNOTE
        Call navigate_to_MAXIS_screen("STAT", "HCRE")                           'we are going to read this information from the HCRE panel.

        hcre_row = 10                   'top row
        household_persons = ""          'starting with a blank string
        Do                              'we are going to look at each row
            EMReadScreen hcre_app_date, 8, hcre_row, 51             'read the app_date
            EMReadScreen hcre_ref_nbr, 2, hcre_row, 24              'read the reference number
            'if the app date matches the app date we are processing, we will save the reference number to the list of all that match
            If hcre_app_date = app_date_with_blanks Then household_persons = household_persons & hcre_ref_nbr & ", "

            hcre_row = hcre_row + 1         'go to the next row.
            If hcre_row = 18 Then           'go to the next page IF we are at the last row
                PF20
                hcre_row = 10
                EMReadScreen last_page_check, 9, 24, 14
                If last_page_check = "LAST PAGE" Then Exit Do   'leave the loop once we have reached the last page of persons on HCRE
            End If
        Loop
        household_persons = trim(household_persons)         'formatting the list of persons requesting HC
        If right(household_persons, 1) = "," THEN household_persons = left(household_persons, len(household_persons) - 1)
    End If
End If

'creating a variable for a shortened form of the application form for the CASE/NOTE header
If application_type = "CAF - 5223" Then short_form_info = "CAF"
If application_type = "MNbenefits CAF - 5223" Then short_form_info = "CAF from MNbenefits"
If application_type = "SNAP App for Seniors - 5223F" Then short_form_info = "Sr SNAP App"
If application_type = "MNsure App for HC - 6696" Then short_form_info = "MNsure HCAPP"
If application_type = "MHCP App for Certain Populations - 3876" Then short_form_info = "HC - Certain Populations"
If application_type = "App for MA for LTC - 3531" Then short_form_info = "LTC HCAPP"
If application_type = "MHCP App for B/C Cancer - 3523" Then short_form_info = "HCAPP for B/C Cancer"
If application_type = "EA/EGA Application" Then short_form_info = "EA/EGA Application"

'NOW WE START CASE NOTING - there are a few
'Initial application CNOTE - all cases get these ones
start_a_blank_case_note
If application_type = "No Application Required" Then
	'this header is for pending a case when no form is received or needed.
	MX_pend_reason = ""
	If Auto_Newborn_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "Auto Newborn & "
	If METS_retro_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "METS Retro Request & "
	If MA_transition_request_checkbox = CHECKED then MX_pend_reason = MX_pend_reason & "MA Transition & "
	MX_pend_reason = trim(MX_pend_reason)
	If right(MX_pend_reason, 1) = "&" Then MX_pend_reason = left(MX_pend_reason, len(MX_pend_reason)-1)
	MX_pend_reason = trim(MX_pend_reason)
	CALL write_variable_in_CASE_NOTE ("~ HC Pended from a METS case for " & MX_pend_reason & " effective " & application_date & " ~")
Else
	If application_type <> "EA/EGA Application" Then application_type = replace(application_type, " - ", " (DHS-") & ")"
	CALL write_variable_in_CASE_NOTE ("~ Application Received (" &  short_form_info & ") pended for " & application_date & " ~")
	CALL write_bullet_and_variable_in_CASE_NOTE("Application Form Received", application_type)
End If
CALL write_bullet_and_variable_in_CASE_NOTE("Requesting HC for MEMBER(S) ", household_persons)
CALL write_bullet_and_variable_in_CASE_NOTE("Request to APPL Form received on ", request_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
Call write_bullet_and_variable_in_CASE_NOTE ("Case Population", population_of_case)
CALL write_bullet_and_variable_in_CASE_NOTE ("Application Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Active Programs", active_programs)
If transfer_to_worker <> "" THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to X127" & transfer_to_worker)
If hh_memb_on_active_case_checkbox = checked Then Call write_variable_in_CASE_NOTE("* A Member on this case is active on another MAXIS Case.")
If only_caf1_recvd_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Case Pended with only information on CAF1 of the Application.")
CALL write_bullet_and_variable_in_CASE_NOTE ("Other Notes", other_notes)

CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

PF3 ' to save Case note


'Functionality to send emails if the case was pended from a 'Request to APPL'
IF how_application_rcvd = "Request to APPL Form" Then
	send_email_to = ""
	cc_email_to = ""
	If team_603_email_checkbox = CHECKED Then send_email_to = "HSPH.EWS.TEAM.603@hennepin.us"

	email_subject = "Request to APPL Form has been processed for MAXIS Case # " & MAXIS_case_number
	email_body = "Request to APPL form has been received and processed."
	email_body = email_body & vbCr & vbCr & "MAXIS Case # " & MAXIS_case_number & " has been pended and is ready to be processed."
	If METS_case_number <> "" Then email_body = email_body & vbCr & "This case is associated with METS Case # " & METS_case_number & "."

	If METS_retro_checkbox = CHECKED and MA_transition_request_checkbox = CHECKED and Auto_Newborn_checkbox = CHECKED THEN
		email_body = email_body & vbCr & vbCr & "Request to APPL was received for:"
		If METS_retro_checkbox = CHECKED Then email_body = email_body & vbCr & "- METS Retro Request"
		If MA_transition_request_checkbox = CHECKED Then email_body = email_body & vbCr & "- MA Transition"
		If Auto_Newborn_checkbox = CHECKED Then email_body = email_body & vbCr & "- Auto Newborn"
	End If
	IF send_appt_ltr = TRUE THEN email_body = email_body & vbCr & vbCr & "A SPEC/MEMO has been created. If the client has completed the interview, please cancel the notice and update STAT/PROG with the interview information. Case Assignment is not tasked with cancelling or preventing this notice from being generated."
	email_body = email_body & vbCr & vbCr & "Case is ready to be processed."

    'Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
	Call create_outlook_email("", send_email_to, cc_email_to, "", email_subject, 1, False, "", "", False, "", email_body, False, "", False)
End If

'Expedited Screening CNOTE for cases where SNAP is pending
If snap_status = "PENDING" Then
	'formatting the numbers to have 2 decimal points, include a leading 0, do not use parenthesis for negatives, do not include a comma
	income = FormatNumber(income, 2, -1, 0, 0)
	assets = FormatNumber(assets, 2, -1, 0, 0)
	rent = FormatNumber(rent, 2, -1, 0, 0)
	utilities = FormatNumber(utilities, 2, -1, 0, 0)
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & right(space(8) & income, 8))			'The ouput to CNOTE will always take up 8 spaces, with blanks leading
    CALL write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & right(space(8) & assets, 8))
    CALL write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & right(space(8) & rent, 8))
    CALL write_variable_in_CASE_NOTE("        Utilities (AMT/HEST claimed): $" & right(space(8) & utilities, 8))
    CALL write_variable_in_CASE_NOTE("---")
    If has_DISQ = True then CALL write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
    If has_DISQ = False then CALL write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
    If in_office_card = False then CALL write_variable_in_CASE_NOTE("Recipient will NOT be able to get an EBT card in an agency office. An EBT card has previously been provided to the household.")
    CALL write_variable_in_CASE_NOTE("---")
    IF expedited_status = "Client Does Not Appear Expedited" THEN CALL write_variable_in_CASE_NOTE("Client does not appear expedited. Application sent to case file.")
    IF expedited_status = "Client Appears Expedited" THEN CALL write_variable_in_CASE_NOTE("Client appears expedited. Application sent to case file.")
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)

    PF3
End If

'IF a transfer is needed (by entry of a transfer_to_worker in the Action dialog) the script will transfer it here
tansfer_message = ""            'some defaults
transfer_case = False
action_completed = TRUE

If transfer_to_worker <> "" Then        'If a transfer_to_worker was entered - we are attempting the transfer
	transfer_case = True
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER
	EMWriteScreen "x", 7, 16                               'transfer within county option
	transmit
	PF9                                                    'putting the transfer in edit mode
	EMreadscreen servicing_worker, 3, 18, 65               'checking to see if the transfer_to_worker is the same as the current_worker (because then it won't transfer)
	servicing_worker = trim(servicing_worker)
	IF servicing_worker = transfer_to_worker THEN          'If they match, cancel the transfer and save the information about the 'failure'
		action_completed = False
        transfer_message = "This case is already in the requested worker's number."
		PF10 'backout
		PF3 'SPEC menu
		PF3 'SELF Menu'
	ELSE                                                   'otherwise we are going for the tranfer
	    EMWriteScreen "X127" & transfer_to_worker, 18, 61  'entering the worker ifnormation
	    transmit                                           'saving - this should then take us to the transfer menu
        EMReadScreen panel_check, 4, 2, 55                 'reading to see if we made it to the right place
        If panel_check = "XWKR" Then
            action_completed = False                       'this is not the right place
            transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            PF10 'backout
            PF3 'SPEC menu
            PF3 'SELF Menu'
        Else                                               'if we are in the right place - read to see if the new worker is the transfer_to_worker
            EMReadScreen new_pw, 3, 21, 20
            If new_pw <> transfer_to_worker Then           'if it is not the transfer_tow_worker - the transfer failed.
                action_completed = False
                transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            End If
        End If
	END IF
END IF

'SENDING a SPEC/MEMO - this happens AFTER the transfer so that the correct team information is on the notice.
'there should not be an issue with PRIV cases because going directly here we shouldn't lose the 'connection/access'
IF send_appt_ltr = TRUE THEN        'If we are supposed to be sending an appointment letter - it will do it here - this matches the information in ON DEMAND functionality
	last_contact_day = DateAdd("d", 30, application_date)
	If DateDiff("d", interview_date, last_contact_day) < 0 Then last_contact_day = interview_date

	'Navigating to SPEC/MEMO and opening a new MEMO
	Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    		'Writes the appt letter into the MEMO.
    Call create_appointment_letter_notice_application(application_date, interview_date, last_contact_day)

    'now we are going to read if a MEMO was created.
    spec_row = 7
    memo_found = False
    Do
        EMReadScreen print_status, 7, spec_row, 67          'we are looking for a WAITING memo - if one is found -d we are going to assume it is the right one.
        If print_status = "Waiting" Then memo_found = True
        spec_row = spec_row + 1
    Loop until print_status = "       "

    If memo_found = True Then                               'CASE NOTING the MEMO sent if it was successful
        start_a_blank_CASE_NOTE
    	Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
        Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
        Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
        Call write_variable_in_CASE_NOTE("* A link to the Domestic Violence Brochure sent to client in SPEC/MEMO as part of notice.")
        Call write_variable_in_CASE_NOTE("---")
        CALL write_variable_in_CASE_NOTE (worker_signature)

    	PF3
    End If

END IF
'If this is an emer app, send the informational notice about rent-help hennepin
If emer_status = "PENDING" Then
    Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True) 
    Call write_variable_in_SPEC_Memo("You recently applied for Emergency Assistance through Hennepin County.") 
    Call write_variable_in_SPEC_Memo("If you are seeking emergency rent assistance (help for past due rent and associated expenses) please access using the RentHelp Hennepin application at:")
    Call write_variable_in_SPEC_Memo("                        ")
    Call write_variable_in_SPEC_Memo("              renthelphennepin.hdsallita.com ")
    Call write_variable_in_SPEC_Memo("                        ")
    Call write_variable_in_SPEC_Memo("Emergency rent assistance for Hennepin County residents is no longer accessed through MNBenefits or the Combined Application Form. Emergency Assistance and Emergency General Assistance programs continue to be available for emergencies not related to past due rent.")
     PF4
End IF 
'THIS IS FUNCTIONALITY WE WILL NEED TO ADD BACK IN WHEN WE RETURN TO IN PERSON.
'removal of in person functionality during the COVID-19 PEACETIME STATE OF EMERGENCY'
'IF same_day_offered = TRUE and how_application_rcvd = "Office" THEN
'   	start_a_blank_CASE_NOTE
'  	Call write_variable_in_CASE_NOTE("~ Same-day interview offered ~")
'  	Call write_variable_in_CASE_NOTE("* Agency informed the client of needed interview.")
'  	Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive 'a denial notice")
'  	Call write_variable_in_CASE_NOTE("* A Domestic Violence Brochure has been offered to client as part of application packet.")
'  	Call write_variable_in_CASE_NOTE("---")
'  	CALL write_variable_in_CASE_NOTE (worker_signature)
'	PF3
'END IF

'Now we create some messaging to explain what happened in the script run.
end_msg = "Application Received has been noted."
end_msg = end_msg & vbCr & "Programs requested: " & programs_applied_for & " on " & application_date
If snap_status = "PENDING" Then end_msg = end_msg & vbCr & vbCr & "Since SNAP is pending, an Expedtied SNAP screening has been completed and noted based on resident reported information from CAF1."

IF send_appt_ltr = TRUE AND memo_found = True THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice has been sent to the resident to alert them to the need for an interview for their requested programs."
IF send_appt_ltr = TRUE AND memo_found = False THEN end_msg = end_msg & vbCr & vbCr & "A SPEC/MEMO Notice about the Interview appears to have failed. Contact QI Knowledge Now to have one sent manually."

If transfer_message = "" Then
    If transfer_case = True Then end_msg = end_msg & vbCr & vbCr & "Case transfer has been completed to x127" & transfer_to_worker
Else
    end_msg = end_msg & vbCr & vbCr & "FAILED CASE TRANSFER:" & vbCr & transfer_message
End If
If transfer_case = False Then end_msg = end_msg & vbCr & vbCr & "NO TRANSFER HAS BEEN REQUESTED."
IF how_application_rcvd = "Request to APPL Form" Then end_msg = end_msg & vbCr & vbCr & "CASE PENDED from a REQUEST TO APPL FORM"
script_run_lowdown = script_run_lowdown & vbCr & "END Message: " & vbCr & end_msg
Call script_end_procedure_with_error_report(end_msg)


'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/10/2021
'--Tab orders reviewed & confirmed----------------------------------------------09/10/2021
'--Mandatory fields all present & Reviewed--------------------------------------09/10/2021
'--All variables in dialog match mandatory fields-------------------------------09/10/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------05/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/10/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------09/10/2021
'--PRIV Case handling reviewed -------------------------------------------------09/10/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/10/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/10/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------09/10/2021
'--Script name reviewed---------------------------------------------------------09/10/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/10/2021
'--comment Code-----------------------------------------------------------------09/13/2021
'--Update Changelog for release/update------------------------------------------09/10/2021
'--Remove testing message boxes-------------------------------------------------09/10/2021
'--Remove testing code/unnecessary code-----------------------------------------05/01/2022                  We were holding old NOTICE details for in person return. Removed as this detail is drastically different.
'--Review/update SharePoint instructions----------------------------------------09/13/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/10/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/10/2021
