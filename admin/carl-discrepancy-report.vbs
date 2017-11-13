''STATS GATHERING----------------------------------------------------------------------------------------------------
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_denomination = "C"       			'C is for each CASE
STATS_manualtime = 25
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("07/18/2017", "Fully tested version with South MPLS & South Sub. regions added.", "Ilse Ferris, Hennepin County")
call changelog_update("07/12/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output 
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	APPL_date = var_month &"/" & var_day & "/" & var_year
end function

'----------------------------------------------------------------------------------------------------The script
EMConnect ""

'dialog and dialog DO...Loop	
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed 
		BeginDialog x_dialog, 0, 0, 266, 110, "CARL Discrepancy dialog"
  			ButtonGroup ButtonPressed
    		PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    		OkButton 145, 90, 50, 15
    		CancelButton 200, 90, 50, 15
  			EditBox 15, 45, 180, 15, file_selection_path
  			GroupBox 10, 5, 250, 80, "Using the CARL DISCREPANCY script"
  			Text 20, 20, 235, 20, "This script should be used needing to reconcile the cases on the NOT IN CARL list with the name of the worker who processed the intake."
  			Text 15, 65, 230, 15, "Select the Excel file that contains the CARL information by selecting the 'Browse' button, and finding the file."
		EndDialog
		err_msg = ""
		Dialog x_dialog
		cancel_confirmation
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
'query_start_time = timer

Dim CARL_array()
REdim CARL_array(4, 0)

const case_number 	= 1
const app_date 		= 2
const worker_numb	= 3
const workers_name	= 4

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do                                                              'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 1).value		're-establishing the case numbers for functions to use
	APPL_date = objExcel.cells(excel_row, 5).value
	
    MAXIS_case_number = trim(MAXIS_case_number)    
	If MAXIS_case_number = "" then exit do
	
	Call ONLY_create_MAXIS_friendly_date(APPL_date)			'reformatting the dates to be MM/DD/YY format to measure against the case numbers

	'Adding client information to the array'
	ReDim Preserve CARL_array(4, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	CARL_array (case_number,     entry_record) = MAXIS_case_number		'The client information is added to the array'
	CARL_array (app_date, 		 entry_record) = APPL_date
	CARL_array (worker_numb, 	 entry_record) = ""
	CARL_array (workers_name, 	 entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	excel_row = excel_row + 1
Loop

objExcel.Quit		'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document and don't want the script to be confused
back_to_self

'Now we will get PMI and Member Number for each client on the array.'
For item = 0 to UBound(CARL_array, 2)
	MAXIS_case_number = CARL_array(case_number, item)
	APPL_date = CARL_array (app_date, item)			

	Call navigate_to_MAXIS_screen("CASE", "NOTE")
	'Checking for PRIV cases
	EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
		PF3
		CARL_array(worker_numb, item) = ""
		CARL_array(workers_name, item) = ""
	ELse
		row = 5
		Do
			EMReadScreen case_note_date, 8, row, 6
			If trim(case_note_date) = "" then exit do
			If case_note_date => appl_date then          'if the case note date is equal to or greater than the application date then the case note header is read
				EMReadScreen case_note_header, 55, row, 25
				case_note_header = trim(case_note_header)
				IF instr(case_note_header, "***Intake") then
					CAF_note_found = True
					exit do
				Elseif instr(case_note_header, "***Reapplication") then
					CAF_note_found = True
					exit do
				Elseif instr(case_note_header, "***Add program") then
					CAF_note_found = True
					exit do
				Elseif instr(case_note_header, "***Addendum") then
					CAF_note_found = True
					exit do	
				Elseif instr(case_note_header, "***Emergency app") then
					CAF_note_found = True
					exit do	
				else 	
					CAF_note_found = False
				END IF
			END IF
			row = row + 1
			If row = 19 then 
				PF8
				row = 5
			End if
		LOOP until case_note_date < appl_date                        'repeats until the case note date is less than the application date
		If CAF_note_found = True then 
			Stats_counter = Stats_counter + 1
			EMReadScreen worker_ID, 7, row, 16
			CARL_array (worker_numb, item) = worker_ID
			
			'South suburban and South Regions are here 
			If worker_ID = "X127D7M" then worker_name = "Ahmed Abdi"
			If worker_ID = "X1272GM" then worker_name = "Faduma Abdi"
			If worker_ID = "X127B0C" then worker_name = "Fowsia Abdi"
			If worker_ID = "X127B0X" then worker_name = "Osman Abdi"
			If worker_ID = "X1275A2" then worker_name = "Qadro Abdi"
			If worker_ID = "X1275H8" then worker_name = "Sharmarke Abdi"
			If worker_ID = "X127HU2" then worker_name = "Mohamed Abdirahman"
			If worker_ID = "X127A0Z" then worker_name = "Jamoda Acevedo"
			If worker_ID = "X127AQ7" then worker_name = "Katie Adams"
			If worker_ID = "X1272KD" then worker_name = "Fatumo Aden"
			If worker_ID = "X1272BK" then worker_name = "Mohamed  Ahmed"
			If worker_ID = "X127AK3" then worker_name = "Abdiweli Ali"
			If worker_ID = "X127HS2" then worker_name = "Osob Ali"
			If worker_ID = "X127B7E" then worker_name = "Maria Ammerman"
			If worker_ID = "X127D5Z" then worker_name = "Jennie Anderson"
			If worker_ID = "X127C1Q" then worker_name = "Marilynn Anderson"
			If worker_ID = "X1275J1" then worker_name = "Alejandra Andrade"
			If worker_ID = "X127Y30" then worker_name = "Emanuel Anyaogu"
			If worker_ID = "X127G07" then worker_name = "Myrna Banham-McKelvy"
			If worker_ID = "X127B33" then worker_name = "Danita Banks"
			If worker_ID = "X127Y04" then worker_name = "Abdullahi Berka"
			If worker_ID = "X127AQ4" then worker_name = "Bethelhem Beyene"
			If worker_ID = "X1275P4" then worker_name = "Ondrenette Blair"
			If worker_ID = "X127Y86" then worker_name = "Rhea Blue Arm"
			If worker_ID = "X127W85" then worker_name = "Kathleen Boswell"
			If worker_ID = "X1275DB" then worker_name = "Diane Boucher"
			If worker_ID = "X127H69" then worker_name = "Phillip Bradbury"
			If worker_ID = "X127DSB" then worker_name = "Douglas Bright"
			If worker_ID = "X127D5D" then worker_name = "Candace Brown"
			If worker_ID = "X127BM6" then worker_name = "Tiarra Buford"
			If worker_ID = "X127R74" then worker_name = "Olga  Bugayev"
			If worker_ID = "X127HN4" then worker_name = "Jamika Burdunice"
			If worker_ID = "X127AK5" then worker_name = "Brandy Canada"
			If worker_ID = "X127F25" then worker_name = "Colleen Canfield"
			If worker_ID = "X1275H7" then worker_name = "Peggy Chavez"
			If worker_ID = "X127A8B" then worker_name = "Shantell Cochran"
			If worker_ID = "X127AY1" then worker_name = "Tammy Coenen"
			If worker_ID = "X1275P1" then worker_name = "Sherry Collins"
			If worker_ID = "X127CSC" then worker_name = "Carina Cortez"
			If worker_ID = "X127FAS" then worker_name = "Sarai Counce"
			If worker_ID = "X127B01" then worker_name = "Terri  Cox"
			If worker_ID = "X127BD1" then worker_name = "Beverly Denman"
			If worker_ID = "X127Y43" then worker_name = "Deborah Diggins"
			If worker_ID = "X1275K0" then worker_name = "Sherry Duggan"
			If worker_ID = "X127776" then worker_name = "Steve Duong"
			If worker_ID = "X127KJE" then worker_name = "Karen Elhindi"
			If worker_ID = "X127AE1" then worker_name = "Ayaan Elmi"
			If worker_ID = "X127L04" then worker_name = "Olga Engebretson"
			If worker_ID = "X127811" then worker_name = "William Engels"
			If worker_ID = "X127B0V" then worker_name = "Ibrahim Farah"
			If worker_ID = "X127ZAE" then worker_name = "Heather Feldmann"
			If worker_ID = "X1272KI" then worker_name = "Shannon Felegy"
			If worker_ID = "X127871" then worker_name = "Kathryn Fitzgerald"
			If worker_ID = "X127FAD" then worker_name = "Daminga Flowers"
			If worker_ID = "X127OA1" then worker_name = "Emily Frazier"
			If worker_ID = "X127AM6" then worker_name = "Jolonna Frieling"
			If worker_ID = "X127C66" then worker_name = "James Fust"
			If worker_ID = "X127FGA" then worker_name = "Fatiya Ganamo"
			If worker_ID = "X127A7O" then worker_name = "Gina Gangelhoff"
			If worker_ID = "X1275F9" then worker_name = "Aaron Gardner-Kocher"
			If worker_ID = "X1271AJ" then worker_name = "Kenneth Garnier"
			If worker_ID = "X127FCD" then worker_name = "Christine Glisczinski"
			If worker_ID = "X127A8F" then worker_name = "Bernardo Gonzalez"
			If worker_ID = "X127GJ3" then worker_name = "Marlenne Gonzalez"
			If worker_ID = "X127Z93" then worker_name = "Huruse Gurhan"
			If worker_ID = "X127AL8" then worker_name = "Madar Hachi"
			If worker_ID = "X127GF5" then worker_name = "Johanne Halvorsen"
			If worker_ID = "X127A16" then worker_name = "Anne Halvorson"
			If worker_ID = "X127ZAH" then worker_name = "Tamika Hannah"
			If worker_ID = "X127Q95" then worker_name = "Shanna Hansen"
			If worker_ID = "X127W40" then worker_name = "Shaquila Harris"
			If worker_ID = "X127B9Q" then worker_name = "Molly Hasbrook"
			If worker_ID = "X127Y37" then worker_name = "Patricia Hegenbarth"
			If worker_ID = "X127GU2" then worker_name = "Alyssa Heise"
			If worker_ID = "X12728S" then worker_name = "Cheryl Heitzinger"
			If worker_ID = "X1275H5" then worker_name = "Valerie Herrera"
			If worker_ID = "X127B18" then worker_name = "Kimberly Hill"
			If worker_ID = "X127995" then worker_name = "Elizabeth Hilpisch"
			If worker_ID = "X127A10" then worker_name = "Christopher Hogan"
			If worker_ID = "X1273M7" then worker_name = "Stephanie Holmes"
			If worker_ID = "X127U53" then worker_name = "Kristine Hopkins"
			If worker_ID = "X127JH4" then worker_name = "Janine Hudson"
			If worker_ID = "X12726O" then worker_name = "Nada Hughes"
			If worker_ID = "X127B9P" then worker_name = "Abdirizak Ibrahim"
			If worker_ID = "X127C1C" then worker_name = "Mark Jacobson"
			If worker_ID = "X127B1B" then worker_name = "Toni Jenkins"
			If worker_ID = "X127B36" then worker_name = "Christine Jernander"
			If worker_ID = "X127FAL" then worker_name = "Ziyad Kadir"
			If worker_ID = "X127A2D" then worker_name = "Kristen Kasim"
			If worker_ID = "X127AP7" then worker_name = "Ryan Kierczynski"
			If worker_ID = "X127JS7" then worker_name = "Andy Knutson"
			If worker_ID = "X127ZAL" then worker_name = "Darren Konsor"
			If worker_ID = "X127T27" then worker_name = "Shirley Korman"
			If worker_ID = "X127663" then worker_name = "Rachel Kuppe"
			If worker_ID = "X127AL5" then worker_name = "Denis Ladeyshchikov"
			If worker_ID = "X127LL1" then worker_name = "Lisa Lampkin"
			If worker_ID = "X127J75" then worker_name = "Michelle Le"
			If worker_ID = "X127Y81" then worker_name = "Shelly Lind"
			If worker_ID = "X127Z34" then worker_name = "Raisa Loevski"
			If worker_ID = "X127MML" then worker_name = "Mubarek Lolo"
			If worker_ID = "X127FAW" then worker_name = "Naasira Looper"
			If worker_ID = "X127CA9" then worker_name = "Sarita Lopez"
			If worker_ID = "X127GR8" then worker_name = "Mali Lor"
			If worker_ID = "X1275K5" then worker_name = "Teng Lor"
			If worker_ID = "X127CAL" then worker_name = "Carrie Lucca"
			If worker_ID = "X127A4E" then worker_name = "Carlotta Madison"
			If worker_ID = "X1272PC" then worker_name = "Ramona Mahadeo"
			If worker_ID = "X127966" then worker_name = "Florence Manley"
			If worker_ID = "X127AG4" then worker_name = "Molly Manley"
			If worker_ID = "X127D6P" then worker_name = "Faith Markel"
			If worker_ID = "X127BN9" then worker_name = "Fawn Marquez"
			If worker_ID = "X127WMM" then worker_name = "Watchen Marshall"
			If worker_ID = "X127A3X" then worker_name = "Amanda Martin"
			If worker_ID = "X127AMA" then worker_name = "Alexandra Marzolf"
			If worker_ID = "X1272A5" then worker_name = "Mary McGuinness"
			If worker_ID = "X127D42" then worker_name = "Jacob Mickelson"
			If worker_ID = "X127201" then worker_name = "Bobbie Miller Thomas"
			If worker_ID = "X1275D0" then worker_name = "Ahmed Mohamed"
			If worker_ID = "X127HS7" then worker_name = "Samsam Mohamed"
			If worker_ID = "X127Y23" then worker_name = "Tracy Mohomes"
			If worker_ID = "X127DM1" then worker_name = "David Montano"
			If worker_ID = "X127Y62" then worker_name = "Jennifer Moses"
			If worker_ID = "X1272LQ" then worker_name = "Kacey Musta"
			If worker_ID = "X127KLT" then worker_name = "Kim Lang Nguyen"
			If worker_ID = "X127HJ9" then worker_name = "Jill Niess"
			If worker_ID = "X127085" then worker_name = "Todd Norling"
			If worker_ID = "X127KMN" then worker_name = "Khadra Nur"
			If worker_ID = "X1272UC" then worker_name = "Kevin Ogburn"
			If worker_ID = "X127D6D" then worker_name = "Abdikadir Omar"
			If worker_ID = "X1275G7" then worker_name = "Dorian Pearson"
			If worker_ID = "X127D45" then worker_name = "Debora Penney"
			If worker_ID = "X127D5T" then worker_name = "Jane Pinkerman"
			If worker_ID = "X12728K" then worker_name = "Kelly Porter"
			If worker_ID = "X127Z40" then worker_name = "Kristina Przybilla"
			If worker_ID = "X127T25" then worker_name = "Lindsey Remus"
			If worker_ID = "X127969" then worker_name = "Joan Rice"
			If worker_ID = "X127D4Y" then worker_name = "Laura Riebe"
			If worker_ID = "X1272ID" then worker_name = "Amorette Robeck"
			If worker_ID = "X127D6J" then worker_name = "LaTrena Robinson"
			If worker_ID = "X127L08" then worker_name = "Deborah Rusnak"
			If worker_ID = "X127B2Z" then worker_name = "Nicole Ryan"
			If worker_ID = "X127HT1" then worker_name = "Alexandra Saenz"
			If worker_ID = "X1272BB" then worker_name = "Darlenne Salinas-Fernandez"
			If worker_ID = "X127A7S" then worker_name = "Soumya Sanyal"
			If worker_ID = "X127T65" then worker_name = "Claudia Saulter"
			If worker_ID = "X12726N" then worker_name = "Gina Schnarr"
			If worker_ID = "X1272GB" then worker_name = "Indranie Singh"
			If worker_ID = "X127U46" then worker_name = "Rayeann St Hubert"
			If worker_ID = "X127GB3" then worker_name = "Jill Sternberg-Adams"
			If worker_ID = "X127CSS" then worker_name = "Cortney Stevens"
			If worker_ID = "X127D4S" then worker_name = "Garrett Stock"
			If worker_ID = "X127A68" then worker_name = "Barbara Sullivan"
			If worker_ID = "X1272B2" then worker_name = "Veronica Suvid"
			If worker_ID = "X1275K1" then worker_name = "Aleen Swanson"
			If worker_ID = "X1272RG" then worker_name = "Olucammi Taliaferro"
			If worker_ID = "X127B9O" then worker_name = "Shakir Taliaferro"
			If worker_ID = "X127AW7" then worker_name = "Tamala Taylor"
			If worker_ID = "X127039" then worker_name = "Julie Thompson"
			If worker_ID = "X127BP2" then worker_name = "Serina Thor"
			If worker_ID = "X127C36" then worker_name = "Lori Timmerman"
			If worker_ID = "X127IRT" then worker_name = "Inez Toles"
			If worker_ID = "X127BP4" then worker_name = "Michael Tronnes"
			If worker_ID = "X127632" then worker_name = "Debra Tucker"
			If worker_ID = "X127D6H" then worker_name = "Baraka Tura"
			If worker_ID = "X127X44" then worker_name = "Kimberly Turner"
			If worker_ID = "X127T21" then worker_name = "Kary Van Slyke"
			If worker_ID = "X127T52" then worker_name = "Lilian Van-Cao"
			If worker_ID = "X127GJ8" then worker_name = "Pahoua Vang"
			If worker_ID = "X127BL9" then worker_name = "Scott Vang"
			If worker_ID = "X127CF6" then worker_name = "Leticia Vasquez"
			If worker_ID = "X127GR3" then worker_name = "Adam Verschoor"
			If worker_ID = "X127AM2" then worker_name = "Kerry Walsh"
			If worker_ID = "X127FAH" then worker_name = "Lorna Welch"
			If worker_ID = "X127X43" then worker_name = "Pamela Whitson"
			If worker_ID = "X127D5W" then worker_name = "Dawn Williams"
			If worker_ID = "X127HT9" then worker_name = "Nellie Woodson"
			If worker_ID = "X1272TV" then worker_name = "Beverly Wyka"
			If worker_ID = "X127FAT" then worker_name = "See Xiong"
			If worker_ID = "X127CB1" then worker_name = "Nee Yang"
			If worker_ID = "X127GD5" then worker_name = "Pakou Yang"
			If worker_ID = "X127B3R" then worker_name = "Gail Yarphel"
			
			'Northwest region
			If worker_ID = "X127Z46" then worker_name = "		Afrah, Muna"
			If worker_ID = "X127FAY" then worker_name = "Ahmed, Lina"
			If worker_ID = "X127X59" then worker_name = "Ali, Osman"
			If worker_ID = "X127T86" then worker_name = "Anderson, Irina"
			If worker_ID = "X127AU2" then worker_name = "Barnes, Michelle"
			If worker_ID = "X127WM1" then worker_name = "Bedoya, Wendy"
			If worker_ID = "X127A3J" then worker_name = "Belland, Jessica"
			If worker_ID = "X127ABN" then worker_name = "Beske, Andrea"
			If worker_ID = "X127HT4" then worker_name = "Blee Alarcon, Julio"
			If worker_ID = "X127F30" then worker_name = "Bommersbach, Lisa"
			If worker_ID = "X127D5B" then worker_name = "Branch, Cheryl"
			If worker_ID = "X127G24" then worker_name = "Brase, Kaye"
			If worker_ID = "X127SCC" then worker_name = "Campbell, Sarah"
			If worker_ID = "X127C1T" then worker_name = "Carlson, Celeste"
			If worker_ID = "X127C0Q" then worker_name = "DeMario, Diana"
			If worker_ID = "X1272EG" then worker_name = "Dickerson, Jessica"
			If worker_ID = "X127Y44" then worker_name = "Ditter, Natalya"
			If worker_ID = "X127B8K" then worker_name = "Eberle, DeAnne"
			If worker_ID = "X1272AF" then worker_name = "Ferguson, Rachel"
			If worker_ID = "X127T50" then worker_name = "Flanigan, Kelly"
			If worker_ID = "X127B7G" then worker_name = "Flasch, Jodynne"
			If worker_ID = "X127CG9" then worker_name = "Garbe, Maria"
			If worker_ID = "X127436" then worker_name = "Greene, Linda"
			If worker_ID = "X127Z71" then worker_name = "Harrell, Sara"
			If worker_ID = "X127651" then worker_name = "Harris, Dianna"
			If worker_ID = "X127C86" then worker_name = "Haubrick, Laura"
			If worker_ID = "X1272LJ" then worker_name = "Haw, Samantha"
			If worker_ID = "X127CA2" then worker_name = "Henry-Bolden, Crystal"
			If worker_ID = "X1271A7" then worker_name = "Holt, Kasey"
			If worker_ID = "X127AK7" then worker_name = "Hullukka-Pargo, Amelise"
			If worker_ID = "X1272CZ" then worker_name = "Hurreh, Abdiaziz"
			If worker_ID = "X127GQ8" then worker_name = "Infante, Deanna"
			If worker_ID = "X127Y76" then worker_name = "Jama, Jamila"
			If worker_ID = "X127L87" then worker_name = "Jibrell, Saeed"
			If worker_ID = "X127E92" then worker_name = "Johnson, Dale"
			If worker_ID = "X127S14" then worker_name = "Johnson, Lora"
			If worker_ID = "X1273D3" then worker_name = "Jourdain, Celeste"
			If worker_ID = "X12746I" then worker_name = "Karlsgodt, Kristine"
			If worker_ID = "X127ZAK" then worker_name = "Kelvie, Amy"
			If worker_ID = "X127Z89" then worker_name = "King, Angela"
			If worker_ID = "X127B4Q" then worker_name = "King, Sylvia"
			If worker_ID = "X127D4X" then worker_name = "Kornmann, Sheri"
			If worker_ID = "X127D6A" then worker_name = "Le, Annie"
			If worker_ID = "X127D2F" then worker_name = "Lee, Linda"
			If worker_ID = "X127C0M" then worker_name = "Lee, Shellie"
			If worker_ID = "X127AX4" then worker_name = "Lelugas, Laura"
			If worker_ID = "X127D7R" then worker_name = "Lenear, Shamikka"
			If worker_ID = "X127Y92" then worker_name = "Lewis, Leticia"
			If worker_ID = "X127CE6" then worker_name = "Lo, Bouakou"
			If worker_ID = "X1275H9" then worker_name = "Manuel, Rashida"
			If worker_ID = "X127A6S" then worker_name = "McDowell, Charice"
			If worker_ID = "X127CA0" then worker_name = "Miller, Sara"
			If worker_ID = "X127AQ8" then worker_name = "Morphew, Teresa"
			If worker_ID = "X127A9X" then worker_name = "Moua, Mailee"
			If worker_ID = "X127X82" then worker_name = "Mrsich, Tiffanie"
			If worker_ID = "X127X04" then worker_name = "Nejo, Ephrem"
			If worker_ID = "X127K82" then worker_name = "Nelson, Lisa"
			If worker_ID = "X1271NG" then worker_name = "Ngene, Innocent"
			If worker_ID = "X127U55" then worker_name = "Niev, Sideth"
			If worker_ID = "X127Z83" then worker_name = "Norman, Kristine"
			If worker_ID = "X127B22" then worker_name = "Olson, Brian"
			If worker_ID = "X127FCA" then worker_name = "Payne, Tanya"
			If worker_ID = "X127L52" then worker_name = "Pha, Susan"
			If worker_ID = "X127S01" then worker_name = "Phelps, Rita"
			If worker_ID = "X127M22" then worker_name = "Rivas-Herrera, Barbara"
			If worker_ID = "X12726L" then worker_name = "Rutkovskaya, Victoria"
			If worker_ID = "X127A3V" then worker_name = "Salazar, Miguel"
			If worker_ID = "X1272B1" then worker_name = "Sarin, Sitha"
			If worker_ID = "X1272RO" then worker_name = "Scott, DiAnne"
			If worker_ID = "X127AY3" then worker_name = "Sebald, Lisa"
			If worker_ID = "X127MLC" then worker_name = "Setodji, Michelle"
			If worker_ID = "X127B1K" then worker_name = "Shaffer, Victoria"
			If worker_ID = "X127C0R" then worker_name = "Shipley, Carol"
			If worker_ID = "X1275F3" then worker_name = "Smith, TyAnn"
			If worker_ID = "X127D6S" then worker_name = "Socha, Monica"
			If worker_ID = "X127Z62" then worker_name = "Steele, Bernita"
			If worker_ID = "X127A66" then worker_name = "Stolpe, Robert"
			If worker_ID = "X127G51" then worker_name = "Szyperski, Deborah"
			If worker_ID = "X127CA8" then worker_name = "Tamba, Louise"
			If worker_ID = "X127FAB" then worker_name = "Tenzin, Dee"
			If worker_ID = "X127J77" then worker_name = "Thai, Tina"
			If worker_ID = "X127PC2" then worker_name = "Thao, Mee"
			If worker_ID = "X127B0B" then worker_name = "Tollund, Cathy"
			If worker_ID = "X127AH5" then worker_name = "Vang, Choua"
			If worker_ID = "X127Z87" then worker_name = "Vang, Judy"
			If worker_ID = "X127GK6" then worker_name = "Vang, Pa"
			If worker_ID = "X1275M3" then worker_name = "Vang, Pa"
			If worker_ID = "X1272QW" then worker_name = "Voskresensky, Oksana"
			If worker_ID = "X127P18" then worker_name = "Vue, Vin"
			If worker_ID = "X127Z45" then worker_name = "Waite, Yasmin"
			If worker_ID = "X127BG6" then worker_name = "Wakeyo, Negesso"
			If worker_ID = "X127D3Y" then worker_name = "Walker, Amanda"
			If worker_ID = "X127520" then worker_name = "Weikum, Laura"
			If worker_ID = "X127B5I" then worker_name = "Welch, Denise"
			If worker_ID = "X127M16" then worker_name = "Weller, Clara"
			If worker_ID = "X1275L2" then worker_name = "Wimberly, Aimee"
			If worker_ID = "X127R63" then worker_name = "Xiong, Andre"
			If worker_ID = "X1272RP" then worker_name = "Xiong, David"
			If worker_ID = "X1272DB" then worker_name = "Xiong, Katie"
			If worker_ID = "X127X0X" then worker_name = "Xiong, Xoua"
			If worker_ID = "X127D6K" then worker_name = "Yang, Alexander"
			If worker_ID = "X127A0U" then worker_name = "Yang, Gaohnou"
			If worker_ID = "X127T81" then worker_name = "Yang, Maytong"
			If worker_ID = "X1275N3" then worker_name = "Yang, Panhia"
			If worker_ID = "X127FAP" then worker_name = "Yang, Yeng"
			
			'North region
			If worker_ID = "X127AAA" then worker_name = "Abdulle, Abdullahi"
			If worker_ID = "X127FAQ" then worker_name = "Aden, Ahmed"
			If worker_ID = "X1275L9" then worker_name = "Ban, Lim"
			If worker_ID = "X127JLB" then worker_name = "Banks, Javette"
			If worker_ID = "X127BM7" then worker_name = "Barrow, Melissa"
			If worker_ID = "X127Q73" then worker_name = "Benfield, Daniel"
			If worker_ID = "X1275J5" then worker_name = "Blue, KaSondra"
			If worker_ID = "X127X41" then worker_name = "Bolden, Deborah"
			If worker_ID = "X127DP3" then worker_name = "Broen, Julie"
			If worker_ID = "X127WLC" then worker_name = "Clark, Wendy"
			If worker_ID = "X1272BT" then worker_name = "Coburn-Paden, Dawn"
			If worker_ID = "X127D8A" then worker_name = "Davis, Elacia"
			If worker_ID = "X127B3T" then worker_name = "Dilday, Delia"
			If worker_ID = "X127Z36" then worker_name = "Doughty- Moore, Buffy"
			If worker_ID = "X127HS4" then worker_name = "Farah, Ahmednor"
			If worker_ID = "X1272F6" then worker_name = "Fredin, Kelly"
			If worker_ID = "X127AX8" then worker_name = "Gelle, Lucky"
			If worker_ID = "X1275G0" then worker_name = "George, Debra"
			If worker_ID = "X127DG1" then worker_name = "Gunter, Danielle"
			If worker_ID = "X1274JO" then worker_name = "Harper, Dawona"
			If worker_ID = "X1275P2" then worker_name = "Her, Tony"
			If worker_ID = "X127W47" then worker_name = "Hill, Christine"
			If worker_ID = "X1272UE" then worker_name = "Hopson, Tasheema"
			If worker_ID = "X127GJ4" then worker_name = "Irwin, Molly"
			If worker_ID = "X127AH3" then worker_name = "Isaac, Zechariye"
			If worker_ID = "X127X29" then worker_name = "Isais, Melissa"
			If worker_ID = "X1274QG" then worker_name = "Jackson, Debrice"
			If worker_ID = "X127A7M" then worker_name = "Jefferson, Stephanie"
			If worker_ID = "X127ZAJ" then worker_name = "John, Lauren"
			If worker_ID = "X127AG5" then worker_name = "Jorgenson, Jessica"
			If worker_ID = "X1275M9" then worker_name = "Lacy, Colanda"
			If worker_ID = "X127AW8" then worker_name = "Larson, Kaeli"
			If worker_ID = "X127D05" then worker_name = "Lawrence, Andrea"
			If worker_ID = "X127D4R" then worker_name = "Lee, Mai"
			If worker_ID = "X127F23" then worker_name = "Lee-Xiong, Xay"
			If worker_ID = "X127Y05" then worker_name = "LisVaj, Cassandra"
			If worker_ID = "X127YL1" then worker_name = "Lor, Yer"
			If worker_ID = "X127D7W" then worker_name = "Mack, Yanisha"
			If worker_ID = "X127L23" then worker_name = "Madison, Paul"
			If worker_ID = "X1275M0" then worker_name = "Magadan, Andre"
			If worker_ID = "X1275H2" then worker_name = "Moore, Brittany"
			If worker_ID = "X1275F4" then worker_name = "Moore, Thomas"
			If worker_ID = "X127JJM" then worker_name = "Munger, Jennifer"
			If worker_ID = "X127FAM" then worker_name = "Pargo, Tiwana"
			If worker_ID = "X127HR4" then worker_name = "Parten, Joann"
			If worker_ID = "X127A0W" then worker_name = "Perkerson, Lakisha"
			If worker_ID = "X127AH4" then worker_name = "Roberson, Lori"
			If worker_ID = "X1275L5" then worker_name = "Shields, Richard"
			If worker_ID = "X127TDL" then worker_name = "Terebenet, Darren"
			If worker_ID = "X127AK6" then worker_name = "Thompson, Melissa"
			If worker_ID = "X127I45" then worker_name = "Tran, Trixy"
			If worker_ID = "X127ZAU" then worker_name = "Wieber, Alexandra"
			If worker_ID = "X1275L3" then worker_name = "Williams, Patricia"
			If worker_ID = "X127AK8" then worker_name = "Xiong, Leigh"
			If worker_ID = "X1275D4" then worker_name = "Yang, Sheng"
			If worker_ID = "X127BP1" then worker_name = "Yang, Sirynoise"
			If worker_ID = "X1274JX" then worker_name = "Young, Mandora"
			If worker_ID = "X127OZA" then worker_name = "Zavala, Omar"
			If worker_ID = "X127HQ6" then worker_name = "Stevenson, Mathilda"
			If worker_ID = "X127HP3" then worker_name = "Lane, Rochelle"
			If worker_ID = "X127JT1" then worker_name = "Brolsma, Alicia"

			'Central NE region
			If worker_ID = "X1271KA" then worker_name = "Abdallah, Khadra"
			If worker_ID = "X1273YA" then worker_name = "Ahmed, Abdi"
			If worker_ID = "X1275B7" then worker_name = "Ahmed, Mohammed"
			If worker_ID = "X127HU4" then worker_name = "Ashiro, Sakaria"
			If worker_ID = "X127CA4" then worker_name = "Assefa, Bezabeh"
			If worker_ID = "X127HU3" then worker_name = "Avila, Yunuen"
			If worker_ID = "X127FAC" then worker_name = "Bailey, Tiffany"
			If worker_ID = "X127W69" then worker_name = "Baker, Terry"
			If worker_ID = "X127HU6" then worker_name = "Batres-Marroquin, Jose"
			If worker_ID = "X1272A6" then worker_name = "Bohun, David"
			If worker_ID = "X127JJB" then worker_name = "Bonebrake, Jennifer"
			If worker_ID = "X127F2F" then worker_name = "Botan, Abdirazak"
			If worker_ID = "X127TMB" then worker_name = "Broomfield, Teisha"
			If worker_ID = "X127X99" then worker_name = "Burnett, Neill"
			If worker_ID = "X127A5G" then worker_name = "Byrd, Roberta"
			If worker_ID = "X127T63" then worker_name = "Cartlidge, Sheryn"
			If worker_ID = "X127J90" then worker_name = "Chang, Kao"
			If worker_ID = "X127CF9" then worker_name = "Charpentier, Jacqueline"
			If worker_ID = "X127A5D" then worker_name = "Clark, Marilyn"
			If worker_ID = "X1272KC" then worker_name = "Colleen, JoAnn"
			If worker_ID = "X127CB3" then worker_name = "Dadi, Miftah"
			If worker_ID = "X127E26" then worker_name = "Davis, Ann"
			If worker_ID = "X1272KG" then worker_name = "Dewji, Raihana"
			If worker_ID = "X127HN7" then worker_name = "Dungey, Shanaya"
			If worker_ID = "X127GH9" then worker_name = "Erickson, Susan"
			If worker_ID = "X127B1T" then worker_name = "Feigum, Melissa"
			If worker_ID = "X1275F1" then worker_name = "Goin, Seth"
			If worker_ID = "X127GS3" then worker_name = "Golden, Rebecca"
			If worker_ID = "X127A9U" then worker_name = "Graves, Jacqueline"
			If worker_ID = "X127W49" then worker_name = "Guzman, Kathryn"
			If worker_ID = "X127F19" then worker_name = "Hampton, Cynthia"
			If worker_ID = "X127AM5" then worker_name = "Hassan, Sartu"
			If worker_ID = "X127T67" then worker_name = "Heard, Trenita"
			If worker_ID = "X127A05" then worker_name = "Hoecherl, Cecelia"
			If worker_ID = "X127JB2" then worker_name = "Holmquist, John"
			If worker_ID = "X127BL8" then worker_name = "Hopson, Rhonda"
			If worker_ID = "X1275H0" then worker_name = "Huerta-Stemper, Remy"
			If worker_ID = "X127A44" then worker_name = "Jenkins, Edward"
			If worker_ID = "X127FAG" then worker_name = "Johnson, Cynthia"
			If worker_ID = "X1275G8" then worker_name = "Kendrick-Stevens, DeNise"
			If worker_ID = "X127B2L" then worker_name = "Kennedy, Debra"
			If worker_ID = "X1275P0" then worker_name = "Khan, Amanda"
			If worker_ID = "X127HS6" then worker_name = "Korenchen, Abby"
			If worker_ID = "X127CF3" then worker_name = "Korynta, Raeann"
			If worker_ID = "X127HT5" then worker_name = "Kravets, Nikolai"
			If worker_ID = "X1273FL" then worker_name = "Lane, Melinda"
			If worker_ID = "X1275K2" then worker_name = "Lee, Pa Nhia"
			If worker_ID = "X127875" then worker_name = "Lessner, Jenny"
			If worker_ID = "X127HM1" then worker_name = "Mahmoud, Hind"
			If worker_ID = "X12729W" then worker_name = "Marx, Jason"
			If worker_ID = "X1275L8" then worker_name = "Messer, Lara"
			If worker_ID = "X127GR4" then worker_name = "Miles, Toni"
			If worker_ID = "X127MHL" then worker_name = "Miller, Heather"
			If worker_ID = "X127E60" then worker_name = "Millhouse, Linda"
			If worker_ID = "X127X96" then worker_name = "Mohamed, Irro"
			If worker_ID = "X127GH6" then worker_name = "Mootz, David"
			If worker_ID = "X127X97" then worker_name = "Nelson, Joseph"
			If worker_ID = "X127GR1" then worker_name = "Nong Van, Jack"
			If worker_ID = "X1272AY" then worker_name = "Nur, Mohamed"
			If worker_ID = "X127GEP" then worker_name = "Parodi, Giovanni"
			If worker_ID = "X127BK2" then worker_name = "Peterson, Diana"
			If worker_ID = "X127GU1" then worker_name = "Peyton, Jordan"
			If worker_ID = "X1275N7" then worker_name = "Pratiwi, Azza"
			If worker_ID = "X127B3J" then worker_name = "Prettyman, Kelly" 
			If worker_ID = "X127A79" then worker_name = "Remus, Gary"
			If worker_ID = "X127D7Z" then worker_name = "Ross, Brittney"
			If worker_ID = "X127Q90" then worker_name = "Sandi, Sahr"
			If worker_ID = "X127GG8" then worker_name = "Scriver, John"
			If worker_ID = "X127GK2" then worker_name = "Semmelink, Keith"
			If worker_ID = "X127A22" then worker_name = "Setterlund, Blair"
			If worker_ID = "X127CE7" then worker_name = "Smeby, Marcia"
			If worker_ID = "X127BH3" then worker_name = "Stewart, Tamara"
			If worker_ID = "X127C52" then worker_name = "Streitz, Dennis"
			If worker_ID = "X127M93" then worker_name = "Ternyak, Alla"
			If worker_ID = "X127A48" then worker_name = "Thompson, Susan"
			If worker_ID = "X127FAA" then worker_name = "Thornton, Kiera"
			If worker_ID = "X127TCT" then worker_name = "Tin, Tinna"
			If worker_ID = "X127Q98" then worker_name = "Tran, Hoa"
			If worker_ID = "X127D5V" then worker_name = "Trembley, Neil"
			If worker_ID = "X127AAT" then worker_name = "Tusa, Abdurezak"
			If worker_ID = "X127AK9" then worker_name = "Tyrrell, Michelle"
			If worker_ID = "X127CH2" then worker_name = "Vang, Houa"
			If worker_ID = "X127D6M" then worker_name = "Vang, Ly"
			If worker_ID = "X1273GQ" then worker_name = "Vang, Mai"
			If worker_ID = "X127HT6" then worker_name = "Villanueva Torres, Gabriel"
			If worker_ID = "X127GS2" then worker_name = "Warmboe, Robert"
			If worker_ID = "X127PC4" then worker_name = "Watkins, Alison"
			If worker_ID = "X1272TZ" then worker_name = "Williams, Angela"
			If worker_ID = "X127D7Y" then worker_name = "Williams, Kimberly"
			If worker_ID = "X1272HY" then worker_name = "Winiarczyk, Jacqueline"
			If worker_ID = "X127AXX" then worker_name = "Xiong, Amy"
			If worker_ID = "X127B0G" then worker_name = "Yang-Xiong, Pa"
			If worker_ID = "X127GX9" then worker_name = "Yezek, Susan"
			If worker_ID = "X127DMY" then worker_name = "Young, Darcy"
			If worker_ID = "X1272IQ" then worker_name = "Zelaya, Rebecca"

			'West region
			If worker_ID = "X127Z91" then worker_name = "Abraham, Bonsi"
			If worker_ID = "X127A8Q" then worker_name = "Awad, Lacosta"
			If worker_ID = "X127F22" then worker_name = "Berthiaume, Betsy"
			If worker_ID = "X127L75" then worker_name = "Boesche, Robert"
			If worker_ID = "X127L18" then worker_name = "Carlson, Sharon"
			If worker_ID = "X127FBX" then worker_name = "Eifert, Becky"
			If worker_ID = "X127JH3" then worker_name = "Hong, Jenny"
			If worker_ID = "X1272GQ" then worker_name = "Jacox, Wendy"
			If worker_ID = "X127R49" then worker_name = "Jones, Nina"
			If worker_ID = "X12730W" then worker_name = "Kabakova, Svetlana"
			If worker_ID = "X127C42" then worker_name = "Katz, Robin"
			If worker_ID = "X127F77" then worker_name = "Kerzman, Cheryl"
			If worker_ID = "X1272C1" then worker_name = "Luetgers, Christine"
			If worker_ID = "X127GH5" then worker_name = "Mack, Ashley"
			If worker_ID = "X127GR9" then worker_name = "Meelberg, Russell"
			If worker_ID = "X127JAC" then worker_name = "Merritt, Jennifer"
			If worker_ID = "X127X63" then worker_name = "Miller, Deedra"
			If worker_ID = "X1275G3" then worker_name = "Mitchell, Marianne"
			If worker_ID = "X1275G4" then worker_name = "Nack, Jerald"
			If worker_ID = "X127AH2" then worker_name = "Parenteau, Michelle"
			If worker_ID = "X127GR2" then worker_name = "Poplavska, Kristina"
			If worker_ID = "X12746P" then worker_name = "Scherer, Clarita"
			If worker_ID = "X127K81" then worker_name = "Teske, Thad"
			If worker_ID = "X127GH4" then worker_name = "Teskey, Benjamin"
			If worker_ID = "X127B7M" then worker_name = "Trimbo, Jenna"
			If worker_ID = "X1272A0" then worker_name = "Vue, Natalie"
			If worker_ID = "X127C09" then worker_name = "Wahlstrom, Elizabeth"
			If worker_ID = "X127FAF" then worker_name = "Watkins, Sean"
			If worker_ID = "X127GH1" then worker_name = "Waychoff, Mollie"
			If worker_ID = "X1275L1" then worker_name = "Womack, Karen"
			If worker_ID = "X127Q85" then worker_name = "Yang, Lor"

			CARL_array (workers_name, item) = worker_name
			worker_name = ""
		End if 	 
	END If
Next 

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

'----------------------------------------------------------------------------------------------------Excel inforamtion
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "CASE #"
ObjExcel.Cells(1, 2).Value = "APP DATE"
ObjExcel.Cells(1, 3).Value = "WORKER #"
ObjExcel.Cells(1, 4).Value = "WORKER NAME"

FOR i = 1 to 4		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
Next

objExcel.Columns(2).NumberFormat = "MM/DD/YY"	'formatting the text
			'
Excel_row = 2

'End of setting up the Excel sheet----------------------------------------------------------------------------------------------------
For i = 0 to Ubound(CARL_array, 2)
	'increased by 1 column for each region
	ObjExcel.Cells(Excel_row,  1).Value = CARL_array (case_number, 	i)
	ObjExcel.Cells(Excel_row,  2).Value = CARL_array (app_date,   	i)
	ObjExcel.Cells(Excel_row,  3).Value = CARL_array (worker_numb, 	i)
	ObjExcel.Cells(Excel_row,  4).Value = CARL_array (workers_name, i)
	excel_row = excel_row + 1
Next

FOR i = 1 to 4		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

script_end_procedure("Complete! Review the list.")