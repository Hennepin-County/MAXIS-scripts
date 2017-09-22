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
			
			'Enter the worker information here!
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