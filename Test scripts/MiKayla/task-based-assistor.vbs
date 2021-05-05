'STATS GATHERING--------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK BASED ASSISTOR.vbs"
start_time = timer
STATS_counter = 1  'sets the stats counter at one
STATS_manualtime = 100 'manual run time in seconds
STATS_denomination = "C" 			   'M is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY================================================================
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
CALL changelog_update("01/15/2021", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK
MiKayla_needs_to_know_this = ""

Function HSR_LIST
	assigned_to = replace(assigned_to, ".", "")
	IF assigned_to = "Katie Aanestad" or  assigned_to = "Katie Aanestad" THEN worker_number =  "X127KN9"
	IF assigned_to = "Gina L Aasgaard" or  assigned_to = "Gina Aasgaard" THEN worker_number =  "X12726N"
	IF assigned_to = "Khadra S Abdallah" or  assigned_to = "Khadra Abdallah" THEN worker_number =  "X1271KA"
	IF assigned_to = "Faduma M Abdi" or  assigned_to = "Faduma Abdi" THEN worker_number =  "X1272GM"
	IF assigned_to = "Qadro A Abdi" or  assigned_to = "Qadro Abdi" THEN worker_number =  "X1275A2"
	IF assigned_to = "Sharmarke Y Abdi" or  assigned_to = "Sharmarke Abdi" THEN worker_number =  "X1275H8"
	IF assigned_to = "Fowsia M Abdi" or  assigned_to = "Fowsia Abdi" THEN worker_number =  "X127B0C"
	IF assigned_to = "Osman M Abdi" or  assigned_to = "Osman Abdi" THEN worker_number =  "X127B0X"
	IF assigned_to = "Ahmed A Abdi" or  assigned_to = "Ahmed Abdi" THEN worker_number =  "X127D7M"
	IF assigned_to = "Ugbad R Abdilahi" or  assigned_to = "Ugbad Abdilahi" THEN worker_number =  "X127URA"
	IF assigned_to = "Mohamed Abdirahman" or  assigned_to = "Mohamed Abdirahman" THEN worker_number =  "X127HU2"
	IF assigned_to = "Nabila S Abdullahi" or  assigned_to = "Nabila Abdullahi" THEN worker_number =  "X127NSA"
	IF assigned_to = "Bonsi A Abraham" or  assigned_to = "Bonsi Abraham" THEN worker_number =  "X127Z91"
	IF assigned_to = "Mihiret A Abrahim" or  assigned_to = "Mihiret Abrahim" THEN worker_number =  "X127KL3"
	IF assigned_to = "Jamoda L Acevedo" or  assigned_to = "Jamoda Acevedo" THEN worker_number =  "X127A0Z"
	IF assigned_to = "Alyssa L Ackert" or  assigned_to = "Alyssa Ackert" THEN worker_number =  "X127GU2"
	IF assigned_to = "Katie M Adams" or  assigned_to = "Katie Adams" THEN worker_number =  "X127AQ7"
	IF assigned_to = "Ahmed M Aden" or  assigned_to = "Ahmed Aden" THEN worker_number =  "X127FAQ"
	IF assigned_to = "Muna M Afrah" or  assigned_to = "Muna Afrah" THEN worker_number =  "X127Z46"
	IF assigned_to = "Abdi Y Ahmed" or  assigned_to = "Abdi Ahmed" THEN worker_number =  "X1273YA"
	IF assigned_to = "Mohammed K Ahmed" or  assigned_to = "Mohammed Ahmed" THEN worker_number =  "X1275B7"
	IF assigned_to = "Fatah A Ahmed" or  assigned_to = "Fatah Ahmed" THEN worker_number =  "X127AAF"
	IF assigned_to = "Lina M Ahmed" or  assigned_to = "Lina Ahmed" THEN worker_number =  "X127FAY"
	IF assigned_to = "Jamiya O Ahmed" or  assigned_to = "Jamiya Ahmed" THEN worker_number =  "X127GAN"
	IF assigned_to = "Olukemi O Adeniyi-Akins" or  assigned_to = "Olukemi Adeniyi-Akins" THEN worker_number =  "X127OLU"
	IF assigned_to = "Angel S Alexander" or  assigned_to = "Angel Alexander" THEN worker_number =  "X127B0Y"
	IF assigned_to = "Osob A Ali" or  assigned_to = "Osob Ali" THEN worker_number =  "X127HS2"
	IF assigned_to = "Rahma Ali" or  assigned_to = "Rahma Ali" THEN worker_number =  "X127RBA"
	IF assigned_to = "Osman I Ali" or  assigned_to = "Osman Ali" THEN worker_number =  "X127X59"
	IF assigned_to = "Betty J Allabough" or  assigned_to = "Betty Allabough" THEN worker_number =  "X127JK6"
	IF assigned_to = "Hollie L Allen" or  assigned_to = "Hollie Allen" THEN worker_number =  "X127JD1"
	IF assigned_to = "Claudia Alvarez" or  assigned_to = "Claudia Alvarez" THEN worker_number =  "X127GM5"
	IF assigned_to = "Tania L Amadi" or  assigned_to = "Tania Amadi" THEN worker_number =  "X127GJ1"
	IF assigned_to = "Filsan A Amin" or  assigned_to = "Filsan Amin" THEN worker_number =  "X127JK1"
	IF assigned_to = "Maria E Ammerman" or  assigned_to = "Maria Ammerman" THEN worker_number =  "X127B7E"
	IF assigned_to = "Marya D Anderson" or  assigned_to = "Marya Anderson" THEN worker_number =  "X127AN5"
	IF assigned_to = "Marilynn R Anderson" or  assigned_to = "Marilynn Anderson" THEN worker_number =  "X127C1Q"
	IF assigned_to = "Jennie E Anderson" or  assigned_to = "Jennie Anderson" THEN worker_number =  "X127D5Z"
	IF assigned_to = "Kari Anderson" or  assigned_to = "Kari Anderson" THEN worker_number =  "X127HN1"
	IF assigned_to = "Scott L Anderson" or  assigned_to = "Scott Anderson" THEN worker_number =  "X127JK3"
	IF assigned_to = "Regina Andrews" or  assigned_to = "Regina Andrews" THEN worker_number =  "X127JK4"
	IF assigned_to = "Jacob P Arco" or  assigned_to = "Jacob Arco" THEN worker_number =  "X1275M2"
	IF assigned_to = "Nicole C Arm" or  assigned_to = "Nicole Arm" THEN worker_number =  "X127HN3"
	IF assigned_to = "Sakaria O Ashiro" or  assigned_to = "Sakaria Ashiro" THEN worker_number =  "X127HU4"
	IF assigned_to = "Bezabeh Assefa" or  assigned_to = "Bezabeh Assefa" THEN worker_number =  "X127CA4"
	IF assigned_to = "Yunuen A Avila" or  assigned_to = "Yunuen Avila" THEN worker_number =  "X127HU3"
	IF assigned_to = "Lacosta L Awad" or  assigned_to = "Lacosta Awad" THEN worker_number =  "X127A8Q"
	IF assigned_to = "Negassa K Ayana" or  assigned_to = "Negassa Ayana" THEN worker_number =  "X127043"
	IF assigned_to = "Tiffany R Bailey" or  assigned_to = "Tiffany Bailey" THEN worker_number =  "X127FAC"
	IF assigned_to = "Terry S Baker" or  assigned_to = "Terry Baker" THEN worker_number =  "X127W69"
	IF assigned_to = "Tameka Ballard" or  assigned_to = "Tameka Ballard" THEN worker_number =  "X127JW3"
	IF assigned_to = "Lim Ban" or  assigned_to = "Lim Ban" THEN worker_number =  "X1275L9"
	IF assigned_to = "Myrna C Banham-McKelvy" or  assigned_to = "Myrna Banham-McKelvy" THEN worker_number =  "X127G07"
	IF assigned_to = "Javette L Banks" or  assigned_to = "Javette Banks" THEN worker_number =  "X127JLB"
	IF assigned_to = "TyAnn Barnes" or  assigned_to = "TyAnn Barnes" THEN worker_number =  "X1275F3"
	IF assigned_to = "Michelle E Barnes" or  assigned_to = "Michelle Barnes" THEN worker_number =  "X127AU2"
	IF assigned_to = "Tammi Barton" or  assigned_to = "Tammi Barton" THEN worker_number =  "X127D4D"
	IF assigned_to = "Diane M Beauchamp" or  assigned_to = "Diane Beauchamp" THEN worker_number =  "X127D3C"
	IF assigned_to = "Ann Becker" or  assigned_to = "Ann Becker" THEN worker_number =  "X127R76"
	IF assigned_to = "Wendy M Bedoya" or  assigned_to = "Wendy Bedoya" THEN worker_number =  "X127WM1"
	IF assigned_to = "Angela S Beljeski" or  assigned_to = "Angela Beljeski" THEN worker_number =  "X127W18"
	IF assigned_to = "Jessica L Belland" or  assigned_to = "Jessica Belland" THEN worker_number =  "X127A3J"
	IF assigned_to = "Daniel D Benfield" or  assigned_to = "Daniel Benfield" THEN worker_number =  "X127Q73"
	IF assigned_to = "Peggy M Benkert" or  assigned_to = "Peggy Benkert" THEN worker_number =  "X127H35"
	IF assigned_to = "James P Berka" or  assigned_to = "James Berka" THEN worker_number =  "X127D3Z"
	IF assigned_to = "Abdullahi A Berka" or  assigned_to = "Abdullahi Berka" THEN worker_number =  "X127Y04"
	IF assigned_to = "Anthony H Berne" or  assigned_to = "Anthony Berne" THEN worker_number =  "X127D4K"
	IF assigned_to = "Bethelhem G Beyene" or  assigned_to = "Bethelhem Beyene" THEN worker_number =  "X127AQ4"
	IF assigned_to = "Cortney S Bhakta" or  assigned_to = "Cortney Bhakta" THEN worker_number =  "X127CSS"
	IF assigned_to = "Erik A Billington" or  assigned_to = "Erik Billington" THEN worker_number =  "X127EAB"
	IF assigned_to = "Ondrenette Blair" or  assigned_to = "Ondrenette Blair" THEN worker_number =  "X1275P4"
	IF assigned_to = "Rhea Blue Arm" or  assigned_to = "Rhea Blue  Arm" THEN worker_number =  "X127Y86"
	IF assigned_to = "Robert Bohr" or  assigned_to = "Robert Bohr" THEN worker_number =  "X127KP3"
	IF assigned_to = "Deborah F Bolden" or  assigned_to = "Deborah Bolden" THEN worker_number =  "X127X41"
	IF assigned_to = "Lisa J Bommersbach" or  assigned_to = "Lisa Bommersbach" THEN worker_number =  "X127F30"
	IF assigned_to = "Abdirazak Botan" or  assigned_to = "Abdirazak Botan" THEN worker_number =  "X127F2F"
	IF assigned_to = "Sabastian Boyle-Mejia" or  assigned_to = "Sabastian Boyle-Mejia" THEN worker_number =  "X127KD3"
	IF assigned_to = "Douglas S Bright" or  assigned_to = "Douglas Bright" THEN worker_number =  "X127DSB"
	IF assigned_to = "Sekena Britt-Nelson" or  assigned_to = "Sekena Britt-Nelson" THEN worker_number =  "X127SBN"
	IF assigned_to = "Julie M Broen" or  assigned_to = "Julie Broen" THEN worker_number =  "X127DP3"
	IF assigned_to = "Hannah E Broman" or  assigned_to = "Hannah Broman" THEN worker_number =  "X127KL6"
	IF assigned_to = "Teisha M Broomfield" or  assigned_to = "Teisha Broomfield" THEN worker_number =  "X127TMB"
	IF assigned_to = "Candace S Brown" or  assigned_to = "Candace Brown" THEN worker_number =  "X127D5D"
	IF assigned_to = "Ariel F Brown" or  assigned_to = "Ariel Brown" THEN worker_number =  "X127KM6"
	IF assigned_to = "Ronisha S Buckner" or  assigned_to = "Ronisha Buckner" THEN worker_number =  "X127JK7"
	IF assigned_to = "Olga P Bugayev" or  assigned_to = "Olga Bugayev" THEN worker_number =  "X127R74"
	IF assigned_to = "Abdiwali B Bulhan" or  assigned_to = "Abdiwali Bulhan" THEN worker_number =  "X127ABB"
	IF assigned_to = "Tyler O Burch" or  assigned_to = "Tyler Burch" THEN worker_number =  "X1275K7"
	IF assigned_to = "Christa J Burdette" or  assigned_to = "Christa Burdette" THEN worker_number =  "X127JS8"
	IF assigned_to = "Terry J Burgess" or  assigned_to = "Terry Burgess" or assigned_to = "Terry J Antonich Burgess" THEN worker_number =  "X127HN5"
	IF assigned_to = "Mariah R Burgess" or  assigned_to = "Mariah Burgess" THEN worker_number =  "X127MBR"
	IF assigned_to = "Neill C Burnett" or  assigned_to = "Neill Burnett" THEN worker_number =  "X127X99"
	IF assigned_to = "Sarah C Campbell" or  assigned_to = "Sarah Campbell" THEN worker_number =  "X127SCC"
	IF assigned_to = "Celeste E Carlson" or  assigned_to = "Celeste Carlson" THEN worker_number =  "X127C1T"
	IF assigned_to = "Sheryn L Cartlidge" or  assigned_to = "Sheryn Cartlidge" THEN worker_number =  "X127T63"
	IF assigned_to = "Lisa M Castile" or  assigned_to = "Lisa Castile" THEN worker_number =  "x127J9O"
	IF assigned_to = "Prophetia Castin" or  assigned_to = "Prophetia Castin" THEN worker_number =  "x127J9N"
	IF assigned_to = "Sim Chang" or  assigned_to = "Sim Chang" THEN worker_number =  "X127GG2"
	IF assigned_to = "Jacqueline Charpentier" or  assigned_to = "Jacqueline Charpentier" THEN worker_number =  "X127CF9"
	IF assigned_to = "Peggy Chavez" or  assigned_to = "Peggy Chavez" THEN worker_number =  "X1275H7"
	IF assigned_to = "Kevin Chavis" or  assigned_to = "Kevin Chavis" THEN worker_number =  "X127089"
	IF assigned_to = "Scott A Chestnut" or  assigned_to = "Scott Chestnut" THEN worker_number =  "X1273DC"
	IF assigned_to = "Ivy Chiinze" or  assigned_to = "Ivy Chiinze" THEN worker_number =  "X127JL7"
	IF assigned_to = "Wendy L Clark" or  assigned_to = "Wendy Clark" THEN worker_number =  "X127WLC"
	IF assigned_to = "Shantell Cochran" or  assigned_to = "Shantell Cochran" THEN worker_number =  "X127A8B"
	IF assigned_to = "Sherry L Collins" or  assigned_to = "Sherry Collins" THEN worker_number =  "X1275P1"
	IF assigned_to = "Carina S Cortez" or  assigned_to = "Carina Cortez" THEN worker_number =  "X127CSC"
	IF assigned_to = "Ikela I Cosey" or  assigned_to = "Ikela Cosey" THEN worker_number =  "x127J9Q"
	IF assigned_to = "Mayra J Cota" or  assigned_to = "Mayra Cota" THEN worker_number =  "x127J9R"
	IF assigned_to = "Sarai R Counce" or  assigned_to = "Sarai Counce" THEN worker_number =  "X127FAS"
	IF assigned_to = "Terri L Cox" or  assigned_to = "Terri Cox" THEN worker_number =  "X127B01"
	IF assigned_to = "Cornel C Culp" or  assigned_to = "Cornel Culp" THEN worker_number =  "X127J9S"
	IF assigned_to = "Miftah M Dadi" or  assigned_to = "Miftah Dadi" THEN worker_number =  "X127CB3"
	IF assigned_to = "Aisha A Dancy" or  assigned_to = "Aisha Dancy" THEN worker_number =  "X127KL7"
	IF assigned_to = "Amber Davis" or  assigned_to = "Amber Davis" THEN worker_number =  "X1275AD"
	IF assigned_to = "Elacia V Davis" or  assigned_to = "Elacia Davis" THEN worker_number =  "X127D8A"
	IF assigned_to = "Ann M Davis" or  assigned_to = "Ann Davis" THEN worker_number =  "X127E26"
	IF assigned_to = "Solange A Davis-Rivera" or  assigned_to = "Solange Davis-Rivera" THEN worker_number =  "X127GG3"
	IF assigned_to = "Cheryl A Deason" or  assigned_to = "Cheryl Deason" THEN worker_number =  "X127S08"
	IF assigned_to = "Kemal T Deko" or  assigned_to = "Kemal Deko" THEN worker_number =  "X127J8H"
	IF assigned_to = "Deanna L Deloach" or  assigned_to = "Deanna Deloach" THEN worker_number =  "X127X27"
	IF assigned_to = "Diana M Demario" or  assigned_to = "Diana Demario" THEN worker_number =  "X127C0Q"
	IF assigned_to = "Sheena A Dempsey" or  assigned_to = "Sheena Dempsey" THEN worker_number =  "X127SD4"
	IF assigned_to = "Beverly A Denman" or  assigned_to = "Beverly Denman" THEN worker_number =  "X127BD1"
	IF assigned_to = "Erick Diaz-Contreras" or  assigned_to = "Erick Diaz" THEN worker_number =  "X127KD5"
	IF assigned_to = "Jessica L Dickerson" or  assigned_to = "Jessica Dickerson" THEN worker_number =  "X1272EG"
	IF assigned_to = "Deborah Diggins" or  assigned_to = "Deborah Diggins" THEN worker_number =  "X127Y43"
	IF assigned_to = "Delia M Dilday" or  assigned_to = "Delia Dilday" THEN worker_number =  "X127B3T"
	IF assigned_to = "Sheryl J Dillenburg" or  assigned_to = "Sheryl Dillenburg" THEN worker_number =  "X1273DD"
	IF assigned_to = "Natalya Ditter" or  assigned_to = "Natalya Ditter" THEN worker_number =  "X127Y44"
	IF assigned_to = "Angela M Docken" or  assigned_to = "Angela Docken" THEN worker_number =  "X127JT3"
	IF assigned_to = "Kimyader M Dodd" or  assigned_to = "Kimyader Dodd" THEN worker_number =  "X127JM4"
	IF assigned_to = "Jonathan N Drogue" or  assigned_to = "Jonathan Drogue" THEN worker_number =  "X127J1D"
	IF assigned_to = "Sherry A Duggan" or  assigned_to = "Sherry Duggan" THEN worker_number =  "X1275K0"
	IF assigned_to = "Stacey Dunham" or  assigned_to = "Stacey Dunham" THEN worker_number =  "X127BM4"
	IF assigned_to = "DeAnne L Eberle" or  assigned_to = "DeAnne Eberle" THEN worker_number =  "X127B8K"
	IF assigned_to = "James B Eckard" or  assigned_to = "James Eckard" THEN worker_number =  "X1273T4"
	IF assigned_to = "Samantha E Haw" or  assigned_to = "Samantha Haw" THEN worker_number =  "X1272LJ"
	IF assigned_to = "Susan M Eeten" or  assigned_to = "Susan Eeten" THEN worker_number =  "X127GX9"
	IF assigned_to = "Christina M Eichorn" or  assigned_to = "Christina Eichorn" THEN worker_number =  "X127GF7"
	IF assigned_to = "Ayaan M Elmi" or  assigned_to = "Ayaan Elmi" THEN worker_number =  "X127AE1"
	IF assigned_to = "Olga V Engebretson" or  assigned_to = "Olga Engebretson" THEN worker_number =  "X127L04"
	IF assigned_to = "Sally R Engstrom" or  assigned_to = "Sally Engstrom" THEN worker_number =  "X1275L6"
	IF assigned_to = "Timothy B Erickson" or  assigned_to = "Timothy Erickson" THEN worker_number =  "X1276TE"
	IF assigned_to = "John M Fandrick" or  assigned_to = "John Fandrick" THEN worker_number =  "X127GAP"
	IF assigned_to = "Hodan K Farah" or  assigned_to = "Hodan Farah" THEN worker_number =  "X127B5A"
	IF assigned_to = "Ahmednor M Farah" or  assigned_to = "Ahmednor Farah" THEN worker_number =  "X127HS4"
	IF assigned_to = "Signe Faulhaber" or  assigned_to = "Signe Faulhaber" THEN worker_number =  "X127KL8"
	IF assigned_to = "Heather J Feldmann" or  assigned_to = "Heather Feldmann" THEN worker_number =  "X127ZAE"
	IF assigned_to = "Rachel A Ferguson" or  assigned_to = "Rachel Ferguson" THEN worker_number =  "X1272AF"
	IF assigned_to = "Shamilia Fisher" or  assigned_to = "Shamilia Fisher" THEN worker_number =  "X127SF1"
	IF assigned_to = "Katie M Flanigan" or  assigned_to = "Katie Flanigan" THEN worker_number =  "X127KL9"
	IF assigned_to = "Kelly C Flanigan" or  assigned_to = "Kelly Flanigan" THEN worker_number =  "X127T50"
	IF assigned_to = "Jodynne D Flasch" or  assigned_to = "Jodynne Flasch" THEN worker_number =  "X127B7G"
	IF assigned_to = "Melissa A Flores" or  assigned_to = "Melissa Flores" THEN worker_number =  "X1272BD"
	IF assigned_to = "Emily J Frazier" or  assigned_to = "Emily Frazier" THEN worker_number =  "X1270A1"
	IF assigned_to = "Iyana R Galloway" or  assigned_to = "Iyana Galloway" THEN worker_number =  "X127HN8"
	IF assigned_to = "Fatiya A Ganamo" or  assigned_to = "Fatiya Ganamo" THEN worker_number =  "X127FGA"
	IF assigned_to = "Gina T Gangelhoff" or  assigned_to = "Gina Gangelhoff" THEN worker_number =  "X127A7O"
	IF assigned_to = "Aaron J Gardner-Kocher" or assigned_to = "Aaron Gardner-Kocher" or assigned_to = "Aaron J Gardner" THEN worker_number =  "X1275F9"
	IF assigned_to = "Kenneth W Garnier" or  assigned_to = "Kenneth Garnier" THEN worker_number =  "X1271AJ"
	IF assigned_to = "Gifti Geleta" or  assigned_to = "Gifti Geleta" THEN worker_number =  "X127090"
	IF assigned_to = "Debra E George" or  assigned_to = "Debra George" THEN worker_number =  "X1275G0"
	IF assigned_to = "Bernardo G Gonzalez" or  assigned_to = "Bernardo Gonzalez" THEN worker_number =  "X127A8F"
	IF assigned_to = "Marlenne Gonzalez" or  assigned_to = "Marlenne Gonzalez" THEN worker_number =  "X127GJ3"
	IF assigned_to = "Olga L Gonzalez" or  assigned_to = "Olga Gonzalez" THEN worker_number =  "x127J9V"
	IF assigned_to = "Tracy A Gorman" or  assigned_to = "Tracy Gorman" THEN worker_number =  "X127N54"
	IF assigned_to = "Penny R Grady" or  assigned_to = "Penny Grady" THEN worker_number =  "X127A6X"
	IF assigned_to = "Jacqueline S Graves" or  assigned_to = "Jacqueline Graves" THEN worker_number =  "X127A9U"
	IF assigned_to = "Eve Gray" or  assigned_to = "Eve Gray" THEN worker_number =  "X1276EG"
	IF assigned_to = "Andrea D Green" or  assigned_to = "Andrea Green" THEN worker_number =  "X127ADG"
	IF assigned_to = "Linda L Greene" or  assigned_to = "Linda Greene" THEN worker_number =  "X127436"
	IF assigned_to = "Josefina G Greene" or  assigned_to = "Josefina Greene" or assigned_to = "Josefina Valverde" THEN worker_number =  "X127KE7"
	IF assigned_to = "Shaneka Greer" or  assigned_to = "Shaneka Greer" THEN worker_number =  "X127HN9"
	IF assigned_to = "Lyvia C Guallpa" or  assigned_to = "Lyvia Guallpa" THEN worker_number =  "X127KM3"
	IF assigned_to = "Huruse M Gurhan" or  assigned_to = "Huruse Gurhan" THEN worker_number =  "X127Z93"
	IF assigned_to = "Diane Ha" or  assigned_to = "Diane Ha" THEN worker_number =  "X1275N1"
	IF assigned_to = "Madar A Hachi" or  assigned_to = "Madar Hachi" THEN worker_number =  "X127AL8"
	IF assigned_to = "Sarah A Haigh" or  assigned_to = "Sarah Haigh" THEN worker_number =  "X127JX9"
	IF assigned_to = "Safiyo A Haji" or  assigned_to = "Safiyo Haji" THEN worker_number =  "X1275P6"
	IF assigned_to = "Jessica A Hall" or  assigned_to = "Jessica Hall" THEN worker_number =  "X127X05"
	IF assigned_to = "Cynthia Hampton" or  assigned_to = "Cynthia Hampton" THEN worker_number =  "X127F19"
	IF assigned_to = "MiKayla Handley" or  assigned_to = "MiKayla Handley" THEN worker_number =  "X127D5X"
	IF assigned_to = "Tamika Hannah" or  assigned_to = "Tamika Hannah" THEN worker_number =  "X127ZAH"
	IF assigned_to = "Shanna C Hansen" or  assigned_to = "Shanna Hansen" THEN worker_number =  "X127Q95"
	IF assigned_to = "Star A Hanson" or  assigned_to = "Star Hanson" THEN worker_number =  "X127BW6"
	IF assigned_to = "Maria J Harald" or  assigned_to = "Maria Harald" THEN worker_number =  "X127JO1"
	IF assigned_to = "Sara K Harrell" or  assigned_to = "Sara Harrell" THEN worker_number =  "X127Z71"
	IF assigned_to = "Inger M Harris" or  assigned_to = "Inger Harris" THEN worker_number =  "X1275N9"
	IF assigned_to = "Jessica R Harris" or  assigned_to = "Jessica Harris" THEN worker_number =  "X127JH1"
	IF assigned_to = "Shaquila Harris" or  assigned_to = "Shaquila Harris" THEN worker_number =  "X127W40"
	IF assigned_to = "Molly C Hasbrook" or  assigned_to = "Molly Hasbrook" THEN worker_number =  "X127B9Q"
	IF assigned_to = "Alisa R Haselhorst" or  assigned_to = "Alisa Haselhorst" THEN worker_number =  "X127AN6"
	IF assigned_to = "Farah A Hassan" or  assigned_to = "Farah Hassan" THEN worker_number =  "X1271FH"
	IF assigned_to = "Sartu A Hassan" or  assigned_to = "Sartu Hassan" THEN worker_number =  "X127AM5"
	IF assigned_to = "LaRae Heard" or  assigned_to = "LaRae Heard" THEN worker_number =  "X127KM4"
	IF assigned_to = "Trenita M Heard" or  assigned_to = "Trenita Heard" THEN worker_number =  "X127T67"
	IF assigned_to = "Patricia A Hegenbarth" or  assigned_to = "Patricia Hegenbarth" THEN worker_number =  "X127Y37"
	IF assigned_to = "Cheryl L Heitzinger" or  assigned_to = "Cheryl Heitzinger" THEN worker_number =  "X12728S"
	IF assigned_to = "Lauren A John" or  assigned_to = "Lauren John" THEN worker_number =  "X127ZAJ"
	IF assigned_to = "Rachel A Henry" or  assigned_to = "Rachel Henry" THEN worker_number =  "X127J8E"
	IF assigned_to = "Crystal M Henry-Bolden" or  assigned_to = "Crystal Henry-Bolden" THEN worker_number =  "X127CA2"
	IF assigned_to = "Tony Her" or assigned_to = "Tony Her" THEN worker_number =  "X1275P2"
	IF assigned_to = "Vila Her" or assigned_to = "Vila Her" THEN worker_number =  "X127D3E"
	IF assigned_to = "Cha Her" or  assigned_to = "Cha Her" THEN worker_number =  "X127JW4"
	IF assigned_to = "Abdifitaah Herei" or  assigned_to = "Abdifitaah Herei" THEN worker_number =  "X127KD7"
	IF assigned_to = "Valerie M Herrera" or  assigned_to = "Valerie Herrera" THEN worker_number =  "X1275H5"
	IF assigned_to = "Janell L Hill" or  assigned_to = "Janell Hill" THEN worker_number =  "X127AM9"
	IF assigned_to = "Kimberly A Hill" or  assigned_to = "Kimberly Hill" THEN worker_number =  "X127B18"
	IF assigned_to = "Cecelia M Hoecherl" or  assigned_to = "Cecelia Hoecherl" THEN worker_number =  "X127A05"
	IF assigned_to = "Nailah Y Holman" or  assigned_to = "Nailah Holman" THEN worker_number =  "X127AQ5"
	IF assigned_to = "Stephanie L Holmes" or  assigned_to = "Stephanie Holmes" THEN worker_number =  "X1273M7"
	IF assigned_to = "John D Holmquist" or  assigned_to = "John Holmquist" THEN worker_number =  "X127JB2"
	IF assigned_to = "Kasey A Holt" or  assigned_to = "Kasey Holt" THEN worker_number =  "X1271A7"
	IF assigned_to = "Jenny Hong" or  assigned_to = "Jenny Hong" THEN worker_number =  "X127JH3"
	IF assigned_to = "Tasheema Hopson" or  assigned_to = "Tasheema Hopson" THEN worker_number =  "X1272UE"
	IF assigned_to = "Rhonda V Hopson" or  assigned_to = "Rhonda Hopson" THEN worker_number =  "X127BL8"
	IF assigned_to = "Andrew J Howard" or  assigned_to = "Andrew Howard" THEN worker_number =  "X127KM5"
	IF assigned_to = "Roberta J Howard" or  assigned_to = "Roberta Howard" THEN worker_number =  "X127RJH"
	IF assigned_to = "Juanita M Hubbard" or  assigned_to = "Juanita Hubbard" THEN worker_number =  "X127CB8"
	IF assigned_to = "Janine M Hudson" or  assigned_to = "Janine Hudson" THEN worker_number =  "X127JH4"
	IF assigned_to = "Remy K Huerta-Stemper" or  assigned_to = "Remy Huerta-Stemper" THEN worker_number =  "X1275H0"
	IF assigned_to = "Abdiaziz M Hurreh" or  assigned_to = "Abdiaziz Hurreh" THEN worker_number =  "X1272CZ"
	IF assigned_to = "Valerie Hurst-Baker" or  assigned_to = "Valerie Hurst-Baker" THEN worker_number =  "X127JM6"
	IF assigned_to = "Abdirizak M Ibrahim" or  assigned_to = "Abdirizak Ibrahim" THEN worker_number =  "X127B9P"
	IF assigned_to = "Molly Irwin" or  assigned_to = "Molly Irwin" THEN worker_number =  "X127GJ4"
	IF assigned_to = "Wendy S Irwin" or  assigned_to = "Wendy Irwin" THEN worker_number =  "X127WI1"
	IF assigned_to = "Zaki M Isaac" or  assigned_to = "Zaki Isaac" or assigned_to = "Zechariye M Isaac"  THEN worker_number =  "X127AH3"
	IF assigned_to = "Melissa M Isais" or  assigned_to = "Melissa Isais" THEN worker_number =  "X127X29"
	IF assigned_to = "Zoye D Jackson" or  assigned_to = "Zoye Jackson" THEN worker_number =  "X127160"
	IF assigned_to = "Debrice L Jackson" or  assigned_to = "Debrice Jackson" THEN worker_number =  "X1274QG"
	IF assigned_to = "Charnice W Jackson" or  assigned_to = "Charnice Jackson" THEN worker_number =  "x127J9X"
	IF assigned_to = "Breauna M Jackson" or  assigned_to = "Breauna Jackson" THEN worker_number =  "X127JM7"
	IF assigned_to = "Mark P Jacobson" or  assigned_to = "Mark Jacobson" THEN worker_number =  "X127C1C"
	IF assigned_to = "Jamila A Jama" or  assigned_to = "Jamila Jama" THEN worker_number =  "X127Y76"
	IF assigned_to = "Sainabou Jaye-Marong" or  assigned_to = "Sainabou Jaye-Marong" THEN worker_number =  "X127KP4"
	IF assigned_to = "Stephanie A Jefferson" or  assigned_to = "Stephanie Jefferson" THEN worker_number =  "X127A7M"
	IF assigned_to = "Toni Jenkins" or  assigned_to = "Toni Jenkins" THEN worker_number =  "X127B1B"
	IF assigned_to = "Antionette L Jenkins" or  assigned_to = "Antionette Jenkins" or assigned_to = "Toni Jenkins" THEN worker_number =  "X127J8I"
	IF assigned_to = "Christine Jernander" or  assigned_to = "Christine Jernander" THEN worker_number =  "X127B36"
	IF assigned_to = "Saeed D Jibrell" or  assigned_to = "Saeed Jibrell" THEN worker_number =  "X127L87"
	IF assigned_to = "Candace S Johnson" or  assigned_to = "Candace Johnson" THEN worker_number =  "X1275L7"
	IF assigned_to = "Maria L Johnson" or  assigned_to = "Maria Johnson" THEN worker_number =  "X1276MJ"
	IF assigned_to = "Cindy Johnson" or  assigned_to = "Cindy Johnson" THEN worker_number =  "X127CDJ"
	IF assigned_to = "Nina Jones" or  assigned_to = "Nina Jones" THEN worker_number =  "X127R49"
	IF assigned_to = "Sandy L Jorgensen" or  assigned_to = "Sandy Jorgensen" THEN worker_number =  "X127KM7"
	IF assigned_to = "Celeste Jourdain" or  assigned_to = "Celeste Jourdain" THEN worker_number =  "X1273D3"
	IF assigned_to = "Svetlana Kabakova" or  assigned_to = "Svetlana Kabakova" THEN worker_number =  "X12730W"
	IF assigned_to = "Ziyad Z Kadir" or  assigned_to = "Ziyad Kadir" THEN worker_number =  "X127FAL"
	IF assigned_to = "Gawa T Kalsang" or  assigned_to = "Gawa Kalsang" THEN worker_number =  "x127J9Y"
	IF assigned_to = "Maggie Karley" or  assigned_to = "Maggie Karley" THEN worker_number =  "X127J8J"
	IF assigned_to = "Kristine Karlsgodt" or  assigned_to = "Kristine Karlsgodt" THEN worker_number =  "X12746I"
	IF assigned_to = "Kristen F Kasim" or  assigned_to = "Kristen Kasim" THEN worker_number =  "X127A2D"
	IF assigned_to = "Abba Bora H Kedir" or  assigned_to = "Abba Bora" THEN worker_number =  "X127D4E"
	IF assigned_to = "Amy N Kelvie" or  assigned_to = "Amy Kelvie" THEN worker_number =  "X127ZAK"
	IF assigned_to = "DeNise Kendrick-Stevens" or  assigned_to = "DeNise Kendrick-Stevens" THEN worker_number =  "X1275G8"
	IF assigned_to = "Debra M Kennedy" or  assigned_to = "Debra Kennedy" THEN worker_number =  "X127B2L"
	IF assigned_to = "Cheryl K Kerzman" or  assigned_to = "Cheryl Kerzman" THEN worker_number =  "X127F77"
	IF assigned_to = "Veronica L Keys" or  assigned_to = "Veronica Keys" THEN worker_number =  "X127VLk"
	IF assigned_to = "Ryan P Kierth" or  assigned_to = "Ryan Kierth" THEN worker_number =  "X127AP7"
	IF assigned_to = "Viola L Kill" or  assigned_to = "Viola Kill" THEN worker_number =  "X1272BR"
	IF assigned_to = "Ronick Kimnong" or  assigned_to = "Ronick Kimnong" THEN worker_number =  "X127BK8"
	IF assigned_to = "Sylvia A King" or  assigned_to = "Sylvia King" THEN worker_number =  "X127B4Q"
	IF assigned_to = "Louise A Kinzer" or  assigned_to = "Louise Kinzer" THEN worker_number =  "X127K85"
	IF assigned_to = "Shauna Kirscht" or  assigned_to = "Shauna Kirscht" THEN worker_number =  "X127KP7"
	IF assigned_to = "Andy A Knutson" or  assigned_to = "Andy Knutson" THEN worker_number =  "X127JS7"
	IF assigned_to = "Darren Konsor" or  assigned_to = "Darren Konsor" THEN worker_number =  "X127ZAL"
	IF assigned_to = "Abby Korenchen" or  assigned_to = "Abby Korenchen" THEN worker_number =  "X127HS6"
	IF assigned_to = "Abdo W Korosso" or  assigned_to = "Abdo Korosso" THEN worker_number =  "X12729T"
	IF assigned_to = "Raeann T Korynta" or  assigned_to = "Raeann Korynta" THEN worker_number =  "X127CF3"
	IF assigned_to = "Kris Koukkari" or  assigned_to = "Kris Koukkari" THEN worker_number =  "X127KEK"
	IF assigned_to = "Nikolai Kravets" or  assigned_to = "Nikolai Kravets" THEN worker_number =  "X127HT5"
	IF assigned_to = "Sarah E LaCoursiere" or  assigned_to = "Sarah LaCoursiere" THEN worker_number =  "X127SEL"
	IF assigned_to = "Colanda R Lacy" or  assigned_to = "Colanda Lacy" THEN worker_number =  "X1275M9"
	IF assigned_to = "Denis L Ladeyshchikov" or  assigned_to = "Denis Ladeyshchikov" THEN worker_number =  "X127AL5"
	IF assigned_to = "Grecia M Lagunes" or  assigned_to = "Grecia Lagunes" THEN worker_number =  "X127KD8"
	IF assigned_to = "Lisa M Lampkin" or  assigned_to = "Lisa Lampkin" THEN worker_number =  "X127LL1"
	IF assigned_to = "Brittany M Lane" or  assigned_to = "Brittany Lane" THEN worker_number =  "X127140"
	IF assigned_to = "Melinda L Lane" or  assigned_to = "Melinda Lane" THEN worker_number =  "X1273FL"
	IF assigned_to = "Rochelle C Lane" or  assigned_to = "Rochelle Lane" THEN worker_number =  "X127HP3"
	IF assigned_to = "Matthew M Lane" or  assigned_to = "Matthew Lane" THEN worker_number =  "X127JD5"
	IF assigned_to = "Kaeli F Larson" or  assigned_to = "Kaeli Larson" THEN worker_number =  "X127AW8"
	IF assigned_to = "Jaime Lavallee" or  assigned_to = "Jaime Lavallee" THEN worker_number =  "X127J1L"
	IF assigned_to = "Andrea S Lawrence" or  assigned_to = "Andrea Lawrence" THEN worker_number =  "X127D05"
	IF assigned_to = "Michelle Le" or  assigned_to = "Michelle " THEN worker_number =  "X127J75"
	IF assigned_to = "Deborah A Lechner" or  assigned_to = "Deborah Lechner" THEN worker_number =  "X127W88"
	IF assigned_to = "Pa Nhia Lee" or  assigned_to = "Pa Nhia Lee" THEN worker_number =  "X1275K2"
	IF assigned_to = "Linda - Lee" or  assigned_to = "Linda Lee"  THEN worker_number =  "X127D2F"
	IF assigned_to = "Payeng Lee" or  assigned_to = "Payeng Lee" THEN worker_number =  "X127D4H"
	IF assigned_to = "Mai C Lee" or  assigned_to = "Mai Lee" THEN worker_number =  "X127D4R"
	IF assigned_to = "Chao Lee" or  assigned_to = "Chao Lee" THEN worker_number =  "X127J8C"
	IF assigned_to = "Mai V Lee" or  assigned_to = "Mai Lee" THEN worker_number =  "X127JD6"
	IF assigned_to = "Bee Lee" or  assigned_to = "Bee Lee" THEN worker_number =  "X127JT7"
	IF assigned_to = "Kia Lee" or  assigned_to = "Kia Lee" THEN worker_number =  "X127Z86"
	IF assigned_to = "Xay L Lee-Xiong" or  assigned_to = "Xay Lee-Xiong" or assigned_to = "Xay L Lee" THEN worker_number =  "X127F23"
	IF assigned_to = "Shamikka S Lenear" or  assigned_to = "Shamikka Lenear" THEN worker_number =  "X127D7R"
	IF assigned_to = "Letitia Lewis" or  assigned_to = "Letitia Lewis" THEN worker_number =  "X127Y92"
	IF assigned_to = "Genni M Lillibridge" or  assigned_to = "Genni Lillibridge" THEN worker_number =  "X127KL4"
	IF assigned_to = "Shelly A Lind" or  assigned_to = "Shelly Lind" THEN worker_number =  "X127Y81"
	IF assigned_to = "True P Lis" or  assigned_to = "TRUE Lis" THEN worker_number =  "X1272BM"
	IF assigned_to = "Cassandra M Lis" or  assigned_to = "Cassandra Lis" THEN worker_number =  "X127Y05"
	IF assigned_to = "Becky A Little" or  assigned_to = "Becky Little" THEN worker_number =  "X127FBX"
	IF assigned_to = "Raisa D Loevski" or  assigned_to = "Raisa Loevski" THEN worker_number =  "X127Z34"
	IF assigned_to = "Nas Looper" or  assigned_to = "Nas Looper" THEN worker_number =  "X127FAW"
	IF assigned_to = "Sarita M Lopez" or  assigned_to = "Sarita Lopez" THEN worker_number =  "X127CA9"
	IF assigned_to = "Teng L Lor" or  assigned_to = "Teng Lor" THEN worker_number =  "X1275K5"
	IF assigned_to = "Mali Lor" or  assigned_to = "Mali Lor" THEN worker_number =  "X127GR8"
	IF assigned_to = "Casey H Love" or  assigned_to = "Casey Love" THEN worker_number =  "X127L1S"
	IF assigned_to = "Amber Lowe" or  assigned_to = "Amber Lowe" THEN worker_number =  "X127JD8"
	IF assigned_to = "Carrie A Lucca" or  assigned_to = "Carrie Lucca" THEN worker_number =  "X127CAL"
	IF assigned_to = "Michelle L Lungelow" or  assigned_to = "Michelle Lungelow" THEN worker_number =  "X127B6J"
	IF assigned_to = "Pajci Ly" or  assigned_to = "Pajci Ly" THEN worker_number =  "X127JX2"
	IF assigned_to = "Yanisha K Mack" or  assigned_to = "Yanisha Mack" THEN worker_number =  "X127D7W"
	IF assigned_to = "Ashley K Mack" or  assigned_to = "Ashley Mack" THEN worker_number =  "X127GH5"
	IF assigned_to = "Paul E Madison" or  assigned_to = "Paul Madison" THEN worker_number =  "X127L23"
	IF assigned_to = "Ramona M Shane" or  assigned_to = "Ramona Shane" THEN worker_number =  "X1272PC"
	IF assigned_to = "Hind S Mahmoud" or  assigned_to = "Hind Mahmoud" THEN worker_number =  "X127HM1"
	IF assigned_to = "Florence A Manley" or  assigned_to = "Florence Manley" THEN worker_number =  "X127966"
	IF assigned_to = "Molly M Manley" or  assigned_to = "Molly Manley" THEN worker_number =  "X127AG4"
	IF assigned_to = "Vickie Mansheim" or  assigned_to = "Vickie Mansheim" THEN worker_number =  "X127VSM"
	IF assigned_to = "Rashida R Manuel" or  assigned_to = "Rashida Manuel" THEN worker_number =  "X1275H9"
	IF assigned_to = "Fawn Marquez" or  assigned_to = "Fawn Marquez" THEN worker_number =  "X127BN9"
	IF assigned_to = "Deja L Martin" or  assigned_to = "Deja Martin" THEN worker_number =  "X127JX5"
	IF assigned_to = "Iliana E Martinez Morales" or  assigned_to = "Iliana Martinez" THEN worker_number =  "X127FBQ"
	IF assigned_to = "Jason A Marx" or  assigned_to = "Jason Marx" THEN worker_number =  "X12729W"
	IF assigned_to = "Angela E Masiello" or  assigned_to = "Angela Masiello" THEN worker_number =  "X127GG5"
	IF assigned_to = "Alejandra Andrade" or  assigned_to = "Alejandra Andrade" THEN worker_number =  "X1275J1"
	IF assigned_to = "Amy M McCall" or  assigned_to = "Amy McCall" THEN worker_number =  "X127130"
	IF assigned_to = "Charice McDowell" or  assigned_to = "Charice McDowell" THEN worker_number =  "X127A6S"
	IF assigned_to = "Matthew M McGovern" or  assigned_to = "Matthew McGovern" THEN worker_number =  "X1274HQ"
	IF assigned_to = "Caire Mckenzie" or  assigned_to = "Caire Mckenzie" THEN worker_number =  "X127HP8"
	IF assigned_to = "Russell D Meelberg" or  assigned_to = "Russell Meelberg" THEN worker_number =  "X127GR9"
	IF assigned_to = "Jennifer A Merritt" or  assigned_to = "Jennifer Merritt" THEN worker_number =  "X127JAC"
	IF assigned_to = "Lara M Messer" or  assigned_to = "Lara Messer" THEN worker_number =  "X1275L8"
	IF assigned_to = "Jacqueline W Miantona" or  assigned_to = "Jacqueline Miantona" THEN worker_number =  "X127BL4"
	IF assigned_to = "Filmon K Michael" or  assigned_to = "Filmon Michael" THEN worker_number =  "X127GG6"
	IF assigned_to = "Toni S Miles" or  assigned_to = "Toni Miles" THEN worker_number =  "X127GR4"
	IF assigned_to = "Sara R Miller" or  assigned_to = "Sara Miller" THEN worker_number =  "X127CA0"
	IF assigned_to = "Deedra C Miller" or  assigned_to = "Deedra Miller" THEN worker_number =  "X127X63"
	IF assigned_to = "Linda K Millhouse" or  assigned_to = "Linda Millhouse" THEN worker_number =  "X127E60"
	IF assigned_to = "Marianne E Simon" or  assigned_to = "Marianne Simon" THEN worker_number =  "X1275G3"
	IF assigned_to = "Alisha E Mitchell" or  assigned_to = "Alisha Mitchell" THEN worker_number =  "X127HP9"
	IF assigned_to = "Samsam A Mohamed" or  assigned_to = "Samsam Mohamed" THEN worker_number =  "X127HS7"
	IF assigned_to = "Basma A Mohamed" or  assigned_to = "Basma Mohamed" THEN worker_number =  "X127JK8"
	IF assigned_to = "Irro A Mohamed" or  assigned_to = "Irro Mohamed" THEN worker_number =  "X127X96"
	IF assigned_to = "Tracy L Mohomes" or  assigned_to = "Tracy Mohomes" THEN worker_number =  "X127Y23"
	IF assigned_to = "Thomas A Moore" or  assigned_to = "Thomas Moore" THEN worker_number =  "X1275F4"
	IF assigned_to = "Stephen S Moore" or  assigned_to = "Stephen Moore" THEN worker_number =  "X127D2T"
	IF assigned_to = "Dave Mootz" or  assigned_to = "Dave Mootz" THEN worker_number =  "X127GH6"
	IF assigned_to = "Ana K Moreno De La Garza" or  assigned_to = "Ana Moreno" THEN worker_number =  "X127KE1"
	IF assigned_to = "Teresa D Morphew" or  assigned_to = "Teresa Morphew" THEN worker_number =  "X127AQ8"
	IF assigned_to = "Jennifer K Moses" or  assigned_to = "Jennifer Moses" THEN worker_number =  "X127Y62"
	IF assigned_to = "Mailee C Moua" or  assigned_to = "Mailee Moua" THEN worker_number =  "X127A9X"
	IF assigned_to = "Tiffanie Mrsich" or  assigned_to = "Tiffanie Mrsich" THEN worker_number =  "X127X82"
	IF assigned_to = "Mai-Ling Mui" or  assigned_to = "Mai-Ling Mui" THEN worker_number =  "X127JV7"
	IF assigned_to = "Sharon Murphy" or  assigned_to = "Sharon Murphy" THEN worker_number =  "X127A0F"
	IF assigned_to = "Jerry Nack" or  assigned_to = "Jerry Nack" THEN worker_number =  "X1275G4"
	IF assigned_to = "Zachary E Nagle" or  assigned_to = "Zachary Nagle" THEN worker_number =  "X127JX7"
	IF assigned_to = "Maribel Navarrete Reyes" or  assigned_to = "Maribel Navarrete" THEN worker_number =  "X127K9G"
	IF assigned_to = "Ephrem Nejo" or  assigned_to = "Ephrem Nejo" THEN worker_number =  "X127X04"
	IF assigned_to = "Lisa A Nelson" or  assigned_to = "Lisa Nelson" THEN worker_number =  "X127K82"
	IF assigned_to = "Joseph Nelson" or  assigned_to = "Joseph Nelson" THEN worker_number =  "X127X97"
	IF assigned_to = "Lamar Salinas-Niemczycki" or  assigned_to = "Lamar Salinas-Niemczycki" THEN worker_number =  "X127LN1"
	IF assigned_to = "John B Niemi" or  assigned_to = "John Niemi" THEN worker_number =  "X127JU5"
	IF assigned_to = "Jill I Niess" or  assigned_to = "Jill Niess" THEN worker_number =  "X127HJ9"
	IF assigned_to = "Sideth D Niev" or  assigned_to = "Sideth Niev" THEN worker_number =  "X127U55"
	IF assigned_to = "Todd A Norling" or  assigned_to = "Todd Norling" THEN worker_number =  "X127085"
	IF assigned_to = "Kristine R Norman" or  assigned_to = "Kristine Norman" THEN worker_number =  "X127Z83"
	IF assigned_to = "Mohamed O Nur" or  assigned_to = "Mohamed Nur" THEN worker_number =  "X1272AY"
	IF assigned_to = "Khadra M Nur" or  assigned_to = "Khadra Nur" THEN worker_number =  "X127KMN"
	IF assigned_to = "Autumn J O'Brien" or  assigned_to = "Autumn O'Brien" THEN worker_number =  "X127JU6"
	IF assigned_to = "Nicole Ocampo" or  assigned_to = "Nicole Ocampo" THEN worker_number =  "X127JV6"
	IF assigned_to = "Jodi R Ojala" or  assigned_to = "Jodi Ojala" THEN worker_number =  "X127K9D"
	IF assigned_to = "Laura A Olson" or  assigned_to = "Laura Olson" THEN worker_number =  "X127AX4"
	IF assigned_to = "Brian D Olson" or  assigned_to = "Brian Olson" THEN worker_number =  "X127B22"
	IF assigned_to = "Christy P Olson" or  assigned_to = "Christy Olson" THEN worker_number =  "X127G50"
	IF assigned_to = "Hangatu Omar" or  assigned_to = "Hangatu Omar" THEN worker_number =  "X127JX8"
	IF assigned_to = "Abraham T Page" or  assigned_to = "Abraham Page" THEN worker_number =  "X127KM9"
	IF assigned_to = "Michelle E Parenteau" or  assigned_to = "Michelle Parenteau" THEN worker_number =  "X127AH2"
	IF assigned_to = "Tiwana L Pargo" or  assigned_to = "Tiwana Pargo" THEN worker_number =  "X127FAM"
	IF assigned_to = "Monica F Parham" or  assigned_to = "Monica Parham" THEN worker_number =  "X127MFP"
	IF assigned_to = "Giovanni E Parodi" or  assigned_to = "Giovanni Parodi" THEN worker_number =  "X127GEP"
	IF assigned_to = "Tanya L Payne" or  assigned_to = "Tanya Payne" THEN worker_number =  "X127FCA"
	IF assigned_to = "Dickyi Peldon" or  assigned_to = "Dickyi Peldon" THEN worker_number =  "X127DP7"
	IF assigned_to = "Gloria Perez Amastal" or  assigned_to = "Gloria Perez" THEN worker_number =  "X127GDP"
	IF assigned_to = "Claudia Perez Selva de Heintz" or  assigned_to = "Claudia Perez" THEN worker_number =  "X127FBJ"
	IF assigned_to = "Lakisha P Perkerson" or  assigned_to = "Lakisha Perkerson" THEN worker_number =  "X127A0W"
	IF assigned_to = "Diana Peterson" or  assigned_to = "Diana Peterson" THEN worker_number =  "X127BK2"
	IF assigned_to = "Sheri L Peterson" or  assigned_to = "Sheri Peterson" THEN worker_number =  "X127D4X"
	IF assigned_to = "Rita M Phelps" or  assigned_to = "Rita Phelps" THEN worker_number =  "X127S01"
	IF assigned_to = "Diki Wangkhang-Phuntsok" or  assigned_to = "Diki Phunts" THEN worker_number =  "X127K9E"
	IF assigned_to = "Kristina Poplavska" or  assigned_to = "Kristina Poplavska" THEN worker_number =  "X127GR2"
	IF assigned_to = "Azza D Pratiwi" or  assigned_to = "Azza Pratiwi" THEN worker_number =  "X1275N7"
	IF assigned_to = "Michelle Pringle" or  assigned_to = "Michelle Pringle" THEN worker_number =  "X127C1Y"
	IF assigned_to = "Kelly A Quigley" or  assigned_to = "Kelly Quigley" THEN worker_number =  "X127K9F"
	IF assigned_to = "Dayanne Quinonez" or  assigned_to = "Dayanne Quinonez" THEN worker_number =  "X127DQ1"
	IF assigned_to = "Brenda M Raygor" or  assigned_to = "Brenda Raygor" THEN worker_number =  "X12746F"
	IF assigned_to = "Mary D Reeck" or  assigned_to = "Mary Reeck" THEN worker_number =  "X127927"
	IF assigned_to = "Jonathan Reeck" or  assigned_to = "Jonathan Reeck" THEN worker_number =  "X127JF1"
	IF assigned_to = "Brooke A Reilley" or  assigned_to = "Brooke Reilley" THEN worker_number =  "X1272HC"
	IF assigned_to = "Timothy J Remme" or  assigned_to = "Timothy Remme" THEN worker_number =  "X127TJR"
	IF assigned_to = "Gary M Remus" or  assigned_to = "Gary Remus" THEN worker_number =  "X127A79"
	IF assigned_to = "Lindsey H Remus" or  assigned_to = "Lindsey Remus" THEN worker_number =  "X127T25"
	IF assigned_to = "Maria Remy" or  assigned_to = "Maria Remy" THEN worker_number =  "X1275J3"
	IF assigned_to = "Laura L Riebe" or  assigned_to = "Laura Riebe" THEN worker_number =  "X127D4Y"
	IF assigned_to = "Barbara Herrera" or  assigned_to = "Barbara Herrera" THEN worker_number =  "X127M22"
	IF assigned_to = "Amorette B Robeck" or  assigned_to = "Amorette Robeck" THEN worker_number =  "X1272ID"
	IF assigned_to = "Lori A Roberson" or  assigned_to = "Lori Roberson" THEN worker_number =  "X127AH4"
	IF assigned_to = "Loranzie S Rogers" or  assigned_to = "Loranzie Rogers" THEN worker_number =  "X127LSR"
	IF assigned_to = "Malinda M Rolack" or  assigned_to = "Malinda Rolack" THEN worker_number =  "X127W12"
	IF assigned_to = "Brittney N Ross" or  assigned_to = "Brittney Ross" THEN worker_number =  "X127D7Z"
	IF assigned_to = "Fabio A Rozo" or  assigned_to = "Fabio Rozo" THEN worker_number =  "X127A1T"
	IF assigned_to = "Dan Rubenstein" or  assigned_to = "Dan Rubenstein" THEN worker_number =  "X127JU8"
	IF assigned_to = "Marnya Rudolph" or  assigned_to = "Marnya Rudolph" THEN worker_number =  "X127MMR"
	IF assigned_to = "Anastacia R Ruiz" or  assigned_to = "Anastacia Ruiz" THEN worker_number =  "X127HQ3"
	IF assigned_to = "Deborah L Rusnak" or  assigned_to = "Deborah Rusnak" THEN worker_number =  "X127L08"
	IF assigned_to = "Victoria Rutkovskaya" or  assigned_to = "Victoria Rutkovskaya" THEN worker_number =  "X12726L"
	IF assigned_to = "Nicole Ryan" or  assigned_to = "Nicole Ryan" THEN worker_number =  "X127B2Z"
	IF assigned_to = "Alyssa A Ryan" or  assigned_to = "Alyssa Ryan" THEN worker_number =  "X127W23"
	IF assigned_to = "Alexandra G Guardado Saenz" or  assigned_to = "Alexandra Guardado"  or assigned_to = "Alexandra Guardado Saenz" THEN worker_number =  "X127HT1"
	IF assigned_to = "Iman Said" or  assigned_to = "Iman Said" THEN worker_number =  "X127KE2"
	IF assigned_to = "Miguel A Salazar" or  assigned_to = "Miguel Salazar" THEN worker_number =  "X127A3V"
	IF assigned_to = "Darlenne Salinas-Fernandez" or  assigned_to = "Darlenne Salinas Fernandez" THEN worker_number =  "X1272BB"
	IF assigned_to = "Yoshauna Sampson" or  assigned_to = "Yoshauna Sampson" THEN worker_number =  "X1276YS"
	IF assigned_to = "Jessica L Sanderson" or  assigned_to = "Jessica Sanderson" THEN worker_number =  "X127BW7"
	IF assigned_to = "Sahr K Sandi" or  assigned_to = "Sahr Sandi" THEN worker_number =  "X127Q90"
	IF assigned_to = "Karina Santana" or  assigned_to = "Karina Santana" THEN worker_number =  "X127KE4"
	IF assigned_to = "Soumya G Sanyal" or  assigned_to = "Soumya Sanyal" THEN worker_number =  "X127A7S"
	IF assigned_to = "Sitha Sarin" or  assigned_to = "Sitha Sarin" THEN worker_number =  "X1272B1"
	IF assigned_to = "Claudia A Saulter" or  assigned_to = "Claudia Saulter" THEN worker_number =  "X127T65"
	IF assigned_to = "Clarita B Scherer" or  assigned_to = "Clarita Scherer" THEN worker_number =  "X12746P"
	IF assigned_to = "Mark X Schmidt" or  assigned_to = "Mark Schmidt" THEN worker_number =  "X127KE5"
	IF assigned_to = "Angela A Schottle" or  assigned_to = "Angela Schottle" THEN worker_number =  "X127KN2"
	IF assigned_to = "Karla Schulz" or  assigned_to = "Karla Schulz" THEN worker_number =  "X127BK7"
	IF assigned_to = "Jaimee M Schwark" or  assigned_to = "Jaimee Schwark" THEN worker_number =  "X127JAI"
	IF assigned_to = "DiAnne A Scott" or  assigned_to = "DiAnne Scott" THEN worker_number =  "X1272RO"
	IF assigned_to = "Kathi N Scott" or  assigned_to = "Kathi Scott" THEN worker_number =  "X127928"
	IF assigned_to = "Lisa L Sebald" or  assigned_to = "Lisa Sebald" THEN worker_number =  "X127AY3"
	IF assigned_to = "Angela C Sebranek" or  assigned_to = "Angela Sebranek" THEN worker_number =  "X127149"
	IF assigned_to = "Keith J Semmelink" or  assigned_to = "Keith Semmelink" THEN worker_number =  "X127GK2"
	IF assigned_to = "Michelle Cline" or assigned_to = "CineSetodjiMichelle" THEN worker_number =  "X127MLC"
	IF assigned_to = "Victoria Shaffer" or  assigned_to = "Victoria Shaffer" THEN worker_number =  "X127B1K"
	IF assigned_to = "Shavon A Johnson" or  assigned_to = "Shavon Johnson" THEN worker_number =  "X127SAJ"
	IF assigned_to = "Nancy N Shevich" or  assigned_to = "Nancy Shevich" THEN worker_number =  "X127NXS"
	IF assigned_to = "Richard D Shields" or  assigned_to = "Richard Shields" THEN worker_number =  "X1275L5"
	IF assigned_to = "Carol T Shipley" or  assigned_to = "Carol Shipley" THEN worker_number =  "X127C0R"
	IF assigned_to = "Tamrat A Shulu" or  assigned_to = "Tamrat Shulu" THEN worker_number =  "X127K9J"
	IF assigned_to = "Indranie B Singh" or  assigned_to = "Indranie Singh" THEN worker_number =  "X1272GB"
	IF assigned_to = "Monica M Socha" or  assigned_to = "Monica Socha" THEN worker_number =  "X127D6S"
	IF assigned_to = "Janelle L Sorenson" or  assigned_to = "Janelle Sorenson" THEN worker_number =  "X127M49"
	IF assigned_to = "Bernita R Steele" or  assigned_to = "Bernita Steele" THEN worker_number =  "X127Z62"
	IF assigned_to = "Jill Sternberg-Adams" or  assigned_to = "Jill Sternberg-Adams" THEN worker_number =  "X127GB3"
	IF assigned_to = "Mathilda M Stevenson" or  assigned_to = "Mathilda Stevenson" THEN worker_number =  "X127HQ6"
	IF assigned_to = "GioVauntai Stewart" or  assigned_to = "GioVauntai Stewart" THEN worker_number =  "X127150"
	IF assigned_to = "Tamara S Stewart" or  assigned_to = "Tamara Stewart" THEN worker_number =  "X127BH3"
	IF assigned_to = "Garrett V Stock" or  assigned_to = "Garrett Stock" THEN worker_number =  "X127D4S"
	IF assigned_to = "Mary F Stone" or  assigned_to = "Mary Stone" THEN worker_number =  "X127A1F"
	IF assigned_to = "Amber L Stone" or  assigned_to = "Amber Stone" THEN worker_number =  "X127ALS"
	IF assigned_to = "Veronica D Suvid" or  assigned_to = "Veronica Suvid" THEN worker_number =  "X1272B2"
	IF assigned_to = "Aleen N Swanson" or  assigned_to = "Aleen Swanson" THEN worker_number =  "X1275K1"
	IF assigned_to = "Louise K Tamba" or  assigned_to = "Louise Tamba" THEN worker_number =  "X127CA8"
	IF assigned_to = "Tamala L Taylor" or  assigned_to = "Tamala Taylor" THEN worker_number =  "X127AW7"
	IF assigned_to = "Dee Tenzin" or  assigned_to = "Dee Tenzin" THEN worker_number =  "X127FAB"
	IF assigned_to = "Darren L Terebenet" or  assigned_to = "Darren Terebenet" THEN worker_number =  "X127TDL"
	IF assigned_to = "Alla Ternyak" or  assigned_to = "Alla Ternyak" THEN worker_number =  "X127M93"
	IF assigned_to = "Thadeus M Teske" or  assigned_to = "Thadeus Teske" THEN worker_number =  "X127K81"
	IF assigned_to = "Ben Teskey" or  assigned_to = "Ben Teskey" THEN worker_number =  "X127GH4"
	IF assigned_to = "DoanTrinh M Thai" or  assigned_to = "DoanTrinh Thai" THEN worker_number =  "X127J77"
	IF assigned_to = "Meena Thao" or  assigned_to = "Meena Thao" THEN worker_number =  "X127MT1"
	IF assigned_to = "Mee Thao" or  assigned_to = "Mee Thao" THEN worker_number =  "X127PC2"
	IF assigned_to = "Shaletha L Thomas" or  assigned_to = "Shaletha Thomas" THEN worker_number =  "X127B0E"
	IF assigned_to = "Julie A Thompson" or  assigned_to = "Julie Thompson" THEN worker_number =  "X127039"
	IF assigned_to = "Janet Thompson" or  assigned_to = "Janet Thompson" THEN worker_number =  "X127J8G"
	IF assigned_to = "Serina M Thor" or  assigned_to = "Serina Thor" THEN worker_number =  "X127BP2"
	IF assigned_to = "Kiera M Thornton" or  assigned_to = "Kiera Thornton" THEN worker_number =  "X127FAA"
	IF assigned_to = "Tinna C Tin" or  assigned_to = "Tinna Tin" THEN worker_number =  "X127TCT"
	IF assigned_to = "Inez R Toles" or  assigned_to = "Inez Toles" THEN worker_number =  "X127IRT"
	IF assigned_to = "Neil G Trembley" or  assigned_to = "Neil Trembley" THEN worker_number =  "X127D5V"
	IF assigned_to = "Jenna C Trimbo" or  assigned_to = "Jenna Trimbo" THEN worker_number =  "X127B7M"
	IF assigned_to = "Kristina T Truong" or  assigned_to = "Kristina Truong" THEN worker_number =  "X127GG7"
	IF assigned_to = "Baraka A Tura" or  assigned_to = "Baraka Tura" THEN worker_number =  "X127D6H"
	IF assigned_to = "Nicole Tyson" or  assigned_to = "Nicole Tyson" THEN worker_number =  "X127NT1"
	IF assigned_to = "Staci Ulmen" or  assigned_to = "Staci Ulmen" THEN worker_number =  "X127JV3"
	IF assigned_to = "Neil Urbanski" or  assigned_to = "Neil Urbanski" THEN worker_number =  "X127UN1"
	IF assigned_to = "Kary J Van Slyke" or  assigned_to = "Kary Van" THEN worker_number =  "X127T21"
	IF assigned_to = "Pa D Vang" or  assigned_to = "Pa Vang" THEN worker_number =  "X1275M3"
	IF assigned_to = "My M Vang" or  assigned_to = "My Vang" THEN worker_number =  "X1275P5"
	IF assigned_to = "Ka Vang" or  assigned_to = "Ka Vang" THEN worker_number =  "X1276KV"
	IF assigned_to = "Youa Vang" or  assigned_to = "Youa Vang" THEN worker_number =  "X1276YV"
	IF assigned_to = "Choua C Vang" or  assigned_to = "Choua Vang" THEN worker_number =  "X127AH5"
	IF assigned_to = "Scott H Vang" or  assigned_to = "Scott Vang" THEN worker_number =  "X127BL9"
	IF assigned_to = "Houa M Vang" or  assigned_to = "Houa Vang" THEN worker_number =  "X127CH2"
	IF assigned_to = "Ly Vang" or  assigned_to = "Ly Vang" THEN worker_number =  "X127D6M"
	IF assigned_to = "Joe B Vang" or  assigned_to = "Joe Vang" THEN worker_number =  "X127D8X"
	IF assigned_to = "Pahoua X Vang" or  assigned_to = "Pahoua Vang" THEN worker_number =  "X127GJ8"
	IF assigned_to = "Kongmeng Vang" or  assigned_to = "Kongmeng Vang" THEN worker_number =  "X127JE6"
	IF assigned_to = "Maria Vang" or  assigned_to = "Maria Vang" THEN worker_number =  "X127JN4"
	IF assigned_to = "Gao J Vang" or  assigned_to = "Gao Vang" THEN worker_number =  "X127KN3"
	IF assigned_to = "Lee Vang" or  assigned_to = "Lee Vang" THEN worker_number =  "X127LV2"
	IF assigned_to = "Judy C Vang" or  assigned_to = "Judy Vang" THEN worker_number =  "X127Z87"
	IF assigned_to = "Leticia C Vasquez Ledesma" or  assigned_to = "Leticia Vasquez" THEN worker_number =  "X127CF6"
	IF assigned_to = "Jaxon P VeeVahn" or  assigned_to = "Jaxon VeeVahn" THEN worker_number =  "X127GR1"
	IF assigned_to = "Michael Vegell" or  assigned_to = "Michael Vegell" THEN worker_number =  "X127MV4"
	IF assigned_to = "Adam R Verschoor" or  assigned_to = "Adam Verschoor" THEN worker_number =  "X127GR3"
	IF assigned_to = "Khristiane C Victorio" or  assigned_to = "Khristiane Victorio" THEN worker_number =  "X127KQ2"
	IF assigned_to = "Mark Vilayrack" or  assigned_to = "Mark Vilayrack" THEN worker_number =  "X127GR5"
	IF assigned_to = "Ana K Villegas Gonzalez" or  assigned_to = "Ana Villegas" THEN worker_number =  "X127KN4"
	IF assigned_to = "Oksana Voskresensky" or  assigned_to = "Oksana Voskresensky" THEN worker_number =  "X1272QW"
	IF assigned_to = "Natalie Vue" or  assigned_to = "Natalie Vue" THEN worker_number =  "X1272A0"
	IF assigned_to = "Pheng Vue" or  assigned_to = "Pheng Vue" THEN worker_number =  "X1276PV"
	IF assigned_to = "Elizabeth J Wahlstrom" or  assigned_to = "Elizabeth Wahlstrom" THEN worker_number =  "X127C09"
	IF assigned_to = "Yasmin F Waite" or  assigned_to = "Yasmin Waite" THEN worker_number =  "X127Z45"
	IF assigned_to = "Negesso B Wakeyo" or  assigned_to = "Negesso Wakeyo" THEN worker_number =  "X127BG6"
	IF assigned_to = "Amanda M Wallace" or  assigned_to = "Amanda Wallace" THEN worker_number =  "X127D3Y"
	IF assigned_to = "Kerry E Walsh" or  assigned_to = "Kerry Walsh" THEN worker_number =  "X127AM2"
	IF assigned_to = "Robert D Warmboe" or  assigned_to = "Robert Warmboe" THEN worker_number =  "X127GS2"
	IF assigned_to = "LaShay C Waters" or  assigned_to = "LaShay Waters" THEN worker_number =  "X127KQ3"
	IF assigned_to = "Sean Watkins" or  assigned_to = "Sean Watkins" THEN worker_number =  "X127FAF"
	IF assigned_to = "Alison L Watkins" or  assigned_to = "Alison Watkins" THEN worker_number =  "X127PC4"
	IF assigned_to = "Denise C Welch" or  assigned_to = "Denise Welch" THEN worker_number =  "X127B5I"
	IF assigned_to = "Lorna J Welch" or  assigned_to = "Lorna Welch" THEN worker_number =  "X127FAH"
	IF assigned_to = "Clar Weller" or  assigned_to = "Clar Weller" THEN worker_number =  "X127M16"
	IF assigned_to = "Jacob E West" or  assigned_to = "Jacob West" THEN worker_number =  "X127GF8"
	IF assigned_to = "Pamela Y Whitson" or  assigned_to = "Pamela Whitson" THEN worker_number =  "X127X43"
	IF assigned_to = "Patricia A Williams" or  assigned_to = "Patricia Williams" THEN worker_number =  "X1275L3"
	IF assigned_to = "Dawn L Williams" or  assigned_to = "Dawn Williams" THEN worker_number =  "X127D5W"
	IF assigned_to = "Kimberly A Williams" or  assigned_to = "Kimberly Williams" THEN worker_number =  "X127D7Y"
	IF assigned_to = "Florence R Williams" or  assigned_to = "Florence Williams" THEN worker_number =  "X127HR1"
	IF assigned_to = "Natasha Williams" or  assigned_to = "Natasha Williams" THEN worker_number =  "X127KN6"
	IF assigned_to = "Terrilyn D Wilson" or  assigned_to = "Terrilyn Wilson" THEN worker_number =  "X1273ES"
	IF assigned_to = "Daminga A Wilson" or  assigned_to = "Daminga Wilson" THEN worker_number =  "X127FAD"
	IF assigned_to = "Aimee J Wimberly" or  assigned_to = "Aimee Wimberly" THEN worker_number =  "X1275L2"
	IF assigned_to = "Jacqueline A Winiarczyk" or  assigned_to = "Jacqueline Winiarczyk" THEN worker_number =  "X1272HY"
	IF assigned_to = "Phyllicia C Wise" or  assigned_to = "Phyllicia Wise" THEN worker_number =  "X127JY3"
	IF assigned_to = "Karen D Womack" or  assigned_to = "Karen Womack" THEN worker_number =  "X1275L1"
	IF assigned_to = "Keenya C Woods" or  assigned_to = "Keenya Woods" THEN worker_number =  "X127PC6"
	IF assigned_to = "Nellie A Woodson" or  assigned_to = "Nellie Woodson" THEN worker_number =  "X127HT9"
	IF assigned_to = "Beverly M Wyka" or  assigned_to = "Beverly Wyka" THEN worker_number =  "X1272TV"
	IF assigned_to = "Houa X Xiong" or  assigned_to = "Houa Xiong" THEN worker_number =  "X1275J8"
	IF assigned_to = "Peter Xiong" or  assigned_to = "Peter Xiong" THEN worker_number =  "X1275PX"
	IF assigned_to = "Amy Xiong" or  assigned_to = "Amy Xiong" THEN worker_number =  "X127AXX"
	IF assigned_to = "See Xiong" or  assigned_to = "See Xiong" THEN worker_number =  "X127FAT"
	IF assigned_to = "Julie Xiong" or  assigned_to = "Julie Xiong" THEN worker_number =  "X127HR2"
	IF assigned_to = "Andre S Xiong" or  assigned_to = "Andre Xiong" THEN worker_number =  "X127R63"
	IF assigned_to = "Xoua Xiong" or  assigned_to = "Xoua Xiong" THEN worker_number =  "X127X0X"
	IF assigned_to = "Zong Yang" or  assigned_to = "Zong Yang" THEN worker_number =  "X127159"
	IF assigned_to = "Panhia Yang" or  assigned_to = "Panhia Yang" THEN worker_number =  "X1275N3"
	IF assigned_to = "Gaohnou Yang" or  assigned_to = "Gaohnou Yang" THEN worker_number =  "X127A0U"
	IF assigned_to = "Sirynoise V Yang" or  assigned_to = "Sirynoise Yang" THEN worker_number =  "X127BP1"
	IF assigned_to = "Yeng Yang" or  assigned_to = "Yeng Yang" THEN worker_number =  "X127FAP"
	IF assigned_to = "Pakou Yang" or  assigned_to = "Pakou Yang" THEN worker_number =  "X127GD5"
	IF assigned_to = "Goua Yang" or  assigned_to = "Goua Yang" THEN worker_number =  "X127JE9"
	IF assigned_to = "Maikou Yang" or  assigned_to = "Maikou Yang" THEN worker_number =  "X127JN5"
	IF assigned_to = "Ashley A Yang" or  assigned_to = "Ashley Yang" THEN worker_number =  "X127KN7"
	IF assigned_to = "Maytong Yang" or  assigned_to = "Maytong Yang" THEN worker_number =  "X127T81"
	IF assigned_to = "Yee Yang" or  assigned_to = "Yee Yang" THEN worker_number =  "X127YY1"
	IF assigned_to = "Mandora Young" or  assigned_to = "Mandora Young" THEN worker_number =  "X1274JX"
	IF assigned_to = "Hailey C Young" or  assigned_to = "Hailey Young" THEN worker_number =  "X127KN5"
	IF assigned_to = "Halane Yussuf" or  assigned_to = "Halane Yussuf" THEN worker_number =  "X127YH1"
	IF assigned_to = "Joan Zangs" or  assigned_to = "Joan Zangs" THEN worker_number =  "X127KN8"
	IF assigned_to = "Omar Zavala" or  assigned_to = "Omar Zavala" THEN worker_number =  "X127OZA"
	IF assigned_to = "Rebecca M Zelaya" or  assigned_to = "Rebecca Zelaya" THEN worker_number =  "X1272IQ"
	IF assigned_to = "Pamela J Zolik" or  assigned_to = "Pamela Zolik" THEN worker_number =  "X127411"

	IF assigned_to = "Anna Mahon" or  assigned_to = "Anna Mahon" THEN worker_number =  "X127J9Z"
	IF assigned_to = "Allicia M Brolsma" or  assigned_to = "Allicia Brolsma" THEN worker_number =  "X127JT1"
	IF assigned_to = "Dawn T Olmstead" or  assigned_to = "Dawn Olmstead" THEN worker_number =  "X127KP9"

 	assigned_to = trim(assigned_to)
	IF worker_number = "" THEN worker_number = "REVIEW"
END FUNCTION
 'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

''----------------------------------------------------------------------------------------------------The current day's assignment

assignment_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(assignment_date)       'finds the most recent previous working day for the file names
assignment_date = assignment_date & "" 'have to make it a string for the script to realize that it is a date '

BeginDialog Dialog1, 0, 0, 236, 115, "TASK BASED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 175, 70, 50, 15, "Browse...", select_a_file_button
    OkButton 120, 95, 50, 15
    CancelButton 175, 95, 50, 15
  EditBox 10, 70, 150, 15, file_selection_path
  EditBox 175, 5, 50, 15, assignment_date
  Text 10, 35, 210, 10, "This script should be used to assist with the task based review."
  Text 10, 5, 155, 20, "Please enter the date the HSR was given the assignment:"
  GroupBox 5, 25, 225, 65, "Using this script:"
  Text 10, 50, 205, 20, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
EndDialog


'dialog and dialog DO...Loop
Do
 	Do
  	    err_msg = ""
  	    dialog Dialog1
  	    cancel_without_confirmation
  	    If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
  	    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "Please select a file to continue."
		IF isdate(assignment_date) = False then err_msg = err_msg & vbnewline & "Please enter an assignment date."
  	    If err_msg <> "" Then MsgBox err_msg
 	Loop until err_msg = ""
 	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'setting the footer month to make the updates in'
CALL convert_date_into_MAXIS_footer_month(assignment_date, MAXIS_footer_month, MAXIS_footer_year)
CALL MAXIS_footer_month_confirmation
CALL ONLY_create_MAXIS_friendly_date(assignment_date)


'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
'objExcel.worksheets("Report 1").Activate   'Activates the initial BOBI report

'Establishing array
DIM task_based_array()  'Declaring the array this is what this list is
ReDim task_based_array(interview_const, 0)  'Resizing the array 'that ,list is goign to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const assignment_date_const 	= 0 '= "Date Assigned"
const excel_row_const			= 1 '=
const maxis_case_number_const 	= 2 '= "Case Number" - pretend this means 2
const case_name_const 			= 3 '= "Case Name"
const assigned_to_const 		= 4 '= "Assigned to"
const worker_number_const		= 5 '= "Assigned Worker X127#"
const case_note_const 			= 6 '=
const DAIL_count_const   		= 7 '= "DAIL Count
const DAIL_Type_const   		= 9 '= "DAIL Count"
const case_status_const 		= 10 '=
const case_note_date_const 		= 11 '
const case_note_key_word_const 	= 12
const interview_const 			= 13 'Interview Completed


'setting the columns - using constant so that we know what is going on'
const excel_col_assignment_date = 1 'A'
const excel_col_case_number 	= 3 'C'
const excel_col_case_name 		= 4 'D'
const excel_col_assigned_to 	= 6 'F'
const excel_col_worker_number 	= 7 'G'
const excel_col_case_note_date	= 10 'J' recommend this is changed/removed
const excel_col_case_note   	= 11 'K' ended up being a true false due to macros was orginally a count
const excel_col_key_word		= 12 'L'
const excel_col_DAIL_count		= 13 'M'
const excel_col_case_status 	= 18 'R'
const excel_col_interview		= 19 'S'

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start based on when picking up the information
entry_record = 0 'incrementor for the array and count

Do 'purpose is to read each excel row and to add into each excel array '
 'Reading information from the Excel
 MAXIS_case_number = objExcel.cells(excel_row, excel_col_case_number).Value
 MAXIS_case_number = trim(MAXIS_case_number)
 IF MAXIS_case_number = "" then exit do

   'Adding client information to the array - this is for READING FROM the excel
 ReDim Preserve task_based_array(interview_const, entry_record)	'This resizes the array based on the number of cases
		task_based_array(maxis_case_number_const,  entry_record) = MAXIS_case_number
		task_based_array(assigned_to_const,  entry_record) = trim(objExcel.cells(excel_row, excel_col_assigned_to).Value)
		task_based_array(excel_row_const, entry_record) = excel_row
		'making space in the array for these variables, but valuing them as "" for now
  entry_record = entry_record + 1			'This increments to the next entry in the array
  stats_counter = stats_counter + 1 'Increment for stats counter
 excel_row = excel_row + 1
Loop

back_to_self 'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
worker_number = "" 'clearing for the function'
'Loading of cases is complete. Reviewing the cases in the array.
For item = 0 to UBound(task_based_array, 2)

 	MAXIS_case_number   = task_based_array(maxis_case_number_const, item)
	assigned_to  		= task_based_array(assigned_to_const,   item)

	CALL HSR_LIST 'this will get our worker number and name defines worker_number'
	IF worker_number = "REVIEW" then MiKayla_needs_to_know_this = MiKayla_needs_to_know_this & "name not found for: " & assigned_to & vbNewLine & vbNewLine
	'MsgBox "worker_number ~" & worker_number & "~"
	task_based_array(worker_number_const,   item) = worker_number ' saves it to our array '


	CALL navigate_to_MAXIS_screen("CASE", "NOTE")
 	MAXIS_row = 5 'Defining row for the search feature.
	interview_done = False

 	EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROG and INQUIRY
	EMReadScreen county_code, 4, 21, 14  'Out of county cases from STAT
	EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to  navigate'
 	IF priv_check = "PRIV" THEN  'PRIV cases
  		EMReadscreen priv_worker, 26, 24, 46
  		task_based_array(case_status_const, item) = trim(priv_worker)
 	ELSEIf county_code <> "X127" THEN
	  task_based_array(case_status_const, item) = "OUT OF COUNTY CASE"
 	ELSEIF instr(case_invalid_error, "IS INVALID") THEN  'CASE xxxxxxxx IS INVALID FOR PERIOD 12/99
		task_based_array(case_status_const, item) = trim(case_invalid_error)
	ELSE
		EMReadScreen MAXIS_case_name, 27, 21, 40
		'MsgBox "MAXIS_case_name ~" & MAXIS_case_name & "~"
		task_based_array(case_name_const, item) = trim(MAXIS_case_name)
		task_based_array(case_note_const, item) = "NO" 'defaulting to no to ensure we increment '
		IF worker_number = "REVIEW" THEN task_based_array(case_note_const, item) = ""
 	    DO
 	    	EMReadscreen case_note_date, 8, MAXIS_row, 6
			'MsgBox assignment_date & " ~ " & case_note_date & "~"
 	    	If trim(case_note_date) = "" THEN
				task_based_array(case_status_const, item) = "NO CASE NOTE"
				exit do
 	    	Else
 	    		IF case_note_date = assignment_date THEN 'weekends and the day prior has the date assigned confirmed by the SSR '
					task_based_array(case_note_date_const, item) = case_note_date

 	    			EMReadScreen case_note_worker_number, 7, MAXIS_row, 16
					'MsgBox worker_number & "~" & case_note_worker_number
 	    			IF worker_number = case_note_worker_number THEN
						task_based_array(case_note_const, item) = "YES"
 	    				case_note_count = case_note_count + 1
 	    				EMReadScreen case_note_header, 55, MAXIS_row, 25
  	    				case_note_header = lcase(trim(case_note_header))
 	    				If instr(case_note_header, "interview completed") then task_based_array(interview_const, item) = "TRUE"
	    			END IF

	    		END IF
 	    	END IF
		MAXIS_row = MAXIS_row + 1
 		IF MAXIS_row = 19 THEN
 			PF8 'moving to next case note page if at the end of the page
 			MAXIS_row = 5
 		END IF
  		LOOP UNTIL cdate(case_note_date) < cdate(assignment_date)   'repeats until the case note date is less than the assignment date
 		task_based_array(case_note_count_const, item) = case_note_count
	'END IF
 	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	DO
 		EMReadScreen dail_check, 4, 2, 48
 		If next_dail_check <> "DAIL" then CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	LOOP UNTIL dail_check = "DAIL"

 	DAIL_count = 0	'these are the actionable DAIL counts only
 	dail_row = 5			'Because the script brings each new case to the top of the page, dail_row starts at 6.
 	DO
 		EmReadscreen number_of_dails, 1, 3, 67	'Reads where there count of dAILS is listed
 		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

 		EMReadScreen DAIL_case_number, 8, dail_row, 73
 		DAIL_case_number = trim(DAIL_case_number)
 		If DAIL_case_number <> MAXIS_case_number then exit do
 	    'Determining if there is a new case number...
 	    EMReadScreen new_case, 8, dail_row, 63
 	    new_case = trim(new_case)
 	    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
 	     Call write_value_and_transmit("T", dail_row, 3)
 	     dail_row = 6
 	    ELSEIF new_case = "CASE NBR" THEN
 	     '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
 	     	Call write_value_and_transmit("T", dail_row + 1, 3)
 	     	dail_row = 6
 	    End if
 	    EMReadScreen dail_type, 4, dail_row, 6
 	    EMReadScreen dail_msg, 61, dail_row, 20
 	    dail_msg = trim(dail_msg)
 	    EMReadScreen dail_month, 8, dail_row, 11
 	    dail_month = trim(dail_month)
 	    Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
 	    IF actionable_dail = True then dail_count = dail_count + 1
 	    dail_row = dail_row + 1
 	Loop
		task_based_array(DAIL_count_const, item)  = DAIL_count
		task_based_array(DAIL_type_const, item)  = DAIL_type

	END IF
	CALL back_to_self
	worker_number = "" 'clearing for the function' need to reset the worker number each time we go into the "next"
Next

objExcel.Columns(1).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
objExcel.Columns(10).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
objExcel.Columns(17).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY

For item = 0 to UBound(task_based_array, 2)
 excel_row = task_based_array(excel_row_const, item)
 objExcel.Cells(excel_row, excel_col_case_name).Value  = task_based_array(case_name_const,   item)
 objExcel.Cells(excel_row, excel_col_worker_number).Value  = task_based_array(worker_number_const,  item)
 objExcel.Cells(excel_row, excel_col_case_note).Value = task_based_array(case_note_const,   item)
 objExcel.Cells(excel_row, excel_col_DAIL_count).Value = task_based_array(DAIL_count_const,  item)
 objExcel.Cells(excel_row, excel_col_case_status).Value = task_based_array(case_status_const, item)
 objExcel.Cells(excel_row, excel_col_interview).Value = task_based_array(interview_const,   item)
Next

'Adrian Andres, Cody Ross, Edward Ukatu, Elmi Elmi, Mohsin Hashi, Joseph Brewer, Kiarah Ray, Lori Clayton, Michele Price, Natasha Williams, Rebecca Hemmans, Renee McGrath, Rey Gonzalez-Perez, Sumaya Omar, Thomas Anderson, Wanda Baker

FOR i = 1 to 20							'formatting the cells'
	objExcel.Columns(i).AutoFit()		'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1   'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
IF MiKayla_needs_to_know_this <> "" THEN CALL create_outlook_email("mikayla.handley@Hennepin.us", "", "Name not found task based list", MiKayla_needs_to_know_this, "", TRUE)

script_end_procedure_with_error_report("Success your list has been updated, please review to ensure accuracy.")
