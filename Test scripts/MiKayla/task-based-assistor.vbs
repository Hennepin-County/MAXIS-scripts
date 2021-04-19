'STATS GATHERING--------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK BASED ASSISTOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 100                      'manual run time in seconds
STATS_denomination = "C"       			   'M is for each CASE
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
'END CHANGELOG BLOCK =======================================================================================================
Function HSR_LIST
    IF assigned_to = "Julie A Thompson" THEN worker_number =  "X127039"
    IF assigned_to = "Negassa K Ayana" THEN worker_number =  "X127043"
    IF assigned_to = "Todd A Norling" THEN worker_number = 	"X127085"
    IF assigned_to = "Kevin Chavis" THEN worker_number =  "X127089"
    IF assigned_to = "Gifti Geleta" THEN worker_number =  "X127090"
    IF assigned_to = "Emily J Frazier" THEN worker_number =  "X1270A1"
    IF assigned_to = "Amy M McCall" THEN worker_number =  "X127130"
    IF assigned_to = "Brittany M Lane" THEN worker_number =  "X127140"
    IF assigned_to = "Angela C Sebranek" THEN worker_number = "X127149"
    IF assigned_to = "GioVauntai Stewart" THEN worker_number =  "X127150"
    IF assigned_to = "Zong Yang" THEN worker_number = "X127159"
    IF assigned_to = "Zoye D Jackson" THEN worker_number = 	  "X127160"
    IF assigned_to = "Kasey A Holt" THEN worker_number =  "X1271A7"
    IF assigned_to = "Kenneth W Garnier" THEN worker_number = "X1271AJ"
    IF assigned_to = "Farah A Hassan" THEN worker_number = 	  "X1271FH"
    IF assigned_to = "Khadra S Abdallah" THEN worker_number = "X1271KA"
    IF assigned_to = "Victoria Rutkovskaya" THEN worker_number =  "X12726L"
    IF assigned_to = "Gina L Aasgaard" THEN worker_number =  "X12726N"
    IF assigned_to = "Cheryl L Heitzinger" THEN worker_number =  "X12728S"
    IF assigned_to = "Abdo W Korosso" THEN worker_number = 	"X12729T"
    IF assigned_to = "Jason A Marx" THEN worker_number =  "X12729W"
    IF assigned_to = "Natalie Vue" THEN worker_number =  "X1272A0"
    IF assigned_to = "Rachel A Ferguson" THEN worker_number = "X1272AF"
    IF assigned_to = "Mohamed O Nur" THEN worker_number = "X1272AY"
    IF assigned_to = "Sitha Sarin" THEN worker_number =  "X1272B1"
    IF assigned_to = "Veronica D Suvid" THEN worker_number =  "X1272B2"
    IF assigned_to = "Darlenne Salinas-Fernandez" THEN worker_number = 	  "X1272BB"
    IF assigned_to = "Melissa A Flores" THEN worker_number =  "X1272BD"
    IF assigned_to = "True P Lis" THEN worker_number = 	  "X1272BM"
    IF assigned_to = "Viola L Kill" THEN worker_number =  "X1272BR"
    IF assigned_to = "Abdiaziz M Hurreh" THEN worker_number = "X1272CZ"
    IF assigned_to = "Jessica L Dickerson" THEN worker_number = 	 "X1272EG"
    IF assigned_to = "Indranie B Singh" THEN worker_number =  "X1272GB"
    IF assigned_to = "Faduma M Abdi" THEN worker_number = "X1272GM"
    IF assigned_to = "Brooke A Reilley" THEN worker_number =  "X1272HC"
    IF assigned_to = "Jacqueline A Winiarczyk" THEN worker_number = 	 "X1272HY"
    IF assigned_to = "Amorette B Robeck" THEN worker_number = "X1272ID"
    IF assigned_to = "Rebecca M Zelaya" THEN worker_number =  "X1272IQ"
    IF assigned_to = "Samantha E Haw" THEN worker_number = "X1272LJ"
    IF assigned_to = "Ramona M Shane" THEN worker_number = "X1272PC"
    IF assigned_to = "Oksana Voskresensky" THEN worker_number = "X1272QW"
    IF assigned_to = "DiAnne A Scott" THEN worker_number = "X1272RO"
    IF assigned_to = "Beverly M Wyka" THEN worker_number = "X1272TV"
    IF assigned_to = "Tasheema Hopson" THEN worker_number = "X1272UE"
    IF assigned_to = "Svetlana Kabakova" THEN worker_number = "X12730W"
    IF assigned_to = "Celeste Jourdain" THEN worker_number =  "X1273D3"
    IF assigned_to = "Scott A Chestnut" THEN worker_number =  "X1273DC"
    IF assigned_to = "Sheryl J Dillenburg" THEN worker_number = "X1273DD"
    IF assigned_to = "Terrilyn D Wilson" THEN worker_number = "X1273ES"
    IF assigned_to = "Melinda L Lane" THEN worker_number = 	  "X1273FL"
    IF assigned_to = "Stephanie L Holmes" THEN worker_number = "X1273M7"
    IF assigned_to = "James B Eckard" THEN worker_number = 	  "X1273T4"
    IF assigned_to = "Abdi Y Ahmed" THEN worker_number =  "X1273YA"
    IF assigned_to = "Pamela J Zolik" THEN worker_number = 	  "X127411"
    IF assigned_to = "Linda L. Greene" THEN worker_number = "X127436"
    IF assigned_to = "Brenda M Raygor" THEN worker_number = "X12746F"
    IF assigned_to = "Kristine Karlsgodt" THEN worker_number = "X12746I"
    IF assigned_to = "Clarita B Scherer" THEN worker_number = "X12746P"
    IF assigned_to = "Matthew M McGovern" THEN worker_number = "X1274HQ"
    IF assigned_to = "Mandora Young" THEN worker_number = "X1274JX"
    IF assigned_to = "Debrice L Jackson" THEN worker_number = "X1274QG"
    IF assigned_to = "Qadro A Abdi" THEN worker_number =  "X1275A2"
    IF assigned_to = "Amber Davis" THEN worker_number = "X1275AD"
    IF assigned_to = "Mohammed K Ahmed" THEN worker_number =  "X1275B7"
    IF assigned_to = "TyAnn Barnes" THEN worker_number =  "X1275F3"
    IF assigned_to = "Thomas A. Moore" THEN worker_number =  "X1275F4"
    IF assigned_to = "Aaron J. Gardner-Kocher" THEN worker_number =  "X1275F9"
    IF assigned_to = "Debra E George" THEN worker_number = 	"X1275G0"
    IF assigned_to = "Marianne E. Simon" THEN worker_number = "X1275G3"
    IF assigned_to = "Jerry Nack" THEN worker_number = 	"X1275G4"
    IF assigned_to = "DeNise Kendrick-Stevens" THEN worker_number =  "X1275G8"
    IF assigned_to = "Remy K. Huerta-Stemper" THEN worker_number = "X1275H0"
    IF assigned_to = "Valerie M. Herrera" THEN worker_number = 	"X1275H5"
    IF assigned_to = "Peggy Chavez" THEN worker_number =  "X1275H7"
    IF assigned_to = "Sharmarke Y. Abdi" THEN worker_number = "X1275H8"
    IF assigned_to = "Rashida R. Manuel" THEN worker_number = "X1275H9"
    IF assigned_to = "Alejandra Andrade" THEN worker_number = "X1275J1"
    IF assigned_to = "Maria Remy" THEN worker_number = 	"X1275J3"
    IF assigned_to = "Houa X Xiong" THEN worker_number =  "X1275J8"
    IF assigned_to = "Sherry A Duggan" THEN worker_number = "X1275K0"
    IF assigned_to = "Aleen N Swanson" THEN worker_number = "X1275K1"
    IF assigned_to = "Pa Nhia Lee" THEN worker_number =  "X1275K2"
    IF assigned_to = "Teng L Lor" THEN worker_number = 	"X1275K5"
    IF assigned_to = "Tyler O Burch" THEN worker_number = "X1275K7"
    IF assigned_to = "Karen D Womack" THEN worker_number = 	  "X1275L1"
    IF assigned_to = "Aimee J Wimberly" THEN worker_number =  "X1275L2"
    IF assigned_to = "Patricia A Williams" THEN worker_number = 	 "X1275L3"
    IF assigned_to = "Richard D Shields" THEN worker_number = "X1275L5"
    IF assigned_to = "Sally R Engstrom" THEN worker_number =  "X1275L6"
    IF assigned_to = "Candace S Johnson" THEN worker_number = "X1275L7"
    IF assigned_to = "Lara M Messer" THEN worker_number = "X1275L8"
    IF assigned_to = "Lim Ban" THEN worker_number = 	 "X1275L9"
    IF assigned_to = "Jacob P Arco" THEN worker_number =  "X1275M2"
    IF assigned_to = "Pa D Vang" THEN worker_number = "X1275M3"
    IF assigned_to = "Colanda R Lacy" THEN worker_number = 	  "X1275M9"
    IF assigned_to = "Diane Ha" THEN worker_number =  "X1275N1"
    IF assigned_to = "Panhia Yang" THEN worker_number = 	 "X1275N3"
    IF assigned_to = "Azza D Pratiwi" THEN worker_number = 	  "X1275N7"
    IF assigned_to = "Inger M Harris" THEN worker_number = 	  "X1275N9"
    IF assigned_to = "Sherry L Collins" THEN worker_number =  "X1275P1"
    IF assigned_to = "Tony Her" THEN worker_number =  "X1275P2"
    IF assigned_to = "Ondrenette Blair" THEN worker_number =  "X1275P4"
    IF assigned_to = "My M Vang" THEN worker_number = "X1275P5"
    IF assigned_to = "Safiyo A Haji" THEN worker_number = "X1275P6"
    IF assigned_to = "Peter Xiong" THEN worker_number = 	 "X1275PX"
    IF assigned_to = "Eve Gray" THEN worker_number =  "X1276EG"
    IF assigned_to = "Ka Vang" THEN worker_number = 	 "X1276KV"
    IF assigned_to = "Maria L Johnson" THEN worker_number = 	 "X1276MJ"
    IF assigned_to = "Pheng Vue" THEN worker_number = "X1276PV"
    IF assigned_to = "Timothy B Erickson" THEN worker_number = 	  "X1276TE"
    IF assigned_to = "Yoshauna Sampson" THEN worker_number =  "X1276YS"
    IF assigned_to = "Youa Vang" THEN worker_number = "X1276YV"
    IF assigned_to = "Mary D Reeck" THEN worker_number =  "X127927"
    IF assigned_to = "Kathi N Scott" THEN worker_number = "X127928"
    IF assigned_to = "Florence A. Manley" THEN worker_number = 	  "X127966"
    IF assigned_to = "Cecelia M Hoecherl" THEN worker_number = 	  "X127A05"
    IF assigned_to = "Sharon Murphy" THEN worker_number = "X127A0F"
    IF assigned_to = "Gaohnou Yang" THEN worker_number =  "X127A0U"
    IF assigned_to = "Lakisha P Perkerson" THEN worker_number = 	 "X127A0W"
    IF assigned_to = "Jamoda L Acevedo" THEN worker_number =  "X127A0Z"
    IF assigned_to = "Mary F Stone" THEN worker_number =  "X127A1F"
    IF assigned_to = "Fabio A Rozo" THEN worker_number =  "X127A1T"
    IF assigned_to = "Kristen F Kasim" THEN worker_number = 	 "X127A2D"
    IF assigned_to = "Jessica L Belland" THEN worker_number = "X127A3J"
    IF assigned_to = "Miguel A Salazar" THEN worker_number =  "X127A3V"
    IF assigned_to = "Charice McDowell" THEN worker_number =  "X127A6S"
    IF assigned_to = "Penny R Grady" THEN worker_number = "X127A6X"
    IF assigned_to = "Gary M. Remus" THEN worker_number = "X127A79"
    IF assigned_to = "Stephanie A Jefferson" THEN worker_number = "X127A7M"
    IF assigned_to = "Gina T Gangelhoff" THEN worker_number = "X127A7O"
    IF assigned_to = "Soumya G Sanyal" THEN worker_number = 	 "X127A7S"
    IF assigned_to = "Shantell Cochran" THEN worker_number =  "X127A8B"
    IF assigned_to = "Bernardo G Gonzalez" THEN worker_number = 	 "X127A8F"
    IF assigned_to = "Lacosta L Awad" THEN worker_number = 	  "X127A8Q"
    IF assigned_to = "Jacqueline S Graves" THEN worker_number = 	 "X127A9U"
    IF assigned_to = "Mailee C Moua" THEN worker_number = "X127A9X"
    IF assigned_to = "Fatah A Ahmed" THEN worker_number = "X127AAF"
    IF assigned_to = "Abdiwali B Bulhan" THEN worker_number = "X127ABB"
    IF assigned_to = "Andrea D Green" THEN worker_number = 	  "X127ADG"
    IF assigned_to = "Ayaan M Elmi" THEN worker_number =  "X127AE1"
    IF assigned_to = "Molly M Manley" THEN worker_number = 	  "X127AG4"
    IF assigned_to = "Michelle E Parenteau" THEN worker_number =  "X127AH2"
    IF assigned_to = "Zaki M Isaac" THEN worker_number =  "X127AH3"
    IF assigned_to = "Lori A Roberson" THEN worker_number = 	 "X127AH4"
    IF assigned_to = "Choua C Vang" THEN worker_number =  "X127AH5"
    IF assigned_to = "Denis L Ladeyshchikov" THEN worker_number = "X127AL5"
    IF assigned_to = "Madar A Hachi" THEN worker_number = "X127AL8"
    IF assigned_to = "Amber L Stone" THEN worker_number = "X127ALS"
    IF assigned_to = "Kerry E Walsh" THEN worker_number = "X127AM2"
    IF assigned_to = "Sartu A Hassan" THEN worker_number = 	  "X127AM5"
    IF assigned_to = "Janell L Hill" THEN worker_number = "X127AM9"
    IF assigned_to = "Marya D Anderson" THEN worker_number =  "X127AN5"
    IF assigned_to = "Alisa R Haselhorst" THEN worker_number = 	  "X127AN6"
    IF assigned_to = "Ryan P Kierth" THEN worker_number = "X127AP7"
    IF assigned_to = "Bethelhem G Beyene" THEN worker_number = 	  "X127AQ4"
    IF assigned_to = "Nailah Y Holman" THEN worker_number = 	 "X127AQ5"
    IF assigned_to = "Katie M Adams" THEN worker_number = "X127AQ7"
    IF assigned_to = "Teresa D Morphew" THEN worker_number =  "X127AQ8"
    IF assigned_to = "Michelle E Barnes" THEN worker_number = "X127AU2"
    IF assigned_to = "Tamala L Taylor" THEN worker_number = 	 "X127AW7"
    IF assigned_to = "Kaeli F Larson" THEN worker_number = 	  "X127AW8"
    IF assigned_to = "Laura A Olson" THEN worker_number = "X127AX4"
    IF assigned_to = "Amy Xiong" THEN worker_number = "X127AXX"
    IF assigned_to = "Lisa L Sebald" THEN worker_number = "X127AY3"
    IF assigned_to = "Terri L Cox" THEN worker_number = 	 "X127B01"
    IF assigned_to = "Fowsia M Abdi" THEN worker_number = "X127B0C"
    IF assigned_to = "Shaletha L Thomas" THEN worker_number = "X127B0E"
    IF assigned_to = "Osman M Abdi" THEN worker_number =  "X127B0X"
    IF assigned_to = "Angel S Alexander" THEN worker_number = "X127B0Y"
    IF assigned_to = "Kimberly A Hill" THEN worker_number = 	 "X127B18"
    IF assigned_to = "Toni Jenkins" THEN worker_number =  "X127B1B"
    IF assigned_to = "Victoria Shaffer" THEN worker_number =  "X127B1K"
    IF assigned_to = "Brian D Olson" THEN worker_number = "X127B22"
    IF assigned_to = "Debra M Kennedy" THEN worker_number = 	 "X127B2L"
    IF assigned_to = "Nicole Ryan" THEN worker_number = 	 "X127B2Z"
    IF assigned_to = "Christine Jernander" THEN worker_number = 	 "X127B36"
    IF assigned_to = "Delia M Dilday" THEN worker_number = 	  "X127B3T"
    IF assigned_to = "Sylvia A King" THEN worker_number = "X127B4Q"
    IF assigned_to = "Hodan K Farah" THEN worker_number = "X127B5A"
    IF assigned_to = "Denise C Welch" THEN worker_number = 	  "X127B5I"
    IF assigned_to = "Michelle L Lungelow" THEN worker_number = 	 "X127B6J"
    IF assigned_to = "Maria E Ammerman" THEN worker_number =  "X127B7E"
    IF assigned_to = "Jodynne D Flasch" THEN worker_number =  "X127B7G"
    IF assigned_to = "Jenna C Trimbo" THEN worker_number = 	  "X127B7M"
    IF assigned_to = "DeAnne L. Eberle" THEN worker_number =  "X127B8K"
    IF assigned_to = "Abdirizak M Ibrahim" THEN worker_number = 	 "X127B9P"
    IF assigned_to = "Molly C Hasbrook" THEN worker_number =  "X127B9Q"
    IF assigned_to = "Beverly A Denman" THEN worker_number =  "X127BD1"
    IF assigned_to = "Negesso B Wakeyo" THEN worker_number =  "X127BG6"
    IF assigned_to = "Tamara S Stewart" THEN worker_number =  "X127BH3"
    IF assigned_to = "Diana Peterson" THEN worker_number = 	  "X127BK2"
    IF assigned_to = "Karla Schulz" THEN worker_number =  "X127BK7"
    IF assigned_to = "Ronick Kimnong" THEN worker_number = 	  "X127BK8"
    IF assigned_to = "Jacqueline W Miantona" THEN worker_number = "X127BL4"
    IF assigned_to = "Rhonda V Hopson" THEN worker_number = 	 "X127BL8"
    IF assigned_to = "Scott H Vang" THEN worker_number =  "X127BL9"
    IF assigned_to = "Stacey Dunham" THEN worker_number = "X127BM4"
    IF assigned_to = "Fawn Marquez" THEN worker_number =  "X127BN9"
    IF assigned_to = "Sirynoise V. Yang" THEN worker_number = "X127BP1"
    IF assigned_to = "Serina M. Thor" THEN worker_number = 	  "X127BP2"
    IF assigned_to = "Star A Hanson" THEN worker_number = "X127BW6"
    IF assigned_to = "Jessica L Sanderson" THEN worker_number = 	 "X127BW7"
    IF assigned_to = "Elizabeth J Wahlstrom" THEN worker_number = "X127C09"
    IF assigned_to = "Diana M Demario" THEN worker_number = 	 "X127C0Q"
    IF assigned_to = "Carol T Shipley" THEN worker_number = 	 "X127C0R"
    IF assigned_to = "Mark P Jacobson" THEN worker_number = 	 "X127C1C"
    IF assigned_to = "Marilynn R Anderson" THEN worker_number = 	 "X127C1Q"
    IF assigned_to = "Celeste E Carlson" THEN worker_number = "X127C1T"
    IF assigned_to = "Michelle Pringle" THEN worker_number =  "X127C1Y"
    IF assigned_to = "Sara R Miller" THEN worker_number = "X127CA0"
    IF assigned_to = "Crystal M Henry-Bolden" THEN worker_number = 	  "X127CA2"
    IF assigned_to = "Bezabeh Assefa" THEN worker_number = 	  "X127CA4"
    IF assigned_to = "Louise K Tamba" THEN worker_number = 	  "X127CA8"
    IF assigned_to = "Sarita M Lopez" THEN worker_number = 	  "X127CA9"
    IF assigned_to = "Carrie A Lucca" THEN worker_number = 	  "X127CAL"
    IF assigned_to = "Miftah M Dadi" THEN worker_number = "X127CB3"
    IF assigned_to = "Juanita M Hubbard" THEN worker_number = "X127CB8"
    IF assigned_to = "Cindy Johnson" THEN worker_number = "X127CDJ"
    IF assigned_to = "Raeann T Korynta" THEN worker_number =  "X127CF3"
    IF assigned_to = "Leticia C Vasquez Ledesma" THEN worker_number = "X127CF6"
    IF assigned_to = "Jacqueline Charpentier" THEN worker_number = 	  "X127CF9"
    IF assigned_to = "Houa M Vang" THEN worker_number = 	 "X127CH2"
    IF assigned_to = "Carina S Cortez" THEN worker_number = 	 "X127CSC"
    IF assigned_to = "Cortney S Bhakta" THEN worker_number =  "X127CSS"
    IF assigned_to = "Andrea S Lawrence" THEN worker_number = "X127D05"
    IF assigned_to = "Linda - Lee" THEN worker_number = 	 "X127D2F"
    IF assigned_to = "Stephen S Moore" THEN worker_number = 	 "X127D2T"
    IF assigned_to = "Diane M Beauchamp" THEN worker_number = "X127D3C"
    IF assigned_to = "Vila Her" THEN worker_number =  "X127D3E"
    IF assigned_to = "Amanda M Wallace" THEN worker_number =  "X127D3Y"
    IF assigned_to = "James P. Berka" THEN worker_number = 	  "X127D3Z"
    IF assigned_to = "Tammi Barton" THEN worker_number =  "X127D4D"
    IF assigned_to = "Abba Bora H Kedir" THEN worker_number = "X127D4E"
    IF assigned_to = "Payeng Lee" THEN worker_number = 	  "X127D4H"
    IF assigned_to = "Anthony H. Berne" THEN worker_number =  "X127D4K"
    IF assigned_to = "Mai C Lee" THEN worker_number = "X127D4R"
    IF assigned_to = "Garrett V Stock" THEN worker_number = 	 "X127D4S"
    IF assigned_to = "Sheri L Peterson" THEN worker_number =  "X127D4X"
    IF assigned_to = "Laura L Riebe" THEN worker_number = "X127D4Y"
    IF assigned_to = "Candace S Brown" THEN worker_number = 	 "X127D5D"
    IF assigned_to = "Neil G. Trembley" THEN worker_number =  "X127D5V"
    IF assigned_to = "Dawn L. Williams" THEN worker_number =  "X127D5W"
    IF assigned_to = "MiKayla Handley" THEN worker_number = 	 "X127D5X"
    IF assigned_to = "Jennie E. Anderson" THEN worker_number = 	  "X127D5Z"
    IF assigned_to = "Baraka A. Tura" THEN worker_number = 	  "X127D6H"
    IF assigned_to = "Ly Vang" THEN worker_number = 	 "X127D6M"
    IF assigned_to = "Monica M. Socha" THEN worker_number = 	 "X127D6S"
    IF assigned_to = "Ahmed A Abdi" THEN worker_number =  "X127D7M"
    IF assigned_to = "Shamikka S Lenear" THEN worker_number = "X127D7R"
    IF assigned_to = "Yanisha K. Mack" THEN worker_number = 	 "X127D7W"
    IF assigned_to = "Kimberly A Williams" THEN worker_number = 	 "X127D7Y"
    IF assigned_to = "Brittney N. Ross" THEN worker_number =  "X127D7Z"
    IF assigned_to = "Elacia V. Davis" THEN worker_number = 	 "X127D8A"
    IF assigned_to = "Joe B Vang" THEN worker_number = 	  "X127D8X"
    IF assigned_to = "Julie M Broen" THEN worker_number = "X127DP3"
    IF assigned_to = "Dickyi Peldon" THEN worker_number = "X127DP7"
    IF assigned_to = "Dayanne Quinonez" THEN worker_number =  "X127DQ1"
    IF assigned_to = "Douglas S. Bright" THEN worker_number = "X127DSB"
    IF assigned_to = "Ann M. Davis" THEN worker_number =  "X127E26"
    IF assigned_to = "Linda K. Millhouse" THEN worker_number = 	  "X127E60"
    IF assigned_to = "Erik A Billington" THEN worker_number = "X127EAB"
    IF assigned_to = "Cynthia Hampton" THEN worker_number = 	 "X127F19"
    IF assigned_to = "Xay L Lee-Xiong" THEN worker_number = 	 "X127F23"
    IF assigned_to = "Abdirazak Botan" THEN worker_number = 	 "X127F2F"
    IF assigned_to = "Lisa J Bommersbach" THEN worker_number = 	  "X127F30"
    IF assigned_to = "Cheryl K Kerzman" THEN worker_number =  "X127F77"
    IF assigned_to = "Kiera M Thornton" THEN worker_number =  "X127FAA"
    IF assigned_to = "Dee Tenzin" THEN worker_number = 	  "X127FAB"
    IF assigned_to = "Tiffany R Bailey" THEN worker_number =  "X127FAC"
    IF assigned_to = "Daminga A Wilson" THEN worker_number =  "X127FAD"
    IF assigned_to = "Sean Watkins" THEN worker_number =  "X127FAF"
    IF assigned_to = "Lorna J Welch" THEN worker_number = "X127FAH"
    IF assigned_to = "Ziyad Z Kadir" THEN worker_number = "X127FAL"
    IF assigned_to = "Tiwana L Pargo" THEN worker_number = 	  "X127FAM"
    IF assigned_to = "Yeng Yang" THEN worker_number = "X127FAP"
    IF assigned_to = "Ahmed M Aden" THEN worker_number =  "X127FAQ"
    IF assigned_to = "Sarai R Counce" THEN worker_number = 	  "X127FAS"
    IF assigned_to = "See Xiong" THEN worker_number = "X127FAT"
    IF assigned_to = "Nas Looper" THEN worker_number = 	  "X127FAW"
    IF assigned_to = "Lina M Ahmed" THEN worker_number =  "X127FAY"
    IF assigned_to = "Claudia Perez Selva de Heintz" THEN worker_number = "X127FBJ"
    IF assigned_to = "Iliana E Martinez Morales" THEN worker_number = "X127FBQ"
    IF assigned_to = "Becky A Little" THEN worker_number = 	  "X127FBX"
    IF assigned_to = "Tanya L Payne" THEN worker_number = "X127FCA"
    IF assigned_to = "Fatiya A Ganamo" THEN worker_number = 	 "X127FGA"
    IF assigned_to = "Myrna C Banham-McKelvy" THEN worker_number = 	  "X127G07"
    IF assigned_to = "Christy P Olson" THEN worker_number = 	 "X127G50"
    IF assigned_to = "Jamiya O Ahmed" THEN worker_number = 	  "X127GAN"
    IF assigned_to = "John M Fandrick" THEN worker_number = 	 "X127GAP"
    IF assigned_to = "Jill Sternberg-Adams" THEN worker_number =  "X127GB3"
    IF assigned_to = "Pakou Yang" THEN worker_number = 	  "X127GD5"
    IF assigned_to = "Gloria Perez Amastal" THEN worker_number =  "X127GDP"
    IF assigned_to = "Giovanni E. Parodi" THEN worker_number = 	  "X127GEP"
    IF assigned_to = "Christina M Eichorn" THEN worker_number = 	 "X127GF7"
    IF assigned_to = "Jacob E West" THEN worker_number =  "X127GF8"
    IF assigned_to = "Sim Chang" THEN worker_number = "X127GG2"
    IF assigned_to = "Solange A Davis-Rivera" THEN worker_number = 	  "X127GG3"
    IF assigned_to = "Angela E Masiello" THEN worker_number = "X127GG5"
    IF assigned_to = "Filmon K Michael" THEN worker_number =  "X127GG6"
    IF assigned_to = "Kristina T Truong" THEN worker_number = "X127GG7"
    IF assigned_to = "Ben Teskey" THEN worker_number = 	  "X127GH4"
    IF assigned_to = "Ashley K Mack" THEN worker_number = "X127GH5"
    IF assigned_to = "Dave Mootz" THEN worker_number = 	  "X127GH6"
    IF assigned_to = "Tania L Amadi" THEN worker_number = "X127GJ1"
    IF assigned_to = "Marlenne Gonzalez" THEN worker_number = "X127GJ3"
    IF assigned_to = "Molly Irwin" THEN worker_number = 	 "X127GJ4"
    IF assigned_to = "Pahoua X Vang" THEN worker_number = "X127GJ8"
    IF assigned_to = "Keith J Semmelink" THEN worker_number = "X127GK2"
    IF assigned_to = "Claudia Alvarez" THEN worker_number = 	 "X127GM5"
    IF assigned_to = "Jaxon P VeeVahn" THEN worker_number = 	 "X127GR1"
    IF assigned_to = "Kristina Poplavska" THEN worker_number = 	  "X127GR2"
    IF assigned_to = "Adam R Verschoor" THEN worker_number =  "X127GR3"
    IF assigned_to = "Toni S Miles" THEN worker_number =  "X127GR4"
    IF assigned_to = "Mark Vilayrack" THEN worker_number = 	  "X127GR5"
    IF assigned_to = "Mali Lor" THEN worker_number =  "X127GR8"
    IF assigned_to = "Russell D Meelberg" THEN worker_number = 	  "X127GR9"
    IF assigned_to = "Robert D Warmboe" THEN worker_number =  "X127GS2"
    IF assigned_to = "Alyssa L Ackert" THEN worker_number = 	 "X127GU2"
    IF assigned_to = "Susan M Eeten" THEN worker_number = "X127GX9"
    IF assigned_to = "Peggy M. Benkert" THEN worker_number =  "X127H35"
    IF assigned_to = "Jill I Niess" THEN worker_number =  "X127HJ9"
    IF assigned_to = "Hind S Mahmoud" THEN worker_number = 	  "X127HM1"
    IF assigned_to = "Kari Anderson" THEN worker_number = "X127HN1"
    IF assigned_to = "Nicole C Arm" THEN worker_number =  "X127HN3"
    IF assigned_to = "Terry J Burgess" THEN worker_number = 	 "X127HN5"
    IF assigned_to = "Iyana R Galloway" THEN worker_number =  "X127HN8"
    IF assigned_to = "Shaneka Greer" THEN worker_number = "X127HN9"
    IF assigned_to = "Rochelle C Lane" THEN worker_number = 	 "X127HP3"
    IF assigned_to = "Caire Mckenzie" THEN worker_number = 	  "X127HP8"
    IF assigned_to = "Alisha E Mitchell" THEN worker_number = "X127HP9"
    IF assigned_to = "Anastacia R Ruiz" THEN worker_number =  "X127HQ3"
    IF assigned_to = "Mathilda M Stevenson" THEN worker_number =  "X127HQ6"
    IF assigned_to = "Florence R Williams" THEN worker_number = 	 "X127HR1"
    IF assigned_to = "Julie Xiong" THEN worker_number = 	 "X127HR2"
    IF assigned_to = "Osob A Ali" THEN worker_number = 	  "X127HS2"
    IF assigned_to = "Ahmednor M Farah" THEN worker_number =  "X127HS4"
    IF assigned_to = "Abby Korenchen" THEN worker_number = 	  "X127HS6"
    IF assigned_to = "Samsam A Mohamed" THEN worker_number =  "X127HS7"
    IF assigned_to = "Alexandra G Guardado Saenz" THEN worker_number = 	  "X127HT1"
    IF assigned_to = "Nikolai Kravets" THEN worker_number = 	 "X127HT5"
    IF assigned_to = "Nellie A Woodson" THEN worker_number =  "X127HT9"
    IF assigned_to = "Mohamed Abdirahman" THEN worker_number = 	  "X127HU2"
    IF assigned_to = "Yunuen A Avila" THEN worker_number = 	  "X127HU3"
    IF assigned_to = "Sakaria O Ashiro" THEN worker_number =  "X127HU4"
    IF assigned_to = "Inez R Toles" THEN worker_number =  "X127IRT"
    IF assigned_to = "Jonathan N Drogue" THEN worker_number = "X127J1D"
    IF assigned_to = "Jaime Lavallee" THEN worker_number = 	  "X127J1L"
    IF assigned_to = "Michelle Le" THEN worker_number = 	 "X127J75"
    IF assigned_to = "DoanTrinh M Thai" THEN worker_number =  "X127J77"
    IF assigned_to = "Chao Lee" THEN worker_number =  "X127J8C"
    IF assigned_to = "Rachel A Henry" THEN worker_number = 	  "X127J8E"
    IF assigned_to = "Janet Thompson" THEN worker_number = 	  "X127J8G"
    IF assigned_to = "Kemal T Deko" THEN worker_number =  "X127J8H"
    IF assigned_to = "Antionette L Jenkins" THEN worker_number =  "X127J8I"
    IF assigned_to = "Maggie Karley" THEN worker_number = "X127J8J"
    IF assigned_to = "Prophetia Castin" THEN worker_number =  "x127J9N"
    IF assigned_to = "Lisa M Castile" THEN worker_number = 	  "x127J9O"
    IF assigned_to = "Ikela I Cosey" THEN worker_number = "x127J9Q"
    IF assigned_to = "Mayra J Cota" THEN worker_number =  "x127J9R"
    IF assigned_to = "Cornel C Culp" THEN worker_number = "X127J9S"
    IF assigned_to = "Olga L Gonzalez" THEN worker_number = 	 "x127J9V"
    IF assigned_to = "Charnice W Jackson" THEN worker_number = 	  "x127J9X"
    IF assigned_to = "Gawa T Kalsang" THEN worker_number = 	  "x127J9Y"
    IF assigned_to = "Anna Mahon" THEN worker_number = 	  "X127J9Z" 'wasnt found in maxis'
    IF assigned_to = "Jennifer A Merritt" THEN worker_number = 	  "X127JAC"
    IF assigned_to = "Jaimee M Schwark" THEN worker_number =  "X127JAI"
    IF assigned_to = "John D Holmquist" THEN worker_number =  "X127JB2"
    IF assigned_to = "Hollie L Allen" THEN worker_number = 	  "X127JD1"
    IF assigned_to = "Matthew M Lane" THEN worker_number = 	  "X127JD5"
    IF assigned_to = "Mai V Lee" THEN worker_number = "X127JD6"
    IF assigned_to = "Amber Lowe" THEN worker_number = 	  "X127JD8"
    IF assigned_to = "Kongmeng Vang" THEN worker_number = "X127JE6"
    IF assigned_to = "Goua Yang" THEN worker_number = "X127JE9"
    IF assigned_to = "Jonathan Reeck" THEN worker_number = 	  "X127JF1"
    IF assigned_to = "Jessica R Harris" THEN worker_number =  "X127JH1"
    IF assigned_to = "Jenny Hong" THEN worker_number = 	  "X127JH3"
    IF assigned_to = "Janine M Hudson" THEN worker_number = 	 "X127JH4"
    IF assigned_to = "Filsan A Amin" THEN worker_number = "X127JK1"
    IF assigned_to = "Scott L Anderson" THEN worker_number =  "X127JK3"
    IF assigned_to = "Regina Andrews" THEN worker_number = 	  "X127JK4"
    IF assigned_to = "Betty J Allabough" THEN worker_number = "X127JK6"
    IF assigned_to = "Ronisha S Buckner" THEN worker_number = "X127JK7"
    IF assigned_to = "Basma A Mohamed" THEN worker_number = 	 "X127JK8"
    IF assigned_to = "Ivy Chiinze" THEN worker_number = 	 "X127JL7"
    IF assigned_to = "Javette L. Banks" THEN worker_number =  "X127JLB"
    IF assigned_to = "Kimyader M Dodd" THEN worker_number = 	 "X127JM4"
    IF assigned_to = "Valerie Hurst-Baker" THEN worker_number = 	 "X127JM6"
    IF assigned_to = "Breauna M Jackson" THEN worker_number = "X127JM7"
    IF assigned_to = "Maria Vang" THEN worker_number = 	  "X127JN4"
    IF assigned_to = "Maikou Yang" THEN worker_number = 	 "X127JN5"
    IF assigned_to = "Maria J Harald" THEN worker_number = 	  "X127JO1"
    IF assigned_to = "Andy A Knutson" THEN worker_number = 	  "X127JS7"
    IF assigned_to = "Christa J Burdette" THEN worker_number = 	  "X127JS8"
    IF assigned_to = "Allicia M Brolsma" THEN worker_number = "X127JT1" 'not found in maxis'
    IF assigned_to = "Angela M Docken" THEN worker_number = 	 "X127JT3"
    IF assigned_to = "Bee Lee" THEN worker_number = 	 "X127JT7"
    IF assigned_to = "John B Niemi" THEN worker_number =  "X127JU5"
    IF assigned_to = "Autumn J O'Brien" THEN worker_number =  "X127JU6"
    IF assigned_to = "Dan Rubenstein" THEN worker_number = 	  "X127JU8"
    IF assigned_to = "Staci Ulmen" THEN worker_number = 	 "X127JV3"
    IF assigned_to = "Nicole Ocampo" THEN worker_number = "X127JV6"
    IF assigned_to = "Mai-Ling Mui" THEN worker_number =  "X127JV7"
    IF assigned_to = "Tameka Ballard" THEN worker_number = 	  "X127JW3"
    IF assigned_to = "Cha Her" THEN worker_number = 	 "X127JW4"
    IF assigned_to = "Pajci Ly" THEN worker_number =  "X127JX2"
    IF assigned_to = "Deja L Martin" THEN worker_number = "X127JX5"
    IF assigned_to = "Zachary E Nagle" THEN worker_number = 	 "X127JX7"
    IF assigned_to = "Hangatu Omar" THEN worker_number =  "X127JX8"
    IF assigned_to = "Sarah A Haigh" THEN worker_number = "X127JX9"
    IF assigned_to = "Phyllicia C Wise" THEN worker_number =  "X127JY3"
    IF assigned_to = "Thadeus M Teske" THEN worker_number = 	 "X127K81"
    IF assigned_to = "Lisa A Nelson" THEN worker_number = "X127K82"
    IF assigned_to = "Louise A Kinzer" THEN worker_number = 	 "X127K85"
    IF assigned_to = "Jodi R Ojala" THEN worker_number =  "X127K9D"
    IF assigned_to = "Diki Wangkhang-Phuntsok" THEN worker_number = 	 "X127K9E"
    IF assigned_to = "Kelly A Quigley" THEN worker_number = 	 "X127K9F"
    IF assigned_to = "Maribel Navarrete Reyes" THEN worker_number = 	 "X127K9G"
    IF assigned_to = "Tamrat A Shulu" THEN worker_number = 	  "X127K9J"
    IF assigned_to = "Sabastian Boyle-Mejia" THEN worker_number = "X127KD3"
    IF assigned_to = "Erick Diaz-Contreras" THEN worker_number =  "X127KD5"
    IF assigned_to = "Abdifitaah Herei" THEN worker_number =  "X127KD7"
    IF assigned_to = "Grecia M Lagunes" THEN worker_number =  "X127KD8"
    IF assigned_to = "Ana K Moreno De La Garza" THEN worker_number =  "X127KE1"
    IF assigned_to = "Iman Said" THEN worker_number = "X127KE2"
    IF assigned_to = "Karina Santana" THEN worker_number = 	  "X127KE4"
    IF assigned_to = "Mark X Schmidt" THEN worker_number = 	  "X127KE5"
    IF assigned_to = "Josefina G Greene" THEN worker_number = "X127KE7"
    IF assigned_to = "Kris Koukkari" THEN worker_number = "X127KEK"
    IF assigned_to = "Mihiret A Abrahim" THEN worker_number = "X127KL3"
    IF assigned_to = "Genni M Lillibridge" THEN worker_number = 	 "X127KL4"
    IF assigned_to = "Hannah E Broman" THEN worker_number = 	 "X127KL6"
    IF assigned_to = "Aisha A Dancy" THEN worker_number = "X127KL7"
    IF assigned_to = "Signe Faulhaber" THEN worker_number = 	 "X127KL8"
    IF assigned_to = "Katie M Flanigan" THEN worker_number =  "X127KL9"
    IF assigned_to = "Lyvia C Guallpa" THEN worker_number = 	 "X127KM3"
    IF assigned_to = "LaRae Heard" THEN worker_number = 	 "X127KM4"
    IF assigned_to = "Andrew J Howard" THEN worker_number = 	 "X127KM5"
    IF assigned_to = "Ariel F Brown" THEN worker_number = "X127KM6"
    IF assigned_to = "Sandy L Jorgensen" THEN worker_number = "X127KM7"
    IF assigned_to = "Abraham T Page" THEN worker_number = 	  "X127KM9"
    IF assigned_to = "Khadra M Nur" THEN worker_number =  "X127KMN"
    IF assigned_to = "Angela A Schottle" THEN worker_number = "X127KN2"
    IF assigned_to = "Gao J Vang" THEN worker_number = 	  "X127KN3"
    IF assigned_to = "Ana K Villegas Gonzalez" THEN worker_number = 	 "X127KN4"
    IF assigned_to = "Hailey C Young" THEN worker_number = 	  "X127KN5"
    IF assigned_to = "Natasha Williams" THEN worker_number =  "X127KN6"
    IF assigned_to = "Ashley A Yang" THEN worker_number = "X127KN7"
    IF assigned_to = "Joan Zangs" THEN worker_number = 	  "X127KN8"
    IF assigned_to = "Katie Aanestad" THEN worker_number = 	  "X127KN9"
    IF assigned_to = "Robert Bohr" THEN worker_number = 	 "X127KP3"
    IF assigned_to = "Sainabou Jaye-Marong" THEN worker_number =  "X127KP4"
    IF assigned_to = "Shauna Kirscht" THEN worker_number = 	  "X127KP7"
    IF assigned_to = "Dawn T Olmstead" THEN worker_number = 	 "X127KP9" 'not found in maxis'
    IF assigned_to = "Khristiane C Victorio" THEN worker_number = "X127KQ2"
    IF assigned_to = "LaShay C Waters" THEN worker_number = 	 "X127KQ3"
    IF assigned_to = "Olga V Engebretson" THEN worker_number = 	  "X127L04"
    IF assigned_to = "Deborah L Rusnak" THEN worker_number =  "X127L08"
    IF assigned_to = "Casey H Love" THEN worker_number =  "X127L1S"
    IF assigned_to = "Paul E. Madison" THEN worker_number = 	 "X127L23"
    IF assigned_to = "Saeed D. Jibrell" THEN worker_number =  "X127L87"
    IF assigned_to = "Lisa M Lampkin" THEN worker_number = 	  "X127LL1"
    IF assigned_to = "Lamar Salinas-Niemczycki" THEN worker_number =  "X127LN1"
    IF assigned_to = "Loranzie S Rogers" THEN worker_number = "X127LSR"
    IF assigned_to = "Lee Vang" THEN worker_number =  "X127LV2"
    IF assigned_to = "Clar Weller" THEN worker_number = 	 "X127M16"
    IF assigned_to = "Barbara Herrera" THEN worker_number = 	 "X127M22"
    IF assigned_to = "Janelle L Sorenson" THEN worker_number = 	  "X127M49"
    IF assigned_to = "Alla Ternyak" THEN worker_number =  "X127M93"
    IF assigned_to = "Mariah R Burgess" THEN worker_number =  "X127MBR"
    IF assigned_to = "Monica F Parham" THEN worker_number = 	 "X127MFP"
    IF assigned_to = "Michelle Cline" THEN worker_number = 	  "X127MLC"
    IF assigned_to = "Marnya Rudolph" THEN worker_number = 	  "X127MMR"
    IF assigned_to = "Meena Thao" THEN worker_number = 	  "X127MT1"
    IF assigned_to = "Michael Vegell" THEN worker_number = 	  "X127MV4"
    IF assigned_to = "Tracy A Gorman" THEN worker_number = 	  "X127N54"
    IF assigned_to = "Nabila S Abdullahi" THEN worker_number = 	  "X127NSA"
    IF assigned_to = "Nicole Tyson" THEN worker_number =  "X127NT1"
    IF assigned_to = "Nancy N Shevich" THEN worker_number = 	 "X127NXS"
    IF assigned_to = "Olukemi O Adeniyi-Akins" THEN worker_number = 	 "X127OLU"
    IF assigned_to = "Omar Zavala" THEN worker_number = 	 "X127OZA"
    IF assigned_to = "Mee Thao" THEN worker_number =  "X127PC2"
    IF assigned_to = "Alison L Watkins" THEN worker_number =  "X127PC4"
    IF assigned_to = "Keenya C Woods" THEN worker_number = 	  "X127PC6"
    IF assigned_to = "Daniel D Benfield" THEN worker_number = "X127Q73"
    IF assigned_to = "Sahr K Sandi" THEN worker_number =  "X127Q90"
    IF assigned_to = "Shanna C Hansen" THEN worker_number = 	 "X127Q95"
    IF assigned_to = "Nina Jones" THEN worker_number = 	  "X127R49"
    IF assigned_to = "Andre S. Xiong" THEN worker_number = 	  "X127R63"
    IF assigned_to = "Olga P Bugayev" THEN worker_number = 	  "X127R74"
    IF assigned_to = "Ann Becker" THEN worker_number = 	  "X127R76"
    IF assigned_to = "Rahma Ali" THEN worker_number = "X127RBA"
    IF assigned_to = "Roberta J Howard" THEN worker_number =  "X127RJH"
    IF assigned_to = "Rita M Phelps" THEN worker_number = "X127S01"
    IF assigned_to = "Cheryl A Deason" THEN worker_number = 	 "X127S08"
    IF assigned_to = "Shavon A Johnson" THEN worker_number =  "X127SAJ"
    IF assigned_to = "Sekena Britt-Nelson" THEN worker_number = 	 "X127SBN"
    IF assigned_to = "Sarah C Campbell" THEN worker_number =  "X127SCC"
    IF assigned_to = "Sheena A Dempsey" THEN worker_number =  "X127SD4"
    IF assigned_to = "Sarah E LaCoursiere" THEN worker_number = 	 "X127SEL"
    IF assigned_to = "Shamilia Fisher" THEN worker_number = 	 "X127SF1"
    IF assigned_to = "Kary J Van Slyke" THEN worker_number =  "X127T21"
    IF assigned_to = "Lindsey H Remus" THEN worker_number = 	 "X127T25"
    IF assigned_to = "Kelly C Flanigan" THEN worker_number =  "X127T50"
    IF assigned_to = "Sheryn L Cartlidge" THEN worker_number = 	  "X127T63"
    IF assigned_to = "Claudia A Saulter" THEN worker_number = "X127T65"
    IF assigned_to = "Trenita M Heard" THEN worker_number = 	 "X127T67"
    IF assigned_to = "Maytong Yang" THEN worker_number =  "X127T81"
    IF assigned_to = "Tinna C Tin" THEN worker_number = 	 "X127TCT"
    IF assigned_to = "Darren L Terebenet" THEN worker_number = 	  "X127TDL"
    IF assigned_to = "Timothy J Remme" THEN worker_number = 	 "X127TJR"
    IF assigned_to = "Teisha M. Broomfield" THEN worker_number =  "X127TMB"
    IF assigned_to = "Sideth D Niev" THEN worker_number = "X127U55"
    IF assigned_to = "Neil Urbanski" THEN worker_number = "X127UN1"
    IF assigned_to = "Ugbad R Abdilahi" THEN worker_number =  "X127URA"
    IF assigned_to = "Veronica L Keys" THEN worker_number = 	 "X127VLk"
    IF assigned_to = "Vickie Mansheim" THEN worker_number = 	 "X127VSM"
    IF assigned_to = "Malinda M Rolack" THEN worker_number =  "X127W12"
    IF assigned_to = "Angela S Beljeski" THEN worker_number = "X127W18"
    IF assigned_to = "Alyssa A Ryan" THEN worker_number = "X127W23"
    IF assigned_to = "Shaquila Harris" THEN worker_number = 	 "X127W40"
    IF assigned_to = "Terry Baker" THEN worker_number = 	 "X127W69"
    IF assigned_to = "Deborah A Lechner" THEN worker_number = "X127W88"
    IF assigned_to = "Wendy S Irwin" THEN worker_number = "X127WI1"
    IF assigned_to = "Wendy L Clark" THEN worker_number = "X127WLC"
    IF assigned_to = "Wendy M Bedoya" THEN worker_number = 	  "X127WM1"
    IF assigned_to = "Ephrem Nejo" THEN worker_number = 	 "X127X04"
    IF assigned_to = "Jessica A Hall" THEN worker_number = 	  "X127X05"
    IF assigned_to = "Xoua Xiong" THEN worker_number = 	  "X127X0X"
    IF assigned_to = "Deanna L Deloach" THEN worker_number =  "X127X27"
    IF assigned_to = "Melissa M Isais" THEN worker_number = 	 "X127X29"
    IF assigned_to = "Deborah F Bolden" THEN worker_number =  "X127X41"
    IF assigned_to = "Pamela Y Whitson" THEN worker_number =  "X127X43"
    IF assigned_to = "Osman I Ali" THEN worker_number = 	 "X127X59"
    IF assigned_to = "Deedra C Miller" THEN worker_number = 	 "X127X63"
    IF assigned_to = "Tiffanie Mrsich" THEN worker_number = 	 "X127X82"
    IF assigned_to = "Irro A Mohamed" THEN worker_number = 	  "X127X96"
    IF assigned_to = "Joseph Nelson" THEN worker_number = "X127X97"
    IF assigned_to = "Neill C Burnett" THEN worker_number = "X127X99"
    IF assigned_to = "Abdullahi A Berka" THEN worker_number = "X127Y04"
    IF assigned_to = "Cassandra M Lis" THEN worker_number = "X127Y05"
    IF assigned_to = "Tracy L Mohomes" THEN worker_number = "X127Y23"
    IF assigned_to = "Patricia A Hegenbarth" THEN worker_number = "X127Y37"
    IF assigned_to = "Deborah Diggins" THEN worker_number = "X127Y43"
    IF assigned_to = "Natalya Ditter" THEN worker_number = 	"X127Y44"
    IF assigned_to = "Jennifer K Moses" THEN worker_number =  "X127Y62"
    IF assigned_to = "Jamila A Jama" THEN worker_number = "X127Y76"
    IF assigned_to = "Shelly A Lind" THEN worker_number = "X127Y81"
    IF assigned_to = "Rhea Blue Arm" THEN worker_number = "X127Y86"
    IF assigned_to = "Letitia Lewis" THEN worker_number = "X127Y92"
    IF assigned_to = "Halane Yussuf" THEN worker_number = "X127YH1"
    IF assigned_to = "Yee Yang" THEN worker_number =  "X127YY1"
    IF assigned_to = "Raisa D Loevski" THEN worker_number = "X127Z34"
    IF assigned_to = "Yasmin F Waite" THEN worker_number = 	 "X127Z45"
    IF assigned_to = "Muna M Afrah" THEN worker_number =  "X127Z46"
    IF assigned_to = "Bernita R Steele" THEN worker_number =  "X127Z62"
    IF assigned_to = "Sara K Harrell" THEN worker_number = "X127Z71"
    IF assigned_to = "Kristine R Norman" THEN worker_number = "X127Z83"
    IF assigned_to = "Kia Lee" THEN worker_number =  "X127Z86"
    IF assigned_to = "Judy C Vang" THEN worker_number =  "X127Z87"
    IF assigned_to = "Bonsi A Abraham" THEN worker_number =  "X127Z91"
    IF assigned_to = "Huruse M Gurhan" THEN worker_number =  "X127Z93"
    IF assigned_to = "Heather J Feldmann" THEN worker_number = 	"X127ZAE"
    IF assigned_to = "Tamika Hannah" THEN worker_number = "X127ZAH"
    IF assigned_to = "Lauren A John" THEN worker_number = "X127ZAJ"
    IF assigned_to = "Amy N Kelvie" THEN worker_number =  "X127ZAK"
    IF assigned_to = "Darren Konsor" THEN worker_number = "X127ZAL"
END FUNCTION
    'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

''----------------------------------------------------------------------------------------------------The current day's assignment
'report_date = replace(date, "/", "-")   'Changing the format of the date to use as file path selection default
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the file names
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
'file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & file_date & ".xlsx"

BeginDialog Dialog1, 0, 0, 266, 115, "TASK BASED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used for task based review on a list of pending SNAP and/or MFIP cases."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'dialog and dialog DO...Loop
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
'objExcel.worksheets("Report 1").Activate                                 'Activates the initial BOBI report

'Establishing array
DIM task_based_array()           'Declaring the array
ReDim task_based_array(18, 0)     'Resizing the array
'Creating constants to value the array elements
const date_assigned_const       = 0 '= "Date Assigned"
const SSR_name_const	        = 1 '= "SSR Name"
const maxis_case_number_const 	= 2 '= "Case Number"
const case_name_const       	= 3 '= "Case Name"
const basket_const  			= 4 '= "Basket"
const assigned_to_const    		= 5 '= "Assigned to"
const worker_number_const       = 6 '= "Assigned Worker X127#"
const do_this_const        		= 7 '= "Does worker log indicate they could work the case?"
const case_logged_const         = 8 '= "Case logged by assigned worker?"
const case_note_date_const      = 9  '= "Case Note Date"
const case_note_match_const     = 10 '= "Worker who made case note same as assigned worker"
const case_note_keyword_const   = 11 '= "Does case note title contain keyword"
const total_DAIL_count_const   	= 12 '= "DAIL Count"
const action_DAIL_count_const   = 13 '= "DAIL Type"
const ECF_type_const            = 14 '= "EWS ECF Item Count"
const ECF_form_const            = 15 '= "ECF Form Types"
const oldest_APPL_date_const    = 16 '= "Oldest ECF APPL Date"
const prev_comments_const       = 17 '= "Comments"
const case_status_const 		= 18 '= "Pending over 30 days"
const interview_const           = 19 '= "Interview Completed"

'Now the script adds all the clients on the excel list into an array
excel_row = 5                   're-establishing the row to start based on when Report 1 starts
entry_record = 0                'incrementor for the array and count
all_case_numbers_array = "*"    'setting up string to find duplicate case numbers
Do
    'Reading information from the Excel
    worker_number = objExcel.cells(excel_row, 6).Value
    worker_number = trim(worker_number)

    MAXIS_case_number = objExcel.cells(excel_row, 2).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do

    application_date = objExcel.cells(excel_row, 15).Value
    interview_date   = objExcel.cells(excel_row, 19).Value

    days_pending = datediff("D", application_date, date)

    'If the case number is found in the string of case numbers, it's not added again.
    If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then
        add_to_array = False
    Else
        'Adding client information to the array
        ReDim Preserve task_based_array(19, entry_record)	'This resizes the array based on the number of cases

		task_based_array(date_assigned_const,      entry_record) = ""
		task_based_array(SSR_name_const, 		   entry_record) = ""
		task_based_array(maxis_case_number_const,  entry_record) = ""
		task_based_array(case_name_const,          entry_record) = ""
		task_based_array(basket_const,  		   entry_record) = ""
		task_based_array(assigned_to_const,        entry_record) = ""
		task_based_array(worker_number_const,      entry_record) = ""
		task_based_array(do_this_const,            entry_record) = ""
		task_based_array(case_logged_const,    	   entry_record) = ""
		task_based_array(case_note_count_const,     entry_record) = ""
		task_based_array(case_note_date_const,     entry_record) = ""
		task_based_array(case_note_match_const,    entry_record) = ""
		task_based_array(case_note_keyword_const,  entry_record) = ""
		task_based_array(total_DAIL_count_const,   entry_record) = ""
		task_based_array(DAIL_count_const,         entry_record) = ""
		task_based_array(ECF_form_const,   		   entry_record) = ""
		task_based_array(ECF_type_const,           entry_record) = ""
		task_based_array(oldest_APPL_date_const,   entry_record) = ""
		task_based_array(prev_comments_const,      entry_record) = ""
		task_based_array(case_status_const,        entry_record) = ""
		task_based_array(interview_const,          entry_record) = ""  '= "Interview Completed"
		'making space in the array for these variables, but valuing them as "" for now

        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter
    End if
    excel_row = excel_row + 1
Loop

back_to_self                            'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'Loading of cases is complete. Reviewing the cases in the array.
For item = 0 to UBound(task_based_array, 2)
    worker_number       = task_based_array(worker_number_const,    item)     're-valuing array variables
    MAXIS_case_number   = task_based_array(case_number_const,      item)
    program_ID          = task_based_array(program_ID_const,       item)
    days_pending        = task_based_array(days_pending_const,     item)
    application_date    = task_based_array(application_date_const, item)
	MAXIS_case_name     = task_based_array(application_date_const, item)
	date_assigned		= task_based_array(date_assigned_const,    item)
	assigned_to     	= task_based_array(assigned_to_const,      item)
	case_note_count		= task_based_array(case_note_count_const,  item)
	DAIL_count   		= task_based_array(DAIL_count_const,       item)
	CALL HSR_LIST 'this will get our worker number and name '

	'setting the footer month to make the updates in'
	CALL convert_date_into_MAXIS_footer_month(date_received, MAXIS_footer_month, MAXIS_footer_year)
	CALL MAXIS_footer_month_confirmation
    Call navigate_to_MAXIS_screen("STAT", "PROG")
    	EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROD and INQUIRY
    IF priv_check = "PRIV" then                                             'PRIV cases
        EMReadscreen priv_worker, 26, 24, 46
        task_based_array(case_status_const, item) = trim(priv_worker)
        task_based_array(do_this_const, item) = "Privileged Cases"
    ELSE
        EMReadScreen county_code, 4, 21, 21                                 'Out of county cases from STAT
        If county_code <> "X127" then
            task_based_array(case_status_const, item) = "OUT OF COUNTY CASE"
        End if
	'ELSE
		'EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to  navigate'
		'task_based_array(case_status_const, item) = trim(case_invalid_error)
		'task_based_array(do_this_const, item) = "Error Message"
		'PF10
	END IF
	'Reading the app date from PROG need to compare for over 30 days and the interview stuffs
	EMReadScreen cash1_app_date, 8, 6, 33
	cash1_app_date = replace(cash1_app_date, " ", "/")
	EMReadScreen cash2_app_date, 8, 7, 33
	cash2_app_date = replace(cash2_app_date, " ", "/")
	EMReadScreen emer_app_date, 8, 8, 33
	emer_app_date = replace(emer_app_date, " ", "/")
	EMReadScreen grh_app_date, 8, 9, 33
	grh_app_date = replace(grh_app_date, " ", "/")
	EMReadScreen snap_app_date, 8, 10, 33
	snap_app_date = replace(snap_app_date, " ", "/")
	EMReadScreen ive_app_date, 8, 11, 33
	ive_app_date = replace(ive_app_date, " ", "/")
	EMReadScreen hc_app_date, 8, 12, 33
	hc_app_date = replace(hc_app_date, " ", "/")
	EMReadScreen cca_app_date, 8, 14, 33
	cca_app_date = replace(cca_app_date, " ", "/")

' was an interview completed on the assingment day '
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
    IF access_denied_check = "ACCESS DENIED" Then
        PF10
        last_name = "UNABLE TO FIND"
        first_name = ""
        mid_initial = ""
    ELSE
        EMReadscreen last_name, 25, 6, 30
        EMReadscreen first_name, 12, 6, 63
        last_name = trim(replace(last_name, "_", ""))
        first_name = trim(replace(first_name, "_", ""))
    	MAXIS_case_name = first_name & " "  & last_name
	END IF
	task_based_array(MAXIS_case_name_const, item) = MAXIS_case_name

	CALL ONLY_create_MAXIS_friendly_date(date_assigned)
    CALL navigate_to_MAXIS_screen("CASE", "NOTE")
    MAXIS_row = 5                      'Defining row for the search feature.
	case_note_count = 0			'setting to zero'
	interview_done = False

	Call ONLY_create_MAXIS_friendly_date(date_assigned)
	Do
		EMReadscreen case_note_date, 8, MAXIS_row, 6
		If trim(case_note_date) = "" then
			exit do
		Else
			IF case_note_date = date_assigned THEN
				EMReadScreen case_note_worker_number, 7, MAXIS_row, 16
				IF worker_number = case_note_worker_number THEN
					case_note_count = case_note_count + 1
					EMReadScreen case_note_header, 55, MAXIS_row, 25
                	case_note_header = lcase(trim(case_note_header))
					If instr(case_note_header, "Interview Completed") then interview_done = True 	'TODO: check the string for case ntoe header & output results to array & excel
                END IF
            End if
		END IF

	    MAXIS_row = MAXIS_row + 1
		IF MAXIS_row = 19 then
            PF8                         'moving to next case note page if at the end of the page
            MAXIS_row = 5
		End if
    LOOP until cdate(case_note_date) < cdate(date_assigned)                        'repeats until the case note date is less than the assignment date

	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
	DO
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
	Loop until dail_check = "DAIL"

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

	'TODO: output dail_count into array

	'1918125' fUNCTIOONALTIY
	EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.

	If message_error = "NO MESSAGES" then 'NO MESSAGES FOR CASE XXXXXXXX SELECTED-PF5 FOR TOP
		CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		Call write_value_and_transmit(worker, 21, 6)
		TRANSMIT  'transmit past 'THIS IS NOT YOUR DAIL REPORT
    	exit do
	End if
    'End if
Next

'Excel output of cases and information in their applicable categories - PRIV, Req EXP Processing, Exp Screening Required, Not Expedited
Msgbox "Output to Excel Starting."      'warning message to whomever is running the script

'time line of actual runs
'todo save as copy and see how long it takes to run their actual list'

    ObjExcel.Worksheets.Add().Name = task_status
	ObjExcel.Cells(1, 1).Value = "Date Assigned"
	ObjExcel.Cells(1, 2).Value = "SSR Name"
	ObjExcel.Cells(1, 3).Value = "Case Number"
	ObjExcel.Cells(1, 4).Value = "Case Name"
	ObjExcel.Cells(1, 5).Value = "Basket"
	ObjExcel.Cells(1, 6).Value = "Assigned to"
	ObjExcel.Cells(1, 7).Value = "Assigned Worker X127#"
	ObjExcel.Cells(1, 8).Value = "Does worker log indicate they could work the case?"
	ObjExcel.Cells(1, 9).Value = "Case logged by assigned worker?"
	ObjExcel.Cells(1, 10).Value = "Case Note Date"
	ObjExcel.Cells(1, 11).Value = "Worker who made case note same as assigned worker"
	ObjExcel.Cells(1, 12).Value = "Does case note title contain keyword"
	ObjExcel.Cells(1, 13).Value = "DAIL Count"
	ObjExcel.Cells(1, 14).Value = "DAIL Type"
	ObjExcel.Cells(1, 15).Value = "EWS ECF Item Count"
	ObjExcel.Cells(1, 16).Value = "ECF Form Types"
	ObjExcel.Cells(1, 17).Value = "Oldest ECF APPL Date"
	ObjExcel.Cells(1, 18).Value = "Comments"
	ObjExcel.Cells(1, 19).Value = "Pending over 30 days"
	ObjExcel.Cells(1, 20).Value = "Interview Completed"

	objExcel.Columns(1).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    objExcel.Columns(10).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    objExcel.Columns(17).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY

    Excel_row = 2

    For item = 0 to UBound(task_based_array, 2)
	    objExcel.Cells(excel_row, 1).Value = task_based_array(date_assigned_const,      item)
	    objExcel.Cells(excel_row, 2).Value = task_based_array(SSR_name_const, 		    item)
	    objExcel.Cells(excel_row, 3).Value = task_based_array(maxis_case_number_const,  item) = MAXIS_case_number
	    objExcel.Cells(excel_row, 4).Value = task_based_array(case_name_const,          item) = MAXIS_case_name
	    objExcel.Cells(excel_row, 5).Value = task_based_array(basket_const,  		    item) = ""
	    objExcel.Cells(excel_row, 6).Value = task_based_array(assigned_to_const,        item) = ""
	    objExcel.Cells(excel_row, 7).Value = task_based_array(worker_number_const,      item) = worker_number
	    objExcel.Cells(excel_row, 8).Value = task_based_array(do_this_const,            item)
	    objExcel.Cells(excel_row, 9).Value = task_based_array(case_logged_const,    	item)
	    objExcel.Cells(excel_row, 10).Value = task_based_array(case_note_date_const,    item)
	    objExcel.Cells(excel_row, 11).Value = task_based_array(case_note_match_const,   item)
	    objExcel.Cells(excel_row, 12).Value = task_based_array(case_note_keyword_const, item)
	    objExcel.Cells(excel_row, 13).Value = task_based_array(total_DAIL_count_const,  item)
	    objExcel.Cells(excel_row, 14).Value = task_based_array(action_DAIL_count_const, item)
	    objExcel.Cells(excel_row, 15).Value = task_based_array(ECF_form_const,   		item)
	    objExcel.Cells(excel_row, 16).Value = task_based_array(ECF_type_const,          item)
	    objExcel.Cells(excel_row, 17).Value = task_based_array(oldest_APPL_date_const,  item) = trim(application_date)
	    objExcel.Cells(excel_row, 18).Value = task_based_array(prev_comments_const,     item) = program_ID
	    objExcel.Cells(excel_row, 19).Value = task_based_array(case_status_const,       item) = days_pending
	    objExcel.Cells(excel_row, 20).Value = task_based_array(interview_const,         item) = trim(interview_date)
	    'making space in the array for these variables, but valuing them as "" for now
        excel_row = excel_row + 1
    Next

    FOR i = 1 to 8		'formatting the cells
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

objWorkbook.Save()  'saves existing workbook as same name
objExcel.Quit

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the run is complete. The workbook has been saved.")
