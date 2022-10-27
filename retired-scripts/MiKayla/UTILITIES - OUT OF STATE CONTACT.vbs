'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - OUT OF STATE CONTACT.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
 'END OF stats block=========================================================================================================

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
'The script----------------------------------------------------------------------------------------------------
EMCONNECT ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 146, 50, "Out-of-state contact information"
  DropListBox 60, 10, 80, 15, "Select one..."+chr(9)+"Alabama"+chr(9)+"Alaska"+chr(9)+"Arizona"+chr(9)+"Arkansas"+chr(9)+"California"+chr(9)+"Colorado"+chr(9)+"Connecticut"+chr(9)+"Delaware"+chr(9)+"District of Columbia"+chr(9)+ _
  "Florida"+chr(9)+"Georga"+chr(9)+"Guam"+chr(9)+"Hawaii"+chr(9)+"Idaho"+chr(9)+"Illinois"+chr(9)+"Indiana"+chr(9)+"Iowa"+chr(9)+"Kansas"+chr(9)+"Kentucky"+chr(9)+"Louisana"+chr(9)+ _
  "Maine"+chr(9)+"Maryland"+chr(9)+"Massachusetts"+chr(9)+"Michigan"+chr(9)+"Mississippi"+chr(9)+"Missouri"+chr(9)+"Montana"+chr(9)+"Nebraska"+chr(9)+"Nevada"+chr(9)+"New Hampshire"+chr(9)+"New Jersey"+chr(9)+ _
  "New Mexico"+chr(9)+"New York"+chr(9)+"North Carolina"+chr(9)+"North Dakota"+chr(9)+"Ohio"+chr(9)+"Oklahoma"+chr(9)+"Oregon"+chr(9)+"Pennsylvania"+chr(9)+"Rhode Island"+chr(9)+"Puerto Rico"+chr(9)+"South Carolina"+chr(9)+  _
  "South Dakota"+chr(9)+"Tennessee"+chr(9)+"Texas"+chr(9)+"Utah"+chr(9)+"Vermont"+chr(9)+"Virginia"+chr(9)+"Viriginia"+chr(9)+"Washington"+chr(9)+"West Virginia"+chr(9)+"Wisconsin"+chr(9)+"Wyoming", state_list
  ButtonGroup ButtonPressed
  OkButton 35, 30, 50, 15
  CancelButton 90, 30, 50, 15
  Text 10, 15, 50, 10, "Select a state:"
EndDialog
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If state_list = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select a state."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'info_array = (state_list, general_header, )

'Alabama
IF state_list = "Alabama" then
	general_header = "Out of State inquiries:"
	general_person = "Betty S.White, Program Supervisor"
	general_pref_contact = ""
	general_phone = "334-242-1745"
	general_fax = "334-353-1363"
	general_addr  = "Alabama Dept. of Human Resources: 50 Ripley St, S. Gordon Persons Bldg. Montgomery, AL 36130-4000"
	general_email = "Betty.White@dhr.alabama.gov"
	general_info = ""

	web_addr = "www.dhr.alabama.gov"
	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = "email"
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email  = "fs@dhr.alabama.gov"
	SNAP_info = ""

	If SNAP_header <> "" then SNAP_info_list = SNAP_header & "|" 'UNSURE WHAT THIS DOES ASK ILSE'

	SNAP_header =
	SNAP_person =
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = ""
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = "334-242-0513"
	CASH_addr  = ""
	CASH_email = "DHR_PA_Helpdesk@dhr.alabama.gov"
	CASH_info = "TANF/FIP-No TANF information given over the phone."

	HC_header = "Medical Benefits:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = "Phone (334) 242-5010"
	HC_fax = ""
	HC_addr  = ""
	HC_email = "PIDParis@medicaid.alabama.gov"
	HC_info = ""

	Claims_header = ""
	Claims_person = "Geraldine Turner"
	Claims_pref_contact = "email"
	Claims_phone = "251-450-7544"
	Claims_fax = "251-450-7544"
	Claims_addr  = "Same"
	Claims_email = "Geraldine.turner@dhr.alabama.gov"
	Claims_info = ""

'Alaska'
If state_list = "ALASKA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "Fax"
	general_phone = "907-465-3347"
    general_addr = "PO Box 110640, Juneau, AK 99811-0640"
    general_email = "DPApolicy@alaska.gov"
    general_info = "Contact the office from which the client last received benefits. A list of offices can be found at:
				    http://dhss.alaska.gov/dpa/Pages/features/org/dpado.aspx"
	web_addr = "http://www.hss.state.ak.us/dpa/"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = "Fax"
	SNAP_phone = "907-465-3347"
	SNAP_fax = "907-465-5254"
	SNAP_addr = "PO Box 110640, Juneau, AK 99811-0640"
	SNAP_email = "DPApolicy@alaska.gov"
	SNAP_info = "Contact the office from which the client last received benefits. A list of offices can be found at:
                     http://dhss.alaska.gov/dpa/Pages/features/org/dpado.aspx"

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assitance Participation"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Arizona'
If state_list = "ARIZONA" then
    general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "EMAIL"
	general_phone = "602-771-2047"
	general_fax = "602-353-5746"
	general_addr = ""
	general_email = "Azstateinquiries@Azdes.go"
	general_info = "Dept. of Economic Security Communication Center"

	web_addr = "www.des.az.gov"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = "EMAIL"
	SNAP_phone = "602-771-2047" 'MSG SYSTEM FOR STATUS REQUEST ONLY'
	SNAP_fax = "602-353-5746"
	SNAP_addr = ""
	SNAP_email = "Azstateinquiries@Azdes.gov"
	SNAP_info = "Dept. of Economic Security Communication Center"

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Arkansas'
If state_list = "ARKANSAS" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = ""
	general_fax= ""
	general_addr = ""
	general_email = ""
	general_info = ""

	web_addr = "www.arkansas.gov/dhhs"

	SNAP_header = "SNAP Participation:"
	SNAP_person = "Beverly Stewart-Coleman, Administrative Specialist II"
	SNAP_pref_contact = ""
	SNAP_phone = "501-682-8993"
	SNAP_fax = ""
	SNAP_addr = " Customer Assistance Unit, PO Box 1437, Slot S-341 Little Rock, Arkansas 72203-1437"
	SNAP_email = "beverly.stewart-coleman@dhs.arkansas.gov"
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'California'
If state_list = "CALIFORNIA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "916-651-8848" "press 1, then press 7"
	general_fax= "916-651-8866"
	general_addr = "California Department of Social Services 744 P Street, MS 8-4-23 Sacramento, CA 95814-6400"
	general_email = ""
	general_info = "The city and/or county in which the client resided
					in California must be provided in order to provide a referral to one of the 58 counties for
					verification of benefits.  See web address to get a Central County Index Listing."

	web_addr = "http://www.cdss.ca.gov/cdssweb/entres/pdf/CountyCentral"   "IndexListing.pdf"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Colorado'
If state_list = "COLORADO" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "1-800-536-5298"
	general_fax = ""
	general_addr = "Colorado Department of Human Services 1575 Sherman St. 3rd Fl Denver, CO 80203"
	general_email = "Outofstateinquiries@state.co.us"
	general_info = ""

	web_addr = "www.cdhs.state.co.us"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = "PARIS match"
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = "1-800-359-1991"
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	calims_fraud_line =
	claims_info = ""

'Connecticut'
If state_list = "CONNECTICUT" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "1-860-424-5030"
	general_fax = ""
	general_addr = "Department of Social Service. 55 Farmington Ave.Hartford, CT 06105-3725"
	general_email = "TPI.EU@ct.gov."
	general_info = "Please send an email from you State or County email	account on your agency’s letter head to: TPI.EU@ct.gov.
			Requests must include names, dates of birth and last 4 digits of SSN for each individual for whom verifications are being sought. Also include what	programs you need
			verified and the application address in your state.	Allow 3 to 5 working days for a response."

	web_addr = ""

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = "1-860-424-5540"
	CASH_fax = "1-860-424-4886"
	CASH_addr = ""
	CASH_email = ""
	CASH_info = "Fax request on your Agency’s Letterhead 1-860-424-4886, allow 3 to 5 days for a response"

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'DELAWARE'
If state_list = "DELAWARE" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "EMAIL"
	general_phone = "302-571-4900"
	general_fax = ""
	general_addr = "Delaware Division of Social Services PO Box 906, New Castle, DE 19720"
	general_email = "DHSS_DSS_Outofstate@state.de.us"
	general_info = ""

	web_addr = "http://www.dhss.delaware.gov/dhss/dss"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'District of Columbia'
If state_list = "DISTRICT OF COLUMBIA" then
	general_header = "Out of State inquiries:"
	general_person = "Vivessia Avent, Program Analyst"
	general_pref_contact = "EMAIL"
	general_phone = "202-535-1145"
	general_fax = "202-671-4409"
	general_addr = "District of Columbia Department of Human Services Office of Program Review, Monitoring and Investigation
					64 New York Avenue, N.E. – 6th Floor Washington, D.C. 20002"
	general_email = "Vivesia.Avent@dc.gov"
	general_info = ""

	web_addr = "www.dhs.dc.gov"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Florida
If state_list = "FLORIDA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "EMAIL"
	general_phone = "866-762-2237"
	general_fax = ""
	general_addr = "Dept. of Children & Families
					1317 Winewood Blvd., Bldg. 3, Room 435
					Tallahassee, FL 32399-0700"
	general_email = "SNR.D11.SFL.CallCenter@myflfamilies.com"
	general_info = ""

	web_addr= "www.dcf.state.fl.us/ess/"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Georgia'
If state_list = "GEORGIA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "EMAIL"
	general_phone = "1-877-423-4746"
	general_fax = "1-888-740-9355"
	general_addr = "DFCS Customer Service Operations, 2 Peachtree Street, Suite 8-268 Atlanta, Georgia 30303"
	general_email = "ga.paris@dhs.ga.gov"
	general_info = ""

    web_addr = ""

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""
'Guam'
If state_list = "Guam" then
	general_header = "Out of State inquiries:"
	general_person = "Maria Cindy F. Malanum, Social Services Supervisor I"
	general_pref_contact = ""
	general_phone = "671-735-7288/7237"
	general_fax = "671-734-7092"
	general_addr = "Bureau of Economic Security
					Division of Public Welfare
					Department of Public Health and Social Services
					123 Chalan Kareta
					Mangilao, Guam 96913"
	general_email = "mariacindy.malanum@dphss.guam.gov"
	general_info = ""

	web_addr = "http://dphss.guam.gov"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'Hawaii'
If state_list = "HAWAII" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "808-586-5735"
	general_fax = ""
	general_addr = "Department of Human Services
					State Office Administrative Assistant (FSP & TANF)
					Benefit, Employment & Support Services Division
					820 Mililani Street, Suite 606
					Honolulu, Hi 96813"
	general_email = ""
	general_info = ""

    web_addr = "http://hawaii.gov/dhs"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = "808-586-5720"
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = "808-586-5732"
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'IDAHO'
If state_list = "IDAHO" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "208-334-5815"
	general_fax = "208-334-5817 or 1-866-434-8278"
	general_addr = "Idaho Department of Health & Welfare
					Division of Welfare, 2nd Floor
					P.O. Box 8372, Boise, ID 83720-0036"
	general_email = "mybenefits@dhw.idaho.gov"
	general_info = ""

	web_addr = "http://www.healthandwelfare.idaho.gov/"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = "welfraud@dhw.idaho.gov."
	claims_info = "To verify benefit status in Idaho for clients already
				   active on benefits in your state (dual participation
				   alerts & investigations, PARIS matches, etc.), please
				   contact the Fraud Unit via e-mail at:"
'ILLINOIS'
 If state_list = "ILLINOIS" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "217-524-4174"
	general_fax = ""
	general_addr = "Illinois Department of Human Services
					Bureau of Customer and Support Services
					600 East Ash St. Bld. 500, Springfield, IL 62703"
	general_email = "DHS.WEBBITS@Illinois.gov" 'please be sure to encrypt ask Ilse if we can build this in'
	general_info = "Allow 3-5 Business days for a response."

	web_addr = "http://www.dhs.state.il.us/page.aspx?module=16&type=2 "

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'INDIANA'
If state_list = "INDIANA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = ""
	general_fax = ""
	general_addr = "Indiana Family and Social Services Administration
					P.O. Box 1810
					Marion, IN 46952"
	general_email = "INoutofstate.inquiries@fssa.IN.gov."
	general_info = "Indiana no longer accepts faxed/telephone inquiries. Requests must include names, dates of birth and last
					4 digits of SSN for each individual for whom verifications are being sought. Also include what programs you need
					verified."

	web_addr = "www.in.gov/fssa/"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'IOWA'
If state_list = "IOWA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "1-877-855-0021"
	general_fax = "515-564-4095"
	general_addr = "Integrated Claims Recovery Unit
					PO Box 36570, Des Moines, IA 50315"
	general_email = "ICRU@dhs.state.ia.us"
	general_info = "To have an active Iowa case closed:
					Clients can call Customer Service Call Center at 1-877-
					347-5678 OR Email: IMCustomerSC@dhs.state.ia.us"

	web_addr = "www.dhs.state.ia.us"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'KANSAS'
If state_list = "KANSAS" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = ""
	general_fax = ""
	general_addr = "Kansas Department for Children and Families
					Economic and Employment Services
		  			555 S Kansas Avenue, 4th Floor, Topeka, KS 66603"
	general_email = "DCF.EBTMAIL@ks.gov"
	general_info = "Kansas no longer accepts faxed inquires. All SNAP and TANF out-of-state inquiries must be emailed"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = "800-792-4884"
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = "For all Medical inquiries, contact Kansas Department of Health and Environment"

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	web_addr = "www.dcf.ks.gov"
'KENTUCKY'
 If state_list = "KENTUCKY" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "502-564-3440"
	general_fax = ""
	general_addr = ""
	general_email = "Outofstateinquiries@ky.gov"
	general_info = ""

	web_addr = "http://cfc.ky.gov"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'MAINE'
If state_list = "MAINE" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "207-624-4130"
	general_fax = "207-287-3455"
	general_addr = "ACES Help Desk – Eligibility Specialists
					Department of Health and Human Services Office for Family Independence
					19 Union Street, SHS#11, Augusta, ME 04333 "
 	general_email = "DESK.ACESHELP@Maine.gov"
	general_info = "A signed release is required to obtain the verification Out-of-State Inquiries"

	web_addr = "http://www.maine.gov/dhhs/ofi/index.html"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'MARAND'
If state_list = "MARYLAND" then
	general_header = "Out of State inquiries:"
	general_person = "Tanisha Williams"
	general_pref_contact = ""
	general_phone = "410-767-7928"
	general_fax = ""
	general_addr = "Maryland Department of Human Resources
	              	311 W. Saratoga St. Baltimore, MD 21201"
	general_email = "dhr.outofstateinquiry@maryland.gov or tanisha.williams@maryland.gov"
	general_info = ""

	web_addr = ""

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

'MASSACHUSET
If state_list = "MASSACHUSETTS" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = ""
		general_fax = "617-889-7847"
		general_addr = "MA Department of Transitional Assistance
				        Data Matching Unit
				  		600 Washington Street, 5th Floor, Boston, MA 02111"
		general_email = ""
		general_info = "MAIL or FAX REQUEST ON AGENCY LETTERHEAD:"

				     web_addr = "www.state.ma.us/DTA"


		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'MICHIG
If state_list = "MICHIGAN" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "1-517-373-3908"
		general_fax = "1-517-335-6054"
		general_addr = "Dept of Human Services
						PO Box 30037, 235 S. Grand Ave, Lansing, MI 48909"
		general_email = "DHS-ICU-Customer-Service@michigan.gov"
		general_info = "Require client’s new address before they will close case in MI"

	    web_addr = "www.michigan.gov/dhs"

		Client_Services "1-855-ASK-MICH" 'add to variable'

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claimsnfo = ""


'MINNESOT
If state_list "MINNESOTA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = ""
		general_fax = ""
		general_addr = "Minnesota Department of Human Services
						Economic Assistance and Employment Services Division
						PO Box 64951 St. Paul, MN 55164-0951"
		general_email = ""
		general_info = "Out of State Inquiries for case status of SNAP and TANF	programs and the number of TANF months expended are
						provided through an automated web service.
						At the web site home page, complete the required fields	for self- registration. In the code field, enter the 			word guest. After accepting the Oath, click the Next 			button which brings up the client information page. When 			all the client information has been entered, click the 			submit
						button which generates the request. A secure email	response will be sent to the requestor with the results 			of
						the benefit verification. No other means for requesting	this information is offered."

		web_addr = "https://mn.gov/snap-tanf-benefit-verification/"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'MISSISSIPPI'
If state_list = "MISSISSIPPI" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "EMAIL"
		general_phone = "1-800-948-3050"
		general_fax = "601-359-4550 "
		general_addr = "Department of Human Services
						Division of Field Operations, P.O. Box 352 Jackson, MS 39205"
		general_email = "ea.CustomerService@mdhs.ms.gov"
		general_info = "Please include email address, agency telephone number and mailing address.
		Also include case member’s name, date of birth, SSN, current mailing address and brief description of verification
		that is needed."

		web_addr = "www.mdhs.state.ms.us"


		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'MISSOURI'
If state_list = "MISSOURI" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "1-800-392-1261"
		general_fax = ""
		general_addr = "Correspondence and Information Unit	Family Support Division
						Department of Social Services
						P.O. Box 2320, Jefferson City, MO 65102-2320"
		general_email = "Cole.CoXIX@dss.mo.gov"
		general_info = "Out-of-State line: The Family Support Division
					    Information Center at 855-FSD-INFO or 855-373-4636,	Option 3."

		web_addr = "http://www.dss.mo.gov"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		phone = ""
		fax = ""
		addr = ""
		email = ""
		info = ""
'MONTANA'
If state_list = "MONTANA" then
		general_header = "Out of State inquiries:"
		general_person = "Mollye Gauer"
		general_pref_contact = ""
		general_phone = "406-444-9401"
		general_fax = ""
		general_addr = "Department of Public Health & Human Services
		                Human & Community Services Division
						PO Box 202925, Helena, MT 59620-2925"
		general_email = "mgauer@mt.gov"
		general_info = ""

		web_addr = "http://www.dphhs.mt.gov/"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
        SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = "Maria Jimenez Gonzalez, TANF Program Coordinator"
		CASH_phone = "406-444-0676"
		CASH_fax = "406-444-0617"
		CASH_addr = "111 N. Jackson, Helena, MT 59601"
		CASH_email = "mjimenezgonzalez@mt.gov"
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""


'NEBRASKA'
If state_list = "NEBRASKA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "1-800- 383-4278"
		general_fax = ""
		general_addr = ""
		general_email = ""
		general_info = "Customer Service Center Economic Assistance Customer Service Center"

		web_addr = "www.accessnebraska.ne.gov"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = "DHHS.EconomicAssistancePolicyQuestions@nebraska.gov"
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = "DHHS.MedicaidPolicyQuestions@nebraska.gov"
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'NEVADA'
If state_list = "NEVADA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = ""
		general_fax = ""
		general_addr = "Dept of Health & Human Services
						Division of Welfare and Supportive Services
						1470 College Parkway
						Carson City, NV 89706"
		general_email = "WELFOOSINQUIRIES@DWSS.NV.GOV"
		general_info = "Nevada requires requests to be sent on agency letterhead. Please include clients; Name, Date of Birth; the SSN or the
						last 4 of the SSN; and a listing of all household members who are applying for assistance in your state. Also be sure to
						include the return information (your name, phone, and fax #’s). Please allow 3 – 5 business days for a response."

		web_addr = "http://dwss.nv.gov"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'NEW HAMPSHIRE'
If state_list = "NEW HAMPSHIRE" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "Katy Halvorsen"
		general_phone = "603-444-0348"
		general_fax = "603-271-423"
		general_addr = "Client Services/SOP_Brown
						Department of Health & Human Services
						Division of Client Services
						129 Pleasant St, Concord, NH 03301"
		general_email = "outofstateinquiries@dhhs.state.nh.us"
		general_info = ""

		web_addr = "http://www.dhhs.nh.gov/"

		Client_Services = "603-271-9700"
		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'NEW JERSEY'
If state_list = "NEW JERSEY" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "EMAIL"
		general_phone = "609-588-2283"
		general_fax = ""
		general_addr = "Department of Human Services Division of Family Development
						Program Assessment and Integrity Unit
						P.O. Box 716, Trenton, NJ 08625-0716"
		general_email = "dfd.paiu@dhs.state.nj.us."
		general_info = "NJ no longer accepts faxed requests. It is also required that the individual’s name, DOB and last 4 numbers of their SSN be included."

		web_addr = "www.state.nj.us/humanservices/dfd/index.html"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = "State Eligibility Policy Liaison for Medicaid: Stephen Myers"
		HC_pref_contact = ""
		HC_phone = "609-588-7758"
		HC_fax = ""
		HC_addr = ""
		HC_email = "Stephen.myers@dhs.state.nj.us."
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

'NEW MEXICO'
If state_list = "NEW MEXICO" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "EMAIL"
		general_phone = "505-827-7250"
		general_fax = "505-827-7203"
		general_addr = "New Mexico Human Services Department
						Income Support Division
						P.O.B 2348 2009 S Pacheco Street
						Santa Fe, NM 87504"
		general_email = "nmhsdinquiry@state.nm.us"
		general_info = ""

	    web_addr = "www.hsd.state.nm.us"

		SNAP_header = "SNAP Participation:"
		SNAP_person= ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""
		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = "1-888-997-2583"
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = "1-800-228-4802"
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = "1-800-283-4465"

'NEW YORK'
If state_list = "NEW YORK" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "Fax"
		general_phone = ""
		general_fax = "518-474-8090"
		general_addr = ""
		general_email = ""
		general_info = "Be sure to include correct return FAX number. No Email/Phone. Include customer’s full name,
						DOB and last 4 digits of SS# - Please include customer’s new address. If names of all household members are not
						included, response may not be completely accurate. If you have not received a response in 5 business days please direct
						an email to:wendy.buell@otda.ny.gov or call 518-486-3460."

		web_addr = "http://www.dfa.state.ny.us"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = "Stephen Bach, Director"
		claims_pref_contact = ""
		claims_phone = "518-402-0117"
		claims_fax = ""
		claims_addr = "Program Integrity, Audit & Quality Improvement NYS OTDA, Riverview Center
						4th Floor, 40 N Pearl St, Albany, NY 12243"
		claims_email = "Stephen.Bach@otda.ny.gov"
		claims_info = ""

		client_services = ""

'NORTH CAROLINA'
If state_list = "NORTH CAROLINA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "1-866-719-0141"
		general_fax = "252-789-5395"
		general_addr = "North Carolina Department of Health &
						Human Services Division of Social Services
						P. O. Box 190 Everetts, NC 27829"
		general_email = "ebt.csc.leads@dhhs.nc.gov"
		general_info = "When forwarding request to the DHHS Call Center, include your agency’s name, mailing address, telephone number, and fax number if not already
						indicated in your request. Due to NC security requirements with SSNs, if your original e-mail contained
						a full SSN, it will be edited so that only the last four digits	remain."

	 	web_addr = "www.ncdhhs.gov/dss"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'NORTH DAKOTA'
If state_list = "NORTH DAKOTA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "701-328-2332 or 701-328-3513"
		general_fax = ""
		general_addr = ""
		general_email = "dhseap@nd.gov"
		general_info = "secure E-mail request with client’s name, full SSN, and DOB on your Agency’s Letterhead. Please allow 1 to 3 business days for a response."

		web_addr = " www.nd.gov/dhs/services"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""


'OKLAHOMA'
If state_list = "OKLAHOMA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "405-521-3444 or 1-866-411-1877"
		general_fax = "405-521-4158"
		general_addr = ""
		general_email = "SNAP@okdhs.org"
		general_info = ""

		web_addr = "www.okdhs.org"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = "eligibility@OKHCA.org"
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'OHIO'
If state_list = "OHIO" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "614-466-4815 option 2, 1"
		general_fax = "614-466-1767"
		general_addr = "Office of Family Assistance
		                Ohio Department of Job & Family Services
	                	P.O. Box 183204, Columbus, Ohio 43218-3204"
		general_email = ""
		general_info = "Our Staff cannot provide benefit information. Because	Ohio is a state supervised, county administered state, any
						eligibility information must be provided by the county	agency. Please FAX request w/clients name, SSN,
						DOB, & the County they lived in Ohio on your Agency letterhead. To receive direct contact information for each
						county agency, see attached."

		web_addr = "www.jfs.ohio.gov/"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""



'OREGON'
If state_list = "OREGON" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "503-945-5600"
		general_fax = "503-373-7032 or 503-581-6198"
		general_addr = "Oregon Department of Human Services
						500 Summer St. NE, E-48
						Salem, OR 97301-1066"
		general_email = ""
		general_info = "TO VERIFY RECEIPT OF SNAP, MEDICAL, AND/OR TANF BENEFITS:
						FAX request w/client's name, SSN & DOB on your Agency's Letterhead"

		web_addr = "http://www.oregon.gov/DHS/assistance/index.shtml"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'PENNSYLVANIA'
If state_list = "PENNSYLVANIA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "717-787-3119"
		general_fax = "717-705-0040"
		general_addr = "Division of Hotline and Correspondence
						Pennsylvania Department of Human Services
						P.O. Box 2675 Harrisburg, PA 17105-2675"
		general_email = "ra-dpwoimnet@pa.gov"
		general_info = "Name of the household members who are applying for benefits, Each members date of birth and last 4 of their social
						security number, The client’s new address. Allow 3 to 5 business days for a response."

		web_addr = "www.DHS.state.pa.us"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'RHODE ISLAND'
If state_list = "RHODE ISLAND" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = ""
		general_fax = "401-415-8563 ATTN Rhode Island IEVS Unit"
		general_addr = "Rhode Island Department of Human Services
						57 Howard Ave  Louis Pastore Building, Cranston, RI 02920"
		general_email = ""
		general_info = ""

		web_addr = "http://www.dhs.ri.gov/"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = "401-462-2175"
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'SOUTH CAROLINA'
If state_list = "SOUTH CAROLINA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "Email"
		general_phone = ""
		general_fax = "803-898-122" 'ATTN: Client Services'
		general_addr = "South Carolina Department of Social Services
						Office of Economic Services
						P.O. Box 1520 Columbia, SC 29202-1520"
		general_email = "ClientServices@dss.sc.gov"
		general_info = "Subject line should read: ‘Out of State Inquiry from ‘name of state’. We will be unable to process your request without the
						following information:Individual’s name, SS#, DOB and current address"

		web_addr = ""

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = "888-549-0820"
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = "Mrs. Angela Clark"
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = "angela.clark@dss.sc.gov."
		claims_info = ""

		client_services = ""

'SOUTH DAKOTA'
If state_list = "SOUTH DAKOTA" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "877-999-5612 or 605-773-4678 "
		general_fax = ""
		general_addr = "Department of Social Services
						700 Governors Drive	Pierre, South Dakota 57501-2291"
		general_email = "SNAP@state.sd.us"
		general_info = ""

	 	web_addr = "http://dss.sd.gov/economicassistance/snap/"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'TENNESSEE'
If state_list = "TENNESSEE" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = "Fax"
		general_phone = ""
		general_fax = "615-687-5535"
		general_addr = ""
		general_email = ""
		general_info = "Your Name, Agency Name and Address, phone number and fax number. Names of household members applying for benefits in
						your State. Complete Social Security Numbers. A Current Address so that closure notices can be sent to
						the client"

		  web_addr = "www.tn.gov/humanservices"
				SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = "Supplemental Nutrition Assistance Program (SNAP) Policy
					Department of Human Services Citizen Plaza Bldg, 8th Floor
					400 Deaderick Street Nashville, TN 37243-7200"
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'TEXAS'
If state_list = "TEXAS" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "877-541-7905"
		general_fax = "877-447-2839"
		general_addr = ""
		general_email = ""
		general_info = ""

	    web_addr = "YourTexasBenefits.com"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'UTAH'
If state_list = "UTAH" then
		general_header = "Out of State inquiries:"
		general_person = ""
		general_pref_contact = ""
		general_phone = "866-435-7414 Option 5"
		general_fax = ""
		general_addr = "Department of Workforce Services
						Eligibility Services Division
						P.O. Box 143245 Salt Lake City, UT 84114-3245"
		general_email = ""
		general_info = ""

		web_addr = "www.jobs.utah.gov"

		SNAP_header = "SNAP Participation:"
		SNAP_person = ""
		SNAP_pref_contact = ""
		SNAP_phone = ""
		SNAP_fax = ""
		SNAP_addr = ""
		SNAP_email = ""
		SNAP_info = ""

		CASH_header = "TANF Participation:"
		CASH_person = ""
		CASH_pref_contact = ""
		CASH_phone = ""
		CASH_fax = ""
		CASH_addr = ""
		CASH_email = ""
		CASH_info = ""

		HC_header = "Medical Assistance Participation:"
		HC_person = ""
		HC_pref_contact = ""
		HC_phone = ""
		HC_fax = ""
		HC_addr = ""
		HC_email = ""
		HC_info = ""

		claims_header = ""
		claims_person = ""
		claims_pref_contact = ""
		claims_phone = ""
		claims_fax = ""
		claims_addr = ""
		claims_email = ""
		claims_info = ""

		client_services = ""

'VERMONT'
If state_list = "VERMONT" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "Phone"
	general_phone = "800- 479-6151 or 802-828-6896"
	general_fax = ""
	general_addr = "Call Center Economic Services Benefits Service Center
					103 South Main St Waterbury, VT 05671-1201"
	general_email = ""
	general_info = ""

	web_addr = " www.mybenefits.vt.gov " (this site can be used, but not preferred).

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

'VIRGINIA'
If state_list = "VIRGINIA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = ""
	general_fax = "804-819-7185 or 804-819-7186"
	general_addr = "Virginia Department of Social Services Division of Benefit Programs – 9th Floor
					801 East Main Street Richmond, VA 23219-2901"
	general_email = "vaoutofstateverifications@dss.virginia.gov"
	general_info = "Include the following:	Client’s name, DOB, SSN (if last four only, you must provide entire DOB), and current address
					All household members who are applying for assistance in your state"

	 web_addr = "http://www.dss.virginia.gov/benefit/foodstamp.cgi"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

'VIRGIN ISLANDS'
If state_list = "VIRGIN ISLANDS" then
	general_header = "Out of State inquiries:"
	general_person = "Emmanueline Archer"
	general_pref_contact = ""
	general_phone = "340-774-0930 ext. 4309"
	general_fax = "340-777-5449"
	general_addr = "Department of Human Services Division of Family Assistance
					1303 Hospital Ground, STE 1 St. Thomas, VI 00801-6722"
	general_email = "emmanueline.archer@dhs.vi.gov"
	general_info = ""

	web_addr = ""
	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

'WASHINGTON'
If state_list = "WASHINGTON" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = "Fax"
	general_phone = "855-927-2747"
	general_fax = "888-212-2319"
	general_addr = "Department of Social and Health Services
					Division of Program Integrity Attn: PARIS Unit
					PO Box 45410 Olympia, WA 98504-5410"
	general_email = ""
	general_info = "Please provide the following information:
					Name of the worker requesting the information Direct contact information for the worker requesting
					information	Name, SSN, DOB of all household members who applied in your state
					Current mailing address for the client(s) Date/programs applied for in the other state"

	web_addr = ""

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

'WEST VIRGINA'
If state_list = "WEST VIRGINA" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "304-356-4619 (Request to speak to the Worker of the Day)"
	general_fax = ""
	general_addr = "Department of Health & Human Resources
					Division of Family Assistance
					350 Capitol St., Room B-18	Charleston, WV 25301-3705"
	general_email = "DHHRbcfbenefitver@wv.gov"
	general_info = ""

	web_addr = "www.dhhr.wv.gov"
	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

'WISCONSIN'
If state_list = "" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = "608-261-6378--Option 3"
	general_fax = "608-327-6125"
	general_addr = "SNAP and HealthCare (MA): WI Department of Health & Family Services
						1 W Wilson St, Madison, WI 53703"
	general_email = "DHSCARESCallCenter@wisconsin.gov"
	general_info = ""

	web_addr = "http://dhs.wisconsin.gov/"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = "608-422-7900"
	CASH_fax = "608-327-6125"
	CASH_addr = ""
	CASH_email = "DCFW2TANFVerify@wisconsin.gov."
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = "608-422-7100"
	claims_fax = ""
	claims_addr = ""
	claims_email = "dcfoig@wisconsin.gov"
	claims_info = ""

	client_services = "1-800-362-3002 or go to dhs.wi.gov/em/customerhelp."

'WYOMING'
If state_list = "WYOMING" then
	general_header = "Out of State inquiries:"
	general_person = "Annette Jones, Administrative Specialist"
	general_pref_contact = ""
	general_phone = "307-777-5846"
	general_fax = "307-777-6276"
	general_addr = "Department of Family Services
					2300 Capitol Ave., Hathaway Bldg, Third Fl.	Cheyenne, WY 82002-0490"
	general_email = "annette.jones@wyo.gov"
	general_info = ""

	web_addr = "http://dfsweb.state.wy.us/foodstampinfo.html"

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = "wesapplications@wyo.gov"
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""

	client_services = ""

End if

 'template'
If state_list = "" then
	general_header = "Out of State inquiries:"
	general_person = ""
	general_pref_contact = ""
	general_phone = ""
	general_fax = ""
	general_addr = ""
	general_email = ""
	general_info = ""

        web_addr = ""

	SNAP_header = "SNAP Participation:"
	SNAP_person = ""
	SNAP_pref_contact = ""
	SNAP_phone = ""
	SNAP_fax = ""
	SNAP_addr = ""
	SNAP_email = ""
	SNAP_info = ""

	CASH_header = "TANF Participation:"
	CASH_person = ""
	CASH_pref_contact = ""
	CASH_phone = ""
	CASH_fax = ""
	CASH_addr = ""
	CASH_email = ""
	CASH_info = ""

	HC_header = "Medical Assistance Participation:"
	HC_person = ""
	HC_pref_contact = ""
	HC_phone = ""
	HC_fax = ""
	HC_addr = ""
	HC_email = ""
	HC_info = ""

	claims_header = ""
	claims_person = ""
	claims_pref_contact = ""
	claims_phone = ""
	claims_fax = ""
	claims_addr = ""
	claims_email = ""
	claims_info = ""
