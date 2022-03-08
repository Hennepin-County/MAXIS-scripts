'GATHERING STATS===========================================================================================
name_of_script = "NOTICES - OUT OF STATE INQUIRY.vbs"
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'------------------start of state list'
function fill_in_the_states()
    IF state_droplist = "Alabama" THEN
    	abbr_state = "AL"
		agency_name = "Betty S. White, Program Manager Alabama Dept. of Human Resources"
		agency_address = "50 Ripley St, S. Gordon Persons Bldg. Montgomery, AL 36130-4000"
		agency_phone = "334-242-5010"
		agency_fax = "334-242-0513"
		agency_email = "DHR_PA_Helpdesk@dhr.alabama.gov"
		agency_website = "www.dhr.alabama.gov"
    	IF OTHER_STATE_FS_CHECKBOX = CHECKED THEN
    		other_state_fs = TRUE
    		agency_email = "fs@dhr.alabama.gov"
    	END IF
    	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
    		other_state_cash = TRUE
    		agency_email = "DHR_PA_Helpdesk@dhr.alabama.gov"
    	END IF
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_phone = "334-242-5010"
    	END IF
    	IF PARIS_CHECKBOX = CHECKED THEN
    		other_state_paris = TRUE
    		agency_email = "PIDParis@medicaid.alabama.gov"
    	END IF
    END IF
    IF	state_droplist = "Alaska" THEN
    	abbr_state = "AK"
    	agency_name = "Alaska Dept of Health & Social Services Division of Public Assistance"
    	agency_address = "PO Box 110640, Juneau, AK 998110640"
    	agency_phone = "907-465-3347"
    	agency_email = "verifications@alaska.gov"
    	agency_website = "http://www.hss.state.ak.us/dpa/"
    END IF
    IF	state_droplist = "Arizona" THEN
    	abbr_state = "AZ"
    	agency_name = "Dept. of Economic Security Communication Center"
    	agency_phone = "602-771-2047"
    	agency_fax = "602-353-5746"
    	agency_email = "Azstateinquiries@Azdes.gov"
    	agency_website = "www.des.az.gov"
    END IF
    IF	state_droplist = "Arkansas"	THEN
    	abbr_state = "AR"
    	agency_name = "Ruthie Broughton DHS Program Manager, Customer Assistance Unit"
    	agency_address = "PO Box 1437, Slot S-341 Little Rock , Arkansas 72203-1437"
    	agency_phone = "501-683-4443"
    	agency_fax = "501-682-8978"
    	agency_email = "Ruthie.broughton@dhs.arkansas.gov"
    	agency_website = "www.arkansas.gov/dhhs"
    END IF
    IF	state_droplist = "California" THEN
    	IF state_county = "" THEN err_msg = err_msg & vbNewLine & "Please select the county of the out of state inquiry."
    	abbr_state = "CA"
    	agency_name = "California Department of Social Services"
    	agency_address = "744 P Street, MS 8-4-23 Sacramento, CA 95814-6400"
    	agency_phone = "844-626-5900"
    	agency_ext = "press 1, then 7"
    	agency_fax = "916-651-8866"
    	agency_email = "cdss.osi@dss.ca.gov"
    	agency_website = "http://www.cdss.ca.gov/cdssweb/entres/pdf/CountyCentralIndexListing.pdf"
    END IF
    IF	state_droplist = "Colorado"	THEN
    	abbr_state = "CO"
    	agency_name = "Colorado Department of Human Services"
    	agency_address = "1575 Sherman St. 3rd Fl, Denver, CO 80203"
    	agency_phone = "1-800-536-5298"
    	agency_ext = "3"
    	agency_email = "cdhsoutofstateinquiries@state.co.us"
    	agency_website = "www.cdhs.state.co.us"
		IF PARIS_CHECKBOX = CHECKED THEN
			agency_phone = "1-800-359-1991"
			other_state_paris = TRUE
		END IF
    END IF
    IF	state_droplist = "Connecticut" THEN
    	abbr_state = "CT"
		agency_name = "Department of Social Service"
		agency_address = "55 Farmington Ave. Hartford, CT 06105-3725"
		agency_phone = "Default - 860-424-5030"
		agency_email = "TPI.EU@ct.gov"
    	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
    		other_state_cash = TRUE
    		If from_paris_match = false Then
				agency_phone = agency_phone & "~Cash - 860-424-5540 (RECOMMENDED)"
			Else
				agency_phone = agency_phone & "~Cash - 860-424-5540"
			End If
    		agency_fax = "860-424-4886"
    	END IF
    	IF PARIS_CHECKBOX = CHECKED THEN
    		other_state_paris = TRUE
    		agency_email = "Contact.Paris@ct.gov"
    		agency_fax = "860-424-5333"
    	END IF
    END IF
    IF	state_droplist = "Delaware"	THEN
    	abbr_state = "DE"
		agency_name = "Delaware Division of Social Service"
		agency_address = "PO Box 906, New Castle, DE 19720"
		agency_phone = "302-571-4900"
		agency_email = "DHSS_DSS_Outofstate@delaware.gov"
		agency_website = "http://www.dhss.delaware.gov/dhss/dss"
    	IF PARIS_CHECKBOX = CHECKED THEN
    		other_state_paris = TRUE
    		agency_email = "DE_PARIS-ARMS@delaware.gov"
    	END IF
    END IF
    IF	state_droplist = "District of Columbia"	THEN
    	abbr_state = "DC"
    	agency_name = "Vivessia Avent, Program Analyst Eligibility Review and Investigations District of Columbia Department of Human Services, Office of Program Review, Monitoring and Investigation"
    	agency_address = "645 H Street, N.E. – 3rd Floor Washington, D.C. 20002"
    	agency_phone = "202-535-1145"
    	agency_website = "www.dhs.dc.gov"
    	agency_fax = "202-645-4197"
    	agency_email = "SNR.D11.SFL.CallCenter@myflfamilies.com"
    	agency_website = "www.dhs.dc.gov"
    END IF
    IF	state_droplist = "Florida"	THEN
    	abbr_state = "FL"
    	agency_name = "Dept. of Children & Families"
    	agency_address = "1317 Winewood Blvd., Bldg. 3, Room 435 Tallahassee, FL 32399-0700"
    	agency_phone = "866-762-2237"
    	agency_website = "www.dcf.state.fl.us/ess/"
    	agency_email = "SNR.D11.SFL.CallCenter@myflfamilies.com"
    END IF
    IF	state_droplist = "Georgia"	THEN
    	abbr_state = "GA"
    	agency_name = "DFCS Customer Service Operations"
    	agency_address = "2 Peachtree Street,  Suite 8-268 Atlanta, Georgia  30303"
    	agency_phone = "1-877-423-4746"
    	agency_fax = "1-888-740-9355"
    	agency_email = "ga.paris@dhs.ga.gov"
		IF PARIS_CHECKBOX = CHECKED THEN
			other_state_paris = TRUE
			agency_email = "ga.paris@dhs.ga.gov"
		END IF
    END IF
    IF	state_droplist = "Hawaii"	THEN
    	abbr_state = "HI"
    	IF OTHER_STATE_FS_CHECKBOX = CHECKED THEN
    		other_state_fs = TRUE
    		agency_phone = "808-586-5720"
    	END IF
    	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
    		other_state_cash = TRUE
    		agency_phone = "808-586-5732"
    	END IF
    	IF OTHER_STATE_CCA_CHECKBOX = CHECKED THEN other_state_cca = TRUE
    	IF COMMOD_CHECKBOX = CHECKED THEN other_state_commod = TRUE
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN other_state_hc = TRUE
    	IF PARIS_CHECKBOX = CHECKED THEN other_state_paris = TRUE
    	agency_name = "Department of Human Services State Office Administrative Assistant Benefit, Employment & Support Services Division"
    	agency_address = "1010 Richards Street, Suite 512 Honolulu, Hi  96813"
    	agency_website = "http://hawaii.gov/dhs"
    	agency_phone = "808-586-5732" 'this is missing'
    END IF
    IF	state_droplist = "Idaho" THEN
    	abbr_state = "ID"
    	agency_name = "Idaho Department of Health & Welfare Division of Welfare, 2nd Floor"
    	agency_address = "P.O. Box 83720 Boise, ID  83720-0036"
    	agency_phone = "208-334-5815"
    	agency_email = "SRIUWFIU@dhw.idaho.gov"
    	agency_website = "http://www.healthandwelfare.idaho.gov/"
    END IF
    IF	state_droplist = "Illinois"	THEN
    	abbr_state = "IL"
    	agency_name = "Illinois Department of Human Services Bureau of Customer and Support Services"
    	agency_address = "600 East Ash St., Bld. 500 Springfield, IL 62703"
    	agency_email = "DHS.OUTOFSTATE@ILLINOIS.GOV"
    	agency_phone = "217-524-4174"
    	agency_website = "www.dhs.state.il.us/contactus"
    END IF
    IF	state_droplist = "Indiana"	THEN
    	abbr_state = "IN"
    	agency_name = "Indiana Family and Social Services Administration"
    	agency_address = "P.O. Box 1810 Marion, IN 46952"
    	agency_website = "www.in.gov/fssa/"
    	agency_email = "INoutofstate.inquiries@fssa.IN.gov."
		IF PARIS_CHECKBOX = CHECKED THEN
			other_state_paris = TRUE
			agency_email = "Paris Unit Inquiries: parisinquiries@fssa.IN.gov"
		END IF
    END IF
    IF	state_droplist = "Iowa"	THEN
    	abbr_state = "IA"
    	agency_name = "Integrated Claims Recovery Unit"
    	agency_phone ="1-877-855-0021"
    	agency_fax = "515-564-4095"
    	agency_fax = "515-564-4095"
    	agency_email = "ICRU@dhs.state.ia.us"
    	agency_website = "www.dhs.state.ia.us"
		IF PARIS_CHECKBOX = CHECKED THEN
			other_state_paris = TRUE
			agency_email = "Paris Unit Inquiries: parisinquiries@fssa.IN.gov"
		END IF
    END IF
    IF	state_droplist = "Kansas"	THEN
    	abbr_state = "KS"
    	agency_name = "Kansas Department for Children and Families Economic and Employment Services"
    	agency_address = "555 S Kansas Avenue, 4th Floor Topeka, KS 66603"
    	agency_email = "DCF.EBTMAIL@ks.gov"
    	agency_website = "www.dcf.ks.gov"
		IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
			other_state_hc = TRUE
			agency_phone = "800-792-4884"
		END IF
		IF PARIS_CHECKBOX = CHECKED THEN
			other_state_paris = TRUE
			agency_email = "DCF.PARIS@KS.Gov"
			agency_phone = "785.296.3874"
			agency_fax = "785.296.6960"
		END IF
    END IF
    IF	state_droplist = "Kentucky"	THEN
    	abbr_state = "KY"
    	agency_name = "Kentucky Cabinet for Health and Family Services"
    	agency_phone = "855-306-8959"
    	agency_email = "Outofstateinquiries@ky.gov"
    	agency_website = "https://chfs.ky.gov/Pages/index.aspx"
    END IF
    IF	state_droplist = "Louisiana"	THEN
    	abbr_state = "LA"
    	agency_name = "Louisiana Department of Children and Family Services Bureau of Communications & Governmental Affairs C/O Cara (Yvette) Shields, Program Specialist"
    	agency_address = "627 North 4th Street, 8th Floor Baton Rouge, Louisiana 70802"
    	agency_phone = "225-342-2342"
    	agency_fax = "225-342-9833"
    	agency_email = "cara.shields@la.gov"
    	agency_website = "www.dcfs.louisiana.gov"
		IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
			other_state_hc = TRUE
			agency_phone = "225-342-9730"
			agency_fax = "225-389-8100"
			agency_email = "OOS@la.gov"
		END IF
    END IF
    IF	state_droplist = "Maine" THEN
    	abbr_state = "ME"
    	MsgBox "The State of Maine requires a signed release is to obtain the verification Out-of-State Inquiries, please attach to the email."
    	agency_name = "ACES Help Desk Eligibility Specialists Department of Health and Human Services Office for Family Independence"
    	agency_address = "109 Capital Street, SHS#11 Augusta, ME 04333"
    	agency_phone = "207-624-4130"
    	agency_fax = "207-287-3455"
    	agency_email = "DESK.ACESHELP@Maine.gov"
    	agency_website = "http://www.maine.gov/dhhs/ofi/index.html"
    END IF
    IF	state_droplist = "Maryland"	THEN
    	abbr_state = "MD"
    	agency_name = "Maryland Department of Human Resources"
    	agency_address = "311 W. Saratoga St. Baltimore, MD 21201"
    	agency_phone = "410-767-7928"
    	agency_email = "dhr.outofstateinquiry@maryland.gov"
		IF PARIS_CHECKBOX = CHECKED THEN
			other_state_paris = TRUE
			agency_email = "paris.inquiries@maryland.gov "
			agency_phone = "410-238-1249"
		END IF
    END IF
    IF	state_droplist = "Massachusetts"	THEN
    	abbr_state = "MA"
    	agency_name = "MA Department of Transitional Assistance Data Matching Unit"
    	agency_address = "600 Washington Street, 5th Floor Boston, MA 02111"
    	agency_fax = "617-889-7847"
    	agency_website = "www.state.ma.us/DTA"
    END IF
    IF	state_droplist = "Michigan"	THEN
    	abbr_state = "MI"
    	agency_name = "Dept of Health and Human Services"
    	agency_address = "PO Box 30037 235 S. Grand Ave. Lansing, MI 48909"
    	agency_phone = "1-517-335-3900"
    	agency_fax = "1-517-335-6054"
    	agency_email = "MDHHS-ICU-Customer-Service@michigan.gov"
    	agency_website = "www.michigan.gov/dhhs"
    END IF
    IF	state_droplist = "Mississippi"	THEN
    	abbr_state = "MS"
    	agency_name = "Department of Human Services Division of Field Operations"
    	agency_address = "P.O. Box 352 Jackson, MS 39205"
    	agency_phone = "1-800-948-3050"
    	agency_email = "ea.CustomerService@mdhs.ms.gov"
    	agency_fax = "601-364-7469 "
    	agency_website = "www.mdhs.state.ms.us"
    END IF
    IF	state_droplist = "Missouri"	THEN
    	abbr_state = "MO"
    	agency_name = "Correspondence and Information Unit Family Support Division Department of Social Services"
    	agency_address = "P.O. Box 2320, Jefferson City, MO 65102-2320"
    	agency_phone = "1-800-392-1261"
    	agency_email = "Cole.CoXIX@dss.mo.gov"
     	agency_website = "http://www.dss.mo.gov"
    END IF
    IF	state_droplist = "Montana"	THEN
    	abbr_state = "MT"
    	agency_name = "Department of Public Health & Human Services Human & Community Services Division C/O Julie Nepine "
    	agency_address = "PO Box 202925, Helena, MT  59620-2925"
    	agency_phone = "406-444-2770"
    	agency_fax = "406-444-2770"
    	agency_email = "hhsparis@mt.gov"
    	agency_website = "http://www.dphhs.mt.gov/"
		IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
			other_state_cash = TRUE
			agency_name = "Montana Department of Public Health and Human Services"
			agency_address = "111 N. Jackson, Helena, MT 59601"
			agency_fax = "406-444-2770"
			agency_email = "TANF@mt.gov"
		END IF
    END IF
    IF	state_droplist = "Nebraska"	THEN
    	abbr_state = "NE"
		agency_name = "Economic Assistance Customer Service Center"
		agency_phone = "1-800- 383-4278"
		agency_website = "www.accessnebraska.ne.gov"
    	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
    		other_state_cash = TRUE
    		agency_email = "DHHS.EconomicAssistancePolicyQuestions@nebraska.gov"
    	END IF
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_email = "DHHS.MedicaidPolicyQuestions@nebraska.gov"
    	END IF
    END IF
    IF	state_droplist = "Nevada"	THEN
    	abbr_state = "NV"
    	agency_name = "Dept of Health & Human Services Division of Welfare and Supportive Services"
    	agency_address = "1470 College Parkway Carson City, NV 89706"
    	agency_fax = "702-631-4487 ATTN: OOS Inquiries."
    	agency_email = "WELFOOSINQUIRIES@DWSS.NV.GOV"
    	agency_website = "http://dwss.nv.gov"
    END IF
    IF	state_droplist = "New Hampshire"	THEN
    	abbr_state = "NH"
    	agency_name = "Department of Health & Human Services Bureau of Family Assistance C/O Molly Fulton"
    	agency_phone = "603-444-3663"
    	agency_email = "outofstateinquiries@dhhs.state.nh.us"
    	agency_fax = "603-444-0348"
    	agency_website = "http://www.dhhs.nh.gov"
    END IF
    IF	state_droplist = "New Jersey"	THEN
    	abbr_state = "NJ"
    	agency_name = "Department of Human Services Division of Family Development (DFD)"
    	agency_address = "P.O. Box 716, Trenton, NJ 08625-0716"
    	agency_phone = "609-588-2283"
    	agency_email = "DFD.FIRM@dhs.nj.gov"
    	agency_website = "http://nj.gov/humanservices/dfd"
    END IF
    IF	state_droplist = "New Mexico" THEN
    	abbr_state = "NM"
    	agency_name = "New Mexico Human Services Department"
    	agency_address = "Income Support Division P.O.Box 2348 2009 S Pacheco Street Santa Fe, NM  87504"
    	agency_email = "outofstate.inquiry@state.nm.us"
    	agency_website = "www.hsd.state.nm.us"
		IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
			other_state_hc = TRUE
			agency_phone = "1-888-997-2583"
		END IF
    END IF
    IF	state_droplist = "New York"	THEN
    	abbr_state = "NY"
    	agency_name = "New York Human Services Department"
    	agency_phone = "518-486-3460"
    	agency_fax = "518-474-8090"
    	agency_email = "wendy.buell@otda.ny.gov"
    END IF
    IF	state_droplist = "North Carolina" THEN
    	abbr_state = "NC"
    	agency_name = "North Carolina Department of Health & Human Services Division of Social Services"
    	agency_address = "DHHS (EBT) Call Center P. O. Box 190 Everetts, NC  27825"
    	agency_phone = "1-866-719-0141"
    	agency_fax = "252-789-5395"
    	agency_email = "ebt.csc.leads@dhhs.nc.gov"
    	agency_website = "www.ncdhhs.gov/dss"
    END IF
    IF	state_droplist = "North Dakota"	THEN
    	abbr_state = "ND"
    	agency_name = "North Dakota Human Services Department"
    	agency_email = "dhseap@nd.gov"
    	agency_phone = "701-328-2332"
    	agency_website = "www.nd.gov/dhs/services"
    END IF
    IF	state_droplist = "Ohio"	THEN
    	abbr_state = "OH"
    	agency_name = "Office of Family Assistance, Ohio Department of Job & Family Services"
    	agency_address = "P.O. Box 183204, Columbus, Ohio 43218-3204"
    	agency_phone = "614-466-4815"
    	agency_ext = "option 2, 1"
    	agency_fax = "614-466-1767"
    	agency_email = "Out_of_State_Inquiries@jfs.ohio.gov"
    	agency_website = "www.jfs.ohio.gov/"
    END IF
    IF	state_droplist = "Oklahoma"	THEN
    	abbr_state = "OK"
		agency_name = "Oklahoma Human Services Department"
		agency_email = "SNAP@okdhs.org"
		agency_phone = "405-521-3444"
		agency_fax = "405-521-4158"
		agency_website = "www.okdhs.org"
    	IF OTHER_STATE_FS_CHECKBOX = CHECKED THEN
    		other_state_fs = TRUE
    		agency_email = "SNAP@okdhs.org"
    	END IF
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_email = "eligibility@OKHCA.org"
    	END IF
    END IF
    IF	state_droplist = "Oregon" THEN
    	abbr_state = "OR"
    	agency_name = "	Oregon Department of Human Services"
    	agency_address = "500 Summer St. NE, E-48 Salem, OR  97301-1066"
    	agency_phone = "503-945-5600"
    	agency_fax = "503-373-7032"
    	agency_email = "benefits.verification@state.or.us"
    	agency_website = "http://www.oregon.gov/DHS/assistance/index.shtml"
    END IF
    IF	state_droplist = "Pennsylvania"	THEN
    	abbr_state = "PA"
    	agency_name = "Pennsylvania Department of Human Services"
    	agency_address = "P.O. Box 2675 Harrisburg, PA 17105-2675"
    	agency_phone = "717-787-3119"
    	agency_fax = "717-705-0040"
    	agency_email = "ra-dpwoimnet@pa.gov"
    	agency_website = "www.DHS.state.pa.us"
    END IF
    IF	state_droplist = "Rhode Island"	THEN
    	abbr_state = "RI"
		agency_name = "RI Department of Human Services"
		agency_address = "Louis Pasteur Building 57 25 Howard Avenue Cranston, RI 02920"
		agency_fax = "401-721-6664"
		agency_email = "DHS.SNAP-Inquiry@dhs.ri.gov "
		agency_website = "http://www.dhs.ri.gov/"
    	IF PARIS_CHECKBOX = CHECKED THEN
    		other_state_paris = TRUE
    		agency_name = "State of Rhode Island and Providence Plantations, Department of Administration"
    		agency_address ="Office of Internal Audits One Capitol Hill – 4th Floor Providence, RI 02908"
    		agency_phone = "401-574-8175"
    		agency_email = "DHS.SNAP-Inquiry@dhs.ri.gov"
    		agency_fax = "401-721-6664"
    	END IF
    END IF
    IF	state_droplist = "South Carolina"	THEN
    	abbr_state = "SC"
		agency_name = "South Carolina Department of Social Services"
    	agency_address = "Out-of-State Inquiries Program Support Unit, Division of County Operations P.O. Box 1520 Columbia, SC 29202-1520"
    	agency_fax = "803-898-1222, ATTN: Program Support Unit"
    	agency_email = "SCDSSVerify@dss.sc.gov"
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_phone " 888-549-0820.  Press #1 for English, #1 for caseworker feedback and #2 to speak with a rep."
    	END IF
    	If other_state_paris = TRUE Then agency_email = agency_email & "; Keshawn.Jacobs@dss.sc.gov"
    END IF
    IF	state_droplist = "South Dakota"	THEN
    	abbr_state = "SD"
    	agency_name = "Department of Social Services"
    	agency_address = "700 Governors Drive Pierre, South Dakota  57501-2291"
    	agency_phone = "1-877-999-5612"
    	agency_email = "SNAP@state.sd.us"
    	agency_website = "http://dss.sd.gov/economicassistance/snap/"
    END IF
    IF	state_droplist = "Tennessee" THEN
    	abbr_state = "TN"
    	agency_name = "TN Supplemental Nutrition Assistance Program (SNAP) Policy"
    	agency_address = "James K. Polk Bldg. 15th Floor 505 Deaderick Street Nashville, TN 37243"
    	agency_email = "Paris.inquiries@tn.gov"
    	agency_website = "https://stateoftennessee.formstack.com/forms/out_of_state_inquiries"
    END IF
    IF	state_droplist = "Texas"	THEN
    	abbr_state = "TX"
    	agency_name = "Texas Health and Human Services Commission"
    	agency_fax = "1-877-447-2839"
    END IF
    IF	state_droplist = "Utah"	THEN
    	abbr_state = "UT"
    	agency_name = "Department of Workforce Services Eligibility Services Division"
    	agency_address = "P.O. Box 143245 Salt Lake City, UT 84114-3245"
    	agency_phone = "866-435-7414"
    	agency_ext = "Option 5"
    	agency_website = "www.jobs.utah.gov"
    END IF
    IF	state_droplist = "Vermont"	THEN
    	abbr_state = "VT"
    	agency_name = "Economic Services Benefits Service Center"
    	agency_address = "1000 River Rd Essex Junction, VT 05452"
    	agency_phone = "1-800- 479-6151"
    	agency_website = "www.mybenefits.vt.gov "
    END IF
    IF	state_droplist = "Virginia"	THEN
    	abbr_state = "VA"
    	agency_name = "Virginia Department of Social Services Division of Benefit Programs"
    	agency_address = "801 East Main St Richmond, VA  23219-2901"
    	agency_phone = "1-800- 479-6151"
    	agency_email = "vaoutofstateverifications@dss.virginia.gov."
    	agency_website = "http://dss.virginia.gov/benefit/snap.cgi"
    END IF
    IF	state_droplist = "Washington"	THEN
    	abbr_state = "WA"
		agency_name = "Washington Dept Human Services"
		agency_phone = "1-855-927-2747"
		agency_email = "dshsparissupport@dshs.wa.gov"
    	If other_state_paris = TRUE Then
    		agency_email = "dshsparissupport@dshs.wa.gov"
    		agency_fax = 1-888-212-2319
    	END IF
    END IF
    IF	state_droplist = "West Virginia"	THEN
    	abbr_state = "WV"
    	agency_name = "Department of Health & Human Resources Division of Family Assistance"
    	agency_address = "350 Capitol St., Room B-18Charleston, WV  25301-3705"
    	agency_phone ="304-356-4619"
    	agency_email = "DHHRbcfbenefitver@wv.gov"
    	agency_website = "www.dhhr.wv.gov"
    END IF
    IF	state_droplist = "Wisconsin"	THEN
    	abbr_state = "WI"
		agency_name = "	WI Department of Health & Family Services"
		agency_address = "1 W Wilson St, Madison, WI 53703"
		agency_phone = "608-261-6378"
		agency_ext = "Option 3"
		agency_fax = "608-267-2269"
		agency_email = "DHSOSBQ@dhs.wisconsin.gov"
    	IF OTHER_STATE_FS_CHECKBOX = CHECKED THEN
    		other_state_fs = TRUE
    		agency_name = "WI Department of Health & Family Services"
    		agency_address = "1 W Wilson St, Madison, WI 53703"
    		agency_phone = "608-261-6378"
    		agency_ext = "Option 3"
    		agency_fax = "608-267-2269"
    		agency_email = "DHSOSBQ@dhs.wisconsin.gov"
    	END IF
    	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN
    		other_state_cash = TRUE 
    		agency_phone = "608-422-7900"
    		agency_fax = "608-327-6125"
    		agency_email = "DCFW2TANFVerify@wisconsin.gov"
    	END IF
    	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_email = "DHSOIGPARIS@dhs.wisconsin.gov"
    		agency_phone ="608-267-0470"
    	END IF
    	IF PARIS_CHECKBOX = CHECKED THEN
    		other_state_hc = TRUE
    		agency_email = "DHSOIGPARIS@dhs.wisconsin.gov"
    		agency_phone ="608-267-0470"
    	END IF
    END IF
    IF	state_droplist = "Wyoming"	THEN
    	abbr_state = "WY"
    	agency_name = "Department of Family Services C/O Ann Bowen, SNAP/TANF Help Desk"
    	agency_address = "2300 Capitol Ave., Hathaway Bld, Third Floor"
    	agency_phone = "307-777-6082"
    	agency_fax = "1-307-777-6276"
    	agency_email = "ann.bowen@wyo.gov"
    	agency_website = "https://dfs.wyo.gov/assistance-programs/food-assistance/"
    END IF '---------------------------------------------end of states
END Function
'blanking out variables for the function'
abbr_state = ""
agency_name = ""
agency_address = ""
agency_phone = ""
agency_fax = ""
agency_email = ""
agency_website = ""

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'back_to_SELF' added to ensure we have the time to update and send the case in the background

Dialog1 = ""
BEGINDIALOG Dialog1, 0, 0, 146, 105, "Out of State Inquiry"
 EditBox 55, 5, 55, 15, MAXIS_case_number
 DropListBox 55, 25, 85, 15, "Send", out_of_state_request '+chr(9)+"Received"+chr(9)+"Unknown/No Response"'
 DropListBox 55, 45, 85, 15, "Select One:"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Phone", how_sent
 DropListBox 55, 65, 85, 15, "Select One:"+chr(9)+"Alabama"+chr(9)+"Alaska"+chr(9)+"Arizona"+chr(9)+"Arkansas"+chr(9)+"California"+chr(9)+"Colorado"+chr(9)+"Connecticut"+chr(9)+"Delaware"+chr(9)+"Florida"+chr(9)+"Georgia"+chr(9)+"Hawaii"+chr(9)+"Idaho"+chr(9)+"Illinois"+chr(9)+"Indiana"+chr(9)+"Iowa"+chr(9)+"Kansas"+chr(9)+"Kentucky"+chr(9)+"Louisiana"+chr(9)+"Maine"+chr(9)+"Maryland"+chr(9)+"Massachusetts"+chr(9)+"Michigan"+chr(9)+"Mississippi"+chr(9)+"Missouri"+chr(9)+"Montana"+chr(9)+"Nebraska"+chr(9)+"Nevada"+chr(9)+"New Hampshire"+chr(9)+"New Jersey"+chr(9)+"New Mexico"+chr(9)+"New York"+chr(9)+"North Carolina"+chr(9)+"North Dakota"+chr(9)+"Ohio"+chr(9)+"Oklahoma"+chr(9)+"Oregon"+chr(9)+"Pennsylvania"+chr(9)+"Rhode Island"+chr(9)+"South Carolina"+chr(9)+"South Dakota"+chr(9)+"Tennessee"+chr(9)+"Texas"+chr(9)+"Utah"+chr(9)+"Vermont"+chr(9)+"Virginia"+chr(9)+"Washington"+chr(9)+"West Virginia"+chr(9)+"Wisconsin"+chr(9)+"Wyoming", state_droplist
 ButtonGroup ButtonPressed
   OkButton 55, 85, 40, 15
   CancelButton 100, 85, 40, 15
   Text 5, 10, 50, 10, "Case Number:"
   Text 20, 30, 30, 10, "Request:"
   Text 20, 50, 30, 10, "Via(How):"
   Text 30, 70, 20, 10, "State:"
ENDDIALOG

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "Enter a valid case number."
		If state_droplist = "Select One:" then err_msg = err_msg & vbnewline & "Select the state."
		If how_sent = "Select One:" then err_msg = err_msg & vbnewline & "Select how the request was sent."
		If out_of_state_request = "Select One:" then err_msg = err_msg & vbNewLine & "Please select the status of the out of state inquiry."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false

'Error proof functions
Call check_for_MAXIS(False)
'msgbox "do i even get in?"
'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)

'Navigate back to self
'Back_to_self
'EMWriteScreen MAXIS_case_number, 18, 43
Call navigate_to_MAXIS_screen("CASE", "CURR")
EMReadScreen CURR_panel_check, 4, 2, 55
EMReadScreen case_status, 8, 8, 9
case_status = trim(case_status)
IF case_status = "ACTIVE" THEN active_status = TRUE
IF case_status = "APP OPEN" THEN active_status = TRUE
IF case_status = "APP CLOS" THEN active_status = TRUE
If case_status = "CAF2 PEN" THEN active_status = TRUE
If case_status = "CAF1 PEN" THEN active_status = TRUE
IF case_status = "REIN" THEN active_status = TRUE
IF case_status = "INACTIVE" THEN active_status = FALSE
Call MAXIS_footer_month_confirmation
EmReadscreen original_MAXIS_footer_month, 2, 20, 43


CALL navigate_to_MAXIS_screen("STAT", "ADDR")
EMreadscreen resi_addr_line_one, 22, 6, 43
resi_addr_line_one = replace(resi_addr_line_one, "_", "")
EMreadscreen resi_addr_line_two, 22, 7, 43
resi_addr_line_two = replace(resi_addr_line_two, "_", "")
EMreadscreen resi_addr_city, 15, 8, 43
resi_addr_city = replace(resi_addr_city, "_", "")
EMreadscreen resi_addr_state, 2, 8, 66
resi_addr_state = replace(resi_addr_state, "_", "")
EMreadscreen resi_addr_zip, 7, 9, 43
resi_addr_zip = replace(resi_addr_zip, "_", "")


EMreadscreen addr_homeless, 1, 10, 43





'this reads clients current mailing address for the letter
Call navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen mail_address, 1, 13, 64
If mail_address = "_" then
	 EMReadScreen client_1staddress, 21, 06, 43
	 EMReadScreen client_2ndaddress, 21, 07, 43
	 EMReadScreen client_city, 14, 08, 43
	 EMReadScreen client_state, 2, 08, 66
	 EMReadScreen client_zip, 7, 09, 43
Else
	 EMReadScreen client_1staddress, 21, 13, 43
	 EMReadScreen client_2ndaddress, 21, 14, 43
	 EMReadScreen client_city, 14, 15, 43
	 EMReadScreen client_state, 2, 16, 43
	 EMReadScreen client_zip, 7, 16, 52
End If
client_address = replace(client_1staddress, "_","") & " " & replace(client_2ndaddress, "_","") & " " & replace(client_city, "_","") & ", " & replace(client_state, "_","") & " " & replace(client_zip, "_","")

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'Checking for PRIV cases.
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
IF priv_check = "PRIV" THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
MN_CASH_STATUS = FALSE 'overall variable'
MN_CCA_STATUS = FALSE
MN_DWP_STATUS = FALSE 'Diversionary Work Program'
MN_ER_STATUS = FALSE
MN_FS_STATUS = FALSE
MN_GA_STATUS = FALSE 'General Assistance'
MN_GRH_STATUS = FALSE
MN_HC_STATUS = FALSE
MN_MS_STATUS = FALSE 'Mn Suppl Aid '
MN_MF_STATUS = FALSE 'Mn Family Invest Program '
MN_RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND" THEN
	MN_FS_STATUS = TRUE
	MN_FS_CHECKBOX = CHECKED
END IF
IF hc_status_check = "ACTV" or hc_status_check = "PEND" THEN
	MN_HC_STATUS = TRUE
	MN_HC_CHECKBOX  = CHECKED
END IF
IF cca_status_check = "ACTV" or cca_status_check = "PEND" THEN
	MN_CCA_STATUS = TRUE
	MN_CCA_CHECKBOX  = CHECKED
END IF
IF cash1_status_check = "ACTV"  or cash1_status_check = "PEND" THEN
	'Msgbox MN_CASH_STATUS
	MN_CASH_STATUS = TRUE
	MN_CASH_CHECKBOX = CHECKED
END IF
IF cash2_status_check = "ACTV"  or cash2_status_check = "PEND" THEN
	'Msgbox MN_CASH_STATUS
	MN_CASH_STATUS = TRUE 'this is because we dont care what cash 2 program it is for out of state '
	MN_CASH_CHECKBOX = CHECKED
END IF
IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN MN_ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN MN_GRH_STATUS = TRUE

IF MN_MF_STATUS = FALSE and MN_FS_STATUS = FALSE and MN_HC_STATUS = FALSE and MN_DWP_STATUS = FALSE and MN_CASH_STATUS = FALSE AND MN_ER_STATUS = FALSE AND MN_GRH_STATUS = FALSE THEN
	active_status = FALSE
	script_end_procedure_with_error_report("It appears no programs are open or pending on this case.")
END IF
'come back to this for the case note need to run the dialog first '
programs_applied_for = ""        'Creates a variable that lists all the active.
IF MN_CASH_STATUS = TRUE THEN programs_applied_for = programs_applied_for & "CASH, "
IF MN_ER_STATUS = TRUE THEN programs_applied_for = programs_applied_for & "Emergency, "
IF MN_GRH_STATUS  = TRUE THEN programs_applied_for = programs_applied_for & "GRH, "
IF MN_FS_STATUS = TRUE THEN programs_applied_for = programs_applied_for & "SNAP, "
'IF MN_IVE_STATUS  = TRUE THEN programs_applied_for = programs_applied_for & "IV-E, "
IF MN_HC_STATUS   = TRUE THEN programs_applied_for = programs_applied_for & "HC, "
IF MN_CCA_STATUS  = TRUE THEN programs_applied_for = programs_applied_for & "CCA"

programs_applied_for = trim(programs_applied_for)  'trims excess spaces of programs_applied_for
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

Const ref_numb_const 			= 0
Const first_name_const 			= 1
Const last_name_const			= 2
Const clt_middle_const 			= 3
Const clt_dob_const 			= 4
Const client_selection_checkbox_const = 5
Const clt_ssn_const 			= 6

Dim ALL_CLT_INFO_ARRAY()
ReDim ALL_CLT_INFO_ARRAY(clt_ssn_const, 0)

the_incrementer = 0
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
	ReDim Preserve ALL_CLT_INFO_ARRAY(clt_ssn_const, the_incrementer)' this is to tell the array to get bigger'
	ALL_CLT_INFO_ARRAY(ref_numb_const, the_incrementer) = ref_nbr 'going back to the first piece of information to hold it in this specific postion'
	'MsgBox access_denied_check
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		last_name = "UNABLE TO FIND"
		first_name = " - Access Denied"
		mid_initial = ""
	Else
	    'Reading info and removing spaces
	    EMReadscreen First_name, 12, 6, 63
	    First_name = replace(First_name, "_", "")
	    ALL_CLT_INFO_ARRAY(first_name_const, the_incrementer) = First_name
	    'Reading Last name and removing spaces
	    EMReadscreen Last_name, 25, 6, 30
	    Last_name = replace(Last_name, "_", "")
	    ALL_CLT_INFO_ARRAY(last_name_const, the_incrementer) = Last_name
	    'Reading Middle initial and replacing _ with a blank if empty.
	    EMReadscreen Middle_initial, 1, 6, 79
	    Middle_initial = replace(Middle_initial, "_", "")
	    ALL_CLT_INFO_ARRAY(clt_middle_const, the_incrementer) = Middle_initial
		 'Reading date of birth and replacing space.
	    Emreadscreen client_dob, 10, 8, 42
	    SSN_number = replace(client_dob, " ", "/")
	    ALL_CLT_INFO_ARRAY(clt_dob_const, the_incrementer) = client_dob
	    'Reads SSN
	    Emreadscreen SSN_number, 11, 7, 42
	    SSN_number = replace(SSN_number, " ", "-")
	    ALL_CLT_INFO_ARRAY(clt_ssn_const, the_incrementer) = SSN_number
		'adds the ref number to the array'
	    ALL_CLT_INFO_ARRAY(ref_numb_const, the_incrementer) = client_ref_number
		'ensuring that the check box is checked for all members in the dialog'
		ALL_CLT_INFO_ARRAY(client_selection_checkbox_const, the_incrementer) = CHECKED
	End If
	the_incrementer = the_incrementer + 1
	TRANSMIT
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'For each path the script takes a different route'
Dialog1 = "" 'runs the dialog that has been dynamically created. Streamlined with new functions.
BEGINDIALOG Dialog1, 0, 0, 241, (50 + (Ubound(ALL_CLT_INFO_ARRAY, 2) * 15)), "Household Member(s) "   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 130, 10, "Select household members to request:"
	FOR the_pers = 0 to Ubound(ALL_CLT_INFO_ARRAY, 2)
		checkbox 10, (20 + (the_pers * 15)), 160, 10, ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(last_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(clt_ssn_const, the_pers), ALL_CLT_INFO_ARRAY(client_selection_checkbox_const, the_pers)
	NEXT
	ButtonGroup ButtonPressed
	OkButton 185, 10, 50, 15
	CancelButton 185, 30, 50, 15
ENDDIALOG

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
	LOOP until err_msg = ""
Loop until are_we_passworded_out = false
Call fill_in_the_states
IF agency_phone = "" THEN agency_phone = "N/A"
IF agency_fax = "" THEN agency_fax = "N/A"
IF agency_email = "" THEN agency_email = "N/A"
date_received = ""

Dialog1 = "" 'runs the dialog that has been dynamically created. Streamlined with new functions.
BeginDialog Dialog1, 0, 0, 231, 140, "OUT OF STATE INQUIRY FOR: "   & Ucase(state_droplist)
  CheckBox 50, 20, 30, 10, "Cash", MN_CASH_CHECKBOX
  CheckBox 80, 20, 25, 10, "CCA", MN_CCA_CHECKBOX
  CheckBox 110, 20, 20, 10, "FS", MN_FS_CHECKBOX
  CheckBox 135, 20, 25, 10, "HC", MN_HC_CHECKBOX
  CheckBox 160, 20, 25, 10, "GRH", MN_GRH_CHECKBOX
  CheckBox 190, 20, 25, 10, "SSI", MN_SSI_CHECKBOX
  CheckBox 50, 45, 30, 10, "Cash", OTHER_STATE_CASH_CHECKBOX
  CheckBox 80, 45, 25, 10, "CCA", OTHER_STATE_CCA_CHECKBOX
  CheckBox 110, 45, 20, 10, "FS", OTHER_STATE_FS_CHECKBOX
  CheckBox 135, 45, 25, 10, "HC", OTHER_STATE_HC_CHECKBOX
  CheckBox 160, 45, 25, 10, "SSI", OTHER_STATE_SSI_CHECKBOX
  CheckBox 185, 45, 40, 10, "OTHER", OTHER_STATE_CHECKBOX
  DropListBox 35, 60, 55, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Closed"+chr(9)+"Unknown", out_of_state_status
  EditBox 175, 60, 45, 15, date_received
  CheckBox 170, 85, 60, 10, "PARIS Match", PARIS_CHECKBOX
  CheckBox 10, 85, 130, 10, "Set an Outlook reminder to follow up ", outlook_remider_CHECKBOX
  EditBox 50, 100, 175, 15, other_notes
  ButtonGroup ButtonPressed
    PushButton 5, 120, 60, 15, "HSR MANUAL", outofstate_button
    OkButton 130, 120, 45, 15
    CancelButton 180, 120, 45, 15
  Text 10, 20, 40, 10, "Programs:"
  GroupBox 5, 5, 220, 30, "Current programs pending or active on in MN"
  Text 120, 65, 50, 10, "Last Received:"
  Text 10, 45, 40, 10, "Programs:"
  Text 10, 65, 25, 10, "Status:"
  GroupBox 5, 35, 220, 45, "Client reported they received assistance (Q5 on CAF):"
  Text 5, 105, 45, 10, "Other Notes:"
EndDialog

IF out_of_state_request = "Sent/Send" THEN
    'DO
    	DO
    		DO  'External resource DO loop
			    Dialog1 = "" 'runs the dialog that has been dynamically created. Streamlined with new functions.
			    BeginDialog Dialog1, 0, 0, 231, 140, "OUT OF STATE INQUIRY FOR: "   & Ucase(state_droplist)
			      CheckBox 50, 20, 30, 10, "Cash", MN_CASH_CHECKBOX
			      CheckBox 80, 20, 25, 10, "CCA", MN_CCA_CHECKBOX
			      CheckBox 110, 20, 20, 10, "FS", MN_FS_CHECKBOX
			      CheckBox 135, 20, 25, 10, "HC", MN_HC_CHECKBOX
			      CheckBox 160, 20, 25, 10, "GRH", MN_GRH_CHECKBOX
			      CheckBox 190, 20, 25, 10, "SSI", MN_SSI_CHECKBOX
			      CheckBox 50, 45, 30, 10, "Cash", OTHER_STATE_CASH_CHECKBOX
			      CheckBox 80, 45, 25, 10, "CCA", OTHER_STATE_CCA_CHECKBOX
			      CheckBox 110, 45, 20, 10, "FS", OTHER_STATE_FS_CHECKBOX
			      CheckBox 135, 45, 25, 10, "HC", OTHER_STATE_HC_CHECKBOX
			      CheckBox 160, 45, 25, 10, "SSI", OTHER_STATE_SSI_CHECKBOX
			      CheckBox 185, 45, 40, 10, "OTHER", OTHER_STATE_CHECKBOX
			      DropListBox 35, 60, 55, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Closed"+chr(9)+"Unknown", out_of_state_status
			      EditBox 175, 60, 45, 15, date_received
			      CheckBox 170, 85, 60, 10, "PARIS Match", PARIS_CHECKBOX
			      CheckBox 10, 85, 130, 10, "Set an Outlook reminder to follow up ", outlook_remider_CHECKBOX
			      EditBox 50, 100, 175, 15, other_notes
			      ButtonGroup ButtonPressed
			    	PushButton 5, 120, 60, 15, "HSR MANUAL", outofstate_button
			    	OkButton 130, 120, 45, 15
			    	CancelButton 180, 120, 45, 15
			      Text 10, 20, 40, 10, "Programs:"
			      GroupBox 5, 5, 220, 30, "Current programs pending or active on in MN"
			      Text 120, 65, 50, 10, "Last Received:"
			      Text 10, 45, 40, 10, "Programs:"
			      Text 10, 65, 25, 10, "Status:"
			      GroupBox 5, 35, 220, 45, "Client reported they received assistance (Q5 on CAF):"
			      Text 5, 105, 45, 10, "Other Notes:"
			    EndDialog

				Dialog Dialog1
    			cancel_confirmation
    			If ButtonPressed = outofstate_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Out_of_State_Inquiry.aspx")
    		Loop until ButtonPressed = -1
    		err_msg = ""
			If agency_state_droplist = "Select One:" THEN  err_msg = err_msg & vbnewline & "Select the state."
            IF out_of_state_status = "Select One:" then err_msg = err_msg & vbnewline & "Please select the reported status regarding the other state's benefits."
            IF out_of_state_status = "Active" AND trim(date_received) = "" then err_msg = err_msg & vbcr & "Enter the date the client reported benefits were last received."
    		IF out_of_state_status = "Closed" AND trim(date_received) = "" then err_msg = err_msg & vbcr & "Enter the date the client reported benefits were last received."
    		IF out_of_state_status = "Unknown" AND other_notes = "" then err_msg = err_msg & vbcr & "Please advise why the status of previous benefits is unknown."

		Dialog1 = "" 'blanking the previous dialog
		BeginDialog Dialog1, 0, 0, 301, 240, "OUT OF STATE INQUIRY FOR: " & Ucase(state_droplist)
		 	Text 10, 10, 290, 10, "Based on the detail entered, the National Directory has the following contact information:"
		 	GroupBox 5, 25, 290, 170, ""
	  	Text 15, 40, 240, 25, "Name: " & Ucase(agency_name)
	  	Text 15, 75, 240, 20, "Address: " & Ucase(agency_address)
	  	Text 15, 105, 100, 10, "Phone: "  & agency_phone
	  	Text 135, 105, 120, 15, "Fax:  " & agency_fax
	  	Text 15, 135, 240, 25, "Email: " & agency_email
	  	ButtonGroup ButtonPressed
	    PushButton 10, 200, 160, 15, "National directory requires an update", change_the_detail_btn
		   OkButton 195, 220, 50, 15
	   		CancelButton 250, 220, 45, 15
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		If ButtonPressed = change_the_detail_btn THEN
			new_addr_detail_entered = TRUE
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 226, 130, "OUT OF STATE INQUIRY"
	          EditBox 45, 20, 165, 15, agency_name
	          EditBox 45, 40, 165, 15, agency_address
	          EditBox 45, 60, 165, 15, agency_email
	          EditBox 45, 80, 50, 15, agency_phone
	          EditBox 160, 80, 50, 15, agency_fax
	          ButtonGroup ButtonPressed
	        	OkButton 125, 110, 45, 15
	        	CancelButton 175, 110, 45, 15
	          GroupBox 5, 5, 215, 95, "Out of State Agency Contact Information"
	          Text 10, 25, 25, 10, "Name:"
	          Text 10, 45, 30, 10, "Address:"
	          Text 10, 65, 25, 10, "Email:"
	          Text 10, 85, 25, 10, "Phone:"
	          Text 140, 85, 15, 10, "Fax:"
	        EndDialog
        	DO      'Password DO loop
        		DO  'Conditional handling DO loop
        			Dialog Dialog1
        			cancel_without_confirmation
        			err_msg = ""
        			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        		LOOP until err_msg = ""
        		CALL check_for_password(are_we_passworded_out)
        	Loop until are_we_passworded_out = false
        	err_msg = "LOOP"
		ELSE
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		END IF
	LOOP until err_msg = ""
	'CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure
'Loop until are_we_passworded_out = false
If OTHER_STATE_CASH_CHECKBOX = Checked THen other_state_cash = TRUE
' If OTHER_STATE_CCA_CHECKBOX = Checked THen
If OTHER_STATE_FS_CHECKBOX = Checked THen other_state_fs = TRUE
If OTHER_STATE_HC_CHECKBOX = Checked THen other_state_hc = TRUE
' If OTHER_STATE_SSI_CHECKBOX = Checked THen
If OTHER_STATE_CHECKBOX = Checked THen other_state_paris = TRUE

	other_state_programs = ""        'Creates a variable that lists all the active.
	IF OTHER_STATE_CASH_CHECKBOX = CHECKED THEN other_state_programs = other_state_programs & "CASH, "
	IF OTHER_STATE_FS_CHECKBOX = CHECKED THEN other_state_programs = other_state_programs & "SNAP, "
	IF OTHER_STATE_HC_CHECKBOX = CHECKED THEN other_state_programs = other_state_programs & "HC, "
	IF OTHER_STATE_CCA_CHECKBOX = CHECKED THEN other_state_programs = other_state_programs & "CCA,"
	IF OTHER_STATE_CHECKBOX or OTHER_STATE_SSI_CHECKBOX  = CHECKED THEN other_state_programs = other_state_programs & "Other,"
	other_state_programs = trim(other_state_programs)  'trims excess spaces of other_state_programs
	If right(other_state_programs, 1) = "," THEN other_state_programs = left(other_state_programs, len(other_state_programs) - 1)


	'Generates Word Doc Form
    Set objWord = CreateObject("Word.Application")
    objWord.Caption = "OUT OF STATE INQUIRY"
 	objWord.Visible = True
    Set objDoc = objWord.Documents.Add()
    Set objSelection = objWord.Selection
    'objSelection.ParagraphFormat.Alignment = 0
    objSelection.ParagraphFormat.LineSpacing = 12
    objSelection.ParagraphFormat.SpaceBefore = 0
    objSelection.ParagraphFormat.SpaceAfter = 0
    objSelection.Font.Name = "Calibri"
    objSelection.Font.Size = "12"
    objSelection.TypeText "OUT OF STATE INQUIRY"
    objSelection.TypeParagraph
	objSelection.TypeParagraph
    objSelection.TypeText "Hennepin County Human Services & Public Health Department"
    objSelection.TypeParagraph
    objSelection.TypeText "PO Box 107 Minneapolis, MN 55440-0107"
    objSelection.TypeParagraph
    objSelection.TypeText "Phone: 612-596-8500"
    objSelection.TypeParagraph
	objSelection.TypeText "Fax: 612-288-2981"
	objSelection.TypeParagraph
    objSelection.TypeText "Email: HHSEWS@hennepin.us"
    objSelection.TypeParagraph
    objSelection.ParagraphFormat.Alignment = 2
    objSelection.ParagraphFormat.LineSpacing = 12
    objSelection.ParagraphFormat.SpaceBefore = 0
    objSelection.ParagraphFormat.SpaceAfter = 0
    objSelection.Font.Name = "Calibri"
    objSelection.Font.Size = "12"
    objSelection.TypeText "Date: " & date()
    objSelection.TypeParagraph
	objSelection.ParagraphFormat.Alignment = 0
    objSelection.ParagraphFormat.LineSpacing = 12
    objSelection.ParagraphFormat.SpaceBefore = 0
    objSelection.ParagraphFormat.SpaceAfter = 0
    objSelection.Font.Name = "Calibri"
    objSelection.Font.Size = "12"
    'objSelection.Font.Bold = True
    objSelection.TypeText "To: " & agency_name
    objSelection.TypeParagraph
    IF agency_address <> "" THEN
		objSelection.TypeText agency_address
    	objSelection.TypeParagraph
	END IF

	IF agency_address <> "" THEN
		objSelection.TypeText "Phone: " & agency_phone
    	objSelection.TypeParagraph
	END IF
	IF agency_address <> "" THEN
		objSelection.TypeText "Fax: " & agency_fax
    	objSelection.TypeParagraph
	END IF
	IF agency_address <> "" THEN
		objSelection.TypeText "Email: " & agency_email
		objSelection.TypeParagraph
	END IF
	objSelection.TypeText " "
    objSelection.TypeParagraph
    objSelection.TypeText "RE: "
    For the_pers = 0 to UBound(ALL_CLT_INFO_ARRAY, 2)
		objSelection.TypeParagraph
    	objSelection.TypeText ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(last_name_const, the_pers)
		objSelection.TypeParagraph
		objSelection.TypeText "  SSN: "  & ALL_CLT_INFO_ARRAY(clt_ssn_const, the_pers) & "  DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers)
		objSelection.TypeParagraph
    NEXT
    objSelection.TypeParagraph
    objSelection.TypeText "Client's current address: " & client_address
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText "Our records indicate that the above individual(s) received or receives assistance from your state.  We need to verify the number of months of federally-funded TANF cash assistance issued by your state that count towards the 60 month lifetime limit.  In addition, we need to know the number of months of TANF assistance from other states that your agency has verified.  "
    objSelection.TypeText "Please indicate if the client is open on SNAP or Medical Assistance in your state or the date these programs most recently closed.  Thank you."
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText "Is CASH currently closed?         ____YES ____NO		Date of closure: "
    objSelection.TypeParagraph
    objSelection.TypeText "Is SNAP currently closed?         ____YES ____NO		Date of closure: "
	objSelection.TypeParagraph
	objSelection.TypeText "Is Medical Assistance closed?  ____YES ____NO		Date of closure: "
	objSelection.TypeParagraph
    'objSelection.TypeParagraph
    'objSelection.TypeText "Total ABAWD months used:"
    'objSelection.TypeParagraph
    'objSelection.TypeText "Please list the month(s)/year(s) of ABAWD months used: "
    'objSelection.TypeParagraph
	objSelection.TypeParagraph
    objSelection.TypeText "Please complete the following:"
    objSelection.TypeParagraph
    objSelection.TypeText "Circle the month(s)/year(s) the person received federally funded TANF cash assistance: "
	objSelection.TypeParagraph
    objSelection.TypeText Year(date)-20 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-19 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-18 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-17 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-16 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-15 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-14 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-13 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-12 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-11 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-10 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-9 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-8 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-7 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-6 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-5 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-4 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-3 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-2 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date)-1 & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
    objSelection.TypeText Year(date) & ":   Jan  Feb  Mar  Apr  May  Jun  Jul  Aug  Sep  Oct  Nov  Dec"
    objSelection.TypeParagraph
	objSelection.TypeParagraph
    objSelection.TypeText "Name of person verifying information: "
    objSelection.TypeParagraph
    objSelection.TypeText "Contact Information: "
    objSelection.TypeParagraph
    objSelection.TypeParagraph
    objSelection.TypeText "Please reply or email your response to: hhsews@hennepin.us"
    objSelection.TypeParagraph
    objSelection.TypeParagraph 'end of word doc'

	start_a_blank_case_note
	Call write_variable_in_CASE_NOTE("---Out of State Inquiry sent via " & how_sent & " to " & abbr_state & "---")
	IF out_of_state_status <> "Unknown" THEN CALL write_variable_in_CASE_NOTE("* Client reported they received " & other_state_programs & " on " & date_received & " the case is currently: " & out_of_state_status)
	CALL write_bullet_and_variable_in_case_note("MN Program(s) applied for", programs_applied_for)
	CALL write_bullet_and_variable_in_CASE_NOTE("Name", agency_name)
	CALL write_bullet_and_variable_in_CASE_NOTE("Address", agency_address)
	CALL write_bullet_and_variable_in_CASE_NOTE("Email", agency_email)
	CALL write_bullet_and_variable_in_CASE_NOTE("Phone", agency_phone)
	CALL write_bullet_and_variable_in_CASE_NOTE("Fax", agency_fax)
	Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3

	FOR the_pers = 0 to UBound(ALL_CLT_INFO_ARRAY, 2)
	    HH_member_array = "RE:" & vbcr & ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(last_name_const, the_pers) & vbcr & "  SSN: "  & ALL_CLT_INFO_ARRAY(clt_ssn_const, the_pers) & "  DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers) & HH_member_array
	NEXT

	message_array = ("Hennepin County Human Services & Public Health Department" & vbcr & "PO Box 107, Minneapolis, MN 55440-0107" & vbcr & "Fax: 612-288-2981" & vbcr & "Phone: 612-596-8500" & vbcr & "Email: HHSEWS@hennepin.us" & vbcr & "Date: " & date & vbcr & vbcr & agency_name & vbcr & agency_address & vbcr & "Email: " & agency_email & vbcr & "Phone: " & agency_phone & vbcr & "Fax: " & agency_fax & vbcr & HH_member_array & vbcr & "Current Address: " & client_address & vbcr & "Our records indicate that the above individual(s) received or receives assistance from your state.  We need to verify the number of months of Federally-funded TANF cash assistance issued by your state that count towards the 60 month lifetime limit.  In addition, we need to know the number of months of TANF assistance from other states that your agency has verified.  " & "Please indicate if the client is open on SNAP or Medical Assistance in your state OR the date these programs most recently closed.  Thank you." & vbcr & "Please list the month(s)/year(s) the person received federally funded TANF cash assistance: " & vbcr & "Is CASH/TANF currently closed? ____YES ____NO   Date of closure: " & vbcr & "Is SNAP currently closed?____YES ____NO    Date of closure: " & vbcr & "Is Medical Assistance closed?____YES ____NO    Date of closure: " & vbcr & vbcr & "Name of Person verifying information: " & vbcr & "Contact Information:" & vbcr & "Please reply or email your response to: Hennepin County Human Services and Public Health Services at hhsews@hennepin.us")

    call create_TIKL("Out of State Inquiry Due", 10, date, TRUE, "")
    'create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
    IF outlook_reminder_CHECKBOX = CHECKED THEN
    	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
    	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "Out of State request for " & MAXIS_case_number, "", "", TRUE, 10, "")
    End if
    IF agency_email <> "" THEN
    	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
    	CALL create_outlook_email(agency_email, "","Out of State Inquiry for case #" &  MAXIS_case_number & " [ENCRYPT]", "Out of State Inquiry" & vbcr & message_array,"", False)
    END IF
END IF 'If for sent/send'

IF out_of_state_request = "Received" THEN ' need to look closer to client, month and program specific for duplicate assistance'
	DO      'Password DO loop
	    DO  'Conditional handling DO loop
			Dialog1 = ""
              	BeginDialog Dialog1, 0, 0, 231, 230, "OUT OF STATE INQUIRY RECEIVED FROM: "  & Ucase(state_droplist)
                 CheckBox 50, 20, 30, 10, "Cash", MN_CASH_CHECKBOX
                 CheckBox 80, 20, 25, 10, "CCA", MN_CCA_CHECKBOX
                 CheckBox 110, 20, 20, 10, "FS", MN_FS_CHECKBOX
                 CheckBox 135, 20, 25, 10, "HC", MN_HC_CHECKBOX
                 CheckBox 160, 20, 25, 10, "GRH", MN_GRH_CHECKBOX
                 CheckBox 190, 20, 25, 10, "SSI", MN_SSI_CHECKBOX
                 CheckBox 50, 45, 30, 10, "Cash", OTHER_STATE_CASH_CHECKBOX
                 CheckBox 80, 45, 25, 10, "CCA", OTHER_STATE_CCA_CHECKBOX
                 CheckBox 110, 45, 20, 10, "FS", OTHER_STATE_FS_CHECKBOX
                 CheckBox 135, 45, 25, 10, "HC", OTHER_STATE_HC_CHECKBOX
                 CheckBox 160, 45, 25, 10, "SSI", OTHER_STATE_SSI_CHECKBOX
                 CheckBox 185, 45, 40, 10, "OTHER", OTHER_STATE_CHECKBOX
                 DropListBox 35, 60, 60, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Closed"+chr(9)+"Set to Close"+chr(9)+"Client not known"+chr(9)+"Other", out_of_state_status
                 EditBox 175, 60, 45, 15, date_received
                 Text 10, 95, 210, 25, "Name: " & Ucase(agency_name)
                 Text 10, 120, 100, 10, "Phone: "  & agency_phone
                 Text 135, 120, 90, 15, "Fax: " & agency_fax
                 Text 10, 135, 205, 15, "Email: "  & agency_email
                 CheckBox 10, 170, 170, 10, "Please confirm that the inquiry was sent to ECF", ECF_checkbox
                 EditBox 50, 190, 175, 15, other_notes
                 ButtonGroup ButtonPressed
                   PushButton 10, 155, 160, 15, "National directory requires an update", change_the_detail_btn
                   OkButton 130, 210, 45, 15
                   CancelButton 180, 210, 45, 15
                 Text 10, 20, 40, 10, "Programs:"
                 GroupBox 5, 5, 220, 30, "Current programs pending or active on in MN:"
                 GroupBox 5, 85, 220, 100, "Out of State Agency Contact"
                 Text 120, 65, 50, 10, "Last Received:"
                 Text 10, 45, 40, 10, "Programs:"
                 Text 10, 65, 25, 10, "Status:"
                 GroupBox 5, 35, 220, 45, "State reported client received(s):"
                 Text 5, 195, 45, 10, "Other Notes:"
               EndDialog

		    	Dialog Dialog1
		    	cancel_without_confirmation
		        err_msg = ""

			If ButtonPressed = change_the_detail_btn THEN
				new_addr_detail_entered = TRUE
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 226, 130, "OUT OF STATE INQUIRY"
			      EditBox 45, 20, 165, 15, agency_name
			      EditBox 45, 40, 165, 15, agency_address
			      EditBox 45, 60, 165, 15, agency_email
			      EditBox 45, 80, 50, 15, agency_phone
			      EditBox 160, 80, 50, 15, agency_fax
			      ButtonGroup ButtonPressed
			        OkButton 125, 110, 45, 15
			        CancelButton 175, 110, 45, 15
			      GroupBox 5, 5, 215, 95, "Out of State Agency Contact Information"
			      Text 10, 25, 25, 10, "Name:"
			      Text 10, 45, 30, 10, "Address:"
			      Text 10, 65, 25, 10, "Email:"
			      Text 10, 85, 25, 10, "Phone:"
			      Text 140, 85, 15, 10, "Fax:"
			    EndDialog
				DO      'Password DO loop
					DO  'Conditional handling DO loop
						Dialog Dialog1
						cancel_without_confirmation
						err_msg = ""
						IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
					LOOP until err_msg = ""
					CALL check_for_password(are_we_passworded_out)
				Loop until are_we_passworded_out = false
				err_msg = "LOOP"

			ELSE
				'IF ECF_checkbox <> CHECKED THEN err_msg = err_msg & vbNewLine & "Please review ECF to ensure that the 'verifcations are there."
				'IF OTHER_STATE_CHECKBOX = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise what 'other benefits the client reported."
				'IF out_of_state_status = "Select One:" then err_msg = err_msg & vbnewline & "Please select the reported 'status regarding the other state's benefits."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
			END IF
	    LOOP until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure
	Loop until are_we_passworded_out = false

	start_a_blank_case_note
	Call write_variable_in_CASE_NOTE("---Out of State Inquiry received via " & how_sent & " from " & abbr_state & "---")
	IF out_of_state_status <> "Unknown" THEN CALL write_variable_in_CASE_NOTE("* " & abbr_state & " reported client received " & other_state_programs & " on " & date_received & " the case is currently: " & out_of_state_status)
	'IF other_state_programs = UNCHECKED THEN CALL write_variable_in_CASE_NOTE("* " & abbr_state & " reported client did not receive benefits in this state. ")
	IF out_of_state_status = "Client not known" THEN CALL write_variable_in_CASE_NOTE("* " & abbr_state & " reported client is not known.")
	CALL write_bullet_and_variable_in_CASE_NOTE("Name", agency_name)
	CALL write_bullet_and_variable_in_CASE_NOTE("Address", agency_address)
	CALL write_bullet_and_variable_in_CASE_NOTE("Email", agency_email)
	CALL write_bullet_and_variable_in_CASE_NOTE("Phone", agency_phone)
	CALL write_bullet_and_variable_in_CASE_NOTE("Fax", agency_fax)
	Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
	CALL write_variable_in_CASE_NOTE("* Updated MAXIS to reflect information received")
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3
END IF 'if for Received'

IF new_addr_detail_entered = TRUE THEN
	'PF11
	MsgBox "PF11 Time"
	'Problem.Reporting
	EMReadScreen nav_check, 4, 1, 27
	IF nav_check = "Prob" THEN
		EMWriteScreen  "Request to update directory of national directory of contacts " & programs_applied_for, 05, 07
		EMWriteScreen "Date: " & date, 06, 07
		IF agency_name <> "" THEN EMWriteScreen "Agency name update: " & agency_name, 07, 07
		IF agency_name <> "" THEN EMWriteScreen "Agency address update: " & agency_address, 08, 07
		IF agency_name <> "" THEN EMWriteScreen "Agency phone update: " & agency_phone, 09, 07
		IF agency_name <> "" THEN EMWriteScreen "Agency fax update: " & agency_fax, 10, 07
		IF agency_name <> "" THEN EMWriteScreen "Agency email update: " & agency_email, 11, 07
		EMWriteScreen "Worker number: X127" & worker_xnumber , 12, 07
	  msgbox "test"
	   TRANSMIT
	   EMReadScreen task_number, 7, 3, 27
	  msgbox task_number
	   TRANSMIT
	   'back_to_self
	   PF3 ''-self'
	   PF3 '- MEMB'
	ELSE
	   script_end_procedure_with_error_report("Could not reach PF11, request to update national directory of contacts has not been sent.")
	END IF
END IF 'end of PF11 action'

IF out_of_state_request = "Unknown/No Response" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 196, 150, "OUT OF STATE INQUIRY NO RESPONSE"
	  EditBox 50, 5, 40, 15, date_closed
	  CheckBox 105, 5, 90, 10, "No response from client", no_response_from_client_CHECKBOX
	  CheckBox 105, 15, 85, 10, "No response from state", no_response_from_state_CHECKBOX
	  CheckBox 10, 35, 85, 10, "Out of State Inquiry", Out_of_State_Inquiry_CHECKBOX
	  CheckBox 10, 45, 90, 10, "Authorization to release", ATR_Verf_CHECKBOX
	  CheckBox 105, 35, 70, 10, "Shelter verification", shel_verf_CHECKBOX
	  CheckBox 105, 45, 80, 10, "Other (please specify)", OTHER_CHECKBOX
	  CheckBox 5, 65, 90, 10, "Contacted other state(s)", other_state_contact_CHECKBOX
	  CheckBox 5, 80, 120, 10, "Unable to close(please explain)", unable_to_close_CHECKBOX
	  CheckBox 5, 95, 180, 10, "Overpayment possible to be reviewed at a later date", overpayment_CHECKBOX
	  EditBox 50, 110, 140, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 105, 130, 40, 15
	    CancelButton 150, 130, 40, 15
	  GroupBox 5, 25, 185, 35, "Verification Requested: "
	  Text 5, 10, 45, 10, "Date closed:"
	  Text 5, 115, 45, 10, "Other Notes:"
	EndDialog
    Do
    	Do
            err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF Isdate(date_closed) = "" and unable_to_close_CHECKBOX <> CHECKED THEN err_msg = err_msg & vbNewLine & "Please enter the closed date."
			IF unable_to_close_CHECKBOX = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please explain why you were unable to close the case."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
     	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
	'-------------------------------------------------------------------pending_verifs
    IF Out_of_State_Inquiry_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Out of State Inquiry, "
    IF shel_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Shelter Verification, "
    IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
    IF OTHER_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Other, "
    pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
    IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("---Out of State Inquiry for " & abbr_state & " no response received---")
	IF no_response_from_client_CHECKBOX = CHECKED THEN
		Call write_variable_in_CASE_NOTE("* No response from client.")
		Call write_variable_in_CASE_NOTE("* Client will need to verify MN residence.")
	END IF
	IF no_response_from_state_CHECKBOX = CHECKED THEN CALL write_variable_in_CASE_NOTE("* No response from state.")
	IF PARIS_CHECKBOX = CHECKED THEN Call write_variable_in_CASE_NOTE("* Agency will need to verify benefits received in the other state prior to reopening case")
	IF other_state_contact_checkbox = CHECKED THEN Call write_variable_in_CASE_NOTE("* " & state_droplist & "  has been contacted")
	IF other_state_contact_checkbox = UNCHECKED THEN Call write_variable_in_CASE_NOTE("* " & state_droplist & "  has not been contacted")
	Call write_bullet_and_variable_in_CASE_NOTE("Date case was closed", date_closed)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verification requested", pending_verifs)
	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3
	IF unable_to_close_CHECKBOX = CHECKED THEN call create_TIKL("Unable to close due to 10 day cutoff", 10, date, TRUE, "")
END IF 'if non received'
'todo do we need to look at memi and update for when received? In MN > 12 Months (Y/N):
'do you want to update the national directory?'
script_end_procedure_with_error_report("Success! Your Out of State Inquiry has been generated, please follow up with the next steps to ensure the request is received timely. The verification request must be reflected in ECF this can be done by saving the word document as a PDF and uploading to ECF.")
