'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - APPOINTMENT LETTER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN 'Shouldn't load FuncLib if it already loaded once
    IF run_locally = FALSE or run_locally = "" THEN     'If the scripts are set to run locally, it skips this and uses an FSO below.
        IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN            'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        ELSEIF beta_agency = "" or beta_agency = True then                          'If you're a beta agency, you should probably use the beta branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        Else                                                                        'Everyone else should use the release branch.
            FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
        End if
        SET req = CreateObject("Msxml2.XMLHttp.6.0")                'Creates an object to get a FuncLib_URL
        req.open "GET", FuncLib_URL, FALSE                          'Attempts to open the FuncLib_URL
        req.send                                                    'Sends request
        IF req.Status = 200 THEN                                    '200 means great success
            Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
            Execute req.responseText                                'Executes the script code
        ELSE                                                        'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
            MsgBox  "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
                    vbCr & _
                    "Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
                    vbCr & _
                    "If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
                    vbTab & "- The name of the script you are running." & vbCr &_
                    vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
                    vbTab & "- The name and email for an employee from your IT department," & vbCr & _
                    vbTab & vbTab & "responsible for network issues." & vbCr &_
                    vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
                    vbCr & _
                    "Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
                    vbCr &_
                    "URL: " & FuncLib_URL
                    script_end_procedure("Script ended due to error connecting to GitHub.")
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

'CLASSES----------------------------------------------------------------------------------------------------------------------
'IF THIS WORKS, CONSIDER INCORPORATING INTO FUNCTIONS LIBRARY

'The following defines a class called "address" which carries several simple address properties, which can be used by scripts.
class address
    public street           'Defines a "street" property.
    public city             'Defines a "city" property.
    public state            'Defines a "state" property.
    public zip              'Defines a "zip" property.

    'Creates a "oneline" property containing the entire address on a single line.
    public property get oneline
        oneline = street & ", " & city & ", " & state & " " & zip
    end property

    'Creates a "twolines" property containing the entire address on two lines, split into an array.
    public property get twolines
        twolines = array(street, city & ", " & state & " " & zip)
    end property
end class

'Declaring variables needed by the script
'First, determining the county code. If it isn't declared, it will ask (proxy)
call worker_county_code_determination(worker_county_code, two_digit_county_code_variable)

if worker_county_code = "x101" then 
    agency_office_array = array("Aitkin")
elseif worker_county_code = "x102" then 
    agency_office_array = array("Anoka", "Blaine", "Columbia Heights", "Lexington")
elseif worker_county_code = "x103" then 
    agency_office_array = array("Becker")
elseif worker_county_code = "x104" then 
    agency_office_array = array("Beltrami")
elseif worker_county_code = "x105" then 
    agency_office_array = array("Benton")
elseif worker_county_code = "x106" then 
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x107" then 
    agency_office_array = array("Blue Earth")
elseif worker_county_code = "x108" then 
    agency_office_array = array("New Ulm", "Sleepy Eye", "Springfield")
elseif worker_county_code = "x109" then 
    agency_office_array = array("Cloquet", "Moose Lake")
elseif worker_county_code = "x110" then 
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x111" then 
    agency_office_array = array("Cass")
elseif worker_county_code = "x112" then 
    agency_office_array = array("Chippewa")
elseif worker_county_code = "x113" then 
    agency_office_array = array("Center City", "North Branch")
elseif worker_county_code = "x114" then 
    agency_office_array = array("Clay")
elseif worker_county_code = "x115" then 
    agency_office_array = array("Clearwater")
elseif worker_county_code = "x116" then 
    agency_office_array = array("Cook")
elseif worker_county_code = "x117" then 
    agency_office_array = array("Cottonwood")
elseif worker_county_code = "x118" then 
    agency_office_array = array("Crow Wing")
elseif worker_county_code = "x119" then 
    agency_office_array = array("Dakota")
elseif worker_county_code = "x120" then 
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x121" then 
    agency_office_array = array("Douglas")
elseif worker_county_code = "x122" then 
    agency_office_array = array("Faribault")
elseif worker_county_code = "x123" then 
    agency_office_array = array("Fillmore")
elseif worker_county_code = "x124" then 
    agency_office_array = array("Freeborn")
elseif worker_county_code = "x125" then 
    agency_office_array = array("Goodhue")
elseif worker_county_code = "x126" then 
    agency_office_array = array("Grant")
elseif worker_county_code = "x127" then 
    agency_office_array = array("Century Plaza", "Northwest", "VEAP", "North Hub", "West Suburban Hub")
elseif worker_county_code = "x128" then 
    agency_office_array = array("Houston")
elseif worker_county_code = "x129" then 
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x130" then 
    agency_office_array = array("Isanti")
elseif worker_county_code = "x131" then 
    agency_office_array = array("Itasca")
elseif worker_county_code = "x132" then 
    agency_office_array = array("Jackson")
elseif worker_county_code = "x133" then 
    agency_office_array = array("Kanabec")
elseif worker_county_code = "x134" then
    agency_office_array = array("Kandiyohi")
elseif worker_county_code = "x135" then     
    agency_office_array = array("Kittson")
elseif worker_county_code = "x136" then     
    agency_office_array = array("Koochiching")
elseif worker_county_code = "x137" then     
    agency_office_array = array("Lac Qui Parle")
elseif worker_county_code = "x138" then     
    agency_office_array = array("Lake")
elseif worker_county_code = "x139" then     
    agency_office_array = array("Lake of the Woods")
elseif worker_county_code = "x140" then     
    agency_office_array = array("LeSueur")
elseif worker_county_code = "x141" then     
    agency_office_array = array("Lincoln")
elseif worker_county_code = "x142" then     
    agency_office_array = array("Lyon")
elseif worker_county_code = "x143" then     
    agency_office_array = array("Mcleod")
elseif worker_county_code = "x144" then     
    agency_office_array = array("Mahnomen")
elseif worker_county_code = "x145" then     
    agency_office_array = array("Marshall")
elseif worker_county_code = "x146" then     
    agency_office_array = array("Martin")
elseif worker_county_code = "x147" then     
    agency_office_array = array("Meeker")
elseif worker_county_code = "x148" then     
    agency_office_array = array("Mille Lacs")
elseif worker_county_code = "x149" then     
    agency_office_array = array("Morrison")
elseif worker_county_code = "x150" then     
    agency_office_array = array("Mower")
elseif worker_county_code = "x151" then     
    agency_office_array = array("Murray")
elseif worker_county_code = "x152" then     
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x153" then     
    agency_office_array = array("Nobles")
elseif worker_county_code = "x154" then     
    agency_office_array = array("Norman")
elseif worker_county_code = "x155" then     
    agency_office_array = array("Olmsted")
elseif worker_county_code = "x156" then     
    agency_office_array = array("Otter Tail")
elseif worker_county_code = "x157" then     
    agency_office_array = array("Pennington")
elseif worker_county_code = "x158" then     
    agency_office_array = array("Pine City", "Sandstone")
elseif worker_county_code = "x159" then     
    agency_office_array = array("Pipestone")
elseif worker_county_code = "x160" then     
    agency_office_array = array("Crookston", "Fosston")
elseif worker_county_code = "x161" then     
    agency_office_array = array("Pope")
elseif worker_county_code = "x162" then     
    agency_office_array = array("Ramsey")
elseif worker_county_code = "x163" then     
    agency_office_array = array("Red Lake")
elseif worker_county_code = "x164" then     
    agency_office_array = array("Redwood")
elseif worker_county_code = "x165" then     
    agency_office_array = array("Renville")
elseif worker_county_code = "x166" then     
    agency_office_array = array("Rice")
elseif worker_county_code = "x167" then     
    agency_office_array = array("Rock")
elseif worker_county_code = "x168" then     
    agency_office_array = array("Roseau")
elseif worker_county_code = "x169" then     
    agency_office_array = array("Duluth", "Virginia", "Hibbing", "Ely")
elseif worker_county_code = "x170" then     
    agency_office_array = array("Scott")
elseif worker_county_code = "x171" then     
    agency_office_array = array("Sherburne")
elseif worker_county_code = "x172" then     
    agency_office_array = array("Sibley")
elseif worker_county_code = "x173" then     
    agency_office_array = array("St. Cloud", "Melrose")
elseif worker_county_code = "x174" then     
    agency_office_array = array("Dodge", "Steele", "Waseca")
elseif worker_county_code = "x175" then     
    agency_office_array = array("Stevens")
elseif worker_county_code = "x176" then     
    agency_office_array = array("Swift")
elseif worker_county_code = "x177" then     
    agency_office_array = array("Long Prairie", "Staples")
elseif worker_county_code = "x178" then     
    agency_office_array = array("Traverse")
elseif worker_county_code = "x179" then     
    agency_office_array = array("Wabasha")
elseif worker_county_code = "x180" then     
    agency_office_array = array("Wadena")
elseif worker_county_code = "x181" then     
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x182" then     
    agency_office_array = array("Cottage Grove", "Forest Lake", "Stillwater", "Woodbury")
elseif worker_county_code = "x183" then     
    agency_office_array = array("Watonwan")
elseif worker_county_code = "x184" then     
    agency_office_array = array("Wilkin")
elseif worker_county_code = "x185" then     
    agency_office_array = array("Winona")
elseif worker_county_code = "x186" then     
    agency_office_array = array("Wright")
elseif worker_county_code = "x187" then     
    agency_office_array = array("Yellow Medicine")
elseif worker_county_code = "x188" then 
    script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
elseif worker_county_code = "x192" then 
    agency_office_array = array("Detroit Lakes", "Naytahwaush", "Bagley", "Mahnomen")
end if






county_office_list = ""     'Blanking this out because it may contain old info from the old global variables (from before this was integrated in this script)

call convert_array_to_droplist_items(agency_office_array, county_office_list)

'DIALOGS----------------------------------------------------------------------------------------------------
'NOTE: this dialog contains a special modification to allow dynamic creation of the county office list. You cannot edit it in
'   Dialog Editor without modifying the commented line.
BeginDialog appointment_letter_dialog, 0, 0, 156, 355, "Appointment letter"
  EditBox 75, 5, 50, 15, case_number
  DropListBox 50, 25, 95, 15, "new application"+chr(9)+"recertification", app_type
  CheckBox 10, 43, 150, 10, "Check here if this is a reschedule.", reschedule_check
  EditBox 50, 55, 95, 15, CAF_date
  CheckBox 10, 75, 130, 10, "Check here if this is a recert and the", no_CAF_check
  DropListBox 70, 100, 75, 15, "Select one..."+chr(9)+"PHONE"+chr(9)+county_office_list, interview_location     'This line dynamically creates itself based on the information in the FUNCTIONS FILE.
  EditBox 70, 120, 75, 15, interview_date
  EditBox 70, 140, 75, 15, interview_time
  EditBox 80, 160, 65, 15, client_phone
  CheckBox 10, 200, 95, 10, "Client appears expedited", expedited_check
  CheckBox 10, 215, 135, 10, "Same day interview offered/declined", same_day_declined_check
  EditBox 10, 250, 135, 15, expedited_explanation
  CheckBox 10, 280, 135, 10, "Check here if you left V/M with client", voicemail_check
  EditBox 85, 305, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 325, 50, 15
    CancelButton 85, 325, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 15, 30, 30, 10, "App type:"
  Text 15, 60, 35, 10, "CAF date:"
  Text 30, 85, 105, 10, "CAF hasn't been received yet."
  Text 15, 105, 55, 10, "Int'vw location:"
  Text 15, 125, 50, 10, "Interview date: "
  Text 15, 145, 50, 10, "Interview time:"
  Text 15, 160, 60, 20, "Client phone (if phone interview):"
  GroupBox 5, 185, 145, 85, "Expedited questions"
  Text 10, 230, 135, 20, "If expedited interview date is not within six days of the application, explain:"
  Text 45, 290, 75, 10, "requesting a call back."
  Text 15, 310, 65, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)


'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
Do
    Do
        Do
            Do
                Do
                    Do
                        Do
                            Do
                                Dialog appointment_letter_dialog
                                If ButtonPressed = cancel then stopscript
                                If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must fill in a valid case number. Please try again."
                            Loop until isnumeric(case_number) = True and len(case_number) <= 8 
                            CAF_date = replace(CAF_date, ".", "/")
                            If no_CAF_check = checked and app_type = "new application" then no_CAF_check = unchecked 'Shuts down "no_CAF_check" so that it will validate the date entered. New applications can't happen if a CAF wasn't provided.
                            If no_CAF_check = unchecked and isdate(CAF_date) = False then Msgbox "You did not enter a valid CAF date (MM/DD/YYYY format). Please try again."
                        Loop until no_CAF_check = checked or isdate(CAF_date) = True
                        if interview_location = "Select one..." then MsgBox "You must select an interview location! Please try again!"
                    Loop until interview_location <> "Select one..."
                    if interview_location = "PHONE" and client_phone = "" then MsgBox "If this is a phone interview, you must enter a phone number! Please try again."
                Loop until interview_location <> "PHONE" or (interview_location = "PHONE" and client_phone <> "")
                interview_date = replace(interview_date, ".", "/")
                If isdate(interview_date) = False then MsgBox "You did not enter a valid interview date (MM/DD/YYYY format). Please try again."
            Loop until isdate(interview_date) = True 
            If interview_time = "" then MsgBox "You must type an interview time. Please try again."
        Loop until interview_time <> ""
        If no_CAF_check = checked then exit do 'If no CAF was turned in, this layer of validation is unnecessary, so the script will skip it.
        If expedited_check = checked and datediff("d", CAF_date, interview_date) > 6 and expedited_explanation = "" then MsgBox "You have indicated that your case is expedited, but scheduled the interview date outside of the six-day window. An explanation of why is required for QC purposes."
    Loop until expedited_check = unchecked or (datediff("d", CAF_date, interview_date) <= 6) or (datediff("d", CAF_date, interview_date) > 6 and expedited_explanation <> "")
    If worker_signature = "" then MsgBox "You must provide a signature for your case note."
Loop until worker_signature <> ""

'Creates a variable to contain the agency addresses. "Address" is a class defined above.
set agency_address = new address

'As these are all MN intake locations, the state for all of them will be MN.
agency_address.state = "MN"

'Determines the address properties based on the county and interview_location dropdown
IF worker_county_code = "x101" THEN
    agency_address.street = "204 1st St NW" 
    agency_address.city = "Aitkin" 
    agency_address.zip = "56431" 
ELSEIF worker_county_code = "x102" THEN
    IF interview_location = "Anoka" THEN
        agency_address.street = "2100 3rd Ave, Suite 400" 
        agency_address.city = "Anoka" 
        agency_address.zip = "55303" 
    ELSEIF interview_location = "Blaine" THEN
        agency_address.street = "1201 89th Ave, Suite 400" 
        agency_address.city = "Blaine" 
        agency_address.zip = "55434" 
    ELSEIF interview_location = "Columbia Heights" THEN
        agency_address.street = "3980 Central Ave NE" 
        agency_address.city = "Columbia Heights" 
        agency_address.zip = "55421" 
    ELSEIF interview_location = "Lexington" THEN
        agency_address.street = "4175 Lovell RD NE" 
        agency_address.city = "Lexington" 
        agency_address.zip = "55014" 
    END IF 
ELSEIF worker_county_code = "x103" THEN
    agency_address.street = "712 Minnesota Ave " 
    agency_address.city = "Detroit Lakes" 
    agency_address.zip = "56501" 
ELSEIF worker_county_code = "x104" THEN
    agency_address.street = "616 America Ave NW, STE 270" 
    agency_address.city = "Bemidji" 
    agency_address.zip = "56601" 
ELSEIF worker_county_code = "x105" THEN
    agency_address.street = "531 Dewey St" 
    agency_address.city = "Foley" 
    agency_address.zip = "56329" 
ELSEIF worker_county_code = "x107" THEN
    agency_address.street = "410 S 5Th Street" 
    agency_address.city = "Mankato" 
    agency_address.zip = "56001" 
ELSEIF worker_county_code = "x108" THEN
    IF interview_location = "New Ulm" THEN
        agency_address.street = "1117 Center ST" 
        agency_address.city = "New Ulm" 
        agency_address.zip = "56073" 
    ELSEIF interview_location = "Sleepy Eye" THEN
        agency_address.street = "300 2nd Ave SW" 
        agency_address.city = "Sleepy Eye" 
        agency_address.zip = "56085" 
    ELSEIF interview_location = "Springfield" THEN
        agency_address.street = "33 N Cass Ave" 
        agency_address.city = "Springfield" 
        agency_address.zip = "56087" 
    END IF 
ELSEIF worker_county_code = "x109" THEN
    IF interview_location = "Cloquet" THEN
        agency_address.street = "14 N 11th St" 
        agency_address.city = "Cloquet" 
        agency_address.zip = "55720" 
    ELSEIF interview_location = "Moose Lake" THEN
        agency_address.street = "316 Elm Ave" 
        agency_address.city = "Moose Lake" 
        agency_address.zip = "55767" 
    END IF 
ELSEIF worker_county_code = "x111" THEN
    agency_address.street = "400 Michigan Ave" 
    agency_address.city = "Walker" 
    agency_address.zip = "56484" 
ELSEIF worker_county_code = "x112" THEN
    agency_address.street = "719 N 7th St Ste 200" 
    agency_address.city = "Montevideo" 
    agency_address.zip = "56265" 
ELSEIF worker_county_code = "x113" THEN
    IF interview_location = "Center City" THEN
        agency_address.street = "313 North Main St – Room 239" 
        agency_address.city = "Center City" 
        agency_address.zip = "55012" 
    ELSEIF interview_location = "North Branch" THEN
        agency_address.street = "6133 402nd Street" 
        agency_address.city = "North Branch" 
        agency_address.zip = "55056" 
    END IF 
ELSEIF worker_county_code = "x114" THEN
    agency_address.street = "715 11th St North #102" 
    agency_address.city = "Moorhead" 
    agency_address.zip = "56560" 
ELSEIF worker_county_code = "x115" THEN
    agency_address.street = "216 Park Ave NW" 
    agency_address.city = "Bagley" 
    agency_address.zip = "56621" 
ELSEIF worker_county_code = "x116" THEN
	agency_address.street = "411 W. 2nd St"
	agency_address.city = "Grand Marais"
	agency_address.zip = "55604"
ELSEIF worker_county_code = "x117" THEN
    agency_address.street = "11 4th St" 
    agency_address.city = "Windom" 
    agency_address.zip = "56101" 
ELSEIF worker_county_code = "x118" THEN
    agency_address.street = "204 Laurel St." 
    agency_address.city = "Brainerd" 
    agency_address.zip = "56401" 
ELSEIF worker_county_code = "x119" THEN
    agency_address.street = "1 Mendota Rd W Ste 100" 
    agency_address.city = "West St Paul" 
    agency_address.zip = "55118" 
ELSEIF worker_county_code = "x121" THEN
    agency_address.street = "809  Elm Street, Ste 1186" 
    agency_address.city = "Alexandria" 
    agency_address.zip = "56308" 
ELSEIF worker_county_code = "x122" THEN
    agency_address.street = "412 N. Nicollet Street" 
    agency_address.city = "Blue Earth" 
    agency_address.zip = "56013" 
ELSEIF worker_county_code = "x123" THEN
    agency_address.street = "902 Houston St NW, Suite 1" 
    agency_address.city = "Preston" 
    agency_address.zip = "55965" 
ELSEIF worker_county_code = "x124" THEN
    agency_address.street = "203 W. Clark Street" 
    agency_address.city = "Albert Lea" 
    agency_address.zip = "56007" 
ELSEIF worker_county_code = "x125" THEN
    agency_address.street = "469 12th Street" 
    agency_address.city = "Red Wing" 
    agency_address.zip = "55066" 
ELSEIF worker_county_code = "x126" THEN
    agency_address.street = "28 Central Ave S" 
    agency_address.city = "Elbow Lake" 
    agency_address.zip = "56531" 
ELSEIF worker_county_code = "x127" THEN
    IF interview_location = "Century Plaza" THEN
        agency_address.street = "330 South 12th Street" 
        agency_address.city = "Minneapolis" 
        agency_address.zip = "55440" 
    ELSEIF interview_location = "Northwest" THEN
        agency_address.street = "7051 Brooklyn Blvd" 
        agency_address.city = "Brooklyn Center" 
        agency_address.zip = "55429" 
    ELSEIF interview_location = "VEAP" THEN
        agency_address.street = "9600 Aldrich Ave" 
        agency_address.city = "Bloomington" 
        agency_address.zip = "55420" 
    ELSEIF interview_location = "North Hub" THEN
        agency_address.street = "1001 Plymouth Ave North" 
        agency_address.city = "Minneapolis" 
        agency_address.zip = "55411" 
    ELSEIF interview_location = "West Suburban Hub" THEN
        agency_address.street = "1011 First Street South, Suite 108" 
        agency_address.city = "Hopkins" 
        agency_address.zip = "55343" 
    END IF 
ELSEIF worker_county_code = "x128" THEN
    agency_address.street = "304 S Marshall St., Room 104" 
    agency_address.city = "Caledonia" 
    agency_address.zip = "55921" 
ELSEIF worker_county_code = "x130" THEN
    agency_address.street = "1700 East Rum River Dr. S. Ste. A" 
    agency_address.city = "Cambridge" 
    agency_address.zip = "55008" 
ELSEIF worker_county_code = "x131" THEN
    agency_address.street = "1209 SE 2nd Ave" 
    agency_address.city = "Grand Rapids" 
    agency_address.zip = "55744" 
ELSEIF worker_county_code = "x132" THEN
    agency_address.street = "407 Fifth St" 
    agency_address.city = "Jackson" 
    agency_address.zip = "56143" 
ELSEIF worker_county_code = "x133" THEN
    agency_address.street = "905 Forest Ave E, Suite 150" 
    agency_address.city = "Mora" 
    agency_address.zip = "55051" 
ELSEIF worker_county_code = "x134" THEN
    agency_address.street = "2200 23rd St NE, Suite 1020" 
    agency_address.city = "Willmar" 
    agency_address.zip = "56201" 
ELSEIF worker_county_code = "x135" THEN
    agency_address.street = " 410 5th Street S Suite 100" 
    agency_address.city = "Hallock" 
    agency_address.zip = "56728" 
ELSEIF worker_county_code = "x136" THEN
    agency_address.street = "1000 5th Street" 
    agency_address.city = "Intl Falls" 
    agency_address.zip = "56649" 
ELSEIF worker_county_code = "x137" THEN
    agency_address.street = "930 1st Ave " 
    agency_address.city = "Madison" 
    agency_address.zip = "56256" 
ELSEIF worker_county_code = "x138" THEN
    agency_address.street = "616 3rd Ave" 
    agency_address.city = "Two Harbors" 
    agency_address.zip = "55616" 
ELSEIF worker_county_code = "x139" THEN
    agency_address.street = "206 8th Ave SE Suite 200" 
    agency_address.city = "Baudette" 
    agency_address.zip = "56623" 
ELSEIF worker_county_code = "x140" THEN
    agency_address.street = "88 S Park Ave" 
    agency_address.city = "Le Center" 
    agency_address.zip = "56057" 
ELSEIF worker_county_code = "x141" THEN
    agency_address.street = "319 N Rebecca St" 
    agency_address.city = "Ivanhoe" 
    agency_address.zip = "56142" 
ELSEIF worker_county_code = "x142" THEN
    agency_address.street = "607 W Main St, Suite 200" 
    agency_address.city = "Marshall" 
    agency_address.zip = "56258" 
ELSEIF worker_county_code = "x143" THEN
    agency_address.street = "1805 Ford Avenue North, Suite 100" 
    agency_address.city = "Glencoe" 
    agency_address.zip = "55336" 
ELSEIF worker_county_code = "x144" THEN
    agency_address.street = "311 N Main St" 
    agency_address.city = "Mahnomen" 
    agency_address.zip = "56557" 
ELSEIF worker_county_code = "x145" THEN
    agency_address.street = "208 E Colvin Ave Ste 14" 
    agency_address.city = "Warren" 
    agency_address.zip = "56762" 
ELSEIF worker_county_code = "x146" THEN
    agency_address.street = "115 W 1st Street" 
    agency_address.city = "Fairmont" 
    agency_address.zip = "56031" 
ELSEIF worker_county_code = "x147" THEN
    agency_address.street = "114 N Holcombe Ave;  Ste 180" 
    agency_address.city = "LItchfield" 
    agency_address.zip = "55355" 
ELSEIF worker_county_code = "x148" THEN
    agency_address.street = "525 2nd Street SE" 
    agency_address.city = "Milaca" 
    agency_address.zip = "56353" 
ELSEIF worker_county_code = "x149" THEN
    agency_address.street = "213 SE 1st Ave" 
    agency_address.city = "Little Falls" 
    agency_address.zip = "56345" 
ELSEIF worker_county_code = "x150" THEN
    agency_address.street = "201 1st St NE Suite 18" 
    agency_address.city = "Austin" 
    agency_address.zip = "55912" 
ELSEIF worker_county_code = "x151" THEN
    agency_address.street = "3001 Maple Road, Suite 100" 
    agency_address.city = "Slayton" 
    agency_address.zip = "56172" 
ELSEIF worker_county_code = "x153" THEN
    agency_address.street = "318 9th St." 
    agency_address.city = "Worthington" 
    agency_address.zip = "56187" 
ELSEIF worker_county_code = "x154" THEN
    agency_address.street = "15 2nd Ave E" 
    agency_address.city = "Ada" 
    agency_address.zip = "56510" 
ELSEIF worker_county_code = "x155" THEN
    agency_address.street = "2117 Campus Dr SE  Suite 100" 
    agency_address.city = "Rochester" 
    agency_address.zip = "55904" 
ELSEIF worker_county_code = "x156" THEN
    agency_address.street = "535 West Fir Avenue" 
    agency_address.city = "Fergus Falls" 
    agency_address.zip = "56537" 
ELSEIF worker_county_code = "x157" THEN
    agency_address.street = "318 Knight Ave N" 
    agency_address.city = "Thief River Falls" 
    agency_address.zip = "56701" 
ELSEIF worker_county_code = "x158" THEN
    IF interview_location = "Pine City" THEN
        agency_address.street = "315 Main St S, Suite 200" 
        agency_address.city = "Pine City" 
        agency_address.zip = "55063" 
    ELSEIF interview_location = "Sandstone" THEN
        agency_address.street = "1610 Highway 23 N" 
        agency_address.city = "Sandstone" 
        agency_address.zip = "55072" 
    END IF 
ELSEIF worker_county_code = "x159" THEN
    agency_address.street = "1091 N Hiawatha Ave" 
    agency_address.city = "Pipestone" 
    agency_address.zip = "56164" 
ELSEIF worker_county_code = "x160" THEN
    IF interview_location = "Crookston" THEN
        agency_address.street = "612 N Broadway, Rm 302" 
        agency_address.city = "Crookston" 
        agency_address.zip = "56716" 
    ELSEIF interview_location = "Fosston" THEN
        agency_address.street = "104 N Kaiser Ave" 
        agency_address.city = "Fosston" 
        agency_address.zip = "56542" 
    END IF 
ELSEIF worker_county_code = "x161" THEN
    agency_address.street = "211 E Minnesota Ave, Suite 200" 
    agency_address.city = "Glenwood" 
    agency_address.zip = "56334" 
ELSEIF worker_county_code = "x162" THEN
    agency_address.street = "160 Kellogg Blvd. E." 
    agency_address.city = "Saint Paul" 
    agency_address.zip = "55101" 
ELSEIF worker_county_code = "x163" THEN
    agency_address.street = "125 Edward Ave" 
    agency_address.city = "Red Lake Falls" 
    agency_address.zip = "56750" 
ELSEIF worker_county_code = "x164" THEN
    agency_address.street = "266 E Bridge St" 
    agency_address.city = "Redwood Falls" 
    agency_address.zip = "56283" 
ELSEIF worker_county_code = "x165" THEN
    agency_address.street = "105 S. 5th St, Suite 203H" 
    agency_address.city = "Olivia" 
    agency_address.zip = "56277" 
ELSEIF worker_county_code = "x166" THEN
    agency_address.street = "320 3rd St N.W." 
    agency_address.city = "Faribault" 
    agency_address.zip = " 55021" 
ELSEIF worker_county_code = "x167" THEN
    agency_address.street = "2 Roundwind Road" 
    agency_address.city = "Luverne" 
    agency_address.zip = "56156" 
ELSEIF worker_county_code = "x168" THEN
    agency_address.street = "208 6th St SW" 
    agency_address.city = "Roseau" 
    agency_address.zip = "56751" 
ELSEIF worker_county_code = "x169" THEN
    IF interview_location = "Duluth" THEN
        agency_address.street = "320 W 2nd St " 
        agency_address.city = "Duluth" 
        agency_address.zip = "55802" 
    ELSEIF interview_location = "Virginia" THEN
        agency_address.street = "307 1st St" 
        agency_address.city = "Virginia" 
        agency_address.zip = "55792" 
    ELSEIF interview_location = "Hibbing" THEN
        agency_address.street = "1814 E 14th Ave" 
        agency_address.city = "Hibbing" 
        agency_address.zip = "55746" 
    ELSEIF interview_location = "Ely" THEN
        agency_address.street = "320 Miners Dr" 
        agency_address.city = "Ely" 
        agency_address.zip = "55731" 
    END IF 
ELSEIF worker_county_code = "x170" THEN
    agency_address.street = "752 Canterbury Rd S" 
    agency_address.city = "Shakopee" 
    agency_address.zip = "55379" 
ELSEIF worker_county_code = "x171" THEN
    agency_address.street = "13880 Business Center Drive NW" 
    agency_address.city = "Elk River" 
    agency_address.zip = "55330" 
ELSEIF worker_county_code = "x172" THEN
    agency_address.street = "111 8th Street" 
    agency_address.city = "Gaylord" 
    agency_address.zip = "55334" 
ELSEIF worker_county_code = "x173" THEN
    IF interview_location = "St. Cloud" THEN
        agency_address.street = "705 Courthouse Square" 
        agency_address.city = "St. Cloud" 
        agency_address.zip = "56302" 
    ELSEIF interview_location = "Melrose" THEN
        agency_address.street = "114 1st Avenue West" 
        agency_address.city = "Melrose" 
        agency_address.zip = "56352" 
    END IF 
ELSEIF worker_county_code = "x174" THEN
    IF interview_location = "Dodge" THEN
        agency_address.street = "22 6TH ST East Dept 401" 
        agency_address.city = "Mantorville" 
        agency_address.zip = "55955" 
    ELSEIF interview_location = "Steele" THEN
        agency_address.street = "630 FLORENCE AVE" 
        agency_address.city = "Owatonna" 
        agency_address.zip = "55060" 
    ELSEIF interview_location = "Waseca" THEN
        agency_address.street = "299 JOHNSON SW STE 160" 
        agency_address.city = "Waseca" 
        agency_address.zip = "56093" 
    END IF 
ELSEIF worker_county_code = "x175" THEN
    agency_address.street = "400 Colorado Ave., Suite 104" 
    agency_address.city = "Morris" 
    agency_address.zip = "56267" 
ELSEIF worker_county_code = "x176" THEN
    agency_address.street = "410 21st St S" 
    agency_address.city = "Benson" 
    agency_address.zip = "56215" 
ELSEIF worker_county_code = "x177" THEN
    IF interview_location = "Long Prairie" THEN
        agency_address.street = "212 2nd Ave S" 
        agency_address.city = "Long Prairie" 
        agency_address.zip = "56347" 
    ELSEIF interview_location = "Staples" THEN
        agency_address.street = "200-1st ST NE Suite 1" 
        agency_address.city = "Staples" 
        agency_address.zip = "56479" 
    END IF 
ELSEIF worker_county_code = "x178" THEN
    agency_address.street = "202 8th Street north" 
    agency_address.city = "Wheaton" 
    agency_address.zip = "56296" 
ELSEIF worker_county_code = "x179" THEN
    agency_address.street = "411 Hiawatha Drive East" 
    agency_address.city = "Wabasha" 
    agency_address.zip = "55981" 
ELSEIF worker_county_code = "x180" THEN
    agency_address.street = "124 First Street SE" 
    agency_address.city = "Wadena" 
    agency_address.zip = "56482" 
ELSEIF worker_county_code = "x182" THEN
    IF interview_location = "Cottage Grove" THEN
        agency_address.street = "13000 Ravine Parkway S" 
        agency_address.city = "Cottage Grove" 
        agency_address.zip = "55016" 
    ELSEIF interview_location = "Forest Lake" THEN
        agency_address.street = "19955 Forest Rd N" 
        agency_address.city = "Forest Lake" 
        agency_address.zip = "55025" 
    ELSEIF interview_location = "Stillwater" THEN
        agency_address.street = "14949 62nd St N" 
        agency_address.city = "Stillwater" 
        agency_address.zip = "55082" 
    ELSEIF interview_location = "Woodbury" THEN
        agency_address.street = "2150 Radio Dr" 
        agency_address.city = "Woodbury" 
        agency_address.zip = "55125" 
    END IF
ELSEIF worker_county_code = "x183" THEN
    agency_address.street = "715 2nd Ave S" 
    agency_address.city = "St. James" 
    agency_address.zip = "56081" 
ELSEIF worker_county_code = "x184" THEN
    agency_address.street = "300 S 5th St" 
    agency_address.city = "Breckenridge" 
    agency_address.zip = "56520" 
ELSEIF worker_county_code = "x185" THEN
    agency_address.street = "202 W 3rd Street" 
    agency_address.city = "Winona" 
    agency_address.zip = "55987" 
ELSEIF worker_county_code = "x186" THEN
    agency_address.street = "10 2nd St NW Room 300" 
    agency_address.city = "Buffalo" 
    agency_address.zip = "55313" 
ELSEIF worker_county_code = "x187" THEN
    agency_address.street = "930 4th Street, Suite 4" 
    agency_address.city = "Granite Falls" 
    agency_address.zip = "56241" 
ELSEIF worker_county_code = "x192" THEN
    IF interview_location = "Detroit Lakes" THEN
        agency_address.street = "210 West State Street" 
        agency_address.city = "Detroit Lakes" 
        agency_address.zip = "56501" 
    ELSEIF interview_location = "Naytahwaush" THEN
        agency_address.street = "2531 310th Avenue" 
        agency_address.city = "Naytahwaush" 
        agency_address.zip = "56566" 
    ELSEIF interview_location = "Bagley" THEN
        agency_address.street = "107 Central Street" 
        agency_address.city = "Bagley" 
        agency_address.zip = "56621" 
    ELSEIF interview_location = "Mahnomen" THEN
        agency_address.street = "2235 College Rd Suite Suite 200" 
        agency_address.city = "Mahnomen" 
        agency_address.zip = "56557" 
    END IF
END IF

'This is a temporary MsgBox that expires 09/01/2015. It is designed to "make sure" that the address is correct. Because this function is new, I want to be ABSOLUTELY SURE it's working before notices get sent out.
If interview_location <> "PHONE" and datediff("d", date, #9/1/2015#) > 0 then
	double_check_MsgBox = MsgBox("Please confirm your chosen office address: " & interview_location & " Office, " & agency_address.oneline & vbNewLine & vbNewLine & "Press OK to continue, or cancel to end the script." & vbNewLine & vbNewLine & "If this info is incorrect, have an alpha user contact Veronica Cary immediately with the correct address.", vbOKCancel)
	If double_check_MsgBox = vbCancel then stopscript
End if

'Checks for MAXIS
call check_for_MAXIS(False)

'Using custom function to assign addresses to the selected office
call assign_county_address_variables(county_address_line_01, county_address_line_02)

'Converting the CAF_date variable to a date, for cases where a CAF was turned in
If no_CAF_check = unchecked then CAF_date = cdate(CAF_date)

'Figuring out the last contact day
If app_type = "recertification" then
    next_month = datepart("m", dateadd("m", 1, interview_date))
    next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
    last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
End if
If app_type = "new application" then last_contact_day = CAF_date + 30
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date 

'This checks to make sure the case is not in background.
MAXIS_background_check

'Navigating to SPEC/MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
    PF5                                                     'PF5s again to initiate the new memo process
END IF
'Checking for SWKR
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
    PF5                                           'PF5s again to initiate the new memo process
END IF
EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
transmit                                                        'Transmits to start the memo writing process

'Writes the MEMO.
call write_variable_in_SPEC_MEMO("***********************************************************")
IF app_type = "new application" then
    call write_variable_in_SPEC_MEMO("You recently applied for assistance in " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
Elseif app_type = "recertification" then
    If no_CAF_check = unchecked then 
        call write_variable_in_SPEC_MEMO("You sent recertification paperwork to " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
    Else
        call write_variable_in_SPEC_MEMO("You asked us to set up an interview for your recertification. Remember to send in your forms before the interview.")
    End if
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "PHONE" then    'Phone interviews have a different verbiage than any other interview type
    call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " at " & interview_time & ".")
Else
    call write_variable_in_SPEC_MEMO("Your in-office interview is scheduled for " & interview_date & " at " & interview_time & ".")
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "PHONE" then
    call write_variable_in_SPEC_MEMO("We will be calling you at this number: " & client_phone & ".") 
    call write_variable_in_SPEC_MEMO("")
    call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview in the office, please call your worker.")
Else
    call write_variable_in_SPEC_MEMO("Your interview is at the " & interview_location & " Office, located at:")
    for each line in agency_address.twolines		'"twolines" is an array, so this will write each line in.
		call write_variable_in_SPEC_MEMO("   " & line)
    next
    call write_variable_in_SPEC_MEMO("")
    call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number.")
End if
call write_variable_in_SPEC_MEMO("")
IF app_type = "new application" then            '"deny your application" vs "your case will auto-close"
    call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " we will deny your application.")
Elseif app_type = "recertification" then
    call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & ", your case will auto-close.")
END IF
call write_variable_in_SPEC_MEMO("***********************************************************")

'Exits the MEMO
PF4

'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note--------------------------------------------

'If it's rescheduled, that header should engage. Otherwise, it uses separate headers for new apps and recerts.
If reschedule_check = checked then
    call write_variable_in_CASE_NOTE("**Client requested rescheduled appointment, appt letter sent in MEMO.**")
ElseIf app_type = "new application" then 
    call write_variable_in_CASE_NOTE("**New CAF received " & CAF_date & ", appt letter sent in MEMO**")
ElseIf app_type = "recertification" then 
    If no_CAF_check = unchecked then        'Uses separate headers for whether-or-not a CAF was received.
        call write_variable_in_CASE_NOTE("**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**")
    Else
        call write_variable_in_CASE_NOTE("**Client requested recert appointment, letter sent in MEMO**")    
    End if
End if

'And the rest...
If same_day_declined_check = checked then write_variable_in_CASE_NOTE("* Same day interview offered and declined.")
call write_bullet_and_variable_in_CASE_NOTE("Appointment date", interview_date) 
call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time) 
call write_bullet_and_variable_in_CASE_NOTE("Why interview is more than six days from now", expedited_explanation)
IF interview_location = "PHONE" then
	call write_variable_in_CASE_NOTE("* Interview will take place by telephone.")
Else
	call write_bullet_and_variable_in_CASE_NOTE("Appointment location", interview_location)
End if
call write_bullet_and_variable_in_CASE_NOTE("Client phone", client_phone)
call write_variable_in_CASE_NOTE("* Client must complete interview by " & last_contact_day & ".")
If voicemail_check = checked then call write_variable_in_CASE_NOTE("* Left client a voicemail requesting a call back.")
If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")              'Defined above
If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")     'Defined above
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
