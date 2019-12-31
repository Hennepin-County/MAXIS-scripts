
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 451, 305, "CAF dialog part 2"        'SEPERATE THIS OUT MORE TO CREATE A BETTER VISUAL FLOW FOR THE DIALOG
  EditBox 65, 35, 375, 15, notes_on_wreg
  Text 65, 60, 375, 10, notes_on_abawd       'Make this a text box
  EditBox 65, 75, 375, 15, notes_on_shel
  DropListBox 65, 95, 100, 15, "Select ALLOWED HEST"+chr(9)+"AC/Heat - Full $493"+chr(9)+"Electric and Phone - $173"+chr(9)+"Electric ONLY - $126"+chr(9)+"Phone ONLY - $47"+chr(9)+"NONE - $0", notes_on_hest        'Make this a dropdown/checkbox or something for HEST then add an ACUT box in the case of DWP
  EditBox 65, 115, 375, 15, notes_on_coex
  EditBox 65, 135, 375, 15, notes_on_dcex
  EditBox 65, 155, 260, 15, notes_on_other_deductions
  EditBox 370, 155, 70, 15, notes_on_cash
  EditBox 65, 175, 375, 15, notes_on_acct
  EditBox 65, 195, 375, 15, notes_on_cars
  EditBox 65, 215, 375, 15, notes_on_rest
  EditBox 105, 235, 335, 15, other_assets
  EditBox 55, 265, 385, 15, verifs_needed
  ButtonGroup ButtonPressed
    PushButton 270, 290, 60, 10, "previous page", previous_to_page_02_button
    PushButton 335, 285, 50, 15, "NEXT", next_to_page_04_button
    CancelButton 390, 285, 50, 15
    PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
    PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
    PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
    PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
    PushButton 240, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 290, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 345, 15, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "next memb", next_memb_button
    PushButton 20, 40, 35, 10, "WREG:", WREG_button
    PushButton 20, 60, 35, 10, "ABAWD:", ABAWD_button   'Make this button open a dashboard to detail ABAWD information for client(s)
    PushButton 25, 80, 25, 10, "SHEL:", SHEL_button
    PushButton 25, 100, 25, 10, "HEST:", HEST_button
    PushButton 25, 120, 25, 10, "COEX:", COEX_button
    PushButton 25, 140, 25, 10, "DCEX:", DCEX_button
    PushButton 340, 160, 25, 10, "CASH:", CASH_button
    PushButton 25, 180, 30, 10, "ACCTs:", ACCT_button
    PushButton 30, 200, 25, 10, "CARS:", CARS_button
    PushButton 30, 220, 25, 10, "REST:", REST_button
    PushButton 5, 240, 25, 10, "SECU/", SECU_button
    PushButton 30, 240, 25, 10, "TRAN/", TRAN_button
    PushButton 55, 240, 45, 10, "other assets:", OTHR_button
  GroupBox 235, 5, 205, 25, "STAT-based navigation"
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  Text 5, 160, 60, 10, "Other Deductions:"
  Text 5, 270, 50, 10, "Verifs needed:"
EndDialog

county_list = "01 Aitkin"
county_list = county_list+chr(9)+"02 Anoka"
county_list = county_list+chr(9)+"03 Becker"
county_list = county_list+chr(9)+"04 Beltrami"
county_list = county_list+chr(9)+"05 Benton"
county_list = county_list+chr(9)+"06 Big Stone"
county_list = county_list+chr(9)+"07 Blue Earth"
county_list = county_list+chr(9)+"08 Brown"
county_list = county_list+chr(9)+"09 Carlton"
county_list = county_list+chr(9)+"10 Carver"
county_list = county_list+chr(9)+"11 Cass"
county_list = county_list+chr(9)+"12 Chippewa"
county_list = county_list+chr(9)+"13 Chisago"
county_list = county_list+chr(9)+"14 Clay"
county_list = county_list+chr(9)+"15 Clearwater"
county_list = county_list+chr(9)+"16 Cook"
county_list = county_list+chr(9)+"17 Cottonwood"
county_list = county_list+chr(9)+"18 Crow Wing"
county_list = county_list+chr(9)+"19 Dakota"
county_list = county_list+chr(9)+"20 Dodge"
county_list = county_list+chr(9)+"21 Douglas"
county_list = county_list+chr(9)+"22 Faribault"
county_list = county_list+chr(9)+"23 Fillmore"
county_list = county_list+chr(9)+"24 Freeborn"
county_list = county_list+chr(9)+"25 Goodhue"
county_list = county_list+chr(9)+"26 Grant"
county_list = county_list+chr(9)+"27 Hennepin"
county_list = county_list+chr(9)+"28 Houston"
county_list = county_list+chr(9)+"29 Hubbard"
county_list = county_list+chr(9)+"30 Isanti"
county_list = county_list+chr(9)+"31 Itasca"
county_list = county_list+chr(9)+"32 Jackson"
county_list = county_list+chr(9)+"33 Kanabec"
county_list = county_list+chr(9)+"34 Kandiyohi"
county_list = county_list+chr(9)+"35 Kittson"
county_list = county_list+chr(9)+"36 Koochiching"
county_list = county_list+chr(9)+"37 Lac Qui Parle"
county_list = county_list+chr(9)+"38 Lake"
county_list = county_list+chr(9)+"39 Lake Of Woods"
county_list = county_list+chr(9)+"40 Le Sueur"
county_list = county_list+chr(9)+"41 Lincoln"
county_list = county_list+chr(9)+"42 Lyon"
county_list = county_list+chr(9)+"43 Mcleod"
county_list = county_list+chr(9)+"44 Mahnomen"
county_list = county_list+chr(9)+"45 Marshall"
county_list = county_list+chr(9)+"46 Martin"
county_list = county_list+chr(9)+"47 Meeker"
county_list = county_list+chr(9)+"48 Mille Lacs"
county_list = county_list+chr(9)+"49 Morrison"
county_list = county_list+chr(9)+"50 Mower"
county_list = county_list+chr(9)+"51 Murray"
county_list = county_list+chr(9)+"52 Nicollet"
county_list = county_list+chr(9)+"53 Nobles"
county_list = county_list+chr(9)+"54 Norman"
county_list = county_list+chr(9)+"55 Olmsted"
county_list = county_list+chr(9)+"56 Otter Tail"
county_list = county_list+chr(9)+"57 Pennington"
county_list = county_list+chr(9)+"58 Pine"
county_list = county_list+chr(9)+"59 Pipestone"
county_list = county_list+chr(9)+"60 Polk"
county_list = county_list+chr(9)+"61 Pope"
county_list = county_list+chr(9)+"62 Ramsey"
county_list = county_list+chr(9)+"63 Red Lake"
county_list = county_list+chr(9)+"64 Redwood"
county_list = county_list+chr(9)+"65 Renville"
county_list = county_list+chr(9)+"66 Rice"
county_list = county_list+chr(9)+"67 Rock"
county_list = county_list+chr(9)+"68 Roseau"
county_list = county_list+chr(9)+"69 St. Louis"
county_list = county_list+chr(9)+"70 Scott"
county_list = county_list+chr(9)+"71 Sherburne"
county_list = county_list+chr(9)+"72 Sibley"
county_list = county_list+chr(9)+"73 Stearns"
county_list = county_list+chr(9)+"74 Steele"
county_list = county_list+chr(9)+"75 Stevens"
county_list = county_list+chr(9)+"76 Swift"
county_list = county_list+chr(9)+"77 Todd"
county_list = county_list+chr(9)+"78 Traverse"
county_list = county_list+chr(9)+"79 Wabasha"
county_list = county_list+chr(9)+"80 Wadena"
county_list = county_list+chr(9)+"81 Waseca"
county_list = county_list+chr(9)+"82 Washington"
county_list = county_list+chr(9)+"83 Watonwan"
county_list = county_list+chr(9)+"84 Wilkin"
county_list = county_list+chr(9)+"85 Winona"
county_list = county_list+chr(9)+"86 Wright"
county_list = county_list+chr(9)+"87 Yellow Medicine"
county_list = county_list+chr(9)+"89 Out-of-State"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 556, 360, "CAF dialog part 2"
  ButtonGroup ButtonPressed
    PushButton 480, 30, 65, 15, "Update ABAWD", abawd_button
  EditBox 40, 50, 505, 15, notes_on_wreg
  ButtonGroup ButtonPressed
    PushButton 235, 90, 50, 15, "Update SHEL", update_shel_button
  DropListBox 45, 140, 100, 45, "Select ALLOWED HEST"+chr(9)+"AC/Heat - Full $493"+chr(9)+"Electric and Phone - $173"+chr(9)+"Electric ONLY - $126"+chr(9)+"Phone ONLY - $47"+chr(9)+"NONE - $0", hest_information
  EditBox 180, 140, 110, 15, notes_on_acut
  EditBox 45, 160, 245, 15, notes_on_coex
  EditBox 45, 180, 245, 15, notes_on_dcex
  EditBox 45, 200, 245, 15, notes_on_other_deduction
  EditBox 45, 220, 245, 15, expense_notes
  CheckBox 320, 85, 125, 10, "Check here to confirm the address.", address_confirmation_checkbox
  DropListBox 345, 150, 85, 45, county_list, addr_county
  DropListBox 480, 150, 30, 45, "No"+chr(9)+"Yes", homeless_yn
  DropListBox 335, 170, 95, 45, "SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"LE - Lease/Rent Doc"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd"+chr(9)+"? - Delayed", shel_verif
  DropListBox 480, 170, 30, 45, "No"+chr(9)+"Yes", reservation_yn
  DropListBox 375, 190, 165, 45, " "+chr(9)+"01 - Own home, lease or roomate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown", List6
  EditBox 315, 220, 230, 15, notes_on_address
  EditBox 35, 255, 405, 15, notes_on_acct
  EditBox 470, 255, 75, 15, notes_on_cash
  EditBox 35, 275, 240, 15, notes_on_cars
  EditBox 305, 275, 240, 15, notes_on_rest
  EditBox 110, 295, 435, 15, notes_on_other_assets
  EditBox 55, 320, 495, 15, verifs_needed
  GroupBox 5, 5, 545, 65, "WREG and ABAWD Information"
  Text 15, 20, 55, 10, "ABAWD Details:"
  Text 75, 20, 470, 10, "notes_on_abawd"
  Text 15, 35, 330, 10, "notes_on_abawd_two"
  GroupBox 5, 75, 290, 165, "Expenses and Deductions"
  Text 15, 95, 50, 10, "Total Shelter:"
  Text 70, 95, 155, 10, "total_shelter_amount"
  Text 15, 110, 275, 10, "shelter_details"
  Text 15, 125, 275, 10, "shelter_details_two"
  Text 20, 205, 20, 10, "Other:"
  Text 20, 225, 25, 10, "Notes:"
  GroupBox 305, 75, 245, 165, "Address"
  Text 350, 100, 175, 10, "addr_line_one"
  Text 350, 115, 175, 10, "addr_line_two"
  Text 350, 130, 175, 10, "city, state zip"
  Text 315, 155, 25, 10, "County:"
  Text 440, 155, 35, 10, "Homeless:"
  Text 315, 175, 20, 10, "Verif:"
  Text 435, 175, 45, 10, "Reservation:"
  Text 315, 195, 55, 10, "Living Situation:"
  Text 315, 210, 75, 10, "Notes on address:"
  GroupBox 5, 245, 545, 70, "Assets"
  Text 5, 325, 50, 10, "Verifs needed:"
  ButtonGroup ButtonPressed
    PushButton 380, 345, 60, 10, "previous page", previous_to_page_02_button
    PushButton 445, 340, 50, 15, "NEXT", next_to_page_04_button
    CancelButton 500, 340, 50, 15
    PushButton 10, 55, 25, 10, "WREG", wreg_button
    PushButton 315, 100, 25, 10, "ADDR", addr_button
    PushButton 15, 145, 25, 10, "HEST", hest_button
    PushButton 150, 145, 25, 10, "ACUT", acut_button
    PushButton 15, 165, 25, 10, "COEX", coex_button
    PushButton 15, 185, 25, 10, "DCEX", dcex_button
    PushButton 10, 260, 25, 10, "ACCT", acct_button
    PushButton 445, 260, 25, 10, "CASH", cash_button
    PushButton 10, 280, 25, 10, "CARS", cars_button
    PushButton 280, 280, 25, 10, "REST", rest_button
    PushButton 10, 300, 25, 10, "SECU", secu_button
    PushButton 35, 300, 25, 10, "TRAN", tran_button
    PushButton 60, 300, 45, 10, "other assets", other_asset_button
EndDialog

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 551, 260, "Dialog"
  GroupBox 5, 10, 540, 95, "Member 01 - Angela Burns"
  Text 15, 30, 70, 10, "FSET WREG Status:"
  DropListBox 90, 25, 130, 45, " "+chr(9)+"03  Unfit for Employment"+chr(9)+"04  Responsible for Care of Another"+chr(9)+"05  Age 60+"+chr(9)+"06  Under Age 16"+chr(9)+"07  Age 16-17, live w/ parent"+chr(9)+"08  Care of Child <6"+chr(9)+"09  Employed 30+ hrs/wk"+chr(9)+"10  Matching Grant"+chr(9)+"11  Unemployment Insurance"+chr(9)+"12  Enrolled in School/Training"+chr(9)+"13 CD Program"+chr(9)+"14  Receiving MFIP"+chr(9)+"20  Pend/Receiving DWP"+chr(9)+"15  Age 16-17 not live w/ Parent"+chr(9)+"16  50-59 Years Old"+chr(9)+"21  Care child < 18"+chr(9)+"17  Receiving RCA or GA"+chr(9)+"30  FSET Participant"+chr(9)+"02  Fail FSET Coop"+chr(9)+"33  Non-coop being referred", wreg_status
  Text 230, 30, 55, 10, "ABAWD Status:"
  DropListBox 285, 25, 110, 45, " "+chr(9)+"01  WREG Exempt"+chr(9)+"02  Under Age 18"+chr(9)+"03  Age 50+"+chr(9)+"04  Caregiver of Minor Child"+chr(9)+"05  Pregnant"+chr(9)+"06  Employed 20+ hrs/wk"+chr(9)+"07  Work Experience"+chr(9)+"08  Other E and T"+chr(9)+"09  Waivered Area"+chr(9)+"10  ABAWD Counted"+chr(9)+"11  Second Set"+chr(9)+"12  RCA or GA Participant"+chr(9)+"13  ABAWD Banked Months", abawd_status
  CheckBox 405, 25, 130, 10, "Check here if this person is the PWE", pwe_checkbox
  Text 15, 50, 145, 10, "Number of ABAWD months used in past 36:"
  EditBox 160, 45, 25, 15, number_abawd_used
  Text 200, 50, 95, 10, "List all ABAWD months used:"
  EditBox 300, 45, 135, 15, list_of_ABAWD_used
  Text 15, 70, 135, 10, "If used, list the first month of Second Set:"
  EditBox 155, 65, 40, 15, initial_second_set_month
  Text 205, 70, 130, 10, "If NOT Eligible for Second Set, Explain:"
  EditBox 335, 65, 200, 15, explain_not_second_set
  Text 15, 90, 115, 10, "Number of BANKED months used:"
  EditBox 130, 85, 25, 15, number_banked_months_used
  Text 170, 90, 45, 10, "Other Notes:"
  EditBox 220, 85, 315, 15, abawd_other_notes
  ButtonGroup ButtonPressed
    PushButton 455, 240, 90, 15, "Return to Main Dialog", return_button
  GroupBox 5, 105, 540, 95, "Member 01 - Angela Burns"
  Text 15, 125, 70, 10, "FSET WREG Status:"
  DropListBox 90, 120, 130, 45, " "+chr(9)+"03  Unfit for Employment"+chr(9)+"04  Responsible for Care of Another"+chr(9)+"05  Age 60+"+chr(9)+"06  Under Age 16"+chr(9)+"07  Age 16-17, live w/ parent"+chr(9)+"08  Care of Child <6"+chr(9)+"09  Employed 30+ hrs/wk"+chr(9)+"10  Matching Grant"+chr(9)+"11  Unemployment Insurance"+chr(9)+"12  Enrolled in School/Training"+chr(9)+"13 CD Program"+chr(9)+"14  Receiving MFIP"+chr(9)+"20  Pend/Receiving DWP"+chr(9)+"15  Age 16-17 not live w/ Parent"+chr(9)+"16  50-59 Years Old"+chr(9)+"21  Care child < 18"+chr(9)+"17  Receiving RCA or GA"+chr(9)+"30  FSET Participant"+chr(9)+"02  Fail FSET Coop"+chr(9)+"33  Non-coop being referred", List3
  Text 230, 125, 55, 10, "ABAWD Status:"
  DropListBox 285, 120, 110, 45, " "+chr(9)+"01  WREG Exempt"+chr(9)+"02  Under Age 18"+chr(9)+"03  Age 50+"+chr(9)+"04  Caregiver of Minor Child"+chr(9)+"05  Pregnant"+chr(9)+"06  Employed 20+ hrs/wk"+chr(9)+"07  Work Experience"+chr(9)+"08  Other E and T"+chr(9)+"09  Waivered Area"+chr(9)+"10  ABAWD Counted"+chr(9)+"11  Second Set"+chr(9)+"12  RCA or GA Participant"+chr(9)+"13  ABAWD Banked Months", List4
  CheckBox 405, 120, 130, 10, "Check here if this person is the PWE", Check3
  Text 15, 145, 145, 10, "Number of ABAWD months used in past 36:"
  EditBox 160, 140, 25, 15, Edit7
  Text 200, 145, 95, 10, "List all ABAWD months used:"
  EditBox 300, 140, 135, 15, Edit8
  Text 15, 165, 135, 10, "If used, list the first month of Second Set:"
  EditBox 155, 160, 40, 15, Edit9
  Text 205, 165, 130, 10, "If NOT Eligible for Second Set, Explain:"
  EditBox 335, 160, 200, 15, Edit10
  Text 15, 185, 115, 10, "Number of BANKED months used:"
  EditBox 130, 180, 25, 15, Edit11
  Text 170, 185, 45, 10, "Other Notes:"
  EditBox 220, 180, 315, 15, Edit12
EndDialog

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 340, 240, "SHEL Detail Dialog"
  DropListBox 60, 10, 125, 45, HH_memb_list, clt_SHEL_is_for
  Text 5, 15, 55, 10, "SHEL for Memb"
  ButtonGroup ButtonPressed
    PushButton 200, 10, 40, 10, "Load", load_button

  DropListBox 85, 30, 30, 45, "Yes"+chr(9)+"No", subsidized_yn
  DropListBox 175, 30, 30, 45, "Yes"+chr(9)+"No", shared_yn
  EditBox 45, 60, 35, 15, retro_rent_amount
  DropListBox 85, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_rent_verif
  EditBox 195, 60, 35, 15, prosp_rent_amount
  DropListBox 235, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_rent_verif
  EditBox 45, 80, 35, 15, retro_lot_amount
  DropListBox 85, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_lot_verif
  EditBox 195, 80, 35, 15, prosp_lot_amount
  DropListBox 235, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_lot_verif
  EditBox 45, 100, 35, 15, retro_mortgage_amount
  DropListBox 85, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_mortgage_verif
  EditBox 195, 100, 35, 15, prosp_mortgage_amount
  DropListBox 235, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_mortgage_verif
  EditBox 45, 120, 35, 15, retro_ins_amount
  DropListBox 85, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_ins_verif
  EditBox 195, 120, 35, 15, prosp_ins_amount
  DropListBox 235, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_ins_verif
  EditBox 45, 140, 35, 15, retro_tax_amount
  DropListBox 85, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_tax_verif
  EditBox 195, 140, 35, 15, prosp_tax_amount
  DropListBox 235, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_tax_verif
  EditBox 45, 160, 35, 15, retro_room_amount
  DropListBox 85, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_room_verif
  EditBox 195, 160, 35, 15, prosp_room_amount
  DropListBox 235, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_room_verif
  EditBox 45, 180, 35, 15, retro_garage_amount
  DropListBox 85, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_garage_verif
  EditBox 195, 180, 35, 15, prosp_garage_amount
  DropListBox 235, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_garage_verif
  EditBox 45, 200, 35, 15, retro_subsity_amount
  DropListBox 85, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", retro_subsidy_verif
  EditBox 195, 200, 35, 15, prosp_subsidy_amount
  DropListBox 235, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", prosp_subsidy_verif
  ButtonGroup ButtonPressed
    PushButton 245, 220, 90, 15, "Return to Main Dialog", return_button
  Text 15, 35, 60, 10, "HUD Subsidized:"
  Text 140, 35, 30, 10, "Shared:"
  Text 45, 50, 50, 10, "Retrospective"
  Text 195, 50, 50, 10, "Prospective"
  Text 20, 65, 20, 10, "Rent:"
  Text 10, 85, 30, 10, "Lot Rent:"
  Text 5, 105, 35, 10, "Mortgage:"
  Text 5, 125, 35, 10, "Insurance:"
  Text 15, 145, 25, 10, "Taxes:"
  Text 15, 165, 25, 10, "Room:"
  Text 10, 185, 30, 10, "Garage:"
  Text 10, 205, 30, 10, "Subsidy:"
EndDialog

' Dialog1 = ""'
' BeginDialog Dialog1, 0, 0, 451, 305, "CAF dialog part 2"
'   EditBox 65, 5, 375, 15, notes_on_wreg
'   Text 65, 30, 375, 10, "notes_on_abawd"
'   EditBox 80, 85, 360, 15, notes_on_shel
'   DropListBox 275, 50, 100, 15, "Select ALLOWED HEST"+chr(9)+"AC/Heat - Full $493"+chr(9)+"Electric and Phone - $173"+chr(9)+"Electric ONLY - $126"+chr(9)+"Phone ONLY - $47"+chr(9)+"NONE - $0", notes_on_hest
'   EditBox 65, 115, 375, 15, notes_on_coex
'   EditBox 65, 135, 375, 15, notes_on_dcex
'   EditBox 65, 155, 260, 15, notes_on_other_deductions
'   EditBox 370, 155, 70, 15, notes_on_cash
'   EditBox 65, 175, 375, 15, notes_on_acct
'   EditBox 65, 195, 375, 15, notes_on_cars
'   EditBox 65, 215, 375, 15, notes_on_rest
'   EditBox 105, 235, 335, 15, other_assets
'   EditBox 55, 265, 385, 15, verifs_needed
'   ButtonGroup ButtonPressed
'     PushButton 270, 290, 60, 10, "previous page", previous_to_page_02_button
'     PushButton 335, 285, 50, 15, "NEXT", next_to_page_04_button
'     CancelButton 390, 285, 50, 15
'     PushButton 25, 10, 35, 10, "WREG:", WREG_button
'     PushButton 25, 30, 35, 10, "ABAWD:", ABAWD_button
'     PushButton 5, 50, 25, 10, "SHEL:", SHEL_button
'     PushButton 245, 50, 25, 10, "HEST:", HEST_button
'     PushButton 25, 120, 25, 10, "COEX:", COEX_button
'     PushButton 25, 140, 25, 10, "DCEX:", DCEX_button
'     PushButton 340, 160, 25, 10, "CASH:", CASH_button
'     PushButton 25, 180, 30, 10, "ACCTs:", ACCT_button
'     PushButton 30, 200, 25, 10, "CARS:", CARS_button
'     PushButton 30, 220, 25, 10, "REST:", REST_button
'     PushButton 5, 240, 25, 10, "SECU/", SECU_button
'     PushButton 30, 240, 25, 10, "TRAN/", TRAN_button
'     PushButton 55, 240, 45, 10, "other assets:", OTHR_button
'   Text 5, 160, 60, 10, "Other Deductions:"
'   Text 5, 270, 50, 10, "Verifs needed:"
'   Text 40, 50, 30, 10, "Amount:"
'   EditBox 75, 45, 165, 15, shel_amount
'   Text 35, 70, 20, 10, "Verif:"
'   EditBox 60, 65, 180, 15, shel_verif
'   ButtonGroup ButtonPressed
'     PushButton 245, 70, 25, 10, "ACUT", ACUT_button
'   EditBox 275, 65, 165, 15, notes_on_acut
'   Text 10, 90, 65, 10, "SHEL/HEST Notes:"
' EndDialog

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 376, 165, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 320, 145, 50, 15
  Text 15, 10, 350, 110, "What does it say?                                                                       How?                                                                                                                                                                                                                                                  The thing here:"
  EditBox 80, 5, 105, 15, Edit1
  EditBox 235, 5, 105, 15, Edit2
  EditBox 80, 25, 105, 15, Edit3
EndDialog


dialog Dialog1
