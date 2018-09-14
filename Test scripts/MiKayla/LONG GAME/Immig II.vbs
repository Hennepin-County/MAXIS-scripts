BeginDialog IMIG_dialog, 0, 0, 366, 300, "Immigration Status"
 EditBox 60, 5, 40, 15, MAXIS_case_number
 EditBox 140, 5, 20, 15, memb_number
 EditBox 210, 5, 40, 15, actual_date
 CheckBox 255, 10, 110, 10, "Address Additional Questions?", second_CHECKBOX
 DropListBox 60, 35, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"US Citizen", immig_status_dropdown
 DropListBox 60, 55, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"N/A", LPR_status_dropdown
 DropListBox 255, 35, 95, 15, "Select One:"+chr(9)+"SAVE Primary"+chr(9)+"SAVE Secondary"+chr(9)+"Alien Card"+chr(9)+"Passport/Visa"+chr(9)+"Re-Entry Prmt"+chr(9)+"INS Correspondence"+chr(9)+"Other Document"+chr(9)+"No Ver Prvd", status_verification
 DropListBox 255, 55, 95, 15, "Select One:"+chr(9)+"AA Amerasian"+chr(9)+"EH Ethnic Chinese"+chr(9)+"EL Ethnic Lao"+chr(9)+"HG Hmong"+chr(9)+"KD Kurd"+chr(9)+"SJ Soviet Jew"+chr(9)+"TT Tinh"+chr(9)+"AF Afghanistan"+chr(9)+"BK Bosnia"+chr(9)+"CB Cambodia"+chr(9)+"CH China,"+chr(9)+"CU Cuba"+chr(9)+"ES El Salvador"+chr(9)+"ER Eritrea"+chr(9)+"ET Ethiopia"+chr(9)+"GT Guatemala"+chr(9)+"HA Haiti"+chr(9)+"HO Honduras"+chr(9)+"IR Iran"+chr(9)+"IZ Iraq"+chr(9)+"LI Liberia"+chr(9)+"MC Micronesia"+chr(9)+"MI Marshall"+chr(9)+"Islands"+chr(9)+"MX Mexico"+chr(9)+"WA Namibia"+chr(9)+"PK Pakistan"+chr(9)+"RP Philippines"+chr(9)+"PL Poland"+chr(9)+"RO Romania"+chr(9)+"RS Russia"+chr(9)+"SO Somalia"+chr(9)+"SF South Africa"+chr(9)+"TH Thailand"+chr(9)+"VM Vietnam"+chr(9)+"OT All Others", nationality_dropdown
 DropListBox 255, 75, 95, 15, "Select One:"+chr(9)+"Certificate of Naturalization"+chr(9)+"Employment Auth Card (I-776 work permit)"+chr(9)+"I-94 Travel Document"+chr(9)+"I-220 B Order of Supervision"+chr(9)+"LPR Card (I-551 green card)"+chr(9)+"SAVE"+chr(9)+"Other", immig_doc_type
 EditBox 305, 95, 45, 15, entry_date
 EditBox 305, 115, 45, 15, status_date
 CheckBox 10, 90, 85, 10, "Inital SAVE Requested?", save_CHECKBOX
 CheckBox 10, 105, 100, 10, "Additional SAVE Requested?", additional_CHECKBOX
 CheckBox 15, 120, 205, 10, "check if a copy of the immig doc was attached to the request", SAVE_docs_check
 OptionGroup RadioGroup1
    RadioButton 15, 155, 25, 10, "No", not_sponsored
    RadioButton 15, 170, 75, 10, "Yes, sponsored by:", sponsored
  EditBox 85, 190, 70, 15, name_sponsor
  EditBox 220, 190, 125, 15, sponsor_addr
  EditBox 85, 210, 70, 15, name_sponsor_two
  EditBox 220, 210, 125, 15, sponsor_addr_two
  EditBox 85, 230, 70, 15, name_sponsor_three
  EditBox 220, 230, 125, 15, sponsor_addr_three
  EditBox 75, 255, 160, 15, other_notes
  EditBox 75, 275, 160, 15, worker_sig
  ButtonGroup ButtonPressed
    CancelButton 305, 275, 45, 15
    OkButton 255, 275, 45, 15
  Text 165, 10, 40, 10, "Actual Date:"
  Text 255, 100, 45, 10, "Date of entry:"
  Text 10, 40, 50, 10, "Immig Status:"
  Text 10, 60, 50, 10, "LPR Adj From:"
  Text 200, 40, 50, 10, "Status Verified:"
  Text 190, 60, 60, 10, "Nationality/Nation:"
  Text 195, 80, 55, 10, "Immig Doc Type:"
  Text 105, 10, 30, 10, "Memb #:"
  GroupBox 5, 140, 350, 110, "Sponsored on I-864 Affidavit of Support? (LPR COA CODE: C, CF, CR, CX, F, FX, IF, IR)"
  Text 80, 155, 245, 10, "*If date of entry was prior to 12/19/1997 sponsor information is not needed"
  Text 120, 170, 205, 10, "*If sponsor is active on MAXIS case information is not needed"
  Text 20, 195, 60, 10, "Name of sponsor:"
  Text 165, 195, 55, 10, "Address/Phone:"
  Text 20, 215, 60, 10, "Name of sponsor:"
  Text 165, 215, 55, 10, "Address/Phone:"
  Text 20, 235, 60, 10, "Name of sponsor:"
  Text 165, 235, 55, 10, "Address/Phone:"
  Text 25, 260, 45, 10, "Other Notes:"
  GroupBox 5, 25, 350, 110, "Immigration Information"
  Text 260, 120, 40, 10, "Status date:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 280, 65, 10, "Worker Signature:"
EndDialog
