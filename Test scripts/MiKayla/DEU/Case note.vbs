IF programs = "Health Care" or programs = "Medical Assistance" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "",
"Claims entered for #" &  MAXIS_case_number, " Member #: " & memb_number & vbcr & " Date Overpayment Created: " & OP_Date & vbcr & "Programs: "
& programs & vbcr & " ", "", False)

EmReadScreen page_line_one, 75, 4, 3
EmReadScreen page_line_two, 75, 5, 3
EmReadScreen page_line_three, 75, 6, 3
EmReadScreen page_line_four, 75, 7, 3
EmReadScreen page_line_five, 75, 8, 3
EmReadScreen page_line_six, 75, 9, 3
EmReadScreen page_line_seven, 75, 10, 3
EmReadScreen page_line_eight, 75, 11, 3
EmReadScreen page_line_nine, 75, 12, 3
EmReadScreen page_line_ten, 75, 13, 3
EmReadScreen page_line_eleven, 75, 14, 3
EmReadScreen page_line_twelve, 75, 15, 3
EmReadScreen page_line_one, 75, 16, 3
EmReadScreen page_line_one, 75, 17, 3

EmReadScreen next_page, 4, 18, 3  If next_page = "more" then
    PF8

    EmReadScreen 75, 4, 3
    EmReadScreen 75, 5, 3
    EmReadScreen 75, 6, 3
    EmReadScreen 75, 7, 3
    EmReadScreen 75, 8, 3
    EmReadScreen 75, 9, 3
    EmReadScreen 75, 10, 3
    EmReadScreen 75, 11, 3
    EmReadScreen 75, 12, 3
    EmReadScreen 75, 13, 3
    EmReadScreen 75, 14, 3
    EmReadScreen 75, 15, 3
    EmReadScreen 75, 16, 3
    EmReadScreen 75, 17, 3
