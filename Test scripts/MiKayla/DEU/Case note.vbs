IF programs = "Health Care" or programs = "Medical Assistance" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "",
"Claims entered for #" &  MAXIS_case_number, " Member #: " & memb_number & vbcr & " Date Overpayment Created: " & OP_Date & vbcr & "Programs: "
& programs & vbcr & " ", "", False)

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

EmReadScreen 75, 18, 3 'more'

EmReadScreen 75, 19, 3
EmReadScreen 75, 20, 3
EmReadScreen 75, 21, 3
EmReadScreen 75, 22, 3
