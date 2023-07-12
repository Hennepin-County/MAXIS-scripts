Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
'--- This function creates a an outlook appointment
'~~~~~ (email_from): email address for sender
'~~~~~ (email_recip): email address for recipient - separated by semicolon
'~~~~~ (email_recip_CC): email address for recipients to cc - separated by semicolon
'~~~~~ (email_recip_bcc): email address for recipients to bcc - separated by semicolon
'~~~~~ (email_subject): subject of email in quotations or a variable
'~~~~~ (email_importance): set importance of email - 0 (low), 1 (normal), or high (2)
'~~~~~ (include_flag): indicate whether to include follow-up flag on email - true or false
'~~~~~ (email_flag_text): set the text of the follow-up flag, if no follow-up flag needed then use ""
'~~~~~ (email_flag_days): set the number of days from today that the flag is due
'~~~~~ (email_flag_reminder): set whether a flag reminder should be set - true or false
'~~~~~ (email_flag_reminder_days): set the number of days from today that a reminder for the flag should be set
'~~~~~ (email_body): body of email in quotations or a variable, function will determine whether HTMLbody is needed based on email_body content
'~~~~~ (include_email_attachment): indicate if any (1 or more) attachments should be included - indicate true to include attachments or false to not include attachments
'~~~~~ (email_attachment_array): if including 1 or more attachments, then enter the array name here to add these attachments to the email
'~~~~~ (send_email): set as TRUE or FALSE
'===== Keywords: MAXIS, PRISM, create, outlook, email

	'Setting up the Outlook application
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    If send_email = False then objMail.Display      'To display message only if the script is NOT sending the email for the user.

    'Adds the information to the email
    objMail.SentOnBehalfOfName = email_from         'email sender
    objMail.to = email_recip                        'email recipient
    objMail.cc = email_recip_CC                     'cc recipient
    objMail.Bcc = email_recip_bcc                   'bcc recipient
    objMail.Subject = email_subject                 'email subject
    objMail.Importance = email_importance           'email importance - 0 (low), 1 (normal), or high (2)

    'Set email follow-up flag
    If include_flag = True Then
        objMail.FlagRequest = email_flag_text
        objMail.FlagDueBy = DateAdd("d", email_flag_days, Date())
        objMail.ReminderSet = email_flag_reminder
        objMail.ReminderTime = DateAdd("d", email_flag_reminder_days, Date()) & " 12:00:00 PM"
    End If
    objMail.Body = email_body                       'Default email body
    'Determines if HTML body is needed based on email_body content
    If instr(email_body, "<p>") OR _
        instr(email_body, "<br>") OR _
        instr(email_body, "<i>") OR _
        instr(email_body, "&emsp") OR _
        instr(email_body, "&ensp") OR _
        instr(email_body, "href") Then 
            objMail.HTMLBody = email_body
    End If

    'Iterates through array of email attachments if indicated
    If include_email_attachment = True Then
      For Each attachment In email_attachment_array
        objMail.Attachments.Add(attachment)
      Next
    End If
    'Sends email
    If send_email = true then objMail.Send	                   'Sends the email
    Set objMail =   Nothing
    Set objOutlook = Nothing
End Function