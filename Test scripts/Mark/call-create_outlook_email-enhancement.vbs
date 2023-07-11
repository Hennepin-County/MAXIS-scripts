Function create_outlook_email(email_from, email_recip, email_recip_CC, email_subject, email_importance, email_body, email_attachment, send_email)
'--- This function creates a an outlook appointment
'~~~~~ (email_from): email address for sender
'~~~~~ (email_recip): email address for recipeint - seperated by semicolon
'~~~~~ (email_recip_CC): email address for recipeints to cc - seperated by semicolon
'~~~~~ (email_subject): subject of email in quotations or a variable
'~~~~~ (email_importance): set importance of email - 0 (low), 1 (normal), or high (2)
'~~~~~ (email_body): body of email in quotations or a variable
'~~~~~ (email_attachment): set as "" if no email or file location
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
    objMail.Subject = email_subject                 'email subject
    objMail.Importance = email_importance           'email importance - 0 (low), 1 (normal), or high (2)
    objMail.Body = email_body                       'email body
    If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
    'Sends email
    If send_email = true then objMail.Send	                   'Sends the email
    Set objMail =   Nothing
    Set objOutlook = Nothing
End Function