# O365-Scripts

This repo will hold different scripts to automate daily task that are done with the Microsoft 365 platform.

Prerequisites(1): Must have the correct version of PyWin32 for your computer. Downloads/Directions can be found @ https://sourceforge.net/projects/pywin32/

Prerequisites(2): Must have Python2.7 install. Downloads/Directions can be found @ https://www.python.org/downloads/

Prerequisites(3): Must have O365 Outlook installed locally in order to function correctly.

Outlook_find_contacts_send_email: 
  This script automates finding all contacts in Outlook then sending one email to one particular contact. This script can easily be rewritten to iterate through each contact in your contacts folder and send each of those contacts an email. If you add the 'Windows Task Scheduler', now you have an automated way to send emails on a daily basis. As of now the script will send a test email to the second person in your contacts folder, but as I said that can easily be adjusted.
  
  updates:
  v2:
    - added try/catch to sendTheEmail() function



