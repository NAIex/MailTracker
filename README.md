# MailTracker
 A python module for tracking mails sent by individuals from the department of External Relations of ASII!

 It was made to make my life easier when it came to tracking who contacted and who didn't! There are multiple ways to call the module, and all of them require one arguement:
 
  - help: tells you what each argument does.
  - login: asks you for your ASII e-mail adress. Then a prompt appears asking for the password.
  - reset: It empties the excel file where all the data is stored. Then a prompt appears asking you for the starting search date.
  - modify_list: It is used to manipulate the txt file storing all the e-mails:
    - add firstname.lastname: Adds the email adress with the format firstname.lastname@asii.ro in the search list.
    - del firstname.lastname: Removes the e-mail adress with the format firstname.lastname@asii.ro from the search list.
    - exit: Saves the file and exits the module.

 