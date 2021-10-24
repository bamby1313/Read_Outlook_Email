# Read Your Outlook Emails
Read outlook email from MATLAB 

* Matlab function which imports the 'readed' or 'unreaded' outlook emails from inbox and their folders - subfolders. 
* Extracts their subjects, bodies and can save their attachments.

% Reads all emails from default inbox
mails = ReadOutlook2;

% Reads all Unread emails from inbox
mails = ReadOutlook2("AccountName", myAccount)

% Reads all Unread emails from inbox
mails = ReadOutlook2("Read", 1);

% Reads all Unread emails from  inbox and mark them as read
mails = ReadOutlook2("Read", 1, "Mark", 1);
