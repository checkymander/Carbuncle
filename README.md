# Carbuncle
Tool for interacting with outlook interop during red team engagements.

# Supported Functions
* Enum - Enumerate e-mails in the users inbox
* Search - Search for E-mails mathching a specific keyword
* Read - Get the contents of an e-mail either by Subject or Number
* Monitor - Monitor and displays new e-mails as they arrive
* Send - Send an e-mail from your target to a person or group of people, can also add attachments for internal phishing.

# Usage
```
carbuncle.exe enum [/display] - Enumerates all e-mails in the users inbox with the option to display the e-mail body (Warning this could be A LOT of output)
carbuncle.exe search /keyword:"password" [/display] - Searches e-mails based on keyword, with the option to display the e-mail body (Searches based on subject and body)
carbuncle.exe read /subject:"Subject of E-mail" - Reads all e-mails that contain the subject
carbuncle.exe read /number:"15" - Reads the 15th e-mail listed via the carbuncle.exe enum command (Index starts at 1)
carbuncle.exe send /body:"Hello World" /subject:"Subject E-mail" /recipients:"test@email.com" - Sends an e-mail to a user or group of users without an attachment
carbuncle.exe send /body:"Test Message to multiple people" /subject:"Hello World" /recipients:"email1@gmail.com,email2@gmail.com,ontothenextone@gmail.com" /attachment:"C:\Users\checkymander\Pictures\checkymander.png" /attachmentname:"checkymander" - Sends an e-mail to a user or group of users with an attachment
carbuncle.exe monitor
```
