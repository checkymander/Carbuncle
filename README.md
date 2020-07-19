# Carbuncle
Tool for interacting with outlook interop during red team engagements.

# Supported Functions
* Enum - Enumerate e-mails in the users inbox
* Read - Get the contents of an e-mail either by Subject or Number
* Monitor - Monitor and displays new e-mails as they arrive
* Send - Send an e-mail from your target to a person or group of people, can also add attachments for internal phishing.


# Enum Usage
```
Search for e-mails from a certain e-mail address
carbuncle.exe enum /email:victim@gmail.com [/display]

Search for e-mails from a certain person
carbuncle.exe enum /name:"Checky Mander" [/display]

Search for e-mails that contain a keyword
carbuncle.exe enum /keyword:"Password" [/display]
```

# Read Usage

Note: When using the Read command, display is enabled by default

```
Read e-mail by subject
carbuncle.exe read /subject:"Important document"

Read e-mail by number
carbuncle.exe read /number:13
```

# Monitor Usage
```
Monitor for new e-mails
carbuncle.exe monitor [/display]
```


# Send Usage
```
Send an e-mail to multiple people
carbuncle.exe send /body:"Test Message to multiple people" /subject:"Hello World" /recipients:"email1@gmail.com,email2@gmail.com,ontothenextone@gmail.com" /attachment:"C:\Users\checkymander\Pictures\checkymander.png" /attachmentname:"checkymander"

Send an e-mail to one person without an attachment
carbuncle.exe send /body:"Hello World" /subject:"Subject E-mail" /recipients:"test@email.com"
````
