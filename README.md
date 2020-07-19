# Carbuncle
Tool for interacting with outlook interop during red team engagements.

# Supported Functions
* Search - Search for E-mails mathching a specific keyword
* Read - Get the contents of an e-mail either by Subject or Number
* Monitor - Monitor and displays new e-mails as they arrive
* Enum - Enumerate e-mails in the users inbox
* Send - Send an e-mail from your target to someone else (ToDo: Attachments)

# Usage
```
carbuncle.exe enum
carbuncle.exe search /keyword:"password"
carbuncle.exe send /body:"Hello World" /subject:"Subject E-mail" /recipient:"test@email.com"
carbuncle.exe read /subject:"Subject of E-mail"
carbuncle.exe read /number:"15"
carbuncle.exe monitor
```
