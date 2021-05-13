	
### Convert Office to PDF

This PowerShell script allows non-admin users to convert Microsoft Ofice Word, Excel, & PowerPoint to PDF.

## Table of Contents
* [General Information](#general-information)
* [Technologies](#technologies)
* [Setup](#setup)

## General Information
Begin with a PowerShell window. Change Directory to the folder location of this script then run it. The user is asked to provide the file types to convert, followed by the folder location of the files. In the current version the user is asked to keep or remove the log file everytime a new file type conversion begins. In future versions I will make this its own function and call it at the end of the process. Additionally, this version asks the user to keep or remove the newly created PDFs. In future versions I will comment this out.
Become familiar with PowerShell by visiting the Microsoft website, PowerShell Documentation: https://docs.microsoft.com/en-us/powershell/

  Many thanks to the contributors on the following sites; without their examples and guides this would have been a laborious endeavour.
  * https://adamtheautomator.com/
  * https://devblogs.microsoft.com/
  * https://morgantechspace.com/
  * https://stackoverflow.com
  * https://superuser.com/


## Technologies
My computing environment is:
* Windows 10.0.18363
* PowerShell version 5.1.18362.1474

## Setup
Open a PowerShell window and Change Directory to the location of this script.
```powershell
cd 'C:\Users\Documents\Scripts'
```
At the prompt run this script (note the ".\\" represents the current folder)
```powershell
.\office2pdf.ps1
```
The process begins by asking the user for the types of files they want to convert (a, p, w, x)
* a  - represents ALL file types - Word, Excel, & PowerPoint
* p  - represents PowerPoint
* w  - represents Word
* x  - represents Excel
After the user enters the file type, they are prompted to enter the folder path.
The conversion process begins and each file is listed in the PS window.
When the conversion process is complete, the PowerShell window displays the number of files converted and the location of the log file.
Error handling catches some of the common errors such as inocrrect user input variable and string vs integer. The conversion process is wrapped in a Try/Catch.
