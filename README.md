# Crack-Protected-VBA

A .xlsm file is an excel file that contains Visual Basic Application (VBA) macros. Oftentimes, these macros can contain sensitive information, depending on the purpose of the file. Because of this, oftentimes, developers will protect their VBA source code with a password (as seen below). This simple script will bypass the password and allow a user to view, manipulate, and save the original VBA code. 

![vba password](https://www.top-password.com/images/vba-project-password.png)

The manual process to crack VBA files can be a little tedious, and will require the use of a hex editor. A great process demo can be found here: (https://www.youtube.com/watch?v=QV56TWT4nKw)

At a high level, the steps to bypass this password protection are as follows:
  1. Convert the file extention from .xlsm to .zip
  2. Extract the .zip archive
  3. Open the ./[archive]/xl/vbaProject.bin file in a hex editor
  4. Find the "DPB" string in hex, and convert the "B" to an "X"
  5. Save the edited hex file
  6. Re-compress the contents back to a .zip file
  7. Convert the new .zip file back to a .xlsm file

The reason this process works is because within the compressed .xlsm file package contain the project files in Open XML format. Specifically, the vbaProject.bin file contains the "DPB" key which defines the encrypted password of the file, however when that key is void by changing its name, the reference for the key breaks, resulting in an exposed project. 

## Usage
Bash: 
`git clone https://github.com/cdnet01/Crack-Protected-VBA.git`
`cd Crack-Protected-VBA`
`mv [VBA File Path] .`
`./crackProtectedVBA.sh [VBA File Name]`

Powershell:
`wget raw.githubusercontent.com/cdnet01/Crack-Protected-VBA/main/crackProtectedVba.ps1`
`Move-Item -Path [VBA File Path] -Destination .`
`.\crackProtectedVBA.ps1 [VBA File Name]`

OR
`git clone https://github.com/cdnet01/Crack-Protected-VBA.git`
`cd Crack-Protected-VBA`
`Move-Item -Path [VBA File Path] -Destination .`
`.\crackProtectedVBA.ps1 [VBA File Name]`

Ultimately, exposing the protected VBA code could reveal useful sensitive information. 

Do no harm. 
