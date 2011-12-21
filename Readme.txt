=== HTML/XML PARSER FOR VB6 ===

DESCRIPTION

This is a simple DOM-compliant (version 1) parser for HTML- and XML-files, written in VB6. 
Licenced under LGPL 2.1.


PROBLEMS WITH ZIP

If you experience problems with opening the files in the ZIP, please download and install GIT
and clone this public repository (with CORE.AUTOCRLF set to TRUE). This means setting "Checkout 
Windows-style, commit Unix-style line endings" in the installer.


INSTRUCTIONS FOR USING THIS LIBRARY

 1. To embed the parser into your project completely, add pHTMLParser.vbp to your project 
    (File -> Add Project .. -> Existing).
 2. OR, to compile the parser as a seperate DLL-file, compile pHTMLParser and include it in
    the installation process of your program.
 3. Add the file Entities.dat to the same directory as your program or the DLL-file.
 4. Execute "Install.bat" for the user if the DLL-file is stored in a different directory.


DEMO PROGRAM

 To just test the capabilities and speed of this parser, compile the demo project using gHTML.vbg


TROUBLESHOOTING

 Problem: 
	Unable to run pHTML.exe. 
 Error message: 
	"You do not have an appropriate license to use this functionality"
 Resolution:
	The library (HTMLParser.dll) is not registered. Please execute "Install.bat" in the same folder.

 
 Problem:
 	Library is not registered when "Install.bat" is run.
 Error messages:
	The module "HTMLParser.dll" was loaded but the call to DllRegisterServer failed 
		with error code 0x80004005.		
	The module "BrowserControl.ocx" was loaded but the call to DllRegisterServer failed 
		with error code 0x80004005.
 Resolution:
	Execute the "Install.bat" as administrator (right-click and choose "Run as Administrator").