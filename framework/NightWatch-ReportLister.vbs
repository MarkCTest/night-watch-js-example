'-------------------------------------------------------------------------------------------
' Nightwatch.js Test Tool Framework Example.

' ## As with ALL examples and code at 'Github > MarkCTest' this is not an optimised solution.
' ## It uses a mix of examples and approaches to demonstrate what's possible and give practice.
' ## This is provided on an As Is basis and you will need to tidy up the code.

' ## What
' Get test reports listed out for ease of access.
' This file uses VBScript to write to both a .csv file and a .html file with .css styling

' ## Where
' Grab all things Nightwatch.js related at http://nightwatchjs.org/
' Grab the Windows Quick Start that precedes this Framework at https://github.com/beatfactor/nightwatch/wiki/Windows-Quick-Start
' Grab these Framework files over at https://github.com/MarkCTest/night-watch-js-example

' ## How
' From the Nightwatch installation, change the parameter below to your reports folder, i.e. in 'GetFiles ("Reports")'.
' Run this VBScript to generate the .csv and .html files examples

'--------------------------------------------------------------------------------------------

Dim oFSO, oOutFile

'Creating File System Object 
Set oFSO = CreateObject("Scripting.FileSystemObject")  

'Create a .csv output file 
Set oOutFile = oFSO.CreateTextFile("nw_List_of_Reports.csv")  

'##############  TEMPLATE ITEMS IN CSV FILE ###########
'Write CSV headers 
oOutFile.WriteLine("Type,File Name,File Path,Status")  

'---------------------------------------------------------------------
'Prepare the HTML file

Dim oFSO1, oOutputFile
Dim sTxtOutput

'Create a .html output file
Set oFSO1 = CreateObject("Scripting.FileSystemObject")
Set sTxtOutput = oFSO1.CreateTextFile("reportViewer.html")

'############# TEMPLATE ITEMS IN HTML FILE ###############

'Create the HTML file with the details the user has entered (move this to a function later)
 sTxtOutput.writeline "<html>"
	  sTxtOutput.writeline "<head>"
		   sTxtOutput.writeline "<style>"
		   sTxtOutput.writeline "h1{font-family:Calibri}"  
		   sTxtOutput.writeline "th{background-color:#502D17; font-family:Calibri; font-size:12pt; color:#FFFFFF}"
		   sTxtOutput.writeline "td{font-family:Calibri}"
		   sTxtOutput.writeline ".footer{font-family:Calibri; font-size:8pt; font-weight:bold}"	   
		   sTxtOutput.writeline "</style>"
  sTxtOutput.writeline "</head>"
    
  sTxtOutput.writeline "<body>"
	  sTxtOutput.writeline "<p style='float: left; clear: left'><img src='http://nightwatchjs.org/img/logo-nightwatch.png' />"
	  sTxtOutput.writeline "<h1>Nightwatch.js - Table of Test Reports</h1></p>"
	  
	   sTxtOutput.writeline "<table border='1'>"
	   
	    sTxtOutput.writeline "<tr>"
		     sTxtOutput.writeline "<th>Item</th>"
		     sTxtOutput.writeline "<th>Report Name</th>"
		     sTxtOutput.writeline "<th>File Path</th>"
		     sTxtOutput.writeline "<th>Test Status</th>"
	    sTxtOutput.writeline "</tr>"

'---------------------------------------------------------------------

'Call the GetFiles function to get all files and folders
GetFiles("reports")  ' full path in Git example is C:\Dev\nightwatch\node_modules\nightwatch\examples\reports

'Close the CSV file 
oOutFile.Close  
WScript.Echo("CSV File created.")  

' Finish writing the HTML file
sTxtOutput.writeline "</tr></table>"
sTxtOutput.writeline "<p class='footer' align='center'>Example framework for NightwatchJS by <a href='#' target='git'>MarkCTest</a>. Use at your own peril.</p>"
sTxtOutput.writeline "</body>"
sTxtOutput.writeline "</html>"

'Close the HTML file
sTxtOutput.Close
WScript.Echo("HTML File created.")

'-----------------------------------------------------------------------

Function GetFiles(FolderName)
	'On Error Resume Next        

	Dim oFolder, oSubFolders
	Dim oFile, oFiles

	Set oFolder = oFSO.GetFolder(FolderName)

	Set oFiles = oFolder.Files

	'Write all FILES to the .csv and .html files   
	For Each oFile In oFiles
		oOutFile.WriteLine("File," & oFile.Name & "," & oFile.Path)
		
		sTxtOutput.writeline "<tr>"
		      sTxtOutput.writeline "<td>File</td>"
		      sTxtOutput.writeline "<td>" & oFile.Name & "</td>"
		      sTxtOutput.writeline "<td>" & oFile.Path & "</td>"
			  sTxtOutput.writeline "<td>Not Set</td>"
		sTxtOutput.writeline "</tr>"
	Next        

	'Get all the subfolders     
	Set oSubFolders = oFolder.SubFolders        

	'Writing SUBFOLDER names to the .csv and .html files 
	For Each oFolder In oSubFolders         
		oOutFile.WriteLine("Folder," & oFolder.Name & "," & oFolder.Path)                 

		sTxtOutput.writeline "<tr style='background-color:#99CCFF; font-weight:bold'>"
		      sTxtOutput.writeline "<td>Folder</td>"
		      sTxtOutput.writeline "<td>" & oFolder.Name &"</td>"
		      sTxtOutput.writeline "<td><a href=' " & oFolder.Path & "' target='reportFolder'>" & oFolder.Path & "</a></td>"
		sTxtOutput.writeline "</tr>"  

		'Getting all Files from subfolder       
		GetFiles(oFolder.Path)    
	Next     

End Function
