' Author: Ben Schneider
' Note: This script was created for People Working Cooperatively
' License: All rights reserved.  Use with permission from Ben Schneider, Cincinnati, OH.
' Modified: 6/4/2015

' Setup Variables
ccRootFolder = "C:\cc\" 'The root where all XML generation job executables live
cXMLFolderName = "pwc" 'The folder name where the XML files are for Content Central to pick up
nfXMLLoadDir = "E:\Natural Forms\Output\10000\" 'Natural Forms root folder
csvFileName = "cc-nf-info.csv"

' Here we go (use caution when modify anything below here)

Const ForReading = 1

Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = Wscript.ScriptFullName
Set objScript = objFSO.GetFile(scriptPath)
csvPath = scriptFolder + csvFileName
Set objFile = objFSO.OpenTextFile(csvPath, ForReading)
scriptFolder = objFSO.GetParentFolderName(objScript)
cXMLFolderPath = ccRootFolder + cXMLFolderName
ccXLMexe = scriptFolder + "\PWC-NaturalFormsXML.exe"
ccZIPdll = scriptFolder + "\Ionic.Zip.dll"

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    arrFields = Split(strLine, ",")

    ' Make life easy, create a variable for the folder name
    createdFolder = arrFields(2) + " " + arrFields(0)
	createdFolderPath = ccRootFolder + createdFolder

    ' Create the folder for the config files to live in
    objFSO.CreateFolder(createdFolderPath)
    
    ' Create and write the contents of the config.ini file to the proper location
    iniFile = ccRootFolder + arrFields(2) + " " + arrFields(0) + "\config.ini"
    Set objINIfile = objFSO.CreateTextFile(iniFile,True)    
    objINIfile.Write "[SETTINGS]" & vbCrLf
    objINIfile.Write "CXMLDIRECTORY=" + cXMLFolderPath & vbCrLf
    objINIfile.Write "CXMLLOADDIRECTORY=" + nfXMLLoadDir + arrFields(0) + "\"  & vbCrLf
    objINIfile.Write "CCATALOG=" + arrFields(1) & vbCrLf
    objINIfile.Write "CDOCTYPE=" + arrFields(2) & vbCrLf
	objINIfile.Write "CATALOGFORIMAGES=" + arrFields(3) & vbCrLf
	objINIfile.Write "DOCTYPEFORIMAGES=" + arrFields(4) & vbCrLf
    objINIfile.Close

    ' Copy the executable program required for the import to run.
    'Set destFolder = createdFolder + "\"
    objFSO.CopyFile ccXLMexe, createdFolderPath + "\", True
	objFSO.CopyFile ccZIPdll, createdFolderPath + "\",True
Loop

' Clean up the folder created by the header of the CSV file
headerFolderName = "CC Doc Type NF ID"
headerFolderPath = ccRootFolder + headerFolderName
If objFSO.FolderExists(headerFolderPath) Then
	objFSO.DeleteFolder(headerFolderPath)
End If

objFile.Close