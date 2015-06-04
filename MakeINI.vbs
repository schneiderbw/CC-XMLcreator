' Author: Ben Schneider
' Note: This script was created for People Working Cooperatively
' License: All rights reserved.  Use with permission from Ben Schneider, Cincinnati, OH.
' Modified: 3/11/2015

Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\cc\cc-nf-info.csv", ForReading)
Const ccXLMexe = "C:\cc\PWC-NaturalFormsXML.exe"

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    arrFields = Split(strLine, ",")

    ' Create the folder for the config files to live in
    objFSO.CreateFolder(arrFields(2) + " " + arrFields(0))

    ' Make life easy, create a variable for the folder name
    'Const createdFolder = arrFields(2) + " " + arrFields(0)
    
    ' Create and write the contents of the config.ini file to the proper location
    iniFile = "C:\cc\" + arrFields(2) + " " + arrFields(0) + "\config.ini"
    Set objINIfile = objFSO.CreateTextFile(iniFile,True)    
    objINIfile.Write "[SETTINGS]" & vbCrLf
    objINIfile.Write "CXMLDIRECTORY=C:\cc\pwc\"  & vbCrLf
    objINIfile.Write "CXMLLOADDIRECTORY=E:\Natural Forms\Output\10000\" + arrFields(0) + "\"  & vbCrLf
    objINIfile.Write "CCATALOG=" + arrFields(1) & vbCrLf
    objINIfile.Write "CDOCTYPE=" + arrFields(2) & vbCrLf
    objINIfile.Close

    ' Copy the executable program required for the import to run.
    'Set destFolder = createdFolder + "\"
    objFSO.CopyFile ccXLMexe, "C:\cc\" + arrFields(2) + " " + arrFields(0) + "\", True
Loop

objFile.Close