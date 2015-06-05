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
	
	'Create the scheduled task XML file
	ccXMLexePath = createdFolderPath + "\PWC-NaturalFormsXML.exe"
	Call genSchedXML(ccXMLexePath)
	objFSO.CopyFile createdFolderPath + "\cc-scheduledtask.xml", ccRootFolder + "\schedulerXML\" + createdFolder + ".xml"
Loop

' Clean up the folder created by the header of the CSV file
headerFolderName = "CC Doc Type NF ID"
headerFolderPath = ccRootFolder + headerFolderName
If objFSO.FolderExists(headerFolderPath) Then
	objFSO.DeleteFolder(headerFolderPath)
End If

objFile.Close

' This subroutine will create the scheduled task XML file
Sub genSchedXML(exePath)
	schedXML = createdFolderPath + "\cc-scheduledtask.xml"
	Set objXMLfile = objFSO.CreateTextFile(schedXML,True)
	objXMLfile.Write "<?xml version=""1.0"" encoding=""UTF-16""?>" & vbCrLf
	objXMLfile.Write "<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">" & vbCrLf
	objXMLfile.Write "  <RegistrationInfo>" & vbCrLf
	objXMLfile.Write "    <Date>2015-03-10T16:00:09.0480063</Date>" & vbCrLf
	objXMLfile.Write "    <Author>PWC\bschneider</Author>" & vbCrLf
	objXMLfile.Write "  </RegistrationInfo>" & vbCrLf
	objXMLfile.Write "  <Triggers>" & vbCrLf
	objXMLfile.Write "    <CalendarTrigger>" & vbCrLf
	objXMLfile.Write "      <Repetition>" & vbCrLf
	objXMLfile.Write "        <Interval>PT5M</Interval>" & vbCrLf
	objXMLfile.Write "        <StopAtDurationEnd>false</StopAtDurationEnd>" & vbCrLf
	objXMLfile.Write "      </Repetition>" & vbCrLf
	objXMLfile.Write "      <StartBoundary>2015-03-10T15:59:32.8707236</StartBoundary>" & vbCrLf
	objXMLfile.Write "      <Enabled>true</Enabled>" & vbCrLf
	objXMLfile.Write "      <ScheduleByDay>" & vbCrLf
	objXMLfile.Write "        <DaysInterval>1</DaysInterval>" & vbCrLf
	objXMLfile.Write "      </ScheduleByDay>" & vbCrLf
	objXMLfile.Write "    </CalendarTrigger>" & vbCrLf
	objXMLfile.Write "  </Triggers>" & vbCrLf
	objXMLfile.Write "  <Principals>" & vbCrLf
	objXMLfile.Write "    <Principal id=""Author"">" & vbCrLf
	objXMLfile.Write "      <UserId>PWC\m0m</UserId>" & vbCrLf
	objXMLfile.Write "      <LogonType>Password</LogonType>" & vbCrLf
	objXMLfile.Write "      <RunLevel>HighestAvailable</RunLevel>" & vbCrLf
	objXMLfile.Write "    </Principal>" & vbCrLf
	objXMLfile.Write "  </Principals>" & vbCrLf
	objXMLfile.Write "  <Settings>" & vbCrLf
	objXMLfile.Write "    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>" & vbCrLf
	objXMLfile.Write "    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>" & vbCrLf
	objXMLfile.Write "    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>" & vbCrLf
	objXMLfile.Write "    <AllowHardTerminate>true</AllowHardTerminate>" & vbCrLf
	objXMLfile.Write "    <StartWhenAvailable>true</StartWhenAvailable>" & vbCrLf
	objXMLfile.Write "    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>" & vbCrLf
	objXMLfile.Write "    <IdleSettings>" & vbCrLf
	objXMLfile.Write "      <StopOnIdleEnd>true</StopOnIdleEnd>" & vbCrLf
	objXMLfile.Write "      <RestartOnIdle>false</RestartOnIdle>" & vbCrLf
	objXMLfile.Write "    </IdleSettings>" & vbCrLf
	objXMLfile.Write "    <AllowStartOnDemand>true</AllowStartOnDemand>" & vbCrLf
	objXMLfile.Write "    <Enabled>true</Enabled>" & vbCrLf
	objXMLfile.Write "    <Hidden>false</Hidden>" & vbCrLf
	objXMLfile.Write "    <RunOnlyIfIdle>false</RunOnlyIfIdle>" & vbCrLf
	objXMLfile.Write "    <WakeToRun>false</WakeToRun>" & vbCrLf
	objXMLfile.Write "    <ExecutionTimeLimit>P3D</ExecutionTimeLimit>" & vbCrLf
	objXMLfile.Write "    <Priority>7</Priority>" & vbCrLf
	objXMLfile.Write "  </Settings>" & vbCrLf
	objXMLfile.Write "  <Actions Context=""Author"">" & vbCrLf
	objXMLfile.Write "    <Exec>" & vbCrLf
	objXMLfile.Write "      <Command>" + chr(34) + exePath + chr(34) + "</Command>"& vbCrLf
	objXMLfile.Write "    </Exec>" & vbCrLf
	objXMLfile.Write "  </Actions>" & vbCrLf
	objXMLfile.Write "</Task>" & vbCrLf
	
End Sub