'//############################
'//  This script by Diagg/OSD-Couture.com
'//   http://www.osd-couture.com/
'// 
'//  Version : 1.61
'//  Release Date : 21/08/2016
'//  Latest Update : 08/11/2017
'//
'//
'//
'//  Change Log : 
'//		06/03/2016	- 1.0	- Initial release
'//					- Added function PinToTaskBar-ActiveSetup , PinToStartMenu-ActiveSetup and Robocopy
'//		12/05/2016	- 1.01	- Added Generate GUID function
'//		21/05/2016	- 1.02	- Added Create Full Shortcut, Pin_Edge and Pin_Store
'//		27/05/2016	- 1.03	- Pin-item function rewritten
'//					- Fixed a bug in Shortcut
'//		04/07/2016	-1.04  - Added function Get_DefaultBrowser
'//					- some rework on shortCut creation
'// 		14/07/2016	-1.05	- Added Disable Services Function (input parameter must be an array  ex:  oSvcs = Array("SERVICE1","SERVICE2","SERVICE3") :  Disable_Services oSvcs )
'//		15/07/2016	-1.06 	- Fixed a bug in shortCut creation
'//		07/08/2016	-1.07	- Added Run as System Function (Requier PsExec)
'//		12/08/2016	-1.08	- Added Pin_IE, Pin_Explorer, Pin_Chrome
'//		21/08/2016	-1.8.1	- Changed all static path to environment variables
'//		02/09/2016	-1.8.2 - Fixed a bug where shortcut were only created on the local admin profil
'//		29/09/2016	-1.8.3 - Fixed bug where Edge & Store were not pinned back
'//		30/09/2016	-1.9	- Fixed an bunch of bugs
'//					- shortCut routine rewritten 
'//		06/12/2016	-1.10	- Added 10 seconds Wait before pinning
'//		07/12/2016	-1.11	- Pinning IE, Edge, Store and Explorer removed and transphered to Set-StartBar.ps1
'//		14/02/2017	-1.12	- Updated Service Function with Start/stop/restart mode and Auto/Disabled/Manual/DelayedAutoStart state  
'//		29/03/2017	-1.20	- Added  WriteToRegistry Function (Allow writting to x64 registry from an x86 environment)
'//		07/04/2017	-1.21	- Fixed a small bug in reg function
'//		19/04/2017	-1.30	- Added New Function AddToTaskBarPinList (add item to pin to the list variabel)
'//					- An argument Was missing in Create_Shortcut function
'//		07/05/2017	-1.40 	- Added CompareVersions function to compare two version numbers
'//		28/07/2017	-1.41 	- Added GetStringbetween function to grab a string between two other known stings
'//					- Added ReadLastLines function to read the X ending lines from a text file
'// 		06/11/2017	-1.50  - Added RunAsTI (Trusted  Installer) Account  Function(Requier PowerRun)
'//					- Added WriteToRegistryAsSystem Function
'//					- Added WriteToRegistryAsTI Function
'//		08/03/2018	-1.60	- Added Support for AutoIT Dll and functions
'//		12/03/2018	-1.61	- Added read retry to ReadLastLines function
'//
'//
'//
'//############################

Dim oAutoIt

Function SetDismFileAssoc (oAppAssocFile)
	oUtility.RunCommand "DISM /online /Import-DefaultAppAssociations:" & oAppAssocFile, 0, True
End Function


Function AddToTaskBarPinList (oPathToAdd)
	Dim dicPinBarList
	Set dicPinBarList = oEnvironment.ListItem("PintoTaskBar")
	dicPinBarList.Add oPathToAdd, ""
	oEnvironment.ListItem("PintoTaskBar") = dicPinBarList
	
End Function


Function PinToTaskBar (oFileToPin)
	PinItem oFileToPin,"TaskBar"
End Function


Function PinToStartMenu (oFileToPin)
	PinItem oFileToPin,"StartMenu"
End Function


Function PinItem (oFileToPin,oPinDest)

	Dim iRet, SyspinPath, sDestinationDrive, oIsLoaded, oCmd, oIsUnLoaded, oDefaultStartupPath,oMyMDTBare

	sDestinationDrive = oUtility.GetOSTargetDriveLetter
	SyspinPath = oEnvironment.item("DeployRoot") &  "\Tools\" & oEnvironment.item("Architecture") & "\syspin.exe"
	oDefaultStartupPath = sDestinationDrive & "\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\startup\"
		
	'Copy Syspin.exe to c:\Windows\System32
	If not oFso.FileExists(sDestinationDrive & "\Windows\system32\Syspin.exe") Then
		If oFso.FileExists(SyspinPath) Then 
			iRet = Robocopy (SyspinPath, sDestinationDrive & "\Windows\system32")
			If iRet = 1 Then 
				oLogging.CreateEntry  "Syspin.exe was copied succesfully to " & sDestinationDrive & "\Windows\system32" , LogTypeInfo
			Else 
				oLogging.CreateEntry "Unable to copy Syspin.exe to " & sDestinationDrive & "\Windows\system32 , unable to pin item, Aborting !!! ", LogTypewarning
				Exit Function			
			End If			
		Else
			oLogging.CreateEntry "Unable to locate : " & SyspinPath & " In MDT deployment Share, unable to pin item, Aborting !!! ", LogTypewarning
			Exit Function
		End If
	End If	
	oLogging.CreateEntry "About to pin : " & oFileToPin & " to " & oPinDest, LogTypeInfo
	' Create Default profile Entry
	oRegName = "ZTIMDT-PinItem-" & Int((Second(Now())* Rnd)*1000) & ".vbs"
	
	If oFso.FileExists(oDefaultStartupPath & "MDT-FullTaskBar.vbs") and oPinDest = "TaskBar" Then
		oLogging.CreateEntry "Found file : " & oDefaultStartupPath & "MDT-FullTaskBar.vbs, Appending to it !", LogTypeInfo
		oMyMDTBare = True
		oRegName = "MDT-FullTaskBar.vbs"
		Set oFile = oFSO.OpenTextFile(oDefaultStartupPath & "MDT-FullTaskBar.vbs" ,8)
	Else
		Set oFile = oFSO.CreateTextFile(oDefaultStartupPath & oRegName ,True)
			oLogging.CreateEntry "Creating file : " & oDefaultStartupPath & oRegName, LogTypeInfo
			oFile.Write "Set oShell = CreateObject(""WScript.Shell"")" & vbCrLf
			oFile.Write "Set oFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & vbCrLf
			oFile.Write "'Wait for Explorer to finish loading" & vbCrLf
			oFile.Write "Wscript.Sleep 10000" & vbCrLf & vbCrLf
			oFile.Write "'Pin Icons" & vbCrLf
	End If		
		
	If oPinDest = "TaskBar" Then
		oFile.Write "oShell.Run ""syspin.exe """"" & oFileToPin & """"" c:5386"",0, True" & vbCrLf
	Else
		oFile.Write "oShell.Run ""syspin.exe """"" & oFileToPin & """"" c:51201"",0, True" & vbCrLf
	End If
		
	If oMyMDTBare <> True Then oFile.Write "oFSO.DeleteFile oShell.ExpandEnvironmentStrings( ""%UserProfile%"" ) & ""\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\startup\"" & WScript.ScriptName, true " & vbCrLf
	
	oFile.Close
	
	oLogging.CreateEntry "Script File " & oRegName & " created to path  " & oDefaultStartupPath, LogTypeInfo

End Function


Function Robocopy (oSourceFile, oDestFolder)
	
	Dim iRet, sFileName, oSourcePath, oCmd
	
	If oFso.FolderExists (oSourceFile) Then
		oCmd = "RoboCopy.Exe " & Chr(34) & oSourceFile & Chr(34) & " " & Chr(34) & oDestFolder & Chr(34) &  " /e /r:1 /w:2"
	ElseIf oFso.FileExists (oSourceFile) Then
		sFileName = oFSO.GetFileName(oSourceFile)
		oSourcePath = oFso.GetParentFolderName(oSourceFile)
		oCmd = "RoboCopy.Exe " & Chr(34) & oSourcePath & Chr(34) & " " & Chr(34) & oDestFolder & Chr(34) & " " & Chr(34) & sFileName & Chr(34) & " /r:1 /w:2"
	Else
		oLogging.CreateEntry "Folder or File : " & oSourceFile & " does not exist, Robocopy unable to processed, Aborting !!! ", LogTypewarning
		Exit Function
	End If	
	
	iRet = oUtility.RunCommand (oCmd, 0, True)
	Robocopy = iRet

End Function

Function Create_Shortcut (oSourceFile, oShortCutName)
	Create_FullShortcut oSourceFile, "", oShortCutName, "", "AllUsersDesktop",""
End Function


Function Create_FullShortcut (oSourceFile, oArguments, oShortCutName, oIconName, oDestination, oIconNumber)

	Dim olnk, oShortCutExt, oAppCommand, oCleanUP, ShortCutPath, sDestinationDrive
	
	oLogging.CreateEntry  "Entering Full Shortcut Function" , LogTypeInfo
	
	sDestinationDrive = oUtility.GetOSTargetDriveLetter
	ShortCutPath = oEnvironment.item("DeployRoot") &  "\Tools\" & oEnvironment.item("Architecture") & "\Shortcut.exe"
		
	'Copy Shortcut.exe to c:\Windows\System32
	If not oFso.FileExists(sDestinationDrive & "\Windows\system32\Shortcut.exe") Then
		If oFso.FileExists(ShortCutPath) Then 
			iRet = Robocopy (ShortCutPath, sDestinationDrive & "\Windows\system32")
			If iRet = 1 Then 
				oLogging.CreateEntry  "Shortcut.exe was copied successfully to " & sDestinationDrive & "\Windows\system32" , LogTypeInfo
			Else 
				oLogging.CreateEntry "Unable to copy Shortcut.exe to " & sDestinationDrive & "\Windows\system32 , unable to create shortcuts, Aborting, Aborting !!! ", LogTypewarning
				Exit Function			
			End If			
		Else
			oLogging.CreateEntry "Unable to locate : " & ShortCutPath & " In MDT deployment Share, unable to create shortcuts, Aborting !!! ", LogTypewarning
			Exit Function
		End If
	End IF
	
	
	' Destination folder can be a fully qualified path or a VBS special folder
	If left(Ucase(oDestination),3) <> "C:\" Then oDestination = oShell.SpecialFolders(oDestination)
	
	If oEnvironment.item("Architecture") = "X64" and left(Ucase(oSourceFile),16) = "C:\PROGRAM FILES" then oSourceFile = ExpandEnvironmentStrings("%ProgramFiles") & Replace (Ucase(oSourceFile), "C:\PROGRAM FILES" ,"" )
	
	If oFso.FileExists (oEnvironment.item("Branding") & "\Icons\" & oIconName) Then
		oLogging.CreateEntry "Copying icon " & oIconName & " from Branding folder to "  & oEnv("ProgramData"), LogTypeInfo
		Robocopy oEnvironment.item("Branding") & "\Icons\" & oIconName, oEnv("ProgramData")
		oIconName = oEnv("ProgramData") & "\" & oIconName
	Else
		oLogging.CreateEntry "Unable to copy icon " & oIconName & " from Branding folder to " & oEnv("ProgramData") & ", assuming it's not necessary !!!", LogTypeInfo
	End If
			
	oLogging.CreateEntry "Creating Short-cut " & oShortCutName & " from source " & oSourceFile & " at location " & oDestination & "\" & oShortCutName & ".lnk with Icon "& oIconName , LogTypeInfo

	'Build command line
	oAppCommand = "Shortcut.exe /F:""" & oDestination & "\" & oShortCutName & ".lnk""" &  " /A:C /T:""" & oSourceFile & """"
	If not oStrings.isNullOrEmpty(oIconName) and not oStrings.isNullOrEmpty(oIconNumber) then oAppCommand = oAppCommand & " /I:""" & oIconName & "," & oIconNumber & """" 
	If not oStrings.isNullOrEmpty(oIconName) then oAppCommand = oAppCommand & " /I:""" & oIconName & """" 
	If not oStrings.isNullOrEmpty(oArguments) then oAppCommand = oAppCommand & " /P:""" & oArguments & """"
	oAppCommand = Replace (oAppCommand, """""", """")

	oUtility.RunCommand oAppCommand, 0, True
	
	oLogging.CreateEntry "--------------------------------------------------------------------------------------------" , LogTypeInfo

End Function


Function GenerateGUID
	
	Set oTypeLib = CreateObject("Scriptlet.TypeLib")
	GenerateGUID = oTypeLib.Guid
	GenerateGUID = replace (GenerateGUID, "{","")
	GenerateGUID = replace (GenerateGUID, "}","")
	Set oTypeLib = Nothing
	
End Function

Function Get_DefaultBrowser
	
	Dim oDefaultBrowser

	'oDefaultBrowser = oShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice\ProgId")
	'If oDefaultBrowser = "AppX4hxtad77fbk3jkkeerkrm0ze94wjf3s9" then 
	'	Get_DefaultBrowser = "start microsoft-edge:"
	'Else
		'oDefaultBrowser = oShell.RegRead("HKEY_CLASSES_ROOT\" & oDefaultBrowser & "\shell\open\command\")
		oDefaultBrowser = oShell.RegRead("HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command\")
		oDefaultBrowser = replace (oDefaultBrowser,"-","")
		oDefaultBrowser = replace (oDefaultBrowser,"%1","")
		oDefaultBrowser = replace (oDefaultBrowser,"""","")
		oDefaultBrowser = Trim(oDefaultBrowser)
		'Get_DefaultBrowser = Chr(34) & oDefaultBrowser & Chr(34)
		Get_DefaultBrowser =  oDefaultBrowser
	'End If
End Function

Function Manage_Services (aTargetSvc,aStartMode,astate)

	'Example: Manage_service MpsSvc, Auto/Disabled/Manual/DelayedAutoStart, Stopped/Running/Restart

	Dim oServiceFound
	oLogging.CreateEntry "Manage services process started", LogTypeInfo
	oLogging.CreateEntry "=========================================================", LogTypeInfo
	Set oWMIService = GetObject("winmgmts:" & "{impersonationlevel=impersonate}!\\.\root\cimv2")
	Set cServices = oWMIService.ExecQuery("SELECT * FROM Win32_Service")
	For Each oService In cServices
		If LCase(oService.Name) = LCase(aTargetSvc) Then
			oServiceFound = 1
			oLogging.CreateEntry "Service " & aTargetSvc & " found on this computer", LogTypeInfo
			oLogging.CreateEntry "Service friendly name is " & oService.DisplayName, LogTypeInfo
			oLogging.CreateEntry "  => Start Mode is  " & oService.StartMode, LogTypeInfo
			oLogging.CreateEntry "  => Status is  " & oService.State, LogTypeInfo
			
			If LCase(aStartMode) = "auto" Then oService.ChangeStartMode("Automatic")
			If LCase(aStartMode) = "disabled" Then oService.ChangeStartMode("Disabled")
			If LCase(aStartMode) = "manual" Then oService.ChangeStartMode("Manual")
			If LCase(aStartMode) = "delayedautostart" Then 
				oService.ChangeStartMode("Automatic")
				oShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\" & aTargetSvc & "\DelayedAutoStart", 1, "REG_DWORD"
			End If	
			
			If LCase(astate) = "stopped" Then oService.StopService()
			If LCase(astate) = "running" Then oService.StartService()
			If LCase(astate) = "restart" Then
				oService.StopService()
				WSCript.Sleep 15000
				oService.StartService()
			End If	
			
			'Wait for action to be done !
			WSCript.Sleep 15000
			
			oLogging.CreateEntry "===== Service " & oService.DisplayName & " updated state is now:", LogTypeInfo
			oLogging.CreateEntry "  => Start Mode is  " & oService.StartMode, LogTypeInfo
			oLogging.CreateEntry "  => Status is  " & oService.State, LogTypeInfo
			Exit For
		Else
			oServiceFound = 0
		End If  
	Next

	If oServiceFound = 0 Then 
		oLogging.CreateEntry "Unable to find Service " & aTargetSvc & " on this computer, Aborting!!!!", LogTypeWarning
		Manage_Services = False
	Else
		Manage_Services = True
	End If	
	
End Function

Function RunAsSystem (oCmdToRun)

	Dim iRet, PsExecPath, sDestinationDrive, oCmd

	sDestinationDrive = oUtility.GetOSTargetDriveLetter
	PsExecPath = oEnvironment.item("DeployRoot") &  "\Tools\" & oEnvironment.item("Architecture") & "\Psexec.exe"
		
	'Copy PsExec to c:\Windows\System32
	If not oFso.FileExists(sDestinationDrive & "\Windows\system32\Psexec.exe") Then
		If oFso.FileExists(PsExecPath) Then 
			iRet = Robocopy (PsExecPath, sDestinationDrive & "\Windows\system32")
			If iRet = 1 Then 
				oLogging.CreateEntry  "Psexec was copied succesfully to " & sDestinationDrive & "\Windows\system32" , LogTypeInfo
			Else 
				oLogging.CreateEntry "Unable to copy Psexec to " & sDestinationDrive & "\Windows\system32 , unable to Run item as system account, Aborting !!! ", LogTypewarning
				Exit Function			
			End If
		Else
			oLogging.CreateEntry "Unable to locate : " & PsExecPath & " In MDT deployment Share, unable to Run item as system account, Aborting !!! ", LogTypewarning
			Exit Function
		End If
	End If

	'Change PsExecPath to the new loaction
	If oFso.FileExists(sDestinationDrive & "\Windows\system32\Psexec.exe") Then
		PsExecPath = sDestinationDrive & "\Windows\system32\Psexec.exe"
	Else
		oLogging.CreateEntry "Unable to locate : " & PsExecPath & " In his destisation folder, unable to Run item as system account, Aborting !!! ", LogTypewarning
		Exit Function
	End If	
	
	'Run command as System
	oLogging.CreateEntry "About to run : " & oCmdToRun & " as system Account " , LogTypeInfo
	oCmd = PsExecPath & " -accepteula -i -s " & oCmdToRun
	
	oLogging.CreateEntry "About to run the followinf command: " & oCmd , LogTypeInfo
	iRet = oUtility.RunWithConsoleLogging (oCmd)
	
	If iRet = 0 Then 
		oLogging.CreateEntry  "Command  " & oCmdToRun & " was run Succefully as System account" , LogTypeInfo
		RunAsSystem = 0
	Else 
		oLogging.CreateEntry "Command  " & oCmdToRun &   "return an error  " & iRet , LogTypewarning
		RunAsSystem = 1
	End If

End Function


Function RunAsTI (oCmdToRun)

	Dim iRet, PsExecPath, sDestinationDrive, oCmd

	sDestinationDrive = oUtility.GetOSTargetDriveLetter
	PsExecPath = oEnvironment.item("DeployRoot") &  "\Tools\" & oEnvironment.item("Architecture") & "\PowerRun.exe"
		
	'Copy PowerRun to c:\Windows\System32
	If not oFso.FileExists(sDestinationDrive & "\Windows\system32\PowerRun.exe") Then
		If oFso.FileExists(PsExecPath) Then 
			iRet = Robocopy (PsExecPath, sDestinationDrive & "\Windows\system32")
			If iRet = 1 Then 
				oLogging.CreateEntry  "PowerRun was copied succesfully to " & sDestinationDrive & "\Windows\system32" , LogTypeInfo
			Else 
				oLogging.CreateEntry "Unable to copy PowerRun to " & sDestinationDrive & "\Windows\system32 , unable to Run item as Trusted Installer account, Aborting !!! ", LogTypewarning
				Exit Function			
			End If
		Else
			oLogging.CreateEntry "Unable to locate : " & PsExecPath & " In MDT deployment Share, unable to Run item as Trusted Installer account, Aborting !!! ", LogTypewarning
			Exit Function
		End If
	End If

	'Change PsExecPath to the new loaction
	If oFso.FileExists(sDestinationDrive & "\Windows\system32\PowerRun.exe") Then
		PsExecPath = sDestinationDrive & "\Windows\system32\PowerRun.exe"
	Else
		oLogging.CreateEntry "Unable to locate : " & PsExecPath & " In his destisation folder, unable to Run item as Trusted Installer account, Aborting !!! ", LogTypewarning
		Exit Function
	End If	
	
	'Run command as Trusted Installer
	oLogging.CreateEntry "About to run : " & oCmdToRun & " as Trusted Installer Account " , LogTypeInfo
	oCmd = PsExecPath & " " & oCmdToRun
	
	oLogging.CreateEntry "About to run the followinf command: " & oCmd , LogTypeInfo
	iRet = oUtility.RunWithConsoleLogging (oCmd)
	
	If iRet = 0 Then 
		oLogging.CreateEntry  "Command  " & oCmdToRun & " was run Succefully as Trusted Installer account" , LogTypeInfo
		RunAsTI = 0
	Else 
		oLogging.CreateEntry "Command  " & oCmdToRun &   "return an error  " & iRet , LogTypewarning
		RunAsTI = 1
	End If

End Function


Function WriteToRegistry (oKeyPath, oKey , oValue, oRegType, oArchTarget)
	WriteToRegistryEx oKeyPath, oKey , oValue, oRegType, oArchTarget, "Admin"
End Function


Function WriteToRegistryAsSystem (oKeyPath, oKey , oValue, oRegType, oArchTarget)
	WriteToRegistryEx oKeyPath, oKey , oValue, oRegType, oArchTarget, "System"
End Function


Function WriteToRegistryAsTI (oKeyPath, oKey , oValue, oRegType, oArchTarget)
	WriteToRegistryEx oKeyPath, oKey , oValue, oRegType, oArchTarget, "TI"
End Function


Function WriteToRegistryEx (oKeyPath, oKey , oValue, oRegType, oArchTarget, oAccount)
	'Example WriteToRegistry "HKEY_CURRENT_USER\SOFTWARE\Diagg", "" , "", "REG_SZ", "x64" -> Create a folder
	'Example  WriteToRegistry "HKCU\SOFTWARE\Diagg", "Motto" , "Understanding Will come with training", "REG_SZ", "x64" -> Create a Key 
	'oRegType = REG_SZ, REG_BINARY, REG_DWORD, REG_MULTI_SZ, REG_EXPAND_SZ
	'oArchTarget =x86/x64 -> if x64 is specified, Allow to write to x64 registry from a x86 Environment (Ex: Sccm or LanDesk)
	
	Dim Ret, oCommand, oExec, OutTxt, oArg		
	
	oLogging.CreateEntry  "###### Writting Key To the Registry ######" , LogTypeInfo
	
	'Check command integrity
	If oRegType = "REG_SZ" Or oRegType  = "REG_BINARY" Or oRegType  = "REG_DWORD" Or oRegType  = "REG_MULTI_SZ" Or oRegType  = "REG_EXPAND_SZ" Then 
		'Ckeck if we need to create a registry Folder
		If oKey= "" and oValue = "" Then
			oLogging.CreateEntry "Detecting a Registry folder Creation" , LogTypeInfo
			oArg = " /ve "
		Else
			oArg = " /v " & chr(34) & oKey & chr(34)
		End if
	Else 
		oLogging.CreateEntry "ERROR: Registry type is not valid (" & oRegType & ") Aborting!" , LogTypewarning
		Exit Function
	End If 
	
	If Ucase(oArchTarget) = "X64" and Ucase(oEnvironment.item("Architecture")) = "X64" Then 
		oCommand = "cmd.exe /C Reg.exe add " & chr(34) & oKeyPath & chr(34) &  oArg  &  " /t "  & oRegType &  " /d " & chr(34) & oValue & chr(34) & " /reg:64 /f"
	Else
		oCommand = "cmd.exe /C Reg.exe add " & chr(34) & oKeyPath & chr(34) &  oArg  &  " /t "  & oRegType &  " /d " & chr(34) & oValue & chr(34) & " /f"
	End If
	
	'#- Run as different Account
	If Ucase(oAccount) = "SYSTEM" Then
		oLogging.CreateEntry "Writing to registry was requested using System Account" , LogTypeInfo
		Ret = RunAsSystem (oCommand)
	ElseIf Ucase(oAccount) = "TI" Then
		oLogging.CreateEntry "Writing to registry was requested using Trusted Installer Account" , LogTypeInfo
		Ret = RunAsTI (oCommand)	
	Else
		oLogging.CreateEntry "Writing to registry was requested using standard Admin Account" , LogTypeInfo
		Ret = oUtility.RunCommand (oCommand, 0, True)
	End If	
		
	'#= Verifiy if the value is written correctly
	If Ret = 0 Then
		wscript.sleep 1000
		If Ucase(oArchTarget) = "X64" and Ucase(oEnvironment.item("Architecture")) = "X64" Then 
			oCommand = "cmd.exe /C Reg.exe query " & chr(34) & oKeyPath & chr(34) &  " /v " & chr(34) & oKey & chr(34) & " /reg:64 "
		Else
			oCommand = "cmd.exe /C Reg.exe Query " & chr(34) & oKeyPath & chr(34) &  " /v " & chr(34) & oKey & chr(34)
		End If
		
		Set oExec = oShell.Exec (oCommand)
		OutTxt = ""
		Do While Not oExec.StdOut.AtEndOfStream
			OutTxt = OutTxt & oExec.StdOut.ReadLine()
		Loop

		oLogging.CreateEntry "Updated Registy value extracted from the registry : " & OutTxt , LogTypeInfo
		
		If InStr(OutTxt,oValue) <> 0 Then
			oLogging.CreateEntry "Registy Key set Succesfully !" , LogTypeInfo
		Else
			oLogging.CreateEntry "ERROR: Unable to set Registy Key!" , LogTypewarning
		End If
		
	Else
		oLogging.CreateEntry "ERROR: Unable to Execute Registry Key, return code is " & Ret , LogTypewarning 
	End If
End Function

Function Lsh(ByVal N, ByVal Bits)
' Bitwise left shift
  Lsh = N * (2 ^ Bits)
End Function

Function GetVersionStringAsArray(ByVal Version)
' Returns a version string "a.b.c.d" as a two-element numeric
' array. The first array element is the most significant 32 bits,
' and the second element is the least significant 32 bits.

	Dim VersionAll, VersionParts, N
	VersionAll = Array(0, 0, 0, 0)
	VersionParts = Split(Version, ".")
	For N = 0 To UBound(VersionParts)
    		VersionAll(N) = CLng(VersionParts(N))
  	Next

	Dim Hi, Lo
  	Hi = Lsh(VersionAll(0), 16) + VersionAll(1)
  	Lo = Lsh(VersionAll(2), 16) + VersionAll(3)

	GetVersionStringAsArray = Array(Hi, Lo)
End Function


Function CompareVersions(ByVal Version1, ByVal Version2)
' Compares two versions "a.b.c.d". If Version1 < Version2,
' returns -1. If Version1 = Version2, returns 0.
' If Version1 > Version2, returns 1.

	Dim Ver1, Ver2, Result
	Ver1 = GetVersionStringAsArray(Version1)
	Ver2 = GetVersionStringAsArray(Version2)
  
	If Ver1(0) < Ver2(0) Then
    		Result = -1
		oLogging.CreateEntry "Version " & Version1 &  " < " & Version2, LogTypeInfo
  	ElseIf Ver1(0) = Ver2(0) Then
    		If Ver1(1) < Ver2(1) Then
      			Result = -1
			oLogging.CreateEntry "Version " & Version1 & " < " & Version2, LogTypeInfo
    		ElseIf Ver1(1) = Ver2(1) Then
      			Result = 0
			oLogging.CreateEntry "Version " & Version1 & " = " & Version2, LogTypeInfo
    		Else
      			Result = 1
			oLogging.CreateEntry "Version " & Version1 & " > " & Version2, LogTypeInfo
    		End If
  	Else
    		Result = 1
		oLogging.CreateEntry "Version " & Version1 & " > " & Version2, LogTypeInfo
  	End If
  	CompareVersions = Result
End Function


 Function GetStringBetween(sSearch , sStart, sStop)
	'Usage: oResult =  GetStringbetween("123 get this 321","123","321")
	' Return the string between the 2 others one 
	
	Dim lSearch, lTemp
 
    lSearch = InStr(1, sSearch, sStart)
    If lSearch > 0 Then
        lSearch = lSearch + Len(sStart)
        lTemp = InStr(lSearch, sSearch, sStop)
        If lTemp > lSearch Then GetStringBetween = Mid(sSearch, lSearch, lTemp - lSearch)
    End If
End Function


Function ReadLastLines (sFile,sAmountOfLines)
	' The function return an array with the last X lines of a file.
	' To be able to call the function with the returning array, 
	' The function must be called using a join or split command to put the receving variable in the correct type
	' Ex 1:   MyLastLines = Join(ReadLastLines ("C:\MININT\SMSOSD\OSDLOGS\ZTIApplications.log",25),"")
	' Ex 2:  MyLastLines = split(MyLastLines,"") : MyLastLines = ReadLastLines ("C:\MININT\SMSOSD\OSDLOGS\ZTIApplications.log",25)
	
  	Dim oFile, j, NumberOfRecords, oContent
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	' Get file's amount of lines
	Set oFile = oFso.OpenTextFile(sFile, ForReading)
		oContent = oFile.ReadAll
		Wscript.sleep 3000
		NumberOfRecords = oFile.Line - 1
	oFile.Close
	oLogging.CreateEntry "Total number of Lines in file " &  sFile & " is "  & NumberOfRecords  , LogTypeInfo

	
	'Check if file is not too small
	If NumberOfRecords < sAmountOfLines Then sAmountOfLines = NumberOfRecords
	
	'Check If File is not empty
	If sAmountOfLines = 0 or NumberOfRecords = 0 Then 
		oLogging.CreateEntry "ERROR: File is empty, unable to proceed!" , LogTypeInfo
		ReadLastLines(0) = Null
		Exit Function
	Else 
		oLogging.CreateEntry "Retrieving last " &  sAmountOfLines & " lines from file.", LogTypeInfo
	End If
		

	'Open file and skip line until we reach the requested amount	
	Set oFile = oFso.OpenTextFile(sFile, ForReading)
		For j = 1 To NumberOfRecords - sAmountOfLines
			oFile.SkipLine
		Next
		
		'Return the last lines in an Array
		ReadLastLines = Split((oFile.readall), vbLf)
	oFile.Close		
 
End Function


Function RegisterAutoIT

	Dim  sAuToAX, oRootTools, oRegister
	
	'//Check if AutoIT is Not already registered
	If oStrings.isNullOrEmpty (oAutoIt) Then

		'// Register AutoIT 
		oRootTools = oEnvironment.Item("DeployRoot") & "\Tools\"
		'oRootTools = "E:\Deploy\Tools\" ' Debug

		If oFSO.FileExists(oRootTools & oEnvironment.Item("Architecture") & "\AutoItX3.dll") then
			sAuToAX = oRootTools & oEnvironment.Item("Architecture") & "\AutoItX3.dll"
			
			'// Register the DLL
			oLogging.CreateEntry "RUN: regsvr32.exe /s """ & sAuToAX & """", LogTypeInfo
			oRegister = oUtility.RunWithConsoleLogging ("regsvr32.exe /s """ & sAuToAX & """") ' Always returns 0 - Success
			
			If oRegister = 0  Then
				Set oAutoIt = CreateObject("AutoItX3.Control")
				RegisterAutoIT = true
				oLogging.CreateEntry "AutoIT Dll registered Successfully !", LogTypeInfo
			Else 		
				Register-AutoIT = false
				oLogging.CreateEntry "[WARNING] Unable to register AutoIT Dll !", LogTypeWarning
			End IF	
		Else 
				RegisterAutoIT = false
				oLogging.CreateEntry "[WARNING] " & oRootTools & oEnvironment.Item("Architecture") & "\AutoItX3.dll File not found, Unable to register AutoIT Dll !", LogTypeWarning
		End if
	
	Else
		RegisterAutoIT = true
		oLogging.CreateEntry "AutoIT Dll, already registered Successfully, Nothing to do !", LogTypeInfo
	End If

End Function