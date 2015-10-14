'============================================================================================== DISCLAIMER

' This script, macro, and other code examples for illustration only, without warranty either expressed or implied, including but not
' limited to the implied warranties of merchantability and/or fitness for a particular purpose. This script is provided 'as is' and the Author does not
' guarantee that the following script, macro, or code can be used in all situations.

'============================================================================================== DECLARATIONS

	Option Explicit
	'On Error Resume Next

'============================================================================================== USER CONFIGURABLE CONSTANTS

	Const Author			=	"David Segura"
	Const AuthorEmail		=	"david@segura.org"
	Const Company			=	""
	Const Script			=	"MyInstall.vbs"
	Const Description		=	"VBScript Wrapper for BAT/EXE/MSI/MSP/MSU/INF Installations"
	Const Release			=	"https://winpeguy.wordpress.com/myinstall/"
	Const Reference			=	"https://winpeguy.wordpress.com/2015/06/24/reference-myinstall-config/"

	Const Title 			=	"MyInstall"
	Const Version 			=	20150922
	Const VersionFull 		=	20150922.01
	Dim TitleVersion		:	TitleVersion = Title & " (" & Version & ")"
	
	Const SupportContact	=	"David Segura"	'I don't really offer support
	Const SupportEmail		=	"david@segura.org"
	Const SupportAction		=	"Submit an Incident for Technical Support"
	Const SupportGroup		=	"SupportGroup"
	Const SupportArea		=	"SupportArea"
	Const SupportSubject	=	"MyInstall Issue"
	Const SupportProblem	=	"Complete Description of Problem Including All Logs"

'============================================================================================== SYSTEM CONSTANTS

	Const ForReading			=	1
	Const ForWriting			=	2
	Const ForAppending			=	8
	Const OverwriteExisting 	=	True

	Const HKEY_CLASSES_ROOT		= 	&H80000000
	Const HKEY_CURRENT_USER		= 	&H80000001
	Const HKEY_LOCAL_MACHINE	= 	&H80000002
	Const HKEY_USERS			= 	&H80000003
	Const HKEY_CURRENT_CONFIG	= 	&H80000005

'============================================================================================== OBJECTS

	Dim objComputer				: 	objComputer				=	"."
	'Dim objComputer			: 	objComputer				=	GetObject("WinNT://.,computer")
	Dim objShell				: 	Set objShell			=	CreateObject("Wscript.Shell")
	Dim objShellApp				: 	Set objShellApp			=	CreateObject("Shell.Application")
	Dim objFSO					: 	Set objFSO 				=	CreateObject("Scripting.FileSystemObject")
	Dim objDictionary			: 	Set objDictionary		=	CreateObject("Scripting.Dictionary")
	Dim objWMIService			: 	Set objWMIService 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\cimv2")
	Dim objRegistry				: 	Set objRegistry 		=	GetObject("winmgmts:{impersonationLevel = impersonate}!\\" & objComputer & "\root\default:StdRegProv")

'============================================================================================== VARIABLES: SYSTEM
	
	Dim MyUserName				: 	MyUserName				= Lcase(objShell.ExpandEnvironmentStrings("%UserName%"))
	Dim MyComputerName			: 	MyComputerName			= Ucase(objShell.ExpandEnvironmentStrings("%ComputerName%"))
	Dim MyTemp					: 	MyTemp					= Lcase(objShell.ExpandEnvironmentStrings("%Temp%"))
	Dim MyWindir				: 	MyWindir				= Lcase(objShell.ExpandEnvironmentStrings("%Windir%"))
	Dim MySystemDrive			: 	MySystemDrive			= Lcase(objShell.ExpandEnvironmentStrings("%SystemDrive%"))
	Dim MyArchitecture			: 	MyArchitecture			= Lcase(objShell.ExpandEnvironmentStrings("%Processor_Architecture%"))
	If MyArchitecture		= "amd64" Then MyArchitecture = "x64"
	Dim MyExitCode				: 	MyExitCode				= 0
	'Alternate Method using Function GetVar
	'Dim MyUserName				: 	MyUserName				= Lcase(GetVar("%UserName%"))
	'Dim MyComputerName			: 	MyComputerName			= Ucase(GetVar("%ComputerName%"))
	'Dim MyTemp					: 	MyTemp					= Lcase(GetVar("%Temp%"))
	'Dim MyWindir				: 	MyWindir				= Lcase(GetVar("%Windir%"))
	'Dim MySystemDrive			: 	MySystemDrive			= Lcase(GetVar("%SystemDrive%"))
	'Dim MyArchitecture			: 	MyArchitecture			= Lcase(GetVar("%Processor_Architecture%"))
	
'============================================================================================== VARIABLES: CURRENT DIRECTORY

	Dim MyScriptFullPath		: 	MyScriptFullPath			= Wscript.ScriptFullName							'Full Path and File Name with Extension
	Dim MyScriptFileName		: 	MyScriptFileName			= objFSO.GetFileName(MyScriptFullPath)				'File Name with Extension
	Dim MyScriptBaseName		: 	MyScriptBaseName			= objFSO.GetBaseName(MyScriptFullPath)				'File Name 
	Dim MyScriptParentFolder	: 	MyScriptParentFolder		= objFSO.GetParentFolderName(MyScriptFullPath)		'Current Directory (Parent Folder)
	Dim MyScriptGParentFolder	: 	MyScriptGParentFolder		= objFSO.GetParentFolderName(MyScriptParentFolder)	'Parent of the Current Directory (Parent of the Parent Folder)
	Dim arrNames				:	arrNames					= Split(MyScriptParentFolder, "\")
	Dim intIndex				:	intIndex					= Ubound(arrNames)
	Dim MyParentFolderName		:	MyParentFolderName			= arrNames(intIndex)
	'Dim MyScriptConfigFile		:	MyScriptConfigFile			= MyScriptParentFolder & "\" & MyScriptBaseName & ".config"

'============================================================================================== VARIABLES: LOGGING

	'Only one line below must be uncommented
	'Dim MyLogFile				: 	MyLogFile					= MyTemp & "\" & Title & ".log"						'Places the LOG in the Temp Directory	
	'Dim MyLogFile				: 	MyLogFile					= MyScriptParentFolder & "\" & Title & ".log"		'Places the LOG in the Script Directory
	'Dim MyLogFile				:	MyLogFile					= MyScriptParentFolder & "\" & "MyInstall Script " & MyParentFolderName & ".log"
	Dim MyLogFile				:	MyLogFile					= MyTemp & "\" & "MyInstall Script " & MyParentFolderName & ".log"
	
	'Only one line below must be uncommented
	Dim DoLogging				: 	DoLogging					= True		'Creates a LOG
	'Dim DoLogging				: 	DoLogging					= False		'Prevents a LOG from being written

	'Only one line below must be uncommented
	'Dim TextFormat				: 	TextFormat					= True		'Results in a TEXT formatted LOG
	Dim TextFormat				: 	TextFormat					= False		'Results in a CMTRACE formatted LOG (default)
	
	LogStart					'Generate the LOG file

	'Identify
	'Dim MyLogInstall			:	MyLogInstall				= MyScriptParentFolder & "\" & "MyInstall " & MyParentFolderName & ".log"
	Dim MyLogInstall			:	MyLogInstall				= MyTemp & "\" & "MyInstall " & MyParentFolderName & ".log"

'==============================================================================================
	TraceLog "================================================================================= Processing System Checks", 1
	'Gets the current date as 8 digit like 20150505
	Dim MyFullDate				:	MyFullDate					= Year(Date) & Right(String(2, "0") & Month(date), 2) & Right(String(2, "0") & Day(date), 2)
	Dim objTextFile
	Dim Return
	Dim Failed
	
	IsAdmin						'Will return IsAdmin = True if it is running with Admin Rights
	IsSystem					'Checks to see if this is running under the System Account
	CheckArguments				'Checks for Command Line Arguments

	
'==============================================================================================
	TraceLog "================================================================================= Processing Operating System", 1
	Dim MyOperatingSystem
	GetMyOperatingSystem		'Checks the Operating System.  We can stop specific OS's in this Sub

'==============================================================================================
	TraceLog "================================================================================= Processing Computer Information", 1
	Dim MyComputerManufacturer, MyComputerModel, MyBIOSVersion
	GetMyComputerInfo

'==============================================================================================
	TraceLog "================================================================================= Processing Elevation", 1
	Dim sArgumentUAC
	'DoElevate					'To force this script to run Elevated, uncomment this line
	
'==============================================================================================
	TraceLog "================================================================================= Processing MDT Information", 1
	
	Dim MDTDeployRoot
	Dim MDTApplicationSuccessCodes
	GetMyMDTInfo
	
	'Dim oTSEnv, oVar
	'On Error Resume Next
	'Set oTSEnv = CreateObject("Microsoft.SMS.TSEnvironment")
	'For Each oVar In oTSEnv.GetVariables
		'TraceLog oVar & "=" & oTSEnv(oVar), 2
	'Next
	
	' /////////////////////////////////////////////////////////
	' Check MDT Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyMDTInfo
		TraceLog "Sub GetMyMDTInfo", 1
		Dim oTSEnv, oVar
		On Error Resume Next
		Set oTSEnv = CreateObject("Microsoft.SMS.TSEnvironment")
		For Each oVar In oTSEnv.GetVariables
			'TraceLog oVar & "=" & oTSEnv(oVar), 2
			
			If oVar = Ucase("DEPLOYROOT") Then
				MDTDeployRoot = oTSEnv(oVar)
				TraceLog "<Variable> MDTDeployRoot = " & MDTDeployRoot, 2
			End If
		Next
		Set oTSEnv = NOTHING
		
	End Sub

'==============================================================================================
	TraceLog "================================================================================= Setting Default Configuration", 1
	'Variables are in the following Priority: (1) Script (2) Config File (3) Models File
	'If a configuration is <> "" in this file, then the config file will be used (if it exists)

	Dim CreateConfigFile		:	CreateConfigFile			= "Yes"					'Do we use a Config File?  Default should be "Yes" so one can be generated on first use
	Dim ConfigFile				:	ConfigFile					= "MyInstall.config"	'This is the file name of the Config File.
	Dim ConfigModels			:	ConfigModels				= "MyInstall.models.txt"'This is the file name of the Models Config File
	Dim ConfigPNPID				:	ConfigPNPID					= "MyInstall.pnpids.txt"

	'   MyInstall Script Properties
	Dim cSimulation				:	cSimulation					= True					'Enter anything between the quotes for Test Only mode.  Only a blank "" entry will allow this to run in Production Mode
	Dim cElevate				:	cElevate					= True					'Controls if UAC Elevation needs to be used on this script.  This needs to be set to "Yes" to run Elevated.  "Yes" is the default in the Config File
	Dim cConfirm				:	cConfirm					= True					'Enabled to prompt before and after the installation
	Dim cLocalPath				:	cLocalPath					= "C:\Sample"			'Copies Content locally before executing
	
	'   Setup File Information
	Dim cSetupFile				:	cSetupFile					= "Source\Sample.exe"	'Use this entry if you have a Single Install File.
	Dim cSetupx86				:	cSetupx86					= ""					'Use this entry if x86 and x64 have separate install files.  cSetupFile must be "" for this to function
	Dim cSetupx64				:	cSetupx64					= ""					'Use this entry if x86 and x64 have separate install files.  cSetupFile must be "" for this to function
	
	'   Setup Switches
	Dim cSetupSwitches			:	cSetupSwitches				= ""					'Switches after the Command Line

	'   Running Processes
	Dim cWaitForProcess			:	cWaitForProcess				= "SampleProcess.exe"	'Enter the name of a file to determine if the installation is complete.  For example, ccmsetup.exe

	'   Restart Actions
	Dim cRebootWithMDT			:	cRebootWithMDT				= False
	Dim cRebootWithOS			:	cRebootWithOS				= False
	Dim cMDTHideProgress		:	cMDTHideProgress			= False
	
	'   Operating System Compatibility
	Dim cConditionOS			:	cConditionOS				= ""					'(Comma Separated List If Multiple) Valid = Windows XP,Windows Vista,Windows 7,Windows 8,Windows 8.1,Server 2012 
	Dim cConditionArch			:	cConditionArch				= ""					'(Comma Separated List If Multiple) Valid = x86,x64
	Dim cOSWindowsXP			:	cOSWindowsXP				= True
	Dim cOSWindowsVista			:	cOSWindowsVista				= True
	Dim cOSWindows7				:	cOSWindows7					= True
	Dim cOSWindows8				:	cOSWindows8					= True
	Dim cOSWindows81			:	cOSWindows81				= True
	Dim cOSWindows10			:	cOSWindows10				= True
	Dim cOSServer2003			:	cOSServer2003				= False
	Dim cOSServer2008			:	cOSServer2008				= False
	Dim cOSServer2008R2			:	cOSServer2008R2				= False
	Dim cOSServer2012			:	cOSServer2012				= False
	Dim cOSServer2012R2			:	cOSServer2012R2				= False
	Dim cOSServer10				:	cOSServer10					= False

	'   Hardware Compatibility
	Dim cOSArchitecturex86		:	cOSArchitecturex86			= True
	Dim cOSArchitecturex64		:	cOSArchitecturex64			= True
	Dim cComputerManufacturer	:	cComputerManufacturer		= ""					'(Comma Separated List If Multiple) Valid = Dell,HP
	Dim cComputerModel			:	cComputerModel				= ""					'(Comma Separated List If Multiple)	Leave cComputerModel "" if you are using a MyModels.txt file in the same directory as the Script
	Dim cComputerPNPID			:	cComputerPNPID				= ""					'(Comma Separated List If Multiple) Valid = HDAUDIO\FUNC_01&VEN_10EC
	Dim cComputerBIOSVerMin		:	cComputerBIOSVerMin			= "A99"					'Run the installation if the BIOS is this version or newer
	Dim cComputerBIOSVerMax		:	cComputerBIOSVerMax			= "A99"					'Run the installation if the BIOS is this version or older
																						'For an A04 BIOS Update, you want to set this to A03 (1 revision back)
	
	'   Shortcut Creation
	Dim cShortcut1Location		:	cShortcut1Location			= ""					'Location where to place the shortcut.  Must be a .lnk file
	Dim cShortcut1TargetPath	:	cShortcut1TargetPath		= ""					'Target of the shortcut
	Dim cShortcut1WorkingDir	:	cShortcut1WorkingDir		= ""					'Working Directory of the shortcut
	Dim cShortcut2Location		:	cShortcut2Location			= ""					'Location where to place the shortcut.  Must be a .lnk file
	Dim cShortcut2TargetPath	:	cShortcut2TargetPath		= ""					'Target of the shortcut
	Dim cShortcut2WorkingDir	:	cShortcut2WorkingDir		= ""					'Working Directory of the shortcut
	Dim cShortcut3Location		:	cShortcut3Location			= ""					'Location where to place the shortcut.  Must be a .lnk file
	Dim cShortcut3TargetPath	:	cShortcut3TargetPath		= ""					'Target of the shortcut
	Dim cShortcut3WorkingDir	:	cShortcut3WorkingDir		= ""					'Working Directory of the shortcut

'==============================================================================================
	TraceLog "================================================================================= Generating Config File if Necessary", 1
	If CreateConfigFile = "Yes" And IsPathWriteable(MyScriptParentFolder) And Not objFSO.FileExists(MyScriptParentFolder & "\" & ConfigFile) Then
		GenerateConfigFile
		GenerateModelsFile
		GeneratePNPIDFile
		CreateLogShortcut
		ConfigFilePrompt
	End If

	Sub GenerateConfigFile
		TraceLog "Generating " & ConfigFile, 2
		Dim objTextStream
		Set objTextStream = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigFile, 2, True)
		objTextStream.WriteLine "'//	This config file is in VBScript format"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== MyInstall Script Properties ===================================="
		objTextStream.WriteLine "	cSimulation                 = "	& cSimulation
		objTextStream.WriteLine "	cConfirm                    = "	& cConfirm
		objTextStream.WriteLine "	cElevate                    = "	& cElevate
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Setup File Information ========================================="
		objTextStream.WriteLine "	cSetupFile                  = """	& cSetupFile			& """"
		objTextStream.WriteLine "	cSetupx86                   = """	& cSetupx86				& """"
		objTextStream.WriteLine "	cSetupx64                   = """	& cSetupx64				& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Setup Switches ================================================="
		objTextStream.WriteLine "	cSetupSwitches              = """	& cSetupSwitches		& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Local Path ====================================================="
		objTextStream.WriteLine "	cLocalPath                  = """	& cLocalPath			& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Running Processes =============================================="
		objTextStream.WriteLine "	cWaitForProcess             = """	& cWaitForProcess		& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Reboot Action =================================================="
		objTextStream.WriteLine "	cRebootWithMDT              = "	& cRebootWithMDT
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== MDT Progress Action ============================================"
		objTextStream.WriteLine "	cMDTHideProgress            = "	& cMDTHideProgress
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Operating System Compatibility ================================="
		objTextStream.WriteLine "	cOSWindowsXP                = "	& cOSWindowsXP
		objTextStream.WriteLine "	cOSWindowsVista             = "	& cOSWindowsVista
		objTextStream.WriteLine "	cOSWindows7                 = "	& cOSWindows7
		objTextStream.WriteLine "	cOSWindows8                 = "	& cOSWindows8
		objTextStream.WriteLine "	cOSWindows81                = "	& cOSWindows81
		objTextStream.WriteLine "	cOSWindows10                = "	& cOSWindows10
		objTextStream.WriteLine "	cOSServer2003               = "	& cOSServer2003
		objTextStream.WriteLine "	cOSServer2008               = "	& cOSServer2008
		objTextStream.WriteLine "	cOSServer2008R2             = "	& cOSServer2008R2
		objTextStream.WriteLine "	cOSServer2012               = "	& cOSServer2012
		objTextStream.WriteLine "	cOSServer2012R2             = "	& cOSServer2012R2
		objTextStream.WriteLine "	cOSServer10                 = "	& cOSServer10
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Hardware Compatibility ========================================="
		objTextStream.WriteLine "	cComputerManufacturer       = """	& cComputerManufacturer	& """"
		objTextStream.WriteLine "	cComputerModel              = """	& cComputerModel		& """"
		objTextStream.WriteLine "	cComputerPNPID              = """	& cComputerPNPID		& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== BIOS Compatibility ============================================="
		objTextStream.WriteLine "	cComputerBIOSVerMin         = """	& cComputerBIOSVerMin	& """"
		objTextStream.WriteLine "	cComputerBIOSVerMax         = """	& cComputerBIOSVerMax	& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Shortcut Creation =============================================="
		objTextStream.WriteLine "	cShortcut1Location          = """	& cShortcut1Location	& """"
		objTextStream.WriteLine "	cShortcut1TargetPath        = """	& cShortcut1TargetPath	& """"
		objTextStream.WriteLine "	cShortcut1WorkingDir        = """	& cShortcut1WorkingDir	& """"
		objTextStream.WriteLine "	'==============================================================="
		objTextStream.WriteLine "	cShortcut2Location          = """	& cShortcut2Location	& """"
		objTextStream.WriteLine "	cShortcut2TargetPath        = """	& cShortcut2TargetPath	& """"
		objTextStream.WriteLine "	cShortcut2WorkingDir        = """	& cShortcut2WorkingDir	& """"
		objTextStream.WriteLine "	'==============================================================="
		objTextStream.WriteLine "	cShortcut3Location          = """	& cShortcut3Location	& """"
		objTextStream.WriteLine "	cShortcut3TargetPath        = """	& cShortcut3TargetPath	& """"
		objTextStream.WriteLine "	cShortcut3WorkingDir        = """	& cShortcut3WorkingDir	& """"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== Script Information (No changes should be made below) ==========="
		objTextStream.WriteLine "	cScriptVersion      = " & Version
		objTextStream.WriteLine ""
		objTextStream.WriteLine "'====== About =========================================================="
		objTextStream.WriteLine "'//	" & Script & " Configuration File"
		objTextStream.WriteLine "'//	" & ConfigFile & " Last Updated " & MyFullDate
		objTextStream.WriteLine "'//	"
		objTextStream.WriteLine "'//	" & Author
		objTextStream.WriteLine "'//	" & Release
		objTextStream.WriteLine "'//	" & Reference
		objTextStream.WriteLine "'//	"
		objTextStream.WriteLine "'//	Supported Extensions for this script are *.*/BAT/CMD/EXE/MSI/MSP/MSU/INF/VBS"
		objTextStream.Close
	End Sub
		
	Sub GenerateModelsFile
		TraceLog "Generating " & ConfigModels & ".sample", 2
		Dim objTextStream
		Set objTextStream = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigModels & ".sample", 2, True)
		objTextStream.WriteLine "Latitude E6430"
		objTextStream.WriteLine "Latitude E6530"
		objTextStream.WriteLine "Latitude E6440"
		objTextStream.WriteLine "Latitude E6540"
		objTextStream.Close
	End Sub
	
	Sub GeneratePNPIDFile
		TraceLog "Generating " & ConfigPNPID & ".sample", 2
		Dim objTextStream
		Set objTextStream = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigPNPID & ".sample", 2, True)
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0269&SUBSYS_1028052C"
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0280&SUBSYS_102805A1"
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0280&SUBSYS_102805D2"
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0280&SUBSYS_102805D3"
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0280&SUBSYS_102805D4"
		objTextStream.WriteLine "HDAUDIO\FUNC_01&VEN_10EC&DEV_0280&SUBSYS_10280620"
		objTextStream.Close
	End Sub
	
	Sub CreateLogShortcut
		TraceLog "Generating shortcut to " & MyTemp & "\" & "MyInstall Script " & MyParentFolderName & ".log", 2
		If objFSO.FileExists(MyScriptParentFolder & "\" & "MyInstall Script " & MyParentFolderName & " Log.lnk") Then
			objFSO.DeleteFile(MyScriptParentFolder & "\" & "MyInstall Script " & MyParentFolderName & " Log.lnk")
		End If

		Dim oShellLink
		Set oShellLink = objShell.CreateShortcut(MyScriptParentFolder & "\" & "MyInstall Script " & MyParentFolderName & " Log.lnk")
			oShellLink.TargetPath = MyTemp & "\" & "MyInstall Script " & MyParentFolderName & ".log"
			oShellLink.WindowStyle = 1
			oShellLink.IconLocation = "shell32.dll, 26"
			oShellLink.Description = "MyInstall Log File"
			oShellLink.WorkingDirectory = MyScriptParentFolder
			oShellLink.Save
	End Sub
	
	Sub ConfigFilePrompt
		TraceLog "Exiting " & Script, 2
		MsgBox	"A Config File has been generated at " & MyScriptParentFolder & "\" & ConfigFile & vbCrLf & "" & vbCrLf & _
				"Complete the information in this file before running this " & Script & " again" & vbCrLf & "" & vbCrLf & _
				"Press OK to open the " & ConfigFile & " file for editing",64,MyParentFolderName
		Dim sCmd
		sCmd = "notepad.exe " & "" & MyScriptParentFolder & "\" & ConfigFile & ""
		objShell.Run sCmd
		Wscript.Quit
	End Sub
	
'==============================================================================================
	TraceLog "================================================================================= Reading Configuration File", 1

	Dim f: Set f = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigFile,ForReading)
	Dim s: s = f.ReadAll()
	ExecuteGlobal s
	
	'Remove Samples from Variables before processing
	cLocalPath				= Replace(cLocalPath,"C:\Sample","")
	cWaitForProcess			= Replace(cWaitForProcess,"SampleProcess.exe","")
	cComputerBIOSVerMin		= Replace(cComputerBIOSVerMin,"A99","")
	cComputerBIOSVerMax		= Replace(cComputerBIOSVerMax,"A99","")
	
	cSetupSwitches			= Replace(cSetupSwitches,"%MyScriptParentFolder%",MyScriptParentFolder)			'Replace the SourceDir variable with the installation parent
	
	TraceLog "<Variable> cSimulation = " 			& cSimulation, 1
	TraceLog "<Variable> cElevate = " 				& cElevate, 1
	TraceLog "<Variable> cConfirm = " 				& cConfirm, 1
	TraceLog "<Variable> cLocalPath = " 			& cLocalPath, 1
	TraceLog "<Variable> cSetupFile = " 			& cSetupFile, 1
	TraceLog "<Variable> cSetupx86 = " 				& cSetupx86, 1
	TraceLog "<Variable> cSetupx64 = " 				& cSetupx64, 1
	TraceLog "<Variable> cSetupSwitches = " 		& cSetupSwitches, 1
	TraceLog "<Variable> cWaitForProcess = " 		& cWaitForProcess, 1
	TraceLog "<Variable> cRebootWithMDT = " 		& cRebootWithMDT, 1
	TraceLog "<Variable> cRebootWithOS = " 			& cRebootWithOS, 1
	TraceLog "<Variable> cMDTHideProgress = " 		& cMDTHideProgress, 1
	TraceLog "<Variable> cOSWindowsXP = " 			& cOSWindowsXP, 1
	TraceLog "<Variable> cOSWindowsVista = " 		& cOSWindowsVista, 1
	TraceLog "<Variable> cOSWindows7 = " 			& cOSWindows7, 1
	TraceLog "<Variable> cOSWindows8 = " 			& cOSWindows8, 1
	TraceLog "<Variable> cOSWindows81 = " 			& cOSWindows81, 1
	TraceLog "<Variable> cOSWindows10 = " 			& cOSWindows10, 1
	TraceLog "<Variable> cOSServer2003 = " 			& cOSServer2003, 1
	TraceLog "<Variable> cOSServer2008 = " 			& cOSServer2008, 1
	TraceLog "<Variable> cOSServer2008R2 = " 		& cOSServer2008R2, 1
	TraceLog "<Variable> cOSServer2012 = " 			& cOSServer2012, 1
	TraceLog "<Variable> cOSServer2012R2 = " 		& cOSServer2012R2, 1
	TraceLog "<Variable> cOSServer10 = " 			& cOSServer2012R2, 1
	TraceLog "<Variable> cComputerManufacturer = " 	& cComputerManufacturer, 1
	TraceLog "<Variable> cComputerModel = " 		& cComputerModel, 1
	TraceLog "<Variable> cComputerPNPID = " 		& cComputerPNPID, 1
	TraceLog "<Variable> cComputerBIOSVerMin = " 	& cComputerBIOSVerMin, 1
	TraceLog "<Variable> cComputerBIOSVerMax = " 	& cComputerBIOSVerMax, 1
	TraceLog "<Variable> cShortcut1Location = " 	& cShortcut1Location, 1
	TraceLog "<Variable> cShortcut1TargetPath = " 	& cShortcut1TargetPath, 1
	TraceLog "<Variable> cShortcut1WorkingDir = " 	& cShortcut1WorkingDir, 1
	TraceLog "<Variable> cShortcut2Location = " 	& cShortcut2Location, 1
	TraceLog "<Variable> cShortcut2TargetPath = " 	& cShortcut2TargetPath, 1
	TraceLog "<Variable> cShortcut2WorkingDir = " 	& cShortcut2WorkingDir, 1
	TraceLog "<Variable> cShortcut3Location = " 	& cShortcut3Location, 1
	TraceLog "<Variable> cShortcut3TargetPath = " 	& cShortcut3TargetPath, 1
	TraceLog "<Variable> cShortcut3WorkingDir = " 	& cShortcut3WorkingDir, 1
	TraceLog "<Variable> cScriptVersion = "			& Version, 1

'==============================================================================================
	TraceLog "================================================================================= Updating Configuration File", 1
	If cScriptVersion <> Version Then
		If CreateConfigFile = "Yes" And IsPathWriteable(MyScriptParentFolder) Then
			GenerateConfigFile
			TraceLog "Updating Config file " & ConfigFile, 2
		Else
			TraceLog "Could not update Config file " & ConfigFile, 3
		End If
	Else
		TraceLog "Update of Config file " & ConfigFile & " is not necessary", 1
	End If


'==============================================================================================
	TraceLog "================================================================================= Checking Elevation Requirements", 1
	If IsAdmin = False And WScript.Arguments.length = 0 And cElevate <> "" Then DoElevate

'==============================================================================================
	TraceLog "================================================================================= Reading Configuration File for Computer Models", 1
	If cComputerModel = "" and objFSO.FileExists(MyScriptParentFolder & "\" & ConfigModels) Then
		Set objTextFile = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigModels, ForReading)
		cComputerModel = objTextFile.ReadAll
		objTextFile.Close
		TraceLog cComputerModel, 1
	End If


'==============================================================================================
	TraceLog "================================================================================= Reading Configuration File for Computer PNPIDS", 1
	If cComputerPNPID = "" and objFSO.FileExists(MyScriptParentFolder & "\" & ConfigPNPID) Then
		Set objTextFile = objFSO.OpenTextFile(MyScriptParentFolder & "\" & ConfigPNPID, ForReading)
		cComputerPNPID = objTextFile.ReadAll
		objTextFile.Close
		TraceLog cComputerPNPID, 1
	End If
	

'==============================================================================================
	TraceLog "================================================================================= Evaluating Setup File", 1
	Dim sMySetupFile
	subGetSetupFile
	
	Sub subGetSetupFile
		If cSetupFile <> "" Then
			TraceLog "<Variable> cSetupFile = " & cSetupFile, 1
			sMySetupFile = cSetupFile
			cConditionArch = "x86 x64"
		Else
			If cSetupx86 <> "" Then cConditionArch = "x86 "
			If cSetupx64 <> "" Then cConditionArch = cConditionArch & "x64"
			If MyArchitecture = "x86" Then sMySetupFile = cSetupx86
			If MyArchitecture = "x64" Then sMySetupFile = cSetupx64
		End If
	End Sub
	TraceLog "<Variable> sMySetupFile = " & sMySetupFile, 1
	
'==============================================================================================
	TraceLog "================================================================================= Building Command Line", 1
	
	Dim sCmdLine
	subBuildCommandLine
	TraceLog "<Variable> sCmdLine = " & sCmdLine, 1
	
	Sub subBuildCommandLine
		Dim sCmdLineMSI		:	sCmdLineMSI		= "msiexec.exe /qb-! /l*vx """ & MyLogInstall & """" & " REBOOT=ReallySuppress UILevel=67 ALLUSERS=2 /i """ & MyScriptParentFolder & "\" & sMySetupFile & """"
		Dim sCmdLineMSP		:	sCmdLineMSP		= "msiexec.exe /q /l*v """ & MyLogInstall & """" & " /p """ & MyScriptParentFolder & "\" & sMySetupFile & """"
		Dim sCmdLineMSU		:	sCmdLineMSU		= """" & MyScriptParentFolder & "\" & sMySetupFile & """" & " /quiet /norestart"
		Dim sCmdLineINF		:	sCmdLineINF		= "RunDll32 advpack.dll,LaunchINFSection """ & MyScriptParentFolder & "\" & sMySetupFile & """,DefaultInstall"
		Dim sCmdLineEXE		:	sCmdLineEXE		= """" & MyScriptParentFolder & "\" & sMySetupFile & """"
		Dim sCmdLineCMD		:	sCmdLineCMD		= "cmd /c " & """" & MyScriptParentFolder & "\" & sMySetupFile & """"
		Dim sCmdLineVBS		:	sCmdLineVBS		= "cscript " & """" & MyScriptParentFolder & "\" & sMySetupFile & """"
		
		If Lcase(Right(sMySetupFile,4))		= ".msi" Then
			sCmdLine = sCmdLineMSI
		ElseIf Lcase(Right(sMySetupFile,4)) = ".msp" Then
			sCmdLine = sCmdLineMSP
		ElseIf Lcase(Right(sMySetupFile,4)) = ".msu" Then
			sCmdLine = sCmdLineMSU
		ElseIf Lcase(Right(sMySetupFile,4)) = ".inf" Then
			sCmdLine = sCmdLineINF
		ElseIf Lcase(Right(sMySetupFile,4)) = ".exe" Then
			sCmdLine = sCmdLineEXE
		ElseIf Lcase(Right(sMySetupFile,4)) = ".cmd" Then
			sCmdLine = sCmdLineCMD
		ElseIf Lcase(Right(sMySetupFile,4)) = ".bat" Then
			sCmdLine = sCmdLineCMD
		ElseIf Lcase(Right(sMySetupFile,4)) = ".vbs" Then
			sCmdLine = sCmdLineVBS
		Else
			sCmdLine = sCmdLineEXE
		End If
		
		If cSetupSwitches <> "" Then sCmdLine = sCmdLine & " " & cSetupSwitches
	End Sub

'==============================================================================================
	TraceLog "================================================================================= Setup File Information", 1
	If NOT objFSO.FileExists(MyScriptParentFolder & "\" & sMySetupFile) Then 
		TraceLog "Setup File does not exist.  Installation will fail.", 3
		TraceLog "Switching to Simulation Mode", 2
		cSimulation = True
	End If
	
	If cSimulation = True Then
		TraceLog "cSimulation has been set to True.  " & Script & " is running in Simulation Mode", 2
	Else
		TraceLog Script & " is NOT running in Simulation Mode", 2
	End If

'==============================================================================================
	TraceLog "================================================================================= Checking to see if we should run from a Local Path", 1
	Dim LocalPathComplete
	If cLocalPath <> "" Then
		'cLocalPath was specified so we have to copy the content to the specified location
		If Lcase(MyScriptParentFolder) = Lcase(cLocalPath & "\" & MyParentFolderName) Then
			'Local Path was specified and we are executing from Local Path, so nothing to copy here
		Else
			'Local Path was specified and we are not executing from Local Path, so we need to copy the content to Local Path
			'Hopefully Robocopy is on the system.  Will look at updating to fall back to xcopy
			objShell.Run "cmd /c Robocopy """ & MyScriptParentFolder & """ """ & cLocalPath & "\" & MyParentFolderName & """" & " *.* /E /NDL /R:0 /W:0 /Z", 1, True
			'Content has been copied now we run the Local Copy of this Script
			objShell.Run "cscript """ & cLocalPath & "\" & MyParentFolderName & "\" & MyScriptFileName & """", 1, True
			'Make a note that this is the initial Script not the spawned Local Copy
			LocalPathComplete = True
		End If
	End If

	'Check to see if we were the Parent or the Spawn MyInstall.vbs
	If LocalPathComplete = True Then
		'So if we ran the spawn Local Copy, this must be the parent, and there is nothing further to complete in the Parent Script
	End If


'==============================================================================================
	TraceLog "================================================================================= Validating Conditions", 1
	If LocalPathComplete <> True Then CheckConditions
	
	Sub CheckConditions
		If cOSWindowsXP = True Then 	cConditionOS = "Windows XP" & vbCrLf
		If cOSWindowsVista = True Then 	cConditionOS = cConditionOS & "Windows Vista" & vbCrLf
		If cOSWindows7 = True Then 		cConditionOS = cConditionOS & "Windows 7" & vbCrLf
		If cOSWindows8 = True Then 		cConditionOS = cConditionOS & "Windows 8" & vbCrLf
		If cOSWindows81 = True Then 	cConditionOS = cConditionOS & "Windows 8.1" & vbCrLf
		If cOSWindows10 = True Then 	cConditionOS = cConditionOS & "Windows 10" & vbCrLf
		If cOSServer2003 = True Then 	cConditionOS = cConditionOS & "Windows Server 2003" & vbCrLf
		If cOSServer2008 = True Then 	cConditionOS = cConditionOS & "Windows Server 2008" & vbCrLf
		If cOSServer2008R2 = True Then 	cConditionOS = cConditionOS & "Windows Server 2008 R2" & vbCrLf
		If cOSServer2012 = True Then 	cConditionOS = cConditionOS & "Windows Server 2012" & vbCrLf
		If cOSServer2012R2 = True Then 	cConditionOS = cConditionOS & "Windows Server 2012 R2"
		If cOSServer10 = True Then 		cConditionOS = cConditionOS & "Windows Server 10"
		
		TraceLog "<Variable> cConditionOS = " & cConditionOS, 1
	
		If cConditionOS <> "" Then
			If Instr(Lcase(cConditionOS),Lcase(MyOperatingSystem)) = 0 Then
				TraceLog "Operating System Condition Failed. " & MyOperatingSystem & " is NOT allowed", 3
				TraceLog "Approved Operating Systems: " & cConditionOS, 3
				Failed = True
			Else
				TraceLog "Operating System Condition Passed. " & MyOperatingSystem & " is allowed", 1
			End If
		Else
			TraceLog "No Operating System Condition Exists. " & MyOperatingSystem & " is allowed", 1
		End If

		TraceLog "<Variable> cConditionArch = " & cConditionArch, 1
		If cConditionArch <> "" Then
			If Instr(Lcase(cConditionArch),Lcase(MyArchitecture)) = 0 Then
				TraceLog "OS Architecture Condition Failed. " & MyArchitecture & " is NOT allowed", 3
				TraceLog "Approved OS Architecture: " & cConditionArch, 3
				Failed = True
			Else
				TraceLog "OS Architecture Condition Passed. " & MyArchitecture & " is allowed", 1
			End If
		Else
			TraceLog "No OS Architecture Condition Exists. " & MyArchitecture & " is allowed", 1
		End If
		
		TraceLog "<Variable> cComputerManufacturer = " & cComputerManufacturer, 1
		If cComputerManufacturer <> "" Then
			If Instr(Lcase(cComputerManufacturer),Lcase(MyComputerManufacturer)) = 0 Then
				TraceLog "Computer Manufacturer Condition Failed. " & MyComputerManufacturer & " is NOT allowed", 3
				Failed = True
			Else
				TraceLog "Computer Manufacturer Condition Passed. " & MyComputerManufacturer & " is allowed", 1
			End If
		Else
			TraceLog "No Computer Manufacturer Condition Exists. " & MyComputerManufacturer & " is allowed", 1
		End If

		TraceLog "<Variable> cComputerModel = " & cComputerModel, 1
		If cComputerModel <> "" Then
			If Instr(Lcase(cComputerModel),Lcase(MyComputerModel)) = 0 Then
				TraceLog "Computer Model Condition Failed. " & MyComputerModel & " is NOT allowed", 3
				Failed = True
			Else
				TraceLog "Computer Model Condition Passed. " & MyComputerModel & " is allowed", 1
			End If
		Else
			TraceLog "No Computer Model Condition Exists. " & MyComputerModel & " is allowed", 1
		End If

		TraceLog "<Variable> cComputerPNPID = " & cComputerPNPID, 1
		If cComputerPNPID <> "" Then
			TraceLog "PNPID Condition Exists.  Checking Hardware", 1
			
			Dim a, i, PNPCheckOK
			a = Split(cComputerPNPID,vbCrLf)
			For i=0 to ubound(a)
				If a(i) = "" Then Exit For
				If PNPCheck(a(i)) = True Then
					PNPCheckOK = True
				End If
			Next

			If PNPCheckOK = True Then
				TraceLog "PNPID was Found.  Installation is allowed", 1
			Else
				TraceLog "PNPID was NOT Found.  Installation is NOT allowed", 3
				Failed = True
			End If
		Else
			TraceLog "No PNPID Condition Exists.  Installation is allowed", 1
		End If
		
		TraceLog "<Variable> MyBIOSVersion = " & MyBIOSVersion, 1
		TraceLog "<Variable> cComputerBIOSVerMin = " & cComputerBIOSVerMin, 1
		TraceLog "<Variable> cComputerBIOSVerMax = " & cComputerBIOSVerMax, 1
	
		If MyBIOSVersion <> "" Then
		'We were able to get our BIOS Version
			If cComputerBIOSVerMin <> "" and MyBIOSVersion < cComputerBIOSVerMin Then
				TraceLog "BIOS Version of this Computer is less than the Minimum BIOS Version for this installation", 3
				TraceLog "Installation will not proceed", 3
				Failed = True
			End If
			If cComputerBIOSVerMax <> "" and MyBIOSVersion > cComputerBIOSVerMax Then
				TraceLog "BIOS Version of this Computer is greater than the Maximum BIOS Version for this installation", 3
				TraceLog "Installation will not proceed", 3
				Failed = True
			End If
		End If
		
		If Failed = True Then
			TraceLog "A required Condition was not met.  Switching to Simulation Mode", 3
			cSimulation = True
			TraceLog "cSimulation has been set to True.  " & Script & " is running in Simulation Mode", 2
		End If
	End Sub

'==============================================================================================
'==============================================================================================
TraceLog "================================================================================= Processing cConfirm Prompt", 1
	'If we required a confirmation, this is where it shall be
	If LocalPathComplete <> True Then
		If cConfirm = True Then
			MsgBox	"Starting Installation of " & MyParentFolderName & vbCrLf & "" & vbCrLf & _
			"Command Line:" & vbCrLf & _
			sCmdLine & vbCrLf & "" & vbCrLf & _
			"Supported Operating Systems:" & vbCrLf & _
			cConditionOS & vbCrLf & "" & vbCrLf & _
			"Supported OS Architectures: " & cConditionArch & vbCrLf & "" & vbCrLf & _
			"Supported Computer Manufacturers:" & vbCrLf & _
			cComputerManufacturer & vbCrLf & "" & vbCrLf & _
			"Supported Computer Models:" & vbCrLf & _
			cComputerModel & vbCrLf & _
			"Supported PNPIDS:" & vbCrLf & _
			cComputerPNPID,64,MyParentFolderName & " Confirmation Prompt"
		End If
	End If
		
'==============================================================================================
'==============================================================================================
TraceLog "================================================================================= Starting Installation", 1
On Error Resume Next
	If cMDTHideProgress = True Then
		Dim oTSProgressUI
		Set oTSProgressUI = CreateObject("Microsoft.SMS.TSProgressUI") 
		oTSProgressUI.CloseProgressDialog 
		Set oTSProgressUI = Nothing
	End If
	
	If LocalPathComplete <> True Then
		If cSimulation = True Then
			TraceLog "Test Only Command Line: " & sCmdLine, 2
		ElseIf cSimulation = False Then
			TraceLog "Running Command Line: " & sCmdLine, 1
			If objFSO.FileExists(MyScriptParentFolder & "\MyInstallAuto.vbs") Then
				objShell.Run "Wscript """ & MyScriptParentFolder & "\MyInstallAuto.vbs""", 1, false
				Wscript.sleep 2000
			ElseIf objFSO.FileExists(MyScriptParentFolder & "\Automatica.ps1") Then
				objShell.Run "PowerShell -ExecutionPolicy Bypass &'" & MyScriptParentFolder & "\Automatica.ps1'", 1, false
				Wscript.sleep 2000
			ElseIf objFSO.FileExists(MyScriptParentFolder & "\MySendKeys.ps1") Then
				objShell.Run "PowerShell -ExecutionPolicy Bypass &'" & MyScriptParentFolder & "\MySendKeys.ps1'", 1, false
				Wscript.sleep 2000
			End If
			Return = objShell.Run(sCmdLine, 1, true)
			
			'If we specified a process to wait for, then we need to account for that here
			If cWaitForProcess <> "" Then subWaitForProcess
			
			CreateAppShortcut
			
			If cRebootWithMDT = True Then
				TraceLog "<Variable> cRebootWithMDT = " & cRebootWithMDT, 1
				TraceLog "Processing MDT Reboot Action", 1
				
				If objFSO.FileExists(MDTDeployRoot & "\Scripts\ZTISetVariable.wsf") Then
					TraceLog "ZTISetVariable.wsf was located at " & MDTDeployRoot & "\Scripts", 2
					
					''''''TraceLog "<Running Command> cscript //nologo """ & MDTDeployRoot & "\Scripts\ZTISetVariable.wsf""" & " /VariableName:cRebootWithMDT /VariableValue:True", 2
					'''''''objShell.Run "cscript """ & MDTDeployRoot & "\Scripts\ZTISetVariable.wsf""" & " /VariableName:cRebootWithMDT /VariableValue:True", 1, false
					
					TraceLog "Setting Application Exit Code 1641 ERROR_SUCCESS_REBOOT_INITIATED", 2
					MyExitCode = 1641
					TraceLog "<Variable> MyExitCode = " & MyExitCode, 2

				Else
					TraceLog "ZTISetVariable.wsf could not be found.  MDT Reboot Action could not be set", 1
				End If
			End If
			
		End If
		'If we required a confirmation, this is where it shall be
		If cConfirm = True Then
			MsgBox	"Completed Installation of " & MyParentFolderName,64,MyParentFolderName & " Confirmation Prompt"
		End If

		If cSimulation = True Then TraceLog "Simulation Complete at " & Now, 1
		If cSimulation = False Then TraceLog "Installation Complete at " & Now, 1
		Wscript.Quit MyExitCode
	End If
'==============================================================================================
'==============================================================================================
'==============================================================================================







'==============================================================================================
	Sub subWaitForProcess
		Wscript.Sleep 2000
		Wscript.Echo "Checking for running process " & cWaitForProcess
		Dim colProcessList, ProcessCount
		Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & cWaitForProcess & "'")
		ProcessCount = colProcessList.count
		Wscript.Sleep 2000

		Do While ProcessCount > 0
			Wscript.Echo "Waiting 10 seconds for " & cWaitForProcess & " to complete"
			Wscript.Sleep 10000
			Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & cWaitForProcess & "'")
			ProcessCount = colProcessList.count
		Loop
		
		Wscript.Sleep 2000
	End Sub
'==============================================================================================
'==============================================================================================
'==============================================================================================
	Sub CreateAppShortcut
		Dim oShellLink
		If cShortcut1Location <> "" Then
			Set oShellLink = objShell.CreateShortcut(cShortcut1Location)
				oShellLink.TargetPath = cShortcut1TargetPath
				oShellLink.WindowStyle = 1
				oShellLink.WorkingDirectory = cShortcut1WorkingDir
				oShellLink.Save
		End If
		If cShortcut2Location <> "" Then
			Set oShellLink = objShell.CreateShortcut(cShortcut2Location)
				oShellLink.TargetPath = cShortcut2TargetPath
				oShellLink.WindowStyle = 1
				oShellLink.WorkingDirectory = cShortcut2WorkingDir
				oShellLink.Save
		End If
		If cShortcut3Location <> "" Then
			Set oShellLink = objShell.CreateShortcut(cShortcut3Location)
				oShellLink.TargetPath = cShortcut3TargetPath
				oShellLink.WindowStyle = 1
				oShellLink.WorkingDirectory = cShortcut3WorkingDir
				oShellLink.Save
		End If
	End Sub
'==============================================================================================
'==============================================================================================




	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
'==============================================================================================
	' /////////////////////////////////////////////////////////
	'	Function to check if Hardware is present
	' /////////////////////////////////////////////////////////
	Function PNPCheck(cComputerPNPID)
		
		Dim PNPID	:	PNPID = Replace(cComputerPNPID,"\","\\")
		Tracelog "Running WMI Query: Select * FROM Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPID & "%'", 1
		
		On Error Resume Next
		Dim colItems, objItem
		Set colItems = objWMIService.ExecQuery("Select * From Win32_PnPEntity WHERE DeviceID LIKE '%" & PNPID & "%'")
		TraceLog "Found " & colItems.Count & " devices that match", 2

		For Each objItem in colItems
			TraceLog "DeviceID:" & objItem.DeviceID, 2
		Next
		If colItems.Count > 0 Then
			'TraceLog "PNPCheck Passed", 3
			PNPCheck = True
		Else
			'TraceLog "PNPCheck Failed", 3
			PNPCheck = False
		End If
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	'	Function to check if Path is Writable
	' /////////////////////////////////////////////////////////
	
	Function IsPathWriteable(Path)
		Dim Temp_Path 'As String
		
		Temp_Path = Path & "\" & objFSO.GetTempName() & ".drs"
		
		On Error Resume Next
			objFSO.CreateTextFile Temp_Path
			IsPathWriteable = Err.Number = 0
			objFSO.DeleteFile Temp_Path
		On Error Goto 0
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	'	Function to save Environment Variable to a Variable
	'
	'	Usage:	MyWindir = Lcase(GetVar("%Windir%"))
	'	Result:	MyWindir = c:\windows
	' /////////////////////////////////////////////////////////
	
	Function GetVar(sVar)
		'Using Windows Shell, return the value of an environment variable
		GetVar = objShell.ExpandEnvironmentStrings(sVar)
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we have Admin Rights
	' /////////////////////////////////////////////////////////
	'	Usage:	If IsAdmin = False Then Wscript.Quit
	'	Result:	Script will exit
	'
	'	Usage:	If IsAdmin = False Then DoElevate
	'	Result:	Script will run the DoElevate Subroutine
	
	Function IsAdmin
		'LogLine
		'TraceLog "Function IsAdmin", 1
		
		Dim RegKey
		IsAdmin = False
		On Error Resume Next
		
		'Try to read a Registry Key that is only readable with Admin Rights
		RegKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\")
		If Err.Number = 0 Then IsAdmin = True
		
		'Log Result
		If IsAdmin = True Then TraceLog "<IsAdmin = True> User has Admin Rights", 1
		If IsAdmin = False Then TraceLog "<IsAdmin = False> User does not have Admin Rights", 1
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check if we are running under SYSTEM Account
	' /////////////////////////////////////////////////////////
	Function IsSystem
		'LogLine
		'TraceLog "Function IsSystem", 1
		
		IsSystem = False
		
		'Determine if we are running this under the System Account and LOG result
		If Lcase(CreateObject("WScript.Network").UserName) = "system" Then
			IsSystem = True
			TraceLog "<IsSystem = True> Script is being run under the SYSTEM context, possibly from SCCM or as a Scheduled Task", 2
		Else
			TraceLog "<IsSystem = False> Script is NOT being run under the SYSTEM context", 1
		End If
	End Function
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check for Command Line Arguments
	' /////////////////////////////////////////////////////////
	Sub CheckArguments
		'LogLine
		TraceLog "Sub CheckArguments", 1

		Dim sArgument, sArguments
		Set sArguments = Wscript.Arguments
		
		If sArguments.Count = 0 Then TraceLog "No Passed Arguments", 1
		
		For Each sArgument in sArguments
			TraceLog "<Variable> sArgument = " & sArgument, 1
			If Lcase(sArgument) = "uac"	Then sArgumentUAC = True
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Relaunch Elevated
	' /////////////////////////////////////////////////////////
	Sub DoElevate
		'LogLine
		'TraceLog "Sub DoElevate", 1
		
		'If IsAdmin = True Then
		'	TraceLog "No need to elevate as we already have the right permissions", 3
		'	Exit Sub
		'End If
		If sArgumentUAC = True Then
			TraceLog "No need to elevate as we already have the right permissions", 3
			Exit Sub
		End If
		
		TraceLog "Relaunching Elevated", 3
		LogLine
		LogLine
		LogLine
		LogLine
		LogLine
		'objShellApp.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
		objShellApp.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 0
		WScript.Quit
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check Operating System Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyOperatingSystem
		'LogLine
		TraceLog "Sub GetMyOperatingSystem", 1
		
		Dim objItem, colItems
		Dim Unsupported
		
		On Error Resume Next
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objItem In colItems
			TraceLog "<Property> Caption: " & objItem.Caption,1
			TraceLog "<Property> OperatingSystemSKU: " & objItem.OperatingSystemSKU,1
			TraceLog "<Property> Organization: " & objItem.Organization,1
			TraceLog "<Property> OSArchitecture: " & objItem.OSArchitecture,1
			TraceLog "<Property> OSProductSuite: " & objItem.OSProductSuite,1
			TraceLog "<Property> OSType: " & objItem.OSType,1
			TraceLog "<Property> ProductType: " & objItem.ProductType,1
			TraceLog "<Property> RegisteredUser: " & objItem.RegisteredUser,1
			TraceLog "<Property> SerialNumber: " & objItem.SerialNumber,1
			TraceLog "<Property> Status: " & objItem.Status,1
			TraceLog "<Property> SuiteMask: " & objItem.SuiteMask,1
			TraceLog "<Property> Version: " & objItem.Version,1
			
			With objItem
			Select Case True
				'Client Operating Systems
				Case Left(.Version,3) = "5.1" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "5.2" and .ProductType = 1
					MyOperatingSystem = "Windows XP"
				Case Left(.Version,3) = "6.0" and .ProductType = 1
					MyOperatingSystem = "Windows Vista"
				Case Left(.Version,3) = "6.1" and .ProductType = 1
					MyOperatingSystem = "Windows 7"
				Case Left(.Version,3) = "6.2" and .ProductType = 1
					MyOperatingSystem = "Windows 8"
				Case Left(.Version,3) = "6.3" and .ProductType = 1
					MyOperatingSystem = "Windows 8.1"
				Case Left(.Version,3) = "10." and .ProductType = 1
					MyOperatingSystem = "Windows 10"
				'Server Operating Systems
				Case Left(.Version,3) = "5.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2003"
				Case Left(.Version,3) = "6.0" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008"
				Case Left(.Version,3) = "6.1" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2008 R2"
				Case Left(.Version,3) = "6.2" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012"
				Case Left(.Version,3) = "6.3" and .ProductType > 1
					MyOperatingSystem = "Windows Server 2012 R2"
				Case Left(.Version,3) = "10." and .ProductType > 1
					MyOperatingSystem = "Windows Server 10"
				End Select
			End With
			
			'If MyOperatingSystem = "" Then MyOperatingSystem = objOperatingSystem.Caption
			If MyOperatingSystem = "" or Unsupported = True Then
				MyOperatingSystem = objItem.Caption
				TraceLog "<Property> MyOperatingSystem = " & MyOperatingSystem, 3
				TraceLog MyOperatingSystem & " is not supported by this Script", 3
				Wscript.Quit
			Else
				TraceLog "<Variable> MyOperatingSystem = " & MyOperatingSystem, 2
				TraceLog "<Variable> MyArchitecture = " & MyArchitecture, 2
			End If
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Check Computer Properties
	' /////////////////////////////////////////////////////////
	Sub GetMyComputerInfo
		'LogLine
		TraceLog "Sub GetMyComputerInfo", 1
		
		Dim objItem, colItems
		Dim Unsupported
		
		On Error Resume Next
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objItem In colItems
			TraceLog "<Property> DNSHostName: "				& objItem.DNSHostName,1
			TraceLog "<Property> Domain: "					& objItem.Domain,1
			TraceLog "<Property> DomainRole: "				& objItem.DomainRole,1
			TraceLog "<Property> Manufacturer: "			& objItem.Manufacturer,1
			TraceLog "<Property> Model: "					& objItem.Model,1
			TraceLog "<Property> PartOfDomain: "			& objItem.PartOfDomain,1
			TraceLog "<Property> PrimaryOwnerName: "		& objItem.PrimaryOwnerName,1
			TraceLog "<Property> TotalPhysicalMemory: "		& objItem.TotalPhysicalMemory,1

			MyComputerManufacturer = Trim(objItem.Manufacturer)
			
			With objItem
			Select Case True
				Case Instr(.Manufacturer, "Dell") > 0
					MyComputerManufacturer	= "Dell"
				Case Instr(.Manufacturer, "Microsoft") > 0
					MyComputerManufacturer	= "Microsoft"
			End Select
			End With
			
			MyComputerModel	= Trim(objItem.Model)
			
			TraceLog "<Variable> MyComputerManufacturer = " & MyComputerManufacturer, 2
			TraceLog "<Variable> MyComputerModel = " 		& MyComputerModel, 2
		Next
		
		Dim BIOS
		For Each BIOS in GetObject("winmgmts:\\.\root\cimv2").InstancesOf("Win32_BIOS")  
			MyBIOSVersion = Trim(Ucase(BIOS.SMBIOSBIOSVERSION))
			TraceLog "<Property> MyBIOSVersion: "		& MyBIOSVersion,1
		Next
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Create a Shortcut
	' /////////////////////////////////////////////////////////
	Sub CreateShortcut(sShortcut,sTargetPath,sWorkingDir)
		On Error Resume Next
	
		'Delete Existing Shortcut
		If objFSO.FileExists(sShortcut) Then objFSO.DeleteFile(sShortcut)

		Dim oShellLink
		Set oShellLink = objShell.CreateShortcut(sShortcut)
		oShellLink.TargetPath = sTargetPath
		oShellLink.WindowStyle = 1
		oShellLink.WorkingDirectory = sWorkingDir
		oShellLink.Save
	End Sub
'==============================================================================================
	' /////////////////////////////////////////////////////////
	' Create a URL Shortcut
	' /////////////////////////////////////////////////////////
	Sub CreateShortcutURL(sShortcut,sTargetPath)
		On Error Resume Next
	
		'Delete Existing Shortcut
		If objFSO.FileExists(sShortcut) Then objFSO.DeleteFile(sShortcut)

		Dim oShellLink
		Set oShellLink = objShell.CreateShortcut(sShortcut)
		oShellLink.TargetPath = sTargetPath
		oShellLink.Save
	End Sub
'==============================================================================================
'==============================================================================================
'==============================================================================================
'==============================================================================================
'============================================================================================== REFERENCE: DIALOG BOXES

	REM Constant			Value			Description
	REM vbOKOnly				0			Display OK button only.
	REM vbOKCancel				1			Display OK and Cancel buttons.
	REM vbAbortRetryIgnore		2			Display Abort, Retry, and Ignore buttons.
	REM vbYesNoCancel			3			Display Yes, No, and Cancel buttons.
	REM vbYesNo					4			Display Yes and No buttons.
	REM vbRetryCancel			5			Display Retry and Cancel buttons.
	REM vbCritical				16			Display Critical Message icon.
	REM vbQuestion				32			Display Warning Query icon.
	REM vbExclamation			48			Display Warning Message icon.
	REM vbInformation			64			Display Information Message icon.
	REM vbDefaultButton1		0			First button is default.
	REM vbDefaultButton2		256			Second button is default.
	REM vbDefaultButton3		512			Third button is default.
	REM vbDefaultButton4		768			Fourth button is default.
	REM vbApplicationModal		0			Application modal; the user must respond to the message box before continuing work in the current application.
	REM vbSystemModal			4096		System modal; all applications are suspended until the user responds to the message box.
	REM vbMsgBoxHelpButton		16384		Adds Help button to the message box
	REM VbMsgBoxSetForeground	65536		Specifies the message box window as the foreground window
	REM vbMsgBoxRight			524288		Text is right aligned
	REM vbMsgBoxRtlReading		1048576		Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
	
'============================================================================================== FUNCTIONS: TRACE LOGGING

	' /////////////////////////////////////////////////////////
	' Logging Function with Trace Log
	' /////////////////////////////////////////////////////////
	Function TraceLog(LogText, LogError)
		Dim LogTemp
		Dim FileOut, MyLogFileX, TitelX, Tst
	
		If DoLogging = False Then Exit Function
		
		If TextFormat = True Then
			If LogError = 0 Then
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForWriting, True)
			Else
				Set FileOut = objFSO.OpenTextFile( MyLogFile, ForAppending, True)
			End If
			FileOut.WriteLine Now()& " - " & LogText
			FileOut.Close
			Set FileOut = Nothing
			Exit Function
		End If
	
		'***********************************************************
		' Write Trace32 / CMTrace compatible log file
		' logfile - syntax (SMS Trace)
		' <![LOG[...]LOG]!>
		' <
		'    time="04:00:54.309+-60"
		'    date="03-14-2008"
		'    component="SrcUpdateMgr"
		'    context=""
		'    type="0"
		'    thread="1812"
		'    file="productpackage.cpp:97"
		' >
		'
		'    "context="		will not display
		'    type="0"		TraceLog-procedure delete logfile an create new logfile
		'    type="1"		display as normally line
		'    type="2"		display as yellow line / warn
		'    type="3"		display as red line / error
		'    type="F"		display as red line / error

		'    "thread="		number, display as "Tread:", example "Tread: 33 (0x21)"
		'    "file="		diplay as "Source:"

		On Error Resume Next
		Tst = KeineLog
		On Error Goto 0
		If UCase( Tst ) = "JA" Then Exit Function

		On Error Resume Next
		TitelX = Titel
		' if not set 'Titel' outside procedure 'TitelX' is empty
		TitelX = title
		' if not set 'title' outside procedure 'TitelX' is empty

		If Len( TitelX ) < 2 Then TitelX = document.title
		' set title in .HTA
		If Len( TitelX ) < 2 Then TitelX = WScript.ScriptName
		' set title in .VBS
		On Error Goto 0

		On Error Resume Next
		MyLogFileX = MyLogFile
		' if not set 'MyLogFile' outside procedure, 'MyLogFileX' is empty
		If Len( MyLogFileX ) < 2    Then MyLogFileX = WScript.ScriptFullName & ".log"' .vbs
		If Len( MyLogFileX ) < 2    Then MyLogFileX = TitelX & ".log"        ' .hta
		On Error Goto 0

		' Enumerate Milliseconds
		Tst = Timer()               ' timer() in USA: 1234.22; dot separation
		Tst = Replace( Tst, "," , ".")        ' timer() in german: 23454,12; comma separation
		If InStr( Tst, "." ) = 0 Then Tst = Tst & ".000"
		Tst = Mid( Tst, InStr( Tst, "." ), 4 )
		If Len( Tst ) < 3 Then Tst = Tst & "0"

		' Enumerate Time Zone
		Dim AktDMTF : Set AktDMTF = CreateObject("WbemScripting.SWbemDateTime")
		AktDMTF.SetVarDate Now(), True : Tst = Tst & Mid( AktDMTF, 22 ) ' : MsgBox Tst, , "099 :: "
		' MsgBox "AktDMTF: '" & AktDMTF & "'", , "100 :: "
		Set AktDMTF = Nothing
		LogTemp = LogText
		LogTemp = "<![LOG[" & LogTemp & "]LOG]!>"
		LogTemp = LogTemp & "<"
		LogTemp = LogTemp & "time=""" & Hour( Time() ) & ":" & Minute( Time() ) & ":" & Second( Time() ) & Tst & """ "
		LogTemp = LogTemp & "date=""" & Month( Date() ) & "-" & Day( Date() ) & "-" & Year( Date() ) & """ "
		LogTemp = LogTemp & "component=""" & TitelX & """ "
		LogTemp = LogTemp & "context="""" "
		LogTemp = LogTemp & "type=""" & LogError & """ "
		LogTemp = LogTemp & "thread=""0"" "
		LogTemp = LogTemp & "file=""David.Segura"" "
		LogTemp = LogTemp & ">"

		Tst = 8							'ForAppending
		If LogError = 0 Then Tst = 2	'ForWriting

		Set FileOut = objFSO.OpenTextFile( MyLogFileX, Tst, True)
		If     LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		If Not LogTemp = vbCRLF Then FileOut.WriteLine ( LogTemp )
		FileOut.Close
		Set FileOut	= Nothing
		'Set objFSO	= Nothing
	End Function
	' /////////////////////////////////////////////////////////
	' Trace Log Solid Line
	' /////////////////////////////////////////////////////////
	Sub LogLine
			TraceLog "=====================================================================================", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Blank Space
	' /////////////////////////////////////////////////////////
	Sub LogSpace
		TraceLog "", 1
	End Sub
	' /////////////////////////////////////////////////////////
	' Trace Log Contents
	' /////////////////////////////////////////////////////////
	Sub LogStart
		'Tracelog "Start a new Log File", 0											'Clears any existing content
		'TraceLog "This is a standard line", 1										'Create an Entry
		'TraceLog "This is a warning line", 2										'Create an Entry and highlight yellow (Warning)
		'TraceLog "This is an error line", 3										'Create an Entry and highlight red (Error or Critical)
		'LogSpace																	'Create a Line without content
		'LogLine																	'Create a Line with =====================================

		If WScript.Arguments.length = 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 0
		If WScript.Arguments.length <> 0 Then TraceLog "Starting "					& WScript.ScriptFullName, 2
		TraceLog "Start Date and Time is "											& Now, 1
		TraceLog "Script Last Modified: " 											& CreateObject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).DateLastModified, 1
		LogLine
		TraceLog "<Constant> Author: " 												& Author, 1
		TraceLog "<Constant> Author Email: " 										& AuthorEmail, 1
		TraceLog "<Constant> Company: " 											& Company, 1
		TraceLog "Do not contact the Author directly for Support", 3
		LogLine
		TraceLog "<Constant> Script: " 												& Script, 1
		TraceLog "<Constant> Description: " 										& Description, 1
		LogLine
		TraceLog "The defined Support process is to Submit an Incident", 3
		TraceLog "This script is provided for Testing Only (No Priority Support)", 3
		TraceLog "<Constant> SupportAction: " 										& SupportAction, 2
		TraceLog "<Constant> Incident Area: " 										& SupportArea, 2
		TraceLog "<Constant> Assign to Group: " 									& SupportGroup, 2
		TraceLog "<Constant> Assignee: " 											& SupportContact, 2
		TraceLog "<Constant> Subject: " 											& SupportSubject, 2
		TraceLog "<Constant> Description: " 										& SupportProblem, 2
		LogLine
		TraceLog "<Constant> Title: " 												& Title, 1
		TraceLog "<Constant> Version: " 											& Version, 2
		TraceLog "<Constant> VersionFull: " 										& VersionFull, 2
		LogLine
		TraceLog "<Variable> MyUserName: " 											& MyUserName, 1
		TraceLog "<Variable> MyComputerName: " 										& MyComputerName, 1
		TraceLog "<Variable> MyWindir: " 											& MyWindir, 1
		TraceLog "<Variable> MyTemp: " 												& MyTemp, 1
		TraceLog "<Variable> MySystemDrive: " 										& MySystemDrive, 1
		TraceLog "<Variable> MyArchitecture: " 										& MyArchitecture, 1
		LogLine
		TraceLog "<Variable> MyScriptFullPath = "									& MyScriptFullPath,			1
		TraceLog "<Variable> MyScriptFileName = "									& MyScriptFileName, 		1
		TraceLog "<Variable> MyScriptBaseName = "									& MyScriptBaseName, 		1
		TraceLog "<Variable> MyScriptParentFolder = "								& MyScriptParentFolder, 	1
		TraceLog "<Variable> MyScriptGParentFolder = "								& MyScriptGParentFolder,	1
		TraceLog "<Variable> MyParentFolderName = "									& MyParentFolderName,		1
	End Sub
'==============================================================================================


























'==============================================================================================
'==============================================================================================
'==============================================================================================
'==============================================================================================
'==============================================================================================










'==============================================================================================
	Function BuildLogScript1
		If WScript.Arguments.length = 0 Then TraceLog "Starting " & WScript.ScriptFullName, 0
		If WScript.Arguments.length <> 0 Then TraceLog "Starting " & WScript.ScriptFullName, 1
		TraceLog "Script Last Modified: " & CreateObject("Scripting.FileSystemObject").GetFile(MyScriptFullPath).DateLastModified, 1
		LogSpace
	End Function
'==============================================================================================
	Function CheckAdminRights
		Dim RegKey
		On Error Resume Next
		
		RegKey = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\")
		
		If err.number <> 0 Then
			HasAdminRights = "NO"
			If MyOperatingSystem = "Windows XP" and Lcase(CreateObject("WScript.Network").UserName) = "administrator" Then HasAdminRights = "YES"
			If Lcase(CreateObject("WScript.Network").UserName) = "system" Then HasAdminRights = "YES"
		Else
			HasAdminRights = "YES"
		End If
	End Function
'==============================================================================================
	Sub PromptAdminRights
	Dim UACValue
		objShellApp.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
		WScript.Quit
	End Sub
'==============================================================================================
	Sub ReservedForLater

		Dim MyFolder, fldr, MyLastNamedDir, MyLastDateDir, MyLastDate
		Set MyFolder = objFSO.GetFolder(MyScriptParentFolder)
		
		For Each fldr In MyFolder.SubFolders
			'Determine Subdirectory by Name
			MyLastNamedDir = fldr.Name
			
			'Determine Subdirectory by Date
			If fldr.DateLastModified > MyLastDate Or IsEmpty(MyLastDate) Then
				MyLastDateDir = fldr.Name
				MyLastDate = fldr.DateLastModified
			End If
		Next

		cSetupFile = Replace(cSetupFile,"%lastdate%",MyLastDateDir)	'Chooses the directory that is last modified date/time
		cSetupFile = Replace(cSetupFile,"%lastname%",MyLastNamedDir)	'Chooses the directory that is last alphabetically
		
		cSetupx86 = Replace(cSetupx86,"%lastdate%",MyLastDateDir)	'Chooses the directory that is last modified date/time
		cSetupx86 = Replace(cSetupx86,"%lastname%",MyLastNamedDir)	'Chooses the directory that is last alphabetically

		cSetupx64 = Replace(cSetupx64,"%lastdate%",MyLastDateDir)	'Chooses the directory that is last modified date/time
		cSetupx64 = Replace(cSetupx64,"%lastname%",MyLastNamedDir)	'Chooses the directory that is last alphabetically

		cSetupSwitches = Replace(cSetupSwitches,"%sourcedir%",MyScriptParentFolder) 'Replace the SourceDir variable with the installation parent
		cSetupSwitches = Replace(cSetupSwitches,"%lastdate%",MyLastDateDir) 'Replace the SourceDir variable with the installation parent
		cSetupSwitches = Replace(cSetupSwitches,"%lastname%",MyLastNamedDir) 'Replace the SourceDir variable with the installation parent
	End Sub
'==============================================================================================
