;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Reads jobs in settings.ini and uses Robocopy (A Standard Windows 7 command) to move files from the job specific servers
; to the destination, also specified in the settings.ini      Only one job is permitted at a time to limit throughput.
; Robocopy Documentation: http://technet.microsoft.com/en-us/library/cc733145.aspx



;~~~~~~~~~~~~~~~~~~~~~
;Compile Options
;~~~~~~~~~~~~~~~~~~~~~
Startup()
The_ProjectName := "ITOps Log Annihilator"
The_VersionName = 2.0.0

;Dependencies
#Include %A_ScriptDir%\Functions
#Include inireadwrite


;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Startup
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

	;Command Line Argument 2 can be a path to a settings file. Otherwise check ScriptDir
	CL_Arg2 = %2%
	If (CL_Arg2 = "")
	{
		;Check settings.ini. Quit if not found
		IfExist, %A_ScriptDir%\settings.ini
		{
		Path_SettingsFile = %A_ScriptDir%\settings.ini
		}
		Else
		{
		Fn_TempMessage("Could not find settings file. Quitting in 10 seconds")
		ExitApp, 1
		}
	}
	Else
	{
	Path_SettingsFile := CL_Arg2
	}


;Self Descriptive. Read settings.ini and set global Vars
Fn_InitializeIni(Path_SettingsFile)
Fn_LoadIni(Path_SettingsFile)
Sb_GlobalVars()
SeperatorLine := "-------------------------------------------"

;Clean all user input and validate
Sb_Sanitize()

;Create Archive director and set File_Log variable to [TodaysArchiveDir]\[Day].txt
if (Settings_dontcreatelogfile = 1) {

} else {
	File_Log := Fn_CreateArchiveDir(Path_Report)
}


;Convert user timeout settings to milliseconds and set quit timer. Keeping as misc feature
Timer_TimeOut := Settings_TimeOut * 60000
SetTimer, Quit, -%Timer_TimeOut%

;Subtract 1 from inisections because [Settings] created 1
inisections -= 1
;Counter_Index is used to increment up/down the ini keys
Counter_Index := 0

;Build GUI and show
GUI_Build()

	;Prompt user to confirm if not run from command line with auto parameter, Note that %1% is the first Command Line Argument
	CL_Arg1 = %1%
	If (CL_Arg1 != "auto") {
		MsgBox, 1, , This will run disk space maintenance on a long list of servers. Must be used from an .adm account. continue?
		IfMsgBox OK
		{
			Fn_TempMessage("Cleanup will start in 30 seconds, close the main GUI to cancel", 30)
		} else {
			Fn_TempMessage("Exiting in 10 seconds")
			ExitApp, 1
		}
	}

;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Main
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

;Header Message in log file
ErrorMessage := "Disk Space Maintenance Started at " A_Hour ":" A_Min ":" A_Sec
FileAppend, %SeperatorLine%`n %ErrorMessage%`n%SeperatorLine%`n, %File_Log%

	;Loop for all source Dirs
	Loop, %inisections%
	{
	;~~~~~~~~~~~~~~~~~~~~~
	;Determine Source and Destination path for current loop
	;~~~~~~~~~~~~~~~~~~~~~
	
	;Path_Source will now hold 1_Path's data; C:\...
	Counter_Index += 1
	
	;Make sure Path=\c$\
	Path_Source := %Counter_Index%_Path
	; RegExMatch(Path_Source, "(\\)[a-zA-Z]\$\\", RE_Path_Source)
	; 	If (RE_Path_Source1 = "")
	; 	{
	; 	ErrorMessage := Path_Source " is missing a leading \ or does not match the expected path pattern. Check settings.ini"
	; 	FileAppend, %SeperatorLine%`n ERROR: %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
	; 	Fn_TempMessage(ErrorMessage)
	; 	Continue
	; 	}
	
	
	StringReplace, Loop_Machines, %Counter_Index%_Machines, %A_SPACE%,, All
	;Split X_Machines into formatted into simple array. `, is escaped comma
	StringSplit, Machine_Array, Loop_Machines, `,,
	Counter_Machine = 0
	
	
	;Clear file types var as it will be looped multiple times
	Type_Files =
	
	
	;Remove all spaces from filetype ;;NO Just assign value, we want spaces to work now.
	Buffer_FileTypes := %Counter_Index%_Types
	;StringReplace, Buffer_FileTypes, %Counter_Index%_Types, %A_SPACE%,, All
	
	
	;Split Xkey_Types into formatted ext,ext,ext into simple array. `, is escaped comma
	StringSplit, file_array, Buffer_FileTypes, `,,
	
	;loop the amount of file types to be handled  %file_array0% = 4 in the case of ext,ext,ext,ext
	Counter_FileExt = 0
		Loop, %file_array0%
		{
		Counter_FileExt += 1
		Current_FileExt := file_array%Counter_FileExt%
		
		;Commenting out space replacement as I need to enter a space into filepattern matching for moving dirs /s
		;StringReplace, Current_FileExt, Current_FileExt, %A_SPACE%,, All
			
			
			;For Filechecking- If the file type already has a period, just remember the file extension after the period
			IfInString, Current_FileExt, `.
			{
			StringSplit, FileExt_Check, Current_FileExt, `.,
			FileExt_Check = %FileExt_Check2%
			}
			Else
			{
			;Else it doesn't have a period. No formatting needed
			FileExt_Check = %Current_FileExt%
			}
			
		Approved_FileTypes = log,log_1,csv,xml,txt,txt /s
			;Skip this filetype if its not part of the approved list
			IfNotInString, Approved_FileTypes, %FileExt_Check%
			{
			ErrorMessage := Current_FileExt . " is not an approved file type"
			FileAppend, %SeperatorLine%`n ERROR: %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
			Fn_TempMessage(ErrorMessage)
			Continue
			}
			
			;If this Ext does !NOT contain a period (.) then make it [*.Ext]
			;This builds the Type_Files var to be formatted for Robocopy EX: "*.txt *.csv"
			IfNotInString, Current_FileExt, `.
			{
			Type_Files = *.%Current_FileExt%%A_Space%%Type_Files%
			}
			Else
			{
			Type_Files = %Current_FileExt%%A_Space%%Type_Files%
			}
		}
	
	
	
		;Loop for each machine, done the same style as extensions. Put into simple array then looped
		Loop, %Machine_Array0%
		{
		Counter_Machine += 1
		Current_Machine := Machine_Array%Counter_Machine%
		Path_Source := "\\" Current_Machine %Counter_Index%_Path
		
		;Get Machine Name out of Source Path. Skip to end of Loop if System Name cannot be determined. Can also use "\\\\(\w+(\d\d))\\[a-zA-Z]\$.+\\" in the future
		RegExMatch(Path_Source, "\\\\(\S+\d\d)\\", RE_Sys_Name)
			If (RE_Sys_Name1 = "")
			{
			ErrorMessage := "A valid system name could not be found in: " Path_Source
			FileAppend, %SeperatorLine%`n ERROR: %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
			Fn_TempMessage(ErrorMessage)
			Continue
			}
		
		;Get System Type out of System Name. Skip to end of Loop if system type is not recognized
		RegExMatch(Path_Source, "\\\\\S+(...)\d\d\\", RE_Sys_Type)
			If (RE_Sys_Type1 = "")
			{
			ErrorMessage := "The source path does not does not contain a recognized system type: " . RE_Sys_Type1
			FileAppend, %SeperatorLine%`n ERROR: %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
			Fn_TempMessage(ErrorMessage)
			Continue
			}
		
		;Uppercase the system Name and Type
		Sys_Name = %RE_Sys_Name1%
		StringUpper, Sys_Name, Sys_Name	
		Sys_Type = %RE_Sys_Type1%
		StringUpper, Sys_Type, Sys_Type
		
		;Update GUI to current Status
		Total_Progress := Counter_Index / (inisections + 1)
		GUI_Update(Sys_Name, Sys_Type, Total_Progress, Path_Source)
		
		;; Move files to the archive or delete them depending on settings
		if (Settings_Delete || %Counter_Index%_Delete) {
			Path_Destination = %Path_Source%\TEMP
		} else {
			Path_Destination = %Settings_Destination%\%Sys_Type%\%Sys_Name%
		}
		
		;Run Robocopy in cmd.exe and /c(lose) when finished   ~~~----------------------------------------------------------------------------------------------------------~~~
		RunWait, %comspec% /c Robocopy "%Path_Source%" "%Path_Destination%" %Type_Files% /MOV /MINLAD:%Age_Days% /LOG+:"%File_Log%" /NP /XX /W:1 /R:0,,hide
		
		if (Settings_Delete || %Counter_Index%_Delete) {
			FileRemoveDir, %Path_Destination%, 1 ;recursive delete the TEMP directory
		}
		
		;Insert some space between each machine
		FileAppend, `n`n`n, %File_Log%
		
		
		;Clear vars so we can ensure it won't be run a second time against a different machine
		Path_Source = 
		Path_Destination = 
		
		;Here is everything that is done after each machines has been completed
		;None
		}
	
	;Here is everything that is done after each job/directory has been completed
	;Insert more space between each system type
	FileAppend, `n `n `n, %File_Log%
	}


if (Settings_DeleteDirectory) { ;; Ok this was a makeshift for QA log clearing. Removable soon
	Loop, Files, %Settings_DeleteDirectory%\*.log*
    {
        FileDelete, % A_LoopFileFullPath
    }
	Loop, Files, %Settings_DeleteDirectory%\*.bak*
    {
        FileDelete, % A_LoopFileFullPath
    }
    Loop, Files, %Settings_DeleteDirectory%\*.txt*
    {
        FileDelete, % A_LoopFileFullPath
    }
	; FileRemoveDir, %Settings_DeleteDirector% , 1
}

;Finished, Tell User, exit afterwards
ErrorMessage := "Disk Space Maintenance Completed Successfully at " A_Hour ":" A_Min ":" A_Sec
FileAppend, %SeperatorLine%`n %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
Fn_TempMessage(ErrorMessage)
ExitApp, 0

;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Functions
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

Fn_CreateArchiveDir(para_PathReport)
{
	FormatTime, CurrentYear,, yyyy
	FormatTime, CurrentMonth,, MM-MMMM
	FormatTime, CurrentDay,, yyyy-MM-dd

	FileCreateDir, %para_PathReport%\%CurrentYear%\%CurrentMonth%\
	Path_Archive = %para_PathReport%\%CurrentYear%\%CurrentMonth%\%CurrentDay%-%A_ComputerName%.txt
		If (Errorlevel != 0) {
			Fn_TempMessage("Could not start log report at " Settings_Report ". Probably unwriteable. Continuing in 20 seconds", 20)
		}
	Return Path_Archive
}


Fn_TempMessage(para_Message, para_Timeout = 10)
{
	MsgBox, 48,, %para_Message%, %para_Timeout%
}


Fn_FindTotalIterations() ;Unfinished function to improve Progress Bar accuracy
{
global
;The_Ini_Buffer = blank

	Loop, %inisections%
	{
	;StringSplit, 
	
	}

}


;/--\--/--\--/--\--/--\
; Subroutines
;\--/--\--/--\--/--\--/

Sb_GlobalVars()
{
global
Path_Report = %Settings_Report%
Age_Days = %Settings_FileToMoveAgeInDays%
;This is an array that holds all messages to be written to the log. X is just used in the Array ex: Array_Log[X]
Array_Log := []
X = 0
}


Startup()
{
#SingleInstance force
#NoEnv
}


Sb_Sanitize()
{
global
;Parent Destination for all servers
RegExMatch(Settings_Destination, "(\\\\)\S+\\", RE_UserInput)
	If (RE_UserInput1 = "")
	{
	ErrorMessage := "Parent destination: " %Settings_Destination% " could not be validated, check settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}
	
;Report Location
RegExMatch(Settings_Report, "(\\\\)\S+\\", RE_UserInput)
	If (RE_UserInput1 = "")
	{
	ErrorMessage := "Report location: " %Settings_Report% " could not be validated, check settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}
	
;TimeOut
RegExMatch(Settings_TimeOut, "(\d+)", RE_UserInput)
	If (RE_UserInput1 = "")
	{
	ErrorMessage := "Timeout setting" %Settings_TimeOut% " could not be validated, check settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}
	If (Settings_TimeOut < 10)
	{
	ErrorMessage := %Settings_TimeOut% "The timeout setting must be greater than 10 mins, check settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}

;FileToMoveAgeInDays
RegExMatch(Settings_FileToMoveAgeInDays, "(\d+)", RE_UserInput)
	If (RE_UserInput1 = "")
	{
	ErrorMessage := "File to move age in days: " %Settings_FileToMoveAgeInDays% " could not be validated, check settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}
	If (Settings_FileToMoveAgeInDays < 2)
	{
	ErrorMessage := "File to move age in days: " %Settings_FileToMoveAgeInDays% "is less than 3 days, that is not allowed. Update settings.ini.  Quitting in 10 Seconds"
	Fn_TempMessage(ErrorMessage)
	ExitApp, 1
	}

}


GUI_Build()
{
global
Gui +AlwaysOnTop
;Title
Gui, Font, s14 w70, Arial
Gui, Add, Text, x2 y0 w480 h40 +Center, Disk Space Cleanup
Gui, Font, s10 w70, Arial
Gui, Add, Text, x420 y0 w200 h20, v%The_VersionName%


;System Name Full
Gui, Add, Text, x2 y56 w140 h20 +Center vGUI_SysName,
Gui, Add, GroupBox, x2 y40 w150 h40 , System Name

;System Type "DDS"
Gui, Add, Text, x14 y96 w30 h20 vGUI_SysType,
Gui, Add, GroupBox, x2 y80 w46 h40 , Type

;Progress
Gui, Add, Text, x64 y96 w60 h20 , 
Gui, Add, Progress, x56 y100 w90 h14 vGUI_ProgressBar, 1
Gui, Add, GroupBox, x52 y80 w100 h40 vGUI_Progress, Progress 0`%

;Full Path
Gui, Add, Text, x172 y60 w290 h50 vGUI_SourcePath +Wrap,
Gui, Add, GroupBox, x162 y40 w310 h80 , Full Path (Source)

;Large Progress Bar UNUSED
;Gui, Add, Progress, x4 y130 w480 h20 , 100

Gui, Show, x127 y87 h130 w488, Disk Space Cleanup
Return
}


GUI_Update(SystemName, SystemType, Progress, Source)
{
;System Name
guicontrol, Text, GUI_SysName, %SystemName%

;System Type
guicontrol, Text, GUI_SysType, %SystemType%

;Full Path. Add newline every 44 characters in case if doesn't fit.
Source := RegexReplace(Source, ".{44}\K", "`n")
guicontrol, Text, GUI_SourcePath, %Source%

;Simple Progress
Progress := Progress * 100
	IfInString, Progress, .
	{
	StringTrimRight, Progress, Progress, 7
	}
guicontrol, Text, GUI_Progress, Progress: %Progress%`%

;Progress Bar
GuiControl,, GUI_ProgressBar, %Progress%
}


GuiClose:
ExitApp, 1

;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Labels
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

;Runs if timeout period is reached.
Quit:
ErrorMessage := "The timeout period " . Settings_TimeOut . " mins was reached. Exiting in 10 seconds."
FileAppend, %SeperatorLine%`n ERROR: %ErrorMessage%`n%SeperatorLine%`n, %File_Log%
Fn_TempMessage(ErrorMessage)
ExitApp, 1
