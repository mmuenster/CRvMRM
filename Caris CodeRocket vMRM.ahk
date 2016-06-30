Startup:         ;MS done
{
validhelpers=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890
letters=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz
numbers=1234567890
#Include %A_ScriptDir%\Caris CodeRocket Internal Functions vMRM.ahk	
#Include %A_ScriptDir%\Caris CodeRocket External Functions vMRM.ahk
#SingleInstance force
#WinActivateForce
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetTitleMatchMode 1
CoordMode, Mouse, Relative
ComObjError(false)

IfExist, %A_MyDocuments%\CarisCodeRocket.ini
	ReadIniValues()
Else
	FirstTimeSetup()

SetTimer, checkForMelanoma, 3600000

Loop Files, %programfiles%\Microsoft Office\*.* , D
	{
	IfInString, A_LoopFileName, Office
		IfExist, %programfiles%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
			OutlookPath= %programfiles%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
		else 
		{
			np := RegExReplace(programfiles, " \(x86\)","")
			IfExist, %np%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
				OutlookPath= %np%\Microsoft Office\%A_LoopFileName%\OUTLOOK.EXE
			else
				Msgbox, Your path to OUTLOOK.EXE could not be determined.  You will not have automatic emailing capabilities.

		}
    }

Gui, Add, Text, vCaseLoaderLbl,Case Loader:
Gui, Add, Edit, vCaseScanBox ys,
Gui, Font, S14, Arial
Gui, Add, Text, r1 w400 xs vPatientLabel, Patient: 
Gui, Add, Text, r1 w400 vDoctorLabel, Doctor: 
Gui, Font, S20, Arial
Gui, Add, Text, r1 cBlue w400 vUsePhotos, <PHOTOS REQUIRED>
Gui, Add, Text, r1 cRED vUseMicros, <MICROS REQUIRED>
;Gui, Add, Text, r1 cGreen vUseICD9s, <ICD9S REQUIRED>
Gui, Add, Text, r1 cBlack w500 vUseMargins, <Margin Preferences appear here>
Gui, Add, ListView, Hidden w600 r5 gMyListView, ID|Code|Category|Subcategory|Dx Line|Comment|Micro|CPT Code|ICD9|ICD10|SNOMED|Premalignant|Malignant|Dysplastic|Melanocytic|Inflammatory|MarginIncluded|Log
Gui, Add, Button, w0 h0 Default, OK

Menu, FileMenu, Add, E&xit, GuiClose
Menu, HelpMenu, Add, Search Diagnosis Codes  (F7), F7
Menu, HelpMenu, Add, Search Extended Phrases  (Shift-F7), +F7
Menu, HelpMenu, Add, Display All Helpers  (F9), F9

Menu, ModeMenu, Add, Manual, ModeManual
Menu, ModeMenu, Add, Basic Automatic, ModeBasicAutomatic
Menu, ModeMenu, Add, Deluxe Automatic, ModeDeluxeAutomatic
Menu, SettingsMenu, Add, Speak Patient Name, SpeakPatientName
Menu, SettingsMenu, Add, Use Send Method, UseSendMethod

Menu, EditMenu, Add, Edit Client Preferences, EditDoctorPreferences

Menu, MyMenuBar, Add, &File, :FileMenu  
Menu, MyMenuBar, Add, &Edit, :EditMenu
Menu, MyMenuBar, Add, &Mode, :ModeMenu
Menu, MyMenuBar, Add, &Settings, :SettingsMenu
Menu, MyMenuBar, Add, &Help, :HelpMenu
Gui, Menu, MyMenuBar

Gui, 2:Font, S12, Verdana
Gui, 2:Add, Text, x18 vDocPreferenceLabel, TestNameXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Gui, 2:Font, S8, Verdana
Gui, 2:Add, Checkbox, vPhotoSelect, Photos Required
Gui, 2:Add, Checkbox, vMicroSelect, Micros Required
;Gui, 2:Add, Checkbox, vICD9Select, ICD9s Required
Gui, 2:Add, Text, , Margin Preferences
Gui, 2:Add, Edit, vMarginSelect w500, 
Gui, 2:Add, Button, gSavePreferences vGo, Save Preferences

Gui, 4:Font, S12, Verdana
Gui, 4:Add, Text, vDisplayCaseNumber , Case Number: XXXXXXXXXXX
Gui, 4:Add, DropDownList, vLoc w500, Boston|Irving
If (HomeLabCasePrefix="C")
	GuiControl, 4:Choose, Loc, Boston
else if (HomeLabCasePrefix="D")
	GuiControl, 4:Choose, Loc, Irving

Gui, 4:Add, DropDownList, AltSubmit vEmailType w500, Patient Double Blind Error|Need Previous Biopsy Report|Clinical Note and Photos|Critical Result Call|Pull Bottles and Blocks
GuiControl, 4:Choose, EmailType, 1
Gui, 4:Add, Text, ,Comments
Gui, 4:Add, Edit, vEmailComments w500, 
Gui, 4:Add, Button, gSendEmail vSendEmail, Send Email


Progress, 0 x400 y1 h130, Preparing for first time use..., Written by Matthew Muenster M.D.`n`nInitializing..., Caris CodeRocket 
Progress, 40, Reading personalized values...
Progress, 60, Getting the diagnosis codes from the database...
ReadDXCodes()
Progress, 80, Getting the helper codes from the database...
ReadHelpers()
Progress, 100, Initialization complete!
Progress, Off
{
	df =FRONT OF DIAGNOSIS HELPERS`n
	df=%df%------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := FrontofDiagnosisHelper%x%
			df = %df%%x% -- %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := FrontofDiagnosisHelper%x%
			df = %df%%x% -- %ph1%`n
		}
	df = %df%----------------------------------------------------`n

	dm =MARGIN HELPERS`n---------------------------------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := BackofDiagnosisHelper%x%
			dm = %dm%%x% -- %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := BackofDiagnosisHelper%x%
			dm = %dm%%x% -- %ph1%`n
		}
	dm = %dm%---------------------------------------------------------------------------------`n

dc =COMMENT HELPERS`n---------------------------------------------------------------------------------`n
	Loop, 10
		{
			x := A_index -1 
			ph1 := CommentHelper%x%
			dc = %dc%%x% - %ph1%`n
	}
	Loop, 26
		{
			x := Chr(64 + A_Index)
			ph1 := CommentHelper%x%
			dc = %dc%%x% - %ph1%`n
		}

	dc = %dc%---------------------------------------------------------------------------------`n
Gui, 3:Font, S8, Verdana
Gui, 3:Add, Text, w150, %df%
Gui, 3:Add, Text, w100 ym, %dm%
Gui, 3:Add, Text, w200 ym, %dc%
}
if OpMode=T
	{
		SplashTextOn, 100,100,,Entering Windowless Operation Mode
		Gosub, ModeTranscription
		Sleep, 1200
		SplashTextOff
	}
Else
	{	
		Gui, 1:+Resize +MinSize400x300 +MaxSize500x400
		Gui, Show, x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%
		SetTimer, WinSURGECaseDataUpdater, 2000
		Gosub, WinSurgeCaseDataUpdater
		If OpMode=D
			Gosub, ModeDeluxeAutomatic
		else if OpMode=B
			Gosub, ModeBasicAutomatic
		Else
			Gosub, ModeManual
	}
If SpeakEnabled
	Menu, SettingsMenu, Check, Speak Patient Name
Else
	Menu, SettingsMenu, UnCheck, Speak Patient Name

If UseSendMethod
	Menu, SettingsMenu, Check, Use Send Method
Else
	Menu, SettingsMenu, UnCheck, Use Send Method

StringLeft, x, A_ScriptDir, 1
if x=C
	x=S
Run, %x%:\CodeRocket\bin\EP\Autohotkey.exe %x%:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk, ,UseErrorLevel , epid
If ErrorLevel
	MsgBox, 4112, Connectivity Error, Your connection to the S:\CodeRocket directory is not present.  Usually`, restarting your computer will correct this.  You can use the CodeRocket program but will have no extended phrase capabilities.


IfExist, %A_MyDocuments%\PersonalExtendedPhrases.ahk
	{
	Run, %x%:\CodeRocket\bin\EP\Autohotkey.exe "%A_MyDocuments%\PersonalExtendedPhrases.ahk", ,UseErrorLevel , ppid	
	If ErrorLevel
		Msgbox, There was an error loading your personal extended phrases file.
	}
return
}

SpeakPatientName:
{
	Menu, SettingsMenu, ToggleCheck, Speak Patient Name
	If SpeakEnabled
		SpeakEnabled := 0
	Else
		SpeakEnabled := 1
	IniWrite, %SpeakEnabled%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
Return
}

UseSendMethod:
{
	Menu, SettingsMenu, ToggleCheck, Use Send Method
	If UseSendMethod
		UseSendMethod := 0
	Else
		UseSendMethod := 1
	IniWrite, %UseSendMethod%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, UseSendMethod

Return
}

ModeTranscription:
{
	Hotkey, F7, On
	Hotkey, +F7, On
	Hotkey, F8, Off
	Hotkey, F9, Off
	Hotkey, +F8, Off
	Hotkey, F12, Off
	Hotkey, ^!s, Off
	Hotkey, ^k, Off
	Hotkey, ^!p, Off
	Return
}

ModeManual:
{
	Menu, ModeMenu, Check, Manual
	Menu, ModeMenu, UnCheck, Basic Automatic	
	Menu, ModeMenu, UnCheck, Deluxe Automatic	
	Menu, EditMenu, Disable, Edit Client Preferences
	Menu, SettingsMenu, Disable, Speak Patient Name
	SpeakEnabled := 0
	
	GuiControl, Text, UsePhotos, MANUAL PREFERENCES
	GuiControl, 1:Show, UsePhotos
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	GuiControl, Disable, CaseScanBox
	GuiControl, Disable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, Manual Mode
	GuiControl, 1:Text, PatientLabel, 
	GuiControl, 1:Text, DoctorLabel, 
	Hotkey, F8, Off
	Hotkey, +F8, Off
	Hotkey, F12, On
	Hotkey, ^!s, Off
	Hotkey, ^k, Off
	Hotkey, ^!p, Off
	OpMode=M
	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	Return
}

ModeBasicAutomatic:
{
	if (!QueueIntoBatchBox OR QueueIntoBatchBox=Error)
		{
			Msgbox, 4, , You have not setup Caris CodeRocket for automated operation.  Do you want to go through setup to enable automation?
			IfMsgBox, Yes
				{
					FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
					Reload
					Return
				}
			Else
				Return
		}
	SpeakEnabled := 0
	Menu, ModeMenu, UnCheck, Manual
	Menu, ModeMenu, Check, Basic Automatic	
	Menu, ModeMenu, UnCheck, Deluxe Automatic	
	Menu, SettingsMenu, Disable, Speak Patient Name
	Menu, EditMenu, Disable, Edit Client Preferences
	Hotkey, F8, On
	Hotkey, +F8, On
	Hotkey, F12, On
	Hotkey, ^!s, On
	Hotkey, ^k, On	
	GuiControl, Text, UsePhotos, MANUAL PREFERENCES
	GuiControl, 1:Text, PatientLabel, 
	GuiControl, 1:Text, DoctorLabel, 
	GuiControl, 1:Show, UsePhotos
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	GuiControl, Enable, CaseScanBox
	GuiControl, Enable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, 
	OpMode=B
	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
Return
}

ModeDeluxeAutomatic:
{
	if (!QueueIntoBatchBox OR QueueIntoBatchBox=Error)
		{
			Msgbox, 4, , You have not setup Caris CodeRocket for automated operation.  Do you want to go through setup to enable automation?
			IfMsgBox, Yes
				{
					FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
					Reload
					Return
				}
			Else
				Return
		}
		
	Menu, ModeMenu, UnCheck, Manual
	Menu, ModeMenu, UnCheck, Basic Automatic	
	Menu, ModeMenu, Check, Deluxe Automatic	
	OpMode=D
	Hotkey, F8, On
	Hotkey, +F8, On
	Hotkey, F12, On
	Hotkey, ^!s, On
	Hotkey, ^k, On	
	Hotkey, ^!p, On
	GuiControl, Text, UsePhotos, <PHOTOS REQUIRED>
	GuiControl, Enable, CaseScanBox
	GuiControl, Enable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, 
	Menu, EditMenu, Enable, Edit Client Preferences
	Menu, SettingsMenu, Enable, Speak Patient Name


	IniWrite, %OpMode%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, OpMode
	lastWinSURGEtitle=  ;Forces Case update to occur.
	Gosub, WinSurgeCaseDataUpdater
Return
}

MyListview:       ;MS done
return	

EditDoctorPreferences:
{
	if OpMode=D
	{
		GuiControl, 2:, PhotoSelect, %UsePhotos%
		GuiControl, 2:, MicroSelect, %UseMicros%
		;GuiControl, 2:, ICD9Select, %UseICD9s%	
		GuiControl, 2:, MarginSelect, %UseMargins%
		GuiControl, 2:Text, PhotoSelect, Photos Required
		GuiControl, 2:Text, MicroSelect, Micros Required
		;GuiControl, 2:Text, ICD9Select, ICD9s Required
		SetTimer, WinSURGECaseDataUpdater, Off
		Gui, 1:Hide
		Gui, 1:+Disabled
		GuiControl, 2:Text, DocPreferenceLabel, Enter preferences for %ClientName%
		;Gosub, GlobalPreferenceSelect
		Gui, 2:Show  
		Gui, 2:-Disabled
		Sleep, 500
		WinActivate, Caris CodeRocket
	}
	Else
		Msgbox, You must be in Deluxe Automatic Mode to use this feature.
	Return
}
	
SavePreferences:   ;MS done
{
Gui, 2:Submit
	
	StringReplace, y, WinSurgeFullName,','', All
	StringReplace, z, ClientName,','',All

If ClinicianFound
	s =	Update CliniPref SET Photo_pref='%PhotoSelect%', Micro_pref ='%MicroSelect%', Icd9_pref='0', Margin_pref='%MarginSelect%', log='%CurrentLog%;%y% %A_mm%-%A_DD%-%A_YYYY%' WHERE Winsurge_id = %ClientWinSurgeId%
Else
	s =	Insert into CliniPref (Winsurge_id,Name,Photo_pref,Micro_pref,Margin_pref,Icd9_pref,log) VALUES (%ClientWinSurgeId%,'%z%','%PhotoSelect%', '%MicroSelect%', '%MarginSelect%', '0', '%y% %A_mm%-%A_DD%-%A_YYYY%')

CodeDatabaseQuery(s) 

lastWinSURGEtitle = 
Gui, 2:Hide
Gui, 2:+Disabled
Gui, 1:-Disabled
Gosub, WinSURGECaseDataUpdater
Gui, 1:Show
WinActivate, WinSURGE - 
WinActivate, Caris CodeRocket
lastWinSURGEtitle=
SetTimer, WinSURGECaseDataUpdater, 2000

Return
}

ButtonOK:
{
	Gui, Submit, NoHide
	If (UndoEnabled AND !DataEntered)
		{
			Msgbox, You must first save or undo the changes to the current case!
			Gosub, F12
			Return
		}
	CaseNumberProblem=0
	StringSplit, x, CaseScanBox, %A_Space%
	StringSplit, y, x1, -
	Stringlen, csbLength, y1
	if (csbLength<>4)
		CaseNumberProblem=1
	StringLeft, csbPrefix, y1, 2
	StringRight, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringLeft, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringRight, csbYear, y1, 2
	csbCaseNum := y2
	StringRight, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLeft, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLen, casenumlength, csbCaseNum
	Loop, %casenumlength%
	{
		StringMid, j, csbCaseNum, A_Index, 1
		IfNotInString, numbers, %j%
			CaseNumberProblem=1
	}

if CaseNumberProblem
		{
			Msgbox, You did not enter a valid case number!
			Gosub, F12
			Return
		}

NewCaseNum=%csbPrefix%%csbYear%-%csbCaseNum%

	If (DataEntered AND UndoEnabled)
		Gosub, F8
	
	If SaveError
		{
			Gosub, F12
			Return
		}
	

	CloseWinSURGEModalWindow("WinSURGE - Final Diagnosis:","","Close")
	;CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
	IfWinNotExist, WinSURGE Case Lookup
		{
		WinActivate, WinSURGE
		Send, !2
	}
	SetTimer, UnblockInput, 5000
	BlockInput, On
	OpenCase(NewCaseNum)
	OpenFinalDiagnosisModal()
	ActivateNextTripleAsterisk() 
	BlockInput, Off
	SetTimer, UnblockInput, Off
	
	ControlSetText, ThunderRT6TextBox19, %NewCaseNum%, Dictation Data, 	
	
	Gosub, WinSURGECaseDataUpdater
	StringSplit, name, PatientName, `,
	p = %name2%%A_Space%%name1%
	If SpeakEnabled
		{
		Run, %A_ScriptDir%\wscript.exe "%A_ScriptDir%\speak.vbs" "%p%"	
		Sleep, 400
		Run, %A_ScriptDir%\wscript.exe "%A_ScriptDir%\speak.vbs" "Age %PatientAge%"
		}
	DataEntered := 0
	Sleep, 100
	gosub, F12
	WinActivate, WinSURGE , 

	Return
}

UnblockInput:
{
	BlockInput, Off
	return
}

GuiClose:
{
	Process, Close, %epid%, 
	Process, Close, %ppid%,
	Process, Close, Autohotkey.exe,
	
	If ErrorLevel
		Process, Close, %ErrorLevel%, 
	ExitApp	
	Return
}

WinSURGECaseDataUpdater:  
{
	WinGetPos, x, y, w, h, Caris CodeRocket
	if((x<>CarisRocketWindowX OR y<>CarisRocketWindowY OR w<>CarisRocketWindowW+8 OR h<>CarisRocketWindowH+54) AND x<>-32000)
	{
	CarisRocketWindowX := x	
	CarisRocketWindowY := y
	CarisRocketWindowW := w - 8
	If CarisRocketWindowW < 100
		Return
	CarisRocketWindowH := h - 54
	IniWrite, %CarisRocketWindowX%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX
	IniWrite, %CarisRocketWindowY%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY
	IniWrite, %CarisRocketWindowW%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW
	IniWrite, %CarisRocketWindowH%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH
	}

;Block for getting UNDO button status
{  
SetTitleMatchMode, 2
ControlGet, UndoEnabled1, Enabled, ,&3 Undo, WinSURGE [, &2 Open Case
ControlGet, UndoEnabled2, Enabled, ,&Undo, WinSURGE - Final Diagnosis:
SetTitleMatchMode, 1
If(UndoEnabled1 = 1 OR UndoEnabled2=1)
	UndoEnabled := 1
Else
	UndoEnabled := 0

if(UndoEnabled AND !DataEntered)
	{
	GuiControl, Disable, CaseScanBox
	GuiControl, Disable, CaseLoaderLbl
	GuiControl, Text, CaseScanBox, Data in Case
}
Else if (OpMode="B" OR OpMode="D")
{
	Gui, Submit, NoHide
	GuiControl, Enable, CaseScanBox
	GuiControl, Enable, CaseLoaderLbl
	if CaseScanBox=Data in Case
		GuiControl, Text, CaseScanBox, 
}
}
	
if (OpMode="M" OR OpMode="B")
	Return

ifWinNotExist, WinSURGE
	{
	GuiControl, Text, PatientLabel, WinSURGE NOT OPEN
	GuiControl, Text, DoctorLabel, 
	GuiControl, Disable, CaseScanBox
	GuiControl, Disable, CaseLoaderLbl
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	GuiControl, 1:Hide, UsePhotos
	Return	
}


WinGetTitle, x, WinSURGE, &2 Open Case
if (x=lastWinSURGEtitle)
	Return
StringGetPos, y, x, No Current Case
StringGetPos, z, x, New
z := y + z
If (z>0 OR x="WinSURGE")
	{
	lastWinSURGEtitle := x
	GuiControl, 1:Text, PatientLabel, No Current Case
	GuiControl, 1:Text, DoctorLabel, 
	GuiControl, 1:Hide, UsePhotos
	GuiControl, 1:Hide, UseMicros
	;GuiControl, 1:Hide, UseICD9s
	Return
	}
lastWinSURGEtitle := x
DataEntered = 0
StringReplace, x, x, Case, |, All
StringSplit, y, x, |, %A_Space%
StringSplit, z, y2, %A_Space%, %A_space%
x := get_filled_case_number(z1)
s := "select s.dx, s.gross, s.numberofspecimenparts, s.custom03, s.clin, p.name, s.clindata, pt.name, s.Computed_PATIENTAGE from specimen s, physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
WinSurgeQuery(s)
CurrentCaseNumber := x
finaldiagnosistext := Result_1
grossdescriptiontext := Result_2
numberofvials := Result_3
OrderedCPTCodes := Result_4
ClientWinSurgeId := Result_5
ClientName := Result_6		
ClinicalData := Result_7
PatientName := Result_8
StringSplit, j, Result_9, .
PatientAge := j1
GuiControl, Text, PatientLabel, Patient: %PatientName% --- Age:%PatientAge%
GuiControl, Text, DoctorLabel, Doctor: %ClientName%

GuiControl, 1:Hide, UseMicros
;GuiControl, 1:Hide, UseICD9s

if OpMode=D
{
	GuiControl, 1:Hide, UsePhotos
	if (ClientWinSurgeId="")
		Return
		
	s = Select top 1 c.name,c.photo_pref,c.micro_pref,c.margin_pref, c.icd9_pref, c.log from clinipref c where c.WinSurge_id=%ClientWinSurgeId%
	CodeDatabaseQuery(s)
	If Result_1  ;If a database entry was found
		{
			ClinicianFound = 1
			CurrentLog := Result_6
			if Result_2  ;Photos Selected\
				{
				GuiControl, 1:Show, UsePhotos
				UsePhotos := 1
			}
			Else	
				{
				GuiControl, 1:Hide, UsePhotos
				UsePhotos := 0
			}

			If Result_3
				{
				GuiControl, 1:Show, UseMicros	
				UseMicros := 1
			}	
			Else	
			{
				GuiControl, 1:Hide, UseMicros
				UseMicros := 0
			}

			If Result_4
				{
				GuiControl, Text, UseMargins ,% Result_4
				UseMargins := Result_4
				}
			Else	
			{
				GuiControl, Text, UseMargins ,
				UseMargins := ""
			}

			If Result_5
				{
				;GuiControl, 1:Show, UseICD9s	
				;UseICD9s := 1
				}
			Else	
			{
				;GuiControl, 1:Hide, UseICD9s
				;UseICD9s := 0
			}

			return 			
		}
	Else    ;Else doctor was not found in the clinipref database
		{
			ClinicianFound = 0
			Gui, 1:Hide
			Gui, 1:+Disabled
			GuiControl, 2:, PhotoSelect, 0
			GuiControl, 2:, MicroSelect, 0
			;GuiControl, 2:, ICD9Select, 0
			GuiControl, 2:Text, DocPreferenceLabel, Enter preferences for %ClientName%
			Gui, 2:Show, , CarisDemo			;x%CarisRocketWindowX% y%CarisRocketWindowY% w%CarisRocketWindowW% h%CarisRocketWindowH%
			Gui, 2:-Disabled +AlwaysOnTop
			SoundBeep
			SoundBeep

			Return
		}
}

Return
}

F7::            ;MS done
{
	InputBox, SearchWord, Diagnostic Code Search, Enter the single word or phrase you want to search for:
	if SearchWord
		{
	
	dt =Diagosis Code Search TABLE`n---Searched for   "%SearchWord%"  ---`n
	counter := 0
	pages := 1

	Loop % LV_GetCount()
		{
		LV_GetText(RetrievedText1, A_Index, 2)
		LV_GetText(RetrievedText2, A_Index, 5)
		LV_GetText(RetrievedText3, A_Index, 6)
		LV_GetText(RetrievedText4, A_Index, 7)
		StringGetPos, t1, RetrievedText1, %SearchWord%
		StringGetPos, t2, RetrievedText2, %SearchWord%
		StringGetPos, t3, RetrievedText3, %SearchWord%
		StringGetPos, t4, RetrievedText4, %SearchWord%

		if (t1>-1 or t2>-1 or t3>-1 or t4>-1)
			{
			counter := counter + 1
			dt = %dt%%RetrievedText1% `n%RetrievedText2%`n  	
			if (RetrievedText3 OR RetrievedText4)
				dt = %dt%Comment:  %RetrievedText3%  %RetrievedText4%`n`n
			if (Floor(counter/9) =  counter/9)
			{
				dt := dt . "¶"
				pages := pages + 1
			}

			}
		}

	Loop, Parse, dt, ¶
		{
			Gui, Font, Courier
			Msgbox, 4097, Diagnostic Code Search Results, % A_LoopField . "`nPage " . A_Index . " of " . pages	
				IfMsgbox, Cancel
					Break
			Gui, Font, Arial
		}

	}
Return
}

+F7::
{
		InputBox, SearchWord, Extended Phrase Search, Enter the single word or phrase you want to search for:
	if SearchWord
		{
	
	dt =Extended Phrase Search`n---Searched for   "%SearchWord%"  ---`n
	dt = %dt%---------------------------------------------------------------------------------`n
	counter := 0
	pages := 1
	perc :="%"
	j='%perc%%SearchWord%%perc%'
	s = Select code,text from extendedphrases where text LIKE %j% OR code LIKE %j%
	
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := s
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	If A_LastError
		{
			Msgbox,  There was an error accessing the Code Database.`n Try again later. `ns=%s%`n  A_LastError = %A_LastError%
			Return
		}

If (rs.EOF=0)
		{
		rs.MoveFirst()
		while rs.EOF = 0{
			DXCodeCount := A_Index
			j := 0
			for field in rs.fields
				{
				j := j + 1
				y := Field.Value
				DxCode%j%=%y%
				}
			dt = %dt%%DxCode1% -- %DxCode2%`n`n
			counter := counter + 1
			if (Floor(counter/12) =  counter/12)
			{
				dt := dt . "¶"
				pages := pages + 1
			}

			rs.MoveNext()
			}
	rs.close()   
	adodb.close()
		}
	dt = %dt%---------------------------------------------------------------------------------`n

	Loop, Parse, dt, ¶
		{
			Gui, Font, Courier
			Msgbox, 4097, Diagnostic Code Search Results, % A_LoopField . "`nPage " . A_Index . " of " . pages	
				IfMsgbox, Cancel
					Break
			Gui, Font, Arial
		}
		}
Return
}

F8::           
{

	SaveError = 0	
	ifWinExist, WinSURGE Case Lookup	
		Return
		
;CPT Code Checking Block
	{
		finaldiag := WinSURGEFinalDiagnosisContents()
		StringGetPos, i, finaldiag, ***
		if i>0
			{
				SaveError = 1
				Msgbox, The final diagnosis is missing critical information (there is a '***' in the box)!
				Return
			}
		
		StringGetPos, i, finaldiag, `%`%88305`%`%
		StringGetPos, j, finaldiag, `%`%88304`%`%
		StringGetPos, k, finaldiag, `%`%88321`%`%
		StringGetPos, l, finaldiag, `%`%88323`%`%
		StringGetPos, m, finaldiag, `%`%88300`%`%
		StringGetPos, n, finaldiag, `%`%NOCHG`%`%
		
		if (i<0 AND j<0 AND k<0 AND l<0 AND m<0 AND n<0)
			{
				SaveError = 1
				Msgbox,4,, The final diagnosis is missing a proper billing code!  Do you wish to continue without further editing?
				IfMsgbox, No
					Return
			}

		StringGetPos, i1, finaldiag, PAS
		StringGetPos, i2, finaldiag, schiff
		StringGetPos, i3, finaldiag, GMS
		StringGetPos, i4, finaldiag, gomori 
		if (i1>0 OR i2>0 OR i3>0 OR i4>0)
			{
				perc := "%%"
				k = %perc%88312%perc%
				StringGetPos, j, finaldiag, %k%
				If j=-1
					{
					Msgbox,4,,PAS or GMS stain mentioned without adding an 88312 billing code.  Do you wish to continue without code editing?
					IfMsgbox, No
						{
						SaveError=1					
						Return	
						}
					}
			}

		
		CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
	}

;Photo presence checking block
	if (OpMode="D" AND UsePhotos)
		{
			x := WinSURGEOpenCasePhoto1()
			y := WinSURGEOpenCasePhoto2()
			if (!x AND !y)
				{
				Msgbox, 4, NO PHOTOS WARNING, Client has requested photos and there are none. Do you want to continue without photos?	
				IfMsgBox No
					{
					SaveError = 1	
					Return
					}
				}
			}
	
	;Progress, x300 y100 h150, , Queuing and Saving the case`n Press Ctrl-Alt-R if progress stops, Working....,
		
		x := SubStr(CurrentCaseNumber, 1, 1)
		if OpMode=D 
			{
			IfNotInString, HomeLabCasePrefix, %x%
				AppendOtherLabComment()
			}

	QueueandAssign()
	CloseandSaveCase()
	DataEntered = 0
	Return
}

+F8::
{
	if OpMode<>M
	{
	CloseandSaveCase()	
	DataEntered = 0
	GuiControl, Hide, StatusLabel
	Gui, Show, NoActivate
	Gosub, F12
	}
	Return
}

F9::
{
Gui, 3:Font, 
Gui, 3:show, ,List of Available Helpers
Return
}

F11::
{
	IfWinExist, WinSURGE - 
	{
		Send, %LastCodeUsed%	
		Gosub, Shift & Enter
	}
	Return
}

F12::
{
	WinActivate, WinSURGE , 	
	WinActivate, Voice -  Boomerang Enterprise Recorder
	WinActivate, Dictation Data
	WinActivate, Caris CodeRocket
	if OpMode<>M
	{
		GuiControl, Text, CaseScanBox,  ;Blanks the data entry textbox 	
		GuiControl, Focus, CaseScanBox,
	}
	
return
}

^k::   ;Special Stain order
{
	
	CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
	CloseWinSURGEModalWindow("WinSURGE - Gross Description","","&Close")
	IfWinNotActive, WinSURGE [, &2 Open , WinActivate, WinSURGE [, &2 Open
	WinWaitActive, WinSURGE [, &2 Open
	Sleep, 400
	DataEntered = 0
	WinMenuSelectItem, WinSURGE [, &2 Open, Tools, Enter Special Stains via Checklist
	Return
}

^!s::  ;Batch Signout
{
	IfWinExist, WinSURGE - Final Diagnosis
	{
		SoundBeep
		Msgbox, Close the current case first
		return
	}
	

	CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")

	;SetTimer, WinSURGECaseDataUpdater, Off

	Progress, x10 y10 h150, Preparing to signout, Obtaining Routine Cases for signout`n Press Ctrl-Alt-R to stop the signout, Working....,

s =	select s.number, s.zaudittraillast, s.yesno07 from specimen s, physician p where  s.sodate < '1950-01-01' and s.path = p.id and p.name ='%WinSurgeFullName%' and s.calculatedslidecountdate>'1950-01-01' order by 2 desc
	WinSurgeQuery(s)
	if !msg
		{
		Msgbox, There are no cases (not including amendments/addendums) to signout!	
		Progress, Off   ;Internal only
		SetTimer, WinSURGECaseDataUpdater, 1000
		Return
		}
	SignOutFileList = 
	SignOutCount := 0

Loop, Parse, msg, `n
		{
			LoopCaseNum = 
			Loop, Parse, A_LoopField, ¥
				{
				if A_Index =2	
					LoopCaseNum = %A_LoopField%
				if A_Index=3
					LoopYesNo07 = %A_LoopField%
				}
				if (LoopCaseNum<>"")
					{
					StringLeft, x, LoopCaseNum, 1
					StringRight, y, LoopCaseNum, 2
					;if ((x="C" OR y="MG") AND LoopYesNo07<>"Y")
						;Continue    ;Skip the C and MG cases where LoopYesNo07 is not set to Y
					SignoutFileList = %SignOutFileList%%LoopCaseNum%`n	
					SignOutCount := SignOutCount + 1
					}
		}

if (SignOutCount>0)
{
	Progress, , Preparing to signout, Signing out the cases`n, Working....,

	Loop, parse, SignOutFileList, `n
	{
		y := SignOutCount - A_Index + 1
		x := 100 * (A_Index / SignOutCount)
		Progress, ,  %y% of %SignOutCount% cases remaining...

		if A_LoopField =  ; Omit the last linefeed (blank item) at the end of the list.
			break
		Open4SignOut(A_LoopField)
		If JumptoCaseNotSet
			{
				Msgbox, You must set the "Jump to Case Entered by User" in the WinSURGE User Preferences menu in order to use the automated signout functions!
				Progress, Off
				Return
			}

		If (!ApprovalPasswordControl OR ApprovalPasswordControl="ERROR")  ;Setup function to get the name of the EsignoutTextBox the first time through	
		{
		appwlabel:
			Progress, Off
			SplashTextOn, 200,100, Setup..., Click your mouse into the "Approved:" password box and press F3.
			KeyWait, F3, d
			KeyWait, F3, u
			WinGetTitle, active_title, A
			ControlGetFocus, x,	%active_title%
			StringGetPos, y, x, TextBox
			if (y>0)
				{
					ApprovalPasswordControl := x	
					IniWrite, %ApprovalPasswordControl%, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ApprovalPasswordControl
				}
			Else
				{
					Msgbox, You did not click into a text box.  Please click your mouse into the "Approved:" password box and press F3.
					Goto, appwlabel
				}
		SplashTextOff
		}
			
		Progress, ,  %y% of %SignOutCount% cases remaining...`n  Press F3 to approve this case.`nPress F4 to skip and go to the next one.`nPress 'End' to End signout loop.

	If (ExpressSignout=1)
		EnterSignoutPasswordandApprove()	
	Else
		{
			Loop,
				{
					GetKeyState, F3state, F3
					GetKeyState, F4state, F4
					GetKeyState, EndState, End
					
					If F3state=D
						{
							EnterSignoutPasswordandApprove()	
							Break
						}
					If F4state=D
						{
							ControlGet, SkipEnabled, Enabled, ,&Skip, WinSURGE E-signout [, &Approve
							if SkipEnabled
								{
								SkipCaseSignout()
								Break
								}
							Else
								{
								Progress, Off   ;Internal only	
								SetTimer, WinSURGECaseDataUpdater, 2000
								Return
								}
						}
					If Endstate=D
						{
						Progress, Off   ;Internal only	
						SetTimer, WinSURGECaseDataUpdater, 2000
						Return
						}
				}
		}
}
Progress, Off   
SetTimer, WinSURGECaseDataUpdater, 2000
Return
}
Return
}

^!e::
{
GuiControl, 4:Text, DisplayCaseNumber, Case Number: %CurrentCaseNumber%
Gui, 4:Show
GuiControl, 4:Focus, EmailType
return
}

SendEmail:
{
		StringSplit, parts, CurrentCaseNumber, -
		np2 := LTrim(parts2, "0")
		ccn=%parts1%-%np2%
	
	Gui, 4:Submit
	EmailComments := RegExReplace(EmailComments, " ", "`%20")
	
	if (Loc="Boston")
	{
		Email1 := "g_adminstaff@cohenderm.com"
		Email2 := "g_clientservices@cohenderm.com"
		Email3 := "g_clientservices@cohenderm.com"
		Email4 := "g_clientservices@cohenderm.com"
		Email5 := "BostonHistoSupsLeads@miracals.com"
	}
	else if (Loc="Irving")
	{
		Email1 := "IRVING-SPCDEPT@MiracaLS.com"
		Email2 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email3 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email4 := "PathologySupportClientServicesIRVPHX@MiracaLS.com"
		Email5 := "IrvingPA-Leads@MiracaLS.com"
	}
	else
		return
	
	if (EmailType=1)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email1%&subject=Patient`%20Double`%20Blind`%20Error`%20on`%20%ccn%&body=(%EmailComments%)
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}

	else if (EmailType=2)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email2%&subject=Need`%20Previous`%20Biopsy`%20Report`%20on`%20%ccn%&body=(%EmailComments%)
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=3)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email3%&subject=Clinical`%20Note`%20and`%20Photos`%20on`%20%ccn%&body=Please`%20obtain`%20from`%20client`%20clinical`%20note`%20and`%20photos.`%20(%EmailComments%)
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=4)
		{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email4%&subject=Critical`%20Result`%20Call`%20%ccn%&body=Please`%20fax`%20and`%20call`%20confirm`%20only!
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}
	else if (EmailType=5)
				{
		SetTitleMatchMode, 2
		Run, "%OutlookPath%" /c ipm.note  /m %Email5%&subject=Please`%20Pull`%20The`%20Bottles`%20And`%20Blocks`%20on`%20%ccn%&body=Please`%20bring`%20them`%20to`%20my`%20office`%20for`%20review`%20with`%20the`%20slides!
		WinWaitActive, (HTML)
		SetTitleMatchMode, 1
		Send, !s
		}

	
	;Msgbox, %EmailComments%
	return
}

^!u::
{
checkForMelanoma:
	SplashTextOn, 100, 100, Melanoma Call Checker, Initializing...

	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)
	pathWinSurgeId := Result_5
	

		
		searchdate = %a_now%
		searchdate += -5, days
		FormatTime, searchdate, %searchdate%, yyyy-MM-dd
		
/* 		lasttime = 12:00:00 AM
 * 
 * 		finishTime := A_Now - 1500
 * 		FormatTime, endtime, %finishTime%, h:mm:ss tt
 * 		StringSplit, j, endtime, %A_Space%:
 */

	s := "select s.number, s.sotime, s.text01, s.sodate, s.dx from specimen s where s.path =" . pathWinSurgeId . " and s.dx LIKE '%MELANOMA%' and s.sodate >= '" . searchdate . "'" ;and s.sotime >= '" . lasttime . "' and s.sotime <= '" . endtime . "'"

	SplashTextOn, 100, 100, Melanoma Call Checker, Searching for Melanomas...

	WinSurgeQuery(s)
	If A_LastError
		{
			Msgbox, There was an error accessing the WinSURGE Database to check your Melanoma cases.`n`ns=%s%`n  A_LastError = %A_LastError%`n
			Return
		}	
		
	SplashTextOn, 100, 100, Melanoma Call Checker, Filtering for Client Service Calls...

	caselist := ""
	Loop, parse, msg, ¥ 
	{
	if (A_Index = 1)
		Continue
	Else
		{
		x := InStr(A_LoopField,"%%mbn%%")
		x := x + Instr(A_LoopField, "%%mis%%")
		x := x + Instr(A_LoopField, "%%misl%%")
		x := x + Instr(A_LoopField, "%%miss%%")
		x := x + Instr(A_LoopField, "%%mm%%")
		x := x + Instr(A_LoopField, "%%mmm%%")
		x := x + Instr(A_LoopField, "%%mis1%%")
		x := x + Instr(A_LoopField, "%%mmaz%%")
		x := x + Instr(A_LoopField, "%%misaz%%")
		x := x + Instr(A_LoopField, "%%lmaz%%")
		x := x + Instr(A_LoopField, "%%aimm1%%")
		x := x + Instr(A_LoopField, "%%aimm2%%")
		x := x + Instr(A_LoopField, "%%asnmm%%")
		x := x + Instr(A_LoopField, "%%pmis%%")
		x := x + Instr(A_LoopField, "%%nmis%%")
		x := x + Instr(A_LoopField, "%%nmm%%")
		x := x + Instr(A_LoopField, "%%omm%%")
		x := x + Instr(A_LoopField, "%%omis%%")
		x := x + Instr(A_LoopField, "%%pamis%%")
		x := x + Instr(A_LoopField, "%%cmm%%")
		x := x + Instr(A_LoopField, "%%misdn%%")
		x := x + Instr(A_LoopField, "%%mmbap%%")
		x := x + Instr(A_LoopField, "%%mmpn%%")
		x := x + Instr(A_LoopField, "%%idmm%%")

		if x>0
			{
				isFaxed := Instr(intcomment, "fax")
				isFaxed := isFaxed + Instr(intcomment, "notified")
				isFaxed := isFaxed + Instr(intcomment, "called")
				isFaxed := isFaxed + Instr(intcomment, "discussed")
				
				if (isFaxed<=0)
					caselist = %caselist%[ %casenumber% ];  %sodate%`n`nInternal Comments:`n%intcomment%`n`n
			}
		casenumber := sotime
		sotime := intcomment
		intcomment := sodate
		sodate := A_LoopField
	}

	}

	if(caselist)
	{
		SplashTextOff
		SoundBeep
		Msgbox, MELANOMA CALL CHECKER`nIt appears you have one or more cases that was/were melanoma but for which client services has not "fax"ed, "called", or "notified" the client.`nPlease check the below information and take appropriate action!`n`n%caselist%
	}
	else
	{
		SplashTextOn, 100, 100, Melanoma Call Checker, All Cases Notified
		Sleep, 500
	}
		

	caselist := ""
	
	SplashTextOff

	return
}

^!c::
{
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)

	jarcount:=casecount:=mxcount:=0
    InputBox, searchdate, Search Date "MM-DD", Please enter the date to search.  Leave blank for today!, , 640, 480
	if (searchdate)
		todaydate = %A_YYYY%-%searchdate%
	else
		todaydate = %A_YYYY%-%A_MM%-%A_DD%
	
	s := "select s.number, s.numberofspecimenparts from specimen s where s.path =" . Result_5 . " and s.sodate = '" . todaydate . "'"
	WinSurgeQuery(s)
	Loop, parse, msg, `n 
		{
			If(A_LoopField)
			{
				StringSplit, res, A_LoopField, ¥
				casecount := casecount + 1
				jarcount := jarcount + res3
				IfInString, res2, MX
					mxcount := mxcount + 1
			}
		}
	Msgbox, % todaydate . "`n" . "Cases: " . casecount . ", Jars: " .  jarcount + mxcount * 2 . "`nNote that counts are only approximate and may not account for all consults and two-tray cases."
	return	
}

^!y::
{
	CaseNumberProblem=0
	InputBox, scancasenum, Scan the case...
	
	;Msgbox, %scancasenum%
	StringSplit, x, scancasenum, %A_Space%
	StringSplit, y, x1, -
	Stringlen, csbLength, y1
	if (csbLength<>4)
		CaseNumberProblem=1
	StringLeft, csbPrefix, y1, 2
	StringRight, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringLeft, j, csbPrefix, 1
	IfNotInString, letters, %j%
		CaseNumberProblem=1
	StringRight, csbYear, y1, 2
	csbCaseNum := y2
	StringRight, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLeft, j, csbYear, 1
	IfNotInString, numbers, %j%
		CaseNumberProblem=1
	StringLen, casenumlength, csbCaseNum
	Loop, %casenumlength%
	{
		StringMid, j, csbCaseNum, A_Index, 1
		IfNotInString, numbers, %j%
			CaseNumberProblem=1
	}

	if CaseNumberProblem
			{
				Msgbox, You did not enter a valid case number!
				Gosub, F12
				Return
			}

	NewCaseNum=%csbPrefix%%csbYear%-%csbCaseNum%
	Msgbox, Nothing
	x := get_filled_case_number(z1)
	s := "select s.dx, s.gross, s.numberofspecimenparts, s.custom03, s.clin, p.name, s.clindata, pt.name, s.Computed_PATIENTAGE from specimen s, physician p, patient pt where s.patient = pt.id and s.clin=p.id and computed_numberfilled='" . x . "'"
	WinSurgeQuery(s)
	CurrentCaseNumber := x
	finaldiagnosistext := Result_1
	grossdescriptiontext := Result_2
	numberofvials := Result_3
	OrderedCPTCodes := Result_4
	ClientWinSurgeId := Result_5
	ClientName := Result_6		
	ClinicalData := Result_7
	PatientName := Result_8
	StringSplit, j, Result_9, .
	PatientAge := j1
	Msgbox, %NewCaseNum%`n%PatientName%`n%ClinicalData%`n%finaldiagnosistext%
	Return
}

^!v::
{
	ListVars
	Return
}

^!l::  ;Hotkey for testing the program
{
	ListLines
	Return
}

^!p::
{
if A_Username=mmuenster
{
if (ExpressSignout=0)
	{
	IniWrite, 1, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ExpressSignout		
	ExpressSignout := 1
	SplashTextOn, 100,100,,Express Mode On
	}
Else
	{
	IniWrite, 0, %A_MyDocuments%\CarisCodeRocket.ini, Window Positions, ExpressSignout		
	ExpressSignout := 0
	SplashTextOn, 100,100,,Express Mode Off
	}
Sleep, 800	
SplashTextOff
}
Return
}

^!1::
{
	If (A_Username<>"mmuenster")
		Return
		
	FileDelete, S:\CodeRocket\bin\EP\_ermpathExtendedPhrases.ahk
	FileMove, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk, S:\CodeRocket\bin\EP\_ermpathExtendedPhrases.ahk
	FileAppend, #NoTrayIcon`n#SingleInstance force`n#Hotstring EndChars  ``t`n#IfWinActive`, WinSURGE - `n, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk

	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	strRequest := "SELECT * FROM extendedphrases"
	adodb.open(connectstring)
	rs := adodb.Execute(strRequest)
	rs.MoveFirst()
	while rs.EOF = 0{
		DXCodeCount := A_Index
		j := 0
		for field in rs.fields
			{
			j := j + 1
			y := Field.Value
			DxCode%j%=%y%
			}
			SplashTextOn, 100, 100, EP convertor, Doing %DXCodeCount%
			FileAppend, ::%DxCode2%::%DxCode3%`n, S:\CodeRocket\bin\EP\DermpathExtendedPhrases.ahk
			rs.MoveNext()
	}
	rs.close()   
	adodb.close()

	SplashTextOff
	Return
}

Pause::Pause
^!r::
{
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

^!q::Reload
^+!r::
{
	FileDelete, %A_MyDocuments%\CarisCodeRocket.ini
	Run, C:\Documents and Settings\All Users\Desktop\Launcher - Caris CodeRocket.exe
	ExitApp
	Return
}

#IfWinActive, WinSURGE - 
Shift & Enter::
{
		Send, {#}{#}
		Sleep, 200
		ControlGetText, x, TX321, WinSURGE - 	
		finaldxcontents := x
		Send, {Backspace}{Backspace}
		StringGetPos, y, x, ##	
		Stringlen, totallen, x
		firsthalflen:= y
		secondhalflen := totallen-y-2

			TempMicros = 0	
			;TempICD9s = 0	

			Loop,parse, x, `n
			{
				CurrentLine := A_LoopField
				IfInString, CurrentLine, ##
					{
					Loop, parse, CurrentLine, %A_Space%, `n `r	
						{
						CurrentWord := A_LoopField
						IfInstring, CurrentWord, ##
							{
								StringSplit, wordpart, CurrentWord, ##
								StringLen, len, wordpart1
								WinActivate, WinSURGE - 
								Loop, %len%
									{
									Send, {Backspace}	
									firsthalflen := firsthalflen -1
									}
								DxCode := wordpart1
								LastCodeUsed := DxCode
								StringLen, len, DxCode
								StringRight, y, DxCode, 2
								IfInstring, y, /
									{
									TempMicros = 1	
									len := len -1
									}
								IfInstring, y, *
									{
									;TempICD9s = 1	
									len := len -1
									}
								StringLeft, DxCode, DxCode, len
							}
						}
					}
			}	


				ErrorLevel := ParseandCheckDXCode()				
				If ErrorLevel
					return


				;Loop to add front helper codes to the diagnosis
				Stringlen, i, fronttophelp
				Loop, %i%
				{
					x := i + 1 - A_Index
					StringMid, j, fronttophelp, x, 1
					p := FrontofDiagnosisHelper%j%
					SelDiagnosis = %p% %SelDiagnosis%
				}

				;Loop to add Back helper codes to the diagnosis
				Stringlen, i, backtophelp
				
				
				


				if (i>1)
					{
					Msgbox, You can only enter one margin code!	
					GuiControl, Text, DXCode,  ;Blanks the data entry textbox 
					GuiControl, Focus, DXCode
					return
					}
				
				if i=1
					{
					IfInString, UseMargins, none
						{
							SoundBeep
							Msgbox,4,,This client has requested the following margin preferences (%UseMargins%) and you have used one!  Do you wish to continue?
							IfMsgbox, No
								Return	
						}
						
						;Msgbox, %basediag%`n%fronttophelp%`n%basediag%`n%backtophelp%`n%comhelp%`n%SelTags%`n
						;z1 := CheckClientPreferences(SelTags, clinPref.Margin_pref)
						;z2 := IsExcision(finaldiagnosistext)
						;z3 := MarginsRequested(ClinicalData)
						;if(z2 OR z3)
							;SoundBeep
						
						;if(z1=1)
							;SoundPlay, c:\Windows\Media\GoodRead.wav
						;else if (z1=-1)
							;SoundPlay, c:\Windows\Media\BadRead.wav

						
						p := BackofDiagnosisHelper%backtophelp%
						IfInString, SelDiagnosis, /-/
						{
							StringGetPos, x, SelDiagnosis, /-/
							StringLen, l, SelDiagnosis
							StringLeft, DiagFirstLine, SelDiagnosis, %x%
							y := l - x
							StringRight, RestofDiag, SelDiagnosis, %y%
							SelDiagnosis = %DiagFirstLine%; %p%%RestofDiag%	
						}
						Else
							SelDiagnosis = %SelDiagnosis%; %p%
					}	

				;Loop to add comment helper codes to the comment
				Stringlen, i, comhelp
				Loop, %i%
				{
					StringMid, j, comhelp, %A_Index%, 1
					p := CommentHelper%j%
					SelComment = %SelComment%  %p%  `
				}


				x3 := SelDiagnosis ;Diagnosis
					StringRight, lastletter, x3, 1
						if (lastletter<>"." AND x3<>"")
							x3 = %x3%.
				x4 := SelComment ;Comment
				x5 := SelMicro ;Micro
				x6 := SelCPTCode ;CPT Code
				cr = `n
				StringReplace, x3, x3, /-/, %cr%, All
				StringReplace, x4, x4, /-/, %cr%, All
				perc := "%%"
				dxtext = %x3%
				If ((TempMicros OR UseMicros) AND !x5)
					Msgbox, Client has requested microscopic descriptions and there is not one for this diagnostic code!  Please enter manually.
				If (x4 or (x5 and (TempMicros OR UseMicros)))
					{
					If (TempMicros OR UseMicros)
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%%A_Space%%A_Space%%x5%
					Else
						dxtext = %dxtext%`n`nComment:%A_Space%%x4%
					}
				
				;if (TempICD9s OR UseICD9s)
					;{
					;if SelICD9	
						;dxtext = %dxtext% (%SelICD9%)
					;Else
						;dxtext = %dxtext% (***)
					;}

				dxtext = %dxtext%`n
				If x6
					{
					Loop, parse, x6,`;
						dxtext = %dxtext%%perc%%A_LoopField%%perc%	
					}
				dxtext = %dxtext%%perc%%SelDXCode%%perc%
					
				SetCapsLockState, Off
			
		StringLeft, firsthalf, finaldxcontents, %firsthalflen%
		StringRight, secondhalf, finaldxcontents, %secondhalflen%
		newtext =%firsthalf%%dxtext%%secondhalf%
		
		if UseSendMethod
			{
			Send, %dxtext%	
			Sleep, 50
			Send, {Shift Down}   ;these lines are to correct the but where the shift key is locked down after a send.
			Sleep, 50
			Send, {Shift Up}
			}
		Else
			ControlSetText, TX321, %newtext%, WinSURGE -  

			DataEntered = 1
			

			ifWinActive,  WinSURGE - Final Diagnosis:
			{
			if OpMode<>M
				{
				Gui, Submit, NoHide	
				GuiControl, Enable, CaseScanBox
				GuiControl, Enable, CaseLoaderLbl
				if CaseScanBox=Data in Case
					GuiControl, Text, CaseScanBox, 
				}

			
			finaldiag := WinSURGEFinalDiagnosisContents()
			StringGetPos, i, finaldiag, ***
			if i>0
				ActivateNextTripleAsterisk()
			Else if OpMode<>M
				Gosub, F12
			}

Return
}

#IfWinActive, Special Stains Checklist
WheelDown::
{
	ControlGetPos, x, y, w, h, Save && Close, Special Stains Checklis
	ControlGetPos, x, y, w, h, Save && Close, Special Stains Checklis
	x := x + w/2
	y := y + h/2
	MouseClick, left, x, y
	WinWaitClose, Special Stains Checklist, , 4
	if ErrorLevel
		return
	
	Loop,
		if (A_Cursor="Arrow" OR A_Cursor="IBeam")
			break
	WinWaitActive, WinSURGE , 
	Sleep, 300
	WinActivate, Caris CodeRocket
return
}

#IfWinActive     ;Resets #IfWin directive so that hotkeys can be turned off
