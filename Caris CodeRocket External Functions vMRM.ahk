;DATA ENTRY FUNTIONS
ActivateNextTripleAsterisk() 
{
		WinActivate, WinSURGE -
		Send, ^g
		Return
}
	
AppendOtherLabComment() 
{
	global
	SetControlDelay, 100
	OpenGrossDescriptionModal()
	ControlFocus, TX321, WinSURGE - Gross Description:
	ControlGetPos, x,y,w,h,TX321, WinSURGE - Gross Description:

	mx := x+w-50
	my := y+h-50
	ControlGetText, grossdescriptiontext, TX321, WinSURGE - Gross Description:
	StringGetPos, z, grossdescriptiontext, Microscopic Examination performed
	if z=-1
		{
		grossdescriptiontext = %grossdescriptiontext%  %HomeLabGrossAddendum%`n
		ControlSetText, TX321, %grossdescriptiontext%, WinSURGE - Gross Description:
		ControlFocus, TX321, WinSURGE - Gross Description:
		MouseClick, left, %mx%, %my%
		Send, {End}
		Sleep, 100
		Send, {Space}
		}
	Else
		{
		StringLeft, grossdescriptiontext, grossdescriptiontext, %z%
		grossdescriptiontext = %grossdescriptiontext%  %HomeLabGrossAddendum%`n
		ControlSetText, TX321, %grossdescriptiontext%, WinSURGE - Gross Description:
		ControlFocus, TX321, WinSURGE - Gross Description:
		MouseClick, left, %mx%, %my%
		Send, {End}
		Sleep, 100
		Send, {Space}
		}
	CloseWinSURGEModalWindow("WinSURGE - Gross Description","","&Close")
	Return 0
}

CloseandSaveCase() 
{
		global
		SetTitleMatchMode, 2
		CloseWinSURGEModalWindow("WinSURGE - Final","","&Close")
		ControlGetPos , X, Y, Width, Height, TX321, WinSURGE [, 
		xscan := Round(x + 0.65*Width)		
		yscan := Round(y + 0.3*height)

		IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open			
		WinWaitActive, WinSURGE [, &7 Save && Open
		Loop,5
		{
			xt := xscan -1 + A_Index
			PixelGetColor, Color, xt, yscan
			If (color=0xC0C0C0 OR color=0xC6C3C6)
				{
				xscan := xt
				Break	
				}
			else if A_Index=5
				{
				Msgbox, No edits to the current case are detected.	
				Return
				}
		}

			Loop,
			{
				If A_Index=1
					LabeledButtonPress("WinSURGE [","&7 Save && Open","&7 Save")
				Sleep, 100
				PixelGetColor, Color, xscan, yscan
				If Color=0xFFFFFF
					Break
			}	
		
Return 
}

CloseWinSURGEModalWindow(WinTitle,WinText,CloseButton)
{
	Loop,
	{
		IfWinExist, %WinTitle%,%WinText%
			WinClose, %WinTitle%, %WinText%
	
		SoundMade := 0
		Sleep, 300
		Loop,
		{
			IfWinExist, Word Not Found In Dictionary	
				{
				if !SoundMade
					{
					SoundBeep, 800, 100
					Sleep, 50
					SoundBeep, 800, 100
					Sleep, 50
					SoundBeep, 800, 100
					SoundMade := 1
					}
				Sleep, 1000	
				}
			Else
				Break
		}
		

		IfWinExist, %WinTitle%,%WinText%
			LabeledButtonPress(WinTitle, WinText, CloseButton)
			
		IfWinNotExist, %WinTitle%,%WinText%
			Break
	}
	Return
}
	
LabeledButtonPress(WinTitle, WinText, ButtonLabel)
{
	IfWinExist, %WinTitle%, %WinText%
		{
			ControlGetPos, x,y,w,h,%ButtonLabel%, %WinTitle%, %WinText%	
			x := Round(w/2 + x)
			y := Round(h/2 + y)
			WinActivate, %WinTitle%, %WinText%
			MouseClick, left, %x%, %y%
		}				
	Return
}

OpenGrossDescriptionModal() 
{
		global
		Loop,
		{
		IfWinNotExist, WinSURGE - Gross Description:
			{
			WinWait, WinSURGE [, &7 Save && Open 
			IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open 
			WinWaitActive, WinSURGE [, &7 Save && Open 
			MouseClick, left, %GrossDescriptionButtonX%, %GrossDescriptionButtonY%	
			}
			
			WinWait, WinSURGE - Gross Description: , ,2
			If ErrorLevel
				Continue
			Else
				Break
		}

		IfWinNotActive, WinSURGE - Gross Description:, , WinActivate, WinSURGE - Gross Description:, 
		WinWaitActive, WinSURGE - Gross Description:, 
		WinMove, WinSURGE - Gross Description:,, %WinSURGEModalWindowX%,%WinSURGEModalWindowY%,%WinSURGEModalWindowW%,%WinSURGEModalWindowH%

		Return 0
	}

OpenFinalDiagnosisModal() 
{
		global
		Loop,
		{
		IfWinNotExist, WinSURGE - Final Diagnosis:
			{
			CloseWinSURGEModalWindow("WinSURGE Case Lookup","","Cancel")
			WinWait, WinSURGE [, &7 Save && Open 
			IfWinNotActive, WinSURGE [, &7 Save && Open , WinActivate, WinSURGE [, &7 Save && Open 
			WinWaitActive, WinSURGE [, &7 Save && Open 
			MouseClick, left, %FinalDiagnosisButtonX%, %FinalDiagnosisButtonY%	
			}
			
			WinWait, WinSURGE - Final , ,2
			If ErrorLevel
				Continue
			Else
				{
				Break
				}
		}

		IfWinNotActive, WinSURGE - Final Diagnosis:, , WinActivate, WinSURGE - Final Diagnosis:, 
		WinWaitActive, WinSURGE - Final Diagnosis:, 
		WinMove, WinSURGE - Final Diagnosis:,, %WinSURGEModalWindowX%,%WinSURGEModalWindowY%,%WinSURGEModalWindowW%,%WinSURGEModalWindowH%
		Return 0
	}

OpenCase(obj) 
{

 	global
/* 	WinActivate, WinSURGE
 * 	Loop, 
 * 		{
 * 			IfWinNotExist, WinSURGE Case Lookup
 * 				Send, !2 ;LabeledButtonPress("WinSURGE","&7 Save && Open","&2 Open Case")
 * 			Else
 * 				Break
 * 				
 * 			Sleep, 50
 * 		}
 */

	WinActivate, WinSURGE Case Lookup
	WinWaitActive, WinSURGE Case Lookup
	If !CaseLookupCaseNumberTextBox   ;Setup function to get the name of the Case Lookup Textbox the first time through
		{
		Loop,
			{
				WinGetTitle, active_title, A
				If (active_title="WinSURGE Case Lookup")
					{
						Sleep, 300
						ControlGetFocus, x,	%active_title%
						CaseLookupCaseNumberTextBox := x
						Break
					}
					
				Sleep, 500
			}
		}
		
	ControlSetText, %CaseLookupCaseNumberTextBox%, %obj%, WinSURGE Case Lookup	
	ControlFocus, %CaseLookupCaseNumberTextBox%, WinSURGE Case Lookup
	Send, {Enter}     ;Need two enters to properly get to QueueIntoBatchbox

	Loop,
	{
		MouseMove, 500, 500
		Sleep, 100
		If (A_Cursor="Wait")
			continue
		else
			break
	}

Return
}

QueueandAssign() 
{
			global
			
			ControlFocus, %QueueIntoBatchBox%, WinSURGE [
			Send, final{TAB}
			Sleep, 10
			Send, {TAB}Muens{TAB}
			Sleep, 10
			



/* 			Loop, 
 * 			{
 * 				SetControlDelay, 200
 * 				ControlSetText, %QueueIntoBatchBox%, Final reports, WinSURGE [
 * 				ControlClick, %PathologistTextBox%, WinSURGE [
 * 				ControlSetText, %PathologistTextBox%, , WinSURGE [
 * 				len := StrLen(WinSurgeFullName)
 * 				len := len -1
 * 				StringLeft, PartialName, WinSurgeFullName, %len%
 * 				ControlSetText, %PathologistTextBox%, %PartialName%, WinSURGE [
 * 				ControlSend, %PathologistTextBox%, {TAB}, WinSURGE [
 * 				ControlClick, %QueueIntoBatchBox%, WinSURGE [
 * 
 * 				Loop,20 
 * 					{
 * 					ControlGetText, t1, %PathologistTextBox%, WinSURGE [
 * 					Sleep, 100
 * 					if t1=%WinSurgeFullName%
 * 						Break
 * 					}
 * 				
 * 				if t1=%WinSurgeFullName%
 * 						Break
 * 			}
 */
return

}

SkipCaseSignout()
{
	global
	WinWaitClose, WinSURGE E-signout, abcdefgABCDEFG 1234567890, , [
	WinActivate, WinSURGE E-signout [, &Approve
	WinWaitActive, WinSURGE E-signout [, &Approve

	Loop,
	{
	LabeledButtonPress("WinSURGE E-signout [","&Approve","&Skip")
	Sleep, 1000

	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		Break
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		Break
	}
Return
}

WinSURGEFinalDiagnosisContents() 
{
	ifWinNotExist, WinSURGE - Final Diagnosis:
		OpenFinalDiagnosisModal()	
	ControlGetText, t, TX321, WinSURGE - Final Diagnosis:	
	Return t
}

WinSURGEOpenCasePhoto1() 
{
	global
	ControlGetText, t1, %Photo1TextBox%, WinSURGE [
	Return t1
}

WinSURGEOpenCasePhoto2() 
{
	global
	ControlGetText, t1, %Photo2TextBox%, WinSURGE [
	Return t1
}

;CASE SIGNOUT FUNCTIONS
EnterSignoutPasswordandApprove() 
{
	global
	WinWaitClose, WinSURGE E-signout, abcdefgABCDEFG 1234567890, , [
	WinActivate, WinSURGE E-signout [, &Approve
	WinWaitActive, WinSURGE E-signout [, &Approve
	

	Loop,    ;Makes sure the signout password textbox has the focus.
		{
			Sleep, 200
			ControlFocus, %ApprovalPasswordControl%, WinSURGE E-signout [
			ControlGetFocus, ctrlfoc, WinSURGE E-signout [
			if ctrlfoc=%ApprovalPasswordControl%
				Break
			Else ifWinExist, WinSURGE E-signout, abcdefgABCDEFG 1234567890, [ ;In case the window for signout just pops up!
				Return	

		}

	
	Loop,    ;Enters the sopw
	{
		ControlSetText, %ApprovalPasswordControl%, , WinSURGE E-signout [	
		ControlSetText, %ApprovalPasswordControl%, %WinSurgeSignoutPassword%, WinSURGE E-signout [	
		ControlGetText, t1, %ApprovalPasswordControl%, WinSURGE E-signout [
		if t1=%WinSurgeSignoutPassword%
			Break
	}
	SetTitleMatchMode, 2
	Loop,
	{
	LabeledButtonPress("WinSURGE E-signout [","&Approve","&Approve")
	Sleep, 1000

	If WinExist("WinSURGE", "OK","E-signout")
		{
			Sleep, 500
			ControlClick, OK, WinSURGE, Ok, left
			SplashTextOn, 200,100, Setup..., Close any error windows.  Then type your signout password into the box and press F3 on your keyboard
			ControlClick, OK, WinSURGE, Ok, left
			KeyWait, F3, d
			KeyWait, F3, u
			SplashTextOff
			ControlGetText, t1, %ApprovalPasswordControl%, WinSURGE E-signout [
			WinSurgeSignoutPassword=%t1%
			IniWrite, %WinSurgeSignoutPassword%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
			Continue
		}
	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		Break
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		Break
	}
SetTitleMatchMode, 1
Return
}

Open4SignOut(obj) 
{
	global
	WinMenuSelectItem, WinSURGE , &7 Save && Open , Tools, E-Signout
	WinWaitActive, WinSURGE E-signout,abcdefgABCDEFG 1234567890,,[
	Send, %obj%
	LabeledButtonPress("WinSURGE E-signout", "abcdefgABCDEFG 1234567890", "OK")

/* 	JumptoCaseNotSet := 0
 * 	IfWinNotExist, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 				{
 * 				WinMenuSelectItem, WinSURGE , &7 Save && Open , Tools, E-Signout	
 * 				Loop, 10
 * 					{
 * 							WinWait, WinSURGE E-signout,abcdefgABCDEFG 1234567890,0.3,[
 * 							If ErrorLevel
 * 								Continue
 * 							Else
 * 								Break
 * 					}
 * 						
 * 				ifWinExist, WinSURGE E-signout [, R&eturn	
 * 					{
 * 					JumptoCaseNotSet := 1	
 * 					Return
 * 					}
 * 				}
 * 	
 * 	;WinWait, WinSURGE E-signout,abcdefgABCDEFG 1234567890,,[
 * 	;IfWinNotActive, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[, , WinActivate, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[,
 * 	WinWaitActive, WinSURGE E-signout,abcdefgABCDEFG 1234567890,,[
 * 	If (!EsignoutTextBox OR EsignoutTextBox="ERROR")  ;Setup function to get the name of the EsignoutTextBox the first time through
 * 		{
 * 		ControlFocus, ThunderRT6TextBox, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 		Sleep, 100
 * 		ControlGetFocus, x,	WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 		StringGetPos, t, x, TextBox
 * 		if t>0
 * 			EsignoutTextBox := x
 * 		Else
 * 			{
 * 				ToolTip, Click in the "Enter case to jump to:" box and press F3!
 * 				KeyWait, F3, d
 * 				KeyWait, F3, u
 * 				ControlGetFocus, x,	WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 				EsignoutTextBox := x
 * 				ToolTip   ;Removes the tooltip
 * 			}
 * 
 * 		IniWrite, %EsignoutTextBox%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, EsignoutTextBox
 * 		}
 * 
 * 	Loop,   ;Enters the case number to the E-signout window
 * 	{
 * 		ControlSetText, %EsignoutTextBox%, %obj%, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[	
 * 		Sleep, 30
 * 		ControlGetText, t1, %EsignoutTextBox%, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 		if t1=%obj%
 * 			Break
 * 	}
 * 
 * 	Loop,   ;Clicks the OK button to the E-signout window
 * 	{
 * 		IfWinExist, WinSURGE E-signout,abcdefgABCDEFG 1234567890,[
 * 			{
 * 			LabeledButtonPress("WinSURGE E-signout", "abcdefgABCDEFG 1234567890", "OK")
 * 			WinWaitClose, WinSURGE E-signout,abcdefgABCDEFG 1234567890, 2, [
 * 			If ErrorLevel
 * 				Continue
 * 			Else
 * 				Break
 * 			
 * 			}
 * 	}
 */
Return
}

ReleaseApprovedCases()
{
	global
	if WinExist("WinSURGE E-signout","abcdefgABCDEFG 1234567890","[","")
		{
			Loop,
			{
				LabeledButtonPress("WinSURGE E-signout","abcdefgABCDEFG 1234567890","Cancel") 
				WinWaitClose, WinSURGE E-signout,abcdefgABCDEFG 1234567890,1,[
				If ErrorLevel
					Continue
				Else
					Break
			}
			
			WinWait, WinSURGE, Retry, 1
			If ErrorLevel
				LabeledButtonPress("WinSURGE E-signout","abcdefgABCDEFG 1234567890","&Release")
			Else
				LabeledButtonPress("WinSURGE","Retry","Yes")
			
			
			WinWait, E-Signout Approval & Release
			WinActivate, E-Signout Approval & Release
			LabeledButtonPress("E-Signout Approval & Release","&Release","&Release")	
			Return
		}
		
	If WinExist("WinSURGE", "abcdefgABCDEFG 1234567890","E-signout","")    
		{
			LabeledButtonPress("WinSURGE", "&Yes","&Yes")
			WinWait, E-Signout Approval & Release
			WinActivate, E-Signout Approval & Release
			LabeledButtonPress("E-Signout Approval & Release","&Release","&Release")	
			Return
		}
}

;DATABASE QUERY FUNCTIONS
{
WinSurgeQuery(s)
{
	global
	Loop, 15          ;Blank the results array
		Result_%A_Index% =

	connectstring := "DRIVER={InterSystems ODBC};SERVER=s-irv-wsg01;DATABASE=wins1csql;uid=_system;pwd=sys;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	adodb.open(connectstring)
RetryWinSURGEDatabase:
	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the WinSURGE Database.`n`ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryWinSURGEDatabase
			else ifMsgBox, Cancel
				ExitApp
		}	
	msg := ""
	txt := rs.state
	If !txt
		return
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		msg = %msg%`n
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}
	
	rs.close()   
	adodb.close()
	return 	
}

CodeDatabaseQuery(s)  ;((((((((((((((((
{
global
RetryCodeDatabase:
	Loop, 15          ;Blank the results array
		Result_%A_Index% =
		
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	adodb.open(connectstring)

	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the Code Database.`n  `ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryCodeDatabase	
			else ifMsgBox, Cancel
				ExitApp
		}	
		
	msg := ""
	txt := rs.state
	If !txt
		return
		
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}

	rs.close()   
	adodb.close()
	return msg	
}
}
