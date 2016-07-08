get_filled_case_number(c)
{
    StringUpper c,c
	ret := ""
    StringSplit,arr,c,-
	if (arr0 = 2) {
        stringlen,len,arr2
		repeatnum := 6-len
		zeros:=""
		cnt := 0
		While (cnt < repeatnum){
			cnt := cnt +1
			zeros:=zeros . "0"
		}
		ret := arr1 . "-" . zeros . arr2
	}		
	return ret
} 

CalculateWindowSizes()
{
	global
	If A_ScreenWidth>1270
	{
	CarisRocketWindowX := 415	
	CarisRocketWindowY := 2
	CarisRocketWindowW = 400
	CarisRocketWindowH = 260
	WinSURGEModalWindowX := 415
	WinSURGEModalWindowY := 313
	WinSURGEModalWindowW := 500
	WinSURGEModalWindowH := 500
	}
	Else if A_ScreenWidth > 1024
	{
	CarisRocketWindowX := 415	
	CarisRocketWindowY := 2
	CarisRocketWindowW = 400
	CarisRocketWindowH = 260
	WinSURGEModalWindowX := 415
	WinSURGEModalWindowY := 313
	WinSURGEModalWindowW := 500
	WinSURGEModalWindowH := 500
	}

	Return
	}
	
ParseQAQCData()
{
	global
	msg = 
	
OrderedCPTCodes:   ;OrderedCPTCodeX  OrderedCPTCount
{
OrderedCPTCount = 0
Loop, Parse, OrderedCPTCodes, CSV	
			{
			OrderedCPTCode%A_Index% = %A_LoopField%	
			OrderedCPTCount := A_Index
			}
		msg = %msg%OrderedCPTCount = %OrderedCPTCount%, %OrderedCPTCode1%, %OrderedCPTCode2%`n
}

GrossDescription:   ;For up to ten vials, grosstext_X, GrossVialCount
{

	StringCaseSense, On
	PositionNotFound := 0
	GrossVialCount := 0
	pos_0 := 0	

	Loop, 26
		{
		grosstext_%A_Index% = 	
		pos_%A_Index% =
		ascii_code := Chr(64+A_Index)
			
		StringGetPos, pos_%A_Index%, grossdescriptiontext, %ascii_code%.%A_Space%		
		If (Errorlevel AND A_Index=1)
			{
			StringCaseSense, Off
			StringGetPos, pos_1, grossdescriptiontext, received
			if Errorlevel
				msg = %msg%The gross description could contain an error because no information for "A. " could be found. `n	
			Else
				{
				GrossVialCount := 1	
				grosstext_1 = %grossdescriptiontext%
				Break
				}
			}

	GrossVialCount := A_Index - 1
	if pos_%A_Index% > 0
		l := pos_%A_Index% - pos_%GrossVialCount% - 1
	Else
		l := StrLen(grossdescriptiontext)
		
	StringMid, grosstext_%GrossVialCount%, grossdescriptiontext, pos_%GrossVialCount%, %l%
	If ErrorLevel
		Break
		}
	StringCaseSense, Off
}

FinalDiagnosis:    ;finaltext_x, biopsytype_X, FinalVialCount
{
		msg=%msg%GrossVialCount=%GrossVialCount%|| %grosstext_1%|| %grosstext_2%`n
		Needle := "%%P%%"
		i := 0
		h := 0
		k := 1

		Loop, 26
		{
			biopsytype_%A_Index% =
			finaltext_%A_Index% = 
		}
		
		Loop, 26
		{
			StringGetPos, i, finaldiagnosistext, %Needle%, , %k%
			if i>0
				{
				j := i - k
				h := h + 1
				StringMid, finaltext_%A_Index%, finaldiagnosistext, %k%+1, %j%
				k := i + 6
				StringGetPOs, y, finaltext_%A_Index%, shave
					if y>0
						biopsytype_%A_Index% = shave
				StringGetPOs, y, finaltext_%A_Index%, punch
					if y>0
						biopsytype_%A_Index% = punch
				StringGetPOs, y, finaltext_%A_Index%, excision
					if y>0
						biopsytype_%A_Index% = excision
				StringGetPOs, y, finaltext_%A_Index%, curettage
					if y>0
						biopsytype_%A_Index% = curettage
				StringReplace, finaltext_%A_Index%, finaltext_%A_Index%,*** ,,All
				}
			Else
				Break
		}
	
	FinalVialCount := h
	msg = %msg%FinalVialCount =%FinalVialCount%|| %finaltext_1%|| %finaltext_2%`n
/*
Loop, %FinalVialCount%
		{
			StringcaseSense, Off
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%R%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%R.%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Rt%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Rt.%A_space%, %A_Space%Right%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%L%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%L.%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Lt%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%Lt.%A_space%, %A_Space%Left%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%abd%A_space%, %A_Space%abdomen%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%lat%A_space%, %A_Space%lateral%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%med%A_space%, %A_Space%medial%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%prox%A_space%, %A_Space%proximal%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%dist%A_space%, %A_Space%distal%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx%A_space%, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx., %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx:, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%bx.:, %A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%post%A_space%, %A_Space%posterior%A_Space%
			StringReplace, finaltext_%A_Index%,finaltext_%A_Index%, %A_Space%ant%A_space%, %A_Space%anterior%A_Space%
		}
*/

		}

ClinicalData:
{
	StringCaseSense, On
	PositionNotFound := 0
	ClinDataVialCount := 0
	pos_0 := 0	

	Loop, 26
		{
		clindata_%A_Index% = 	
		pos_%A_Index% =
		ascii_code := Chr(64+A_Index)
			
		StringGetPos, pos_%A_Index%, ClinicalData, %ascii_code%.%A_Space%		
		If (Errorlevel AND A_Index=1)
			{
				StringCaseSense, Off
				ClinDataVialCount := 0	
				clindata_1 = %ClinicalData%
				Break				
			}

	ClinDataVialCount := A_Index - 1
	if (pos_%A_Index% > 0 AND pos_%A_Index% > pos_%ClinDataVialCount%)
		l := pos_%A_Index% - pos_%ClinDataVialCount% - 1
	Else
		l := StrLen(ClinicalData)
		
	StringMid, clindata_%ClinDataVialCount%, ClinicalData, pos_%ClinDataVialCount%, %l%
	If ErrorLevel
		Break
	}
	StringCaseSense, Off
	If ClinDataVialCount =0
		ClinDataVialCount = 1
	msg = %msg%%ClinDataVialCount%||%clindata_1%||%clindata_2%`n	
}
		
Return
}

ParseandCheckDXCode()
{
	global
				SelectedCodeIndex := 0
				LV_Modify(0, "-Select") 
				
				Stringlen, ltot, DxCode
				StringGetPos, j, DxCode,.
				StringGetPos, k, DxCode,;
				StringGetPos, i, DxCode,:
					
				if k=-1
					comhelp = 
				Else
					{
						x := ltot - k - 1
						StringRight, comhelp, DxCode, x
					}
				
				Stringlen, comlength, comhelp
				If (comlength>0)
					comlength := comlength + 1
					
				if j=-1
					backtophelp = 
				Else
					{
					if k=-1
						x :=  ltot - j -1
					Else
						x := k - j - 1
						
					xstart := j + 2
					StringMid, backtophelp, DxCode, xstart, x
					}

				Stringlen, backlength, backtophelp
				If (backlength>0)
					backlength := backlength + 1

				if i=-1
					{
					fronttophelp =	
					baselength := ltot - comlength - backlength
					StringLeft, basediag, DxCode, baselength
					}					
				Else
					{
					StringLeft, fronttophelp, DxCode, i	
					Stringlen, frontlength, fronttophelp
					frontlength := frontlength + 1
					xstart := i + 2
					x := ltot - comlength - backlength - frontlength
					StringMid, basediag, DxCode, xstart, x
					}
									
				Loop % LV_GetCount()
					{
					LV_GetText(RetrievedText, A_Index, 2)
					if basediag=%RetrievedText%
						{
						LV_Modify(A_Index, "Select")  ; Select each row whose first field contains the filter-text.
						SelectedCodeIndex = %A_Index%
						Break
						}
					}

				If (SelectedCodeIndex = 0 OR ltot=0)
					{
					Msgbox, That is not a valid diagnosis code!
					return, 1
					}
				
				
				LV_GetText(SelDXCode,SelectedCodeIndex,2)  
				LV_GetText(SelDiagnosis,SelectedCodeIndex,5)  
				LV_GetText(SelComment,SelectedCodeIndex,6)
				LV_GetText(SelMicro,SelectedCodeIndex,7)
				LV_GetText(SelCPTCode,SelectedCodeIndex,8)
				LV_GetText(SelICD9,SelectedCodeIndex,9)
				LV_GetText(SelICD10,SelectedCodeIndex,10)
				LV_GetText(SelSnomed,SelectedCodeIndex,11)
				LV_GetText(SelPre,SelectedCodeIndex,12)
				LV_GetText(SelMal,SelectedCodeIndex,13)
				LV_GetText(SelDys,SelectedCodeIndex,14)
				LV_GetText(SelMel,SelectedCodeIndex,15)				
				LV_GetText(SelInf,SelectedCodeIndex,16)
				LV_GetText(SelMargInc,SelectedCodeIndex,17)
				LV_GetText(SelLog,SelectedCodeIndex,18)
				
/*
Stringlen, x, backtophelp
				If (x > 1)    ;MULTIPLE MARGIN CODES WERE ENTERED
					{
					Msgbox, You may only enter one margin code!
					return, 1
					}
				else if (x=1)  ;MARGIN CODE WAS ENTERED
					{
						if SelInf
						{
							Msgbox, 4, ,The diagnosis code you selected is "inflammatory" and you gave a margin.  Are you sure you wish to continue?
							IfMsgbox No
								Return 1
						}
						else if SelMal
						{
							if (MarginNoPreference OR MarginMalignant OR MarginAll)
								Return 0
							Else
							{
								Msgbox, 4, , The client has not requested to have margins on 
							}
						}
						else if SelDys
						{
						}
						else if SelMel
						{
						}
						else if SelPre
						{
						}
							
					}
				else if (x=0)   ;MARGIN CODE WAS NOT ENTERED
					{
						;Margin Checking to ensure lack of margin information is ok goes here.
					}
*/
return 0
}

ZeroMarginFlags()
{
	global
	MarginAll = 0
	MarginOnRequestOnly = 0
	MarginNoPreference = 0
	MarginExcision = 0
	MarginMelanocytic = 0
	MarginDysplastic = 0
	MarginMalignant = 0
	MarginPremalignant = 0
	MarginShave = 0
	Return
}

ReadDXCodes()
{
	global
	LV_Delete()
	Loop, read, %A_ScriptDir%\dxcodes.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
	
	LV_Add("",DXCode1,DXCode2,DXCode3,DXCode4,DXCode5,DXCode6,DXCode7,DXCode8,DXCode9,DXCode10,DXCode11,DXCode12,DXCode13,DXCode14,DXCode15,DXCode16,DXCode17,DXCode18)
	}
	LV_ModifyCol()  ; Auto-size each column to fit its contents. 
	LV_Modify(1, "Sort")

}

ReadHelpers()
{

	global

;Front Helper Load	
	Loop, read, %A_ScriptDir%\fronthelpers.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		FrontofDiagnosisHelper%DxCode2% = %DxCode3%
	}

;Margin Load	
	Loop, read, %A_ScriptDir%\margins.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		BackofDiagnosisHelper%DxCode2% = %DxCode3%
	}

;Comment Helper Load	
	Loop, read, %A_ScriptDir%\commenthelpers.csv
	{
		Loop, parse, A_LoopReadLine, CSV
			DxCode%A_Index%=%A_LoopField%
		
		CommentHelper%DxCode2% = %DxCode3%
	}

}

FirstTimeSetup()
{
	global
	IfWinNotExist, WinSURGE
		{
			Msgbox, WinSURGE must be running the first time you fire the Caris CodeRocket.  Please login to WinSURGE and restart the Caris CodeRocket.
			ExitApp
		}
		
	perc := "%"
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)

	if !Result_1
	{
		s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.ttwentyfive03 LIKE '%y%%perc%'
	CodeDataBaseQuery(s)
		if !Result_1
			{
				OpMode = T
				IniWrite, %A_Username%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
				IniWrite, %OpMode%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, OpMode
				Return
			}
			
	}
	UseSendMethod := 1
	WinSurgeLoginID = %y%
	WinSurgeSignoutPassword = %y%SO
	WinSurgeLoginPassword = %y%PW
	WinSurgeFullName = %Result_1%
	SecurityUserId = %Result_4%
	WinSURGEPathologistID = %Result_5%
	CarisRocketWindowX := 0
	CarisRocketWindowY := 0
	CarisRocketWindowW := 0
	CarisRocketWindowH := 0
	
	
	OpMode = M
	
		IfInString, A_ComputerName, DFW
			{
			HomeLabCasePrefix =D
			HomeLabGrossAddendum = This case was interpreted at Miraca Life Sciences, 6655 North MacArthur Blvd, Irving, Texas 75039.  CLIA #45D0975010.
		}
		Else IfInString, A_ComputerName, PHX
			{
			HomeLabCasePrefix =P
			HomeLabGrossAddendum = This case was interpreted at Miraca Life Sciences. 4207 Cotton Ctr. Blvd., Bldg. 10, Phoenix, AZ 85040-8893.  CLIA#03D1064744
		}
		Else IfInString, A_ComputerName, BOS
			{
			HomeLabCasePrefix =C
			HomeLabGrossAddendum = This case was interpreted at Miraca Life Sciences, 320 Needham Street, Suite 200, Newton, MA  02464.  CLIA #22D957540.
		}
		Else IfInString, A_ComputerName, MAR
			{
			HomeLabCasePrefix =M
			HomeLabGrossAddendum = This case was interpreted at Miraca Life Sciences, 810 Landmark Drive, Suite 217-219, Glen Burnie, MD  21061.  Pathologists: 410.766.4175. Client Services: 617.969.4100  CLIA #21D1077515.
		}
		Else
			{
				Msgbox, Your location could not be determined.  Please email Matt Muenster with the following information:`nA_UserName=%A_UserName%`nA_ComputerName=%A_ComputerName%`nA_IPAddress1=%A_IPAddress1%`n`nSetup will continue...
			}
	IniWrite, %A_Username%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
	IniWrite, %WinSurgeSignoutPassword%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
	IniWrite, %WinSurgeLoginPassword%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeLoginPassword
	IniWrite, %WinSurgeFullName%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniWrite, %OpMode%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, OpMode
	IniWrite, %SecurityUserId%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, SecurityUserId
	IniWrite, %WinSURGEPathologistID%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniWrite, %HomeLabCasePrefix%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, HomeLabCasePrefix
	IniWrite, %HomeLabGrossAddendum%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, HomeLabGrossAddendum
	IniWrite, %CarisRocketWindowX%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX
	IniWrite, %CarisRocketWindowY%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY
	IniWrite, %CarisRocketWindowW%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW
	IniWrite, %CarisRocketWindowH%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH

	Msgbox, 4, Setup..., Caris CodeRocket is installed but is not setup for "automated" modes on your machine.  You must go through setup to use the automated modes or you can skip setup and use in manual mode.  Continue to setup for automation? 
	ifMsgbox, No
		{
			OpMode=M
			IniWrite, %OpMode%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, OpMode			
			Return
		}
		
			

	OpenCase("CD11-1000")
	Sleep, 3000

qiblabel:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Queue Into Batch" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			QueueIntoBatchBox := x	
			IniWrite, %QueueIntoBatchBox%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, QueueIntoBatchBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Queue Into Batch" box and press F3.
			Goto, qiblabel
		}
		
p1label:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Image01:" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinWait, WinSURGE WinsIMAGE
	WinClose, WinSURGE WinsIMAGE
	WinWaitClose, WinSURGE WinsIMAGE
	Sleep, 500
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			Photo1TextBox := x	
			IniWrite, %Photo1TextBox%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, Photo1TextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Image01" box and press F3.
			Goto, p1label
		}

p2label:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Image02:" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinWait, WinSURGE WinsIMAGE
	WinClose, WinSURGE WinsIMAGE
	WinWaitClose, WinSURGE WinsIMAGE
	Sleep, 500
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			Photo2TextBox := x	
			IniWrite, %Photo2TextBox%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, Photo2TextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Image02" box and press F3.
			Goto, p2label
		}

pathlabel:
	SplashTextOn, 200,100, Setup..., Click your mouse into the "Pathologist" box and press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	WinGetTitle, active_title, A
	ControlGetFocus, x,	%active_title%
	StringGetPos, y, x, TextBox
	if (y>0)
		{
			PathologistTextBox := x	
			IniWrite, %PathologistTextBox%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, PathologistTextBox
		}
	Else
		{
			Msgbox, You did not click into a text box.  Please click your mouse into the "Pathologist" box and press F3.
			Goto, pathlabel
		}

finaldxlabel:
	SplashTextOn, 200,100, Setup..., Hover your mouse over the "Final Diagnosis" button BUT DONT CLICK IT!  Then press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	MouseGetPos, xpos, ypos	
	FinalDiagnosisButtonX := xpos
	FinalDiagnosisButtonY := ypos
	IniWrite, %FinalDiagnosisButtonX%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonX
	IniWrite, %FinalDiagnosisButtonY%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonY

grossdesclabel:
	SplashTextOn, 200,100, Setup..., Hover your mouse over the "Gross Description" button BUT DONT CLICK IT!  Then press F3.
	KeyWait, F3, d
	KeyWait, F3, u
	MouseGetPos, xpos, ypos	
	GrossDescriptionButtonX := xpos
	GrossDescriptionButtonY := ypos
	IniWrite, %GrossDescriptionButtonX%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonX
	IniWrite, %GrossDescriptionButtonY%, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonY
	
	
	SplashTextOn, 200, 100, Setup Complete!
	Sleep, 1500
	SplashTextOff
Return
}

ReadIniValues()
{
	global

	IniRead, WindowsLoginId, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WindowsLoginId
	IniRead, OpMode, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, OpMode
	IniRead, WinSurgeSignoutPassword, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeSignoutPassword
	IniRead, WinSurgeLoginPassword, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeLoginPassword
	IniRead, WinSurgeFullName, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSurgeFullName
	IniRead, HomeLabCasePrefix, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, HomeLabCasePrefix
	IniRead, HomeLabGrossAddendum, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, HomeLabGrossAddendum
	IniRead, WinSURGEPathologistID, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, WinSURGEPathologistID
	IniRead, SecurityUserId, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, SecurityUserId

	IniRead, QueueIntoBatchBox, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, QueueIntoBatchBox 	
	IniRead, Photo1TextBox, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, Photo1TextBox 	
	IniRead, Photo2TextBox, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, Photo2TextBox 	
	IniRead, PathologistTextBox, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, PathologistTextBox 	
	IniRead, FinalDiagnosisButtonX, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonX
	IniRead, FinalDiagnosisButtonY, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, FinalDiagnosisButtonY 	
	IniRead, GrossDescriptionButtonX, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonX 	
	IniRead, GrossDescriptionButtonY, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, GrossDescriptionButtonY 	
	IniRead, DictationSendX, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, DictationSendX 	
	IniRead, DictationSendY, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, DictationSendY 	

	IniRead, CarisRocketWindowX, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowX 	
	IniRead, CarisRocketWindowY, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowY 
	IniRead, CarisRocketWindowW, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowW 
	IniRead, CarisRocketWindowH, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, CarisRocketWindowH

	IniRead, EsignoutTextBox, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, EsignoutTextBox
	IniRead, ApprovalPasswordControl, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, ApprovalPasswordControl 

	IniRead, ExpressSignout, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, ExpressSignout
	IniRead, SpeakEnabled, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, SpeakEnabled
	IniRead, UseSendMethod, %A_ScriptDir%\CarisCodeRocket.ini, Window Positions, UseSendMethod

Return
}

WinSurgeSetup() 
{
	global
	IfWinNotExist, WinSURGE
	{
		Run, %WinSurgeFilePath%		
		WinWait, WinSURGE Environment Login, , 2
		If ErrorLevel
			{
				Msgbox, Error Starting WinSurge
				ExitApp
			}
		Else
			{
			IfWinNotActive, WinSURGE Environment Login, , WinActivate, WinSURGE Environment Login, 	
			WinWaitActive, WinSURGE Environment Login, 
			Send, %WinSurgeLoginPassword%{ENTER}
			WinWait, Login Message, 
			IfWinNotActive, Login Message, , WinActivate, Login Message, 
			WinWaitActive, Login Message, 
			Send, {ENTER}
			}
	}
	Else
		Msgbox, WinSurge is currently running on your machine.  Some WinSurge errors close the visible windows but do not close the program.  If you get this message but do not see the windows, hit Ctrl-Alt-Delete and then Task Manager and use it to close WinSurge when this occurs.  Then reopen WinSURGE.
	Return 0
}

OpenAutoAssign() 
{
	global
	Run, http://autoassign/autoassign2/default.php
	WinWait, Caris AutoAssign Main Menu
	IfWinNotActive, Caris AutoAssign Main Menu
		WinActivate, Caris AutoAssign Main Menu	
	WinWaitActive, Caris AutoAssign Main Menu
	Send, %AutoAssignUsername%{Tab}
	Sleep, 500
	Send, {Enter}
	Sleep, 500
	Run, http://autoassign/autoassign2/report_path_case_status.php
	Return
}

