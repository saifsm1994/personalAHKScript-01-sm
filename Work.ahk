;Work AHK Scripts


;Macropad
	f7 & 2::
		gosub lookupUrlIDinInternalServer
	return
	
	f7 & 3::
		gosub lookupUrlID
	return
	
	f7 & 4::
		gosub lookupUrlIDEdit
	return

	f7 & 5::
		gosub chainedcountwordsfreeTextWithoutLinks
	return

	f7 & 6::
		gosub chainedDiffcheckerTextWithoutLinks
	return
	
	f7 & 7::
		gosub chainedDraftableSaveWithoutLinks
	return
	
	f7 & 8::
		gosub chainedDraftableSaveWithoutLinksPDF
	return
	
	f8 & 9::
		gosub openMainPages
	return
	
	f8 & 2::
		msgbox, unassigned
	return
	
	f8 & 3::
		GetKeyState, state, Ctrl
		if(state = "D"){
			gosub makeLinkfromcellsp1
		}else{
			gosub makeLinkfromcellsp2
		}
	return
	
	f8 & 4::
		gosub googleEnterFormula
	return

	f8 & 5::
		gosub draftableComPdf
	return

	f8 & 6::
		gosub chainedDraftableSaveAlreadyOpen
	return
	
	f8 & 7::
		gosub chainedDiffcheckerTextAlreadyOpen
	return
	
	f8 & 8::
		gosub chainedcountwordsfreeTextAlreadyOpen
	return
	
	f9 & 9::
	send  ^!{Right}
	return
	
	f9 & 2::
	msgbox, unassigned
	return
	
	f9 & 3::
		send ^{PgUp}
	return
	
	f9 & 4::
		send ^{PgDn}
	return

	f9 & 5::
	clearFunctionOriginal:
		send ^1
		sleep 222
		send ^c
		send ^+8
		sleep 333
		gosub startOfCopyBuiltLinks
		sleep 100
		send {enter}
		sleep 100
		gosub middleOfCopyBuiltLinks
		sleep 100
		send {enter}
		sleep 1500
		gosub openNoteURL1
		sleep 1055
		gosub openUrlidURL1
		sleep 1055
		gosub findCopyInURLID1
		send ^2
		sleep 100
	;	gosub saveABandDraft
	return


	f9 & 6::
	clearFunctionPDF:
	^numpadClear::
		send ^1
		sleep 222
		send ^c
		send ^+8
		sleep 333
		gosub startOfCopyBuiltLinks
		sleep 100
		gosub changeLinksToPDF
		sleep 100
		send {enter}
		sleep 100
		gosub middleOfCopyBuiltLinks
		sleep 100
		gosub changeLinksToPDF
		sleep 100
		send {enter}
		sleep 1500
		gosub openNoteURL1
		sleep 1055
		gosub openUrlidURL1
		sleep 1055
		gosub findCopyInURLID1
		send ^2
		sleep 100
	;	gosub saveABandDraft
	return
	
	f9 & 7::
	clearFunctionText:
	+numpadClear::
		send ^1
		sleep 222
		send ^c
		send ^+8
		sleep 333
		gosub startOfCopyBuiltLinks
		sleep 100
		gosub changeLinksToText
		sleep 100
		send {enter}
		sleep 100
		gosub middleOfCopyBuiltLinks
		sleep 100
		gosub changeLinksToText
		sleep 100
		send {enter}
		sleep 1500
		gosub openNoteURL1
		sleep 1055
		gosub openUrlidURL1
		sleep 1055
		gosub findCopyInURLID1
		send ^2
		sleep 100
		gosub chainedDiffcheckerTextAlreadyOpen
	return
	
	f9 & 8::
		msgbox, unassigned
	return
	
	f10 & 9::
		msgbox, unassigned
	return
	
	f10 & 2::
		gosub Apdf
	return
	
	f10 & 3::
		gosub Bpdf
	return
	
	f10 & 4::
		send ^f
		sleep 100
		send ^v
	return

	f10 & 5::
		gosub open1Original
	return

	f10 & 6::
		gosub open2Original
	return
	
	f10 & 7::
		gosub open1PDF
	return
	
	f10 & 8::
		gosub open2PDF
	return
	
	f11 & 9::
		send {ctrl}
	return
	
	f11 & 2::
		gosub findURLPageInLocalServer
	return
	
	f11 & 3::
		gosub findIndexPageInLocalServer
	return
	
	f11 & 4::
		send ^t
		sleep 100
		send https://policyreporter.github.io/brother-bear-online/ExcelToolsNew.html{enter}
	return

	f11 & 5::
		gosub open1Text
	return

	f11 & 6::
		gosub open2Text
	return
	
	f11 & 7::
		gosub openNote
	return
	
	f11 & 8::
		gosub openNote2
	return
	
	f12 & 9::
	loop 5
	{
		gosub lookupUrlID
		sleep 100
		send ^{PgUp}
		sleep 200
		send {Down}
		sleep 100
	}
	return
	
	f12 & 2::
		msgbox, unassigned
	return
	
	f12 & 3::
		msgbox, unassigned
	return
	
	f12 & 4::
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub clearGoogle
		else
		If OutNameNoExt = excel
		 gosub clearExcel
		else
		Msgbox this only works on chrome or excel
	return

	f12 & 5::
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub greyGoogle
		else
		If OutNameNoExt = excel
		 gosub greyExcel
		else
		Msgbox this only works on chrome or excel
	return

	f12 & 6::
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub yellowGoogle
		else
		If OutNameNoExt = excel
		 gosub yellowExcel
		else
		Msgbox this only works on chrome or excel
	return
	
	f12 & 7::
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub greenGoogle
		else
		If OutNameNoExt = excel
		 gosub greenExcel
		else
		Msgbox this only works on chrome or excel
	return
	
	f12 & 8::
	red:
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub redGoogle
		else
		If OutNameNoExt = excel
		 gosub redExcel
		else
		Msgbox this only works on chrome or excel
	return
	
	
	
	
	
;Hotkeys

	$~Esc::	
	f7 & 9::
		gosub preventHotkeyStuck
		gosub resetNumber
		Send, {Esc}
		Reload
	Return

;;;corsair

	ScrollLock::
		gosub lookupUrlIDinInternalServer
	return
	
	+ScrollLock::
		gosub lookupUrlIDinInternalServerCard
	return


	pause::
		gosub lookupUrlID
	return

	+pause::
		gosub lookupUrlIDEdit
	return

	PrintScreen::
		gosub findIndexPageInLocalServer
	return

	+PrintScreen::
		gosub findURLPageInLocalServer
	return
	
	numpadHome::
		send ^{PgUp}
	return
	
	numpadUp::
		send ^{PgDn}
	return
	
	numpadPgUp::
		gosub chainedDiffcheckerTextAlreadyOpen
	return
	
	numpadClear::
		gosub clearFunctionOriginal
	return
	
	numpadLeft::
		loop 4
		{
			Send {WheelDown}
			Sleep 40
		}
	return
	
	numpadIns::
		gosub green
	return
	
	numpadDel::
		gosub yellow
	return
	
	
	
	

;;; RAZER RAZOR KEYBOARD SHORTCUTS

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




	f3 & 1::
		send ^t
		sleep 100
		send https://app.slack.com/client/T04S79KDL{enter}
		return
	return

	f3 & 2::
		send ^t
		sleep 100
		send https://m.wuxiaworld.co/Lord-of-the-Mysteries/{enter}
	return

	f3 & 3::
		send ^t
		sleep 100
		send youtube.com{enter}
	return

	f3 & 4::
		send ^t
		sleep 100
		send https://calendar.google.com/calendar/r{enter}
	return

	f3 & 5::
		send ^t
		sleep 100
		send mail.google.com{enter}
	return

	f3 & 6::
		send ^t
		sleep 100
		send http://192.168.1.19:8080{enter}
		sleep 100
		Send {Delete}{enter}
	return


	global test := Object() 
	test[1] := "1adsfa" 


	; RAZER NUMPAD HERE

	f3 & numpad0::

	return 

	f3 & numpaddel::
		
	return 



	;;HIGHLIGHT COLORS
		f3 & numpad1::
		f3 & numpadEnd::
		f3 & End::
		yellow:
			WinGet, OutputVar, ProcessName, A
			SplitPath, OutputVar,,,, OutNameNoExt
			If OutNameNoExt = chrome
			 gosub yellowGoogle
			else
			If OutNameNoExt = excel
			 gosub yellowExcel
			else
			Msgbox this only works on chrome or excel
		return 




		f3 & numpad2::
		f3 & numpadDown::
		f3 & Down::
		green:
			WinGet, OutputVar, ProcessName, A
			SplitPath, OutputVar,,,, OutNameNoExt
			If OutNameNoExt = chrome
			 gosub greenGoogle
			else
			If OutNameNoExt = excel
			 gosub greenExcel
			else
			Msgbox this only works on chrome or excel	
		return


		f3 & numpad3::
		f3 & numpadPgDn::
		f3 & PgDn::
			WinGet, OutputVar, ProcessName, A
			SplitPath, OutputVar,,,, OutNameNoExt
			If OutNameNoExt = chrome
			 gosub redGoogle
			else
			If OutNameNoExt = excel
			 gosub redExcel
			else
			Msgbox this only works on chrome or excel
		return


		f3 & numpad4::
		f3 & numpadLeft::
		f3 & Left::
		grey:
			WinGet, OutputVar, ProcessName, A
			SplitPath, OutputVar,,,, OutNameNoExt
			If OutNameNoExt = chrome
			 gosub greyGoogle
			else
			If OutNameNoExt = excel
			 gosub greyExcel
			else
			Msgbox this only works on chrome or excel
		return

		f3 & numpad5::
		f3 & numpadClear::
		clear:
		WinGet, OutputVar, ProcessName, A
		SplitPath, OutputVar,,,, OutNameNoExt
		If OutNameNoExt = chrome
		 gosub clearGoogle
		else
		If OutNameNoExt = excel
		 gosub clearExcel
		else
		Msgbox this only works on chrome or excel
		return


	;;OPEN LINKS TO PR
		f3 & numpad6::
		f3 & numpadRight::
		f3 & Right::
			gosub lookupUrlID
		return
		
		
		f3 & numpad7::
		f3 & Home::
		f3 & numpadHome::
			gosub lookupUrlIDinInternalServer
		return

		f3 & numpad8::
		f3 & numpadUp::
		f3 & Up::
			gosub copyBuiltLinksandOpenOriginal
		return
		

		f3 & numpad9::
		f3 & numpadPgUp::
		f3 & PgUp::
			gosub copyBuiltLinksandOpenPDF
		return
		



	;; Chrome function
	f3 & numLock::
		;Chrome class all other tabs
		gosub findURLPageInLocalServer
	return

	f3 & numpadDiv::
		gosub findIndexPageInLocalServer
	return




	f3 & numpadMult::
		gosub googleEnterFormula
	return

	f3 & numpadSub::
		gosub saveABandDraft
	return

	f3 & numpadAdd::
		gosub draftableComPdf
	return



	f3 & numpadEnter::
		gosub openOriginal2Fast
	return



	f3 & f6::
	Ins:
		gosub markdownTable
	return

	f3 & f9::
	del:
		gosub linebreak
	return 

	f3 & f7::
	home:
		gosub spaceNbsp
	return

	f3 & f10::
	end:
		gosub spaceNormal
	return

	f3 & f11::
	pgdn:
		gosub endTable
	return 

	f3 & f8::
	pgup:
		gosub startTable
	return


	f3 & f2::
	prntScr:
		gosub removeQuotesFromCell		
	return

	f3 & f5::
	pausebreak:
		gosub paste9
	return 

	f3 & f4::
	scrlk:
		gosub cut9
	return


	f3 & [::
	sqbrktleft:
		gosub openPdfFast
	return

	f3 & ]::
	sqbrktright:
		gosub openPdf2Fast
	return 


	f3 & f1::
	colon:
		gosub opentext
	return

	f3 & '::
	quote:
		gosub opentext2
	return 



	f3 & .::
	rightbrkt:
		gosub openOriginalFast
	return


	f3 & /::
	question:
		gosub openOriginal2Fast
	return 


	setTitleMatchMode,2













	;MAINKEYS



	f3 & l::
		GetKeyState, state, Ctrl
		if(state = "D"){
			send {Enter}
			sleep 100
			loop 1
			{
			send +{Home}
			send {Backspace}
			sleep 5
			}
			sendInput '=countif(c:c,c2)'
			sleep 100
			send {backspace}
			sleep 100
			send {Home}
			sleep 100
			send {Delete}
			send {Enter}
			}else{
				send {Enter}
				sleep 100
				loop 1
				{
				send +{Home}
				send {Backspace}
				sleep 5
				}
				sendInput '=join("","https://portal.policyreporter.com/policy/review/",A2)'
				sleep 100
				send {backspace}
				sleep 100
				send {Home}
				sleep 100
				send {Delete}
				send {Enter}
			}
	return


	f3 & m::
		gosub Apdf
	return

	makeLinkfromcells:
	f3 & p::
		GetKeyState, state, Ctrl
		if(state = "D"){
			gosub makeLinkfromcellsp1
		}else{
			gosub makeLinkfromcellsp2
		}
	return
	
	f3 & z::
		gosub heading1
	return
	
	f3 & x::
		gosub heading2
	return
	
	f3 & c::
		gosub heading3
	return
	
	f3 & v::
		gosub heading4
	return




;Functions
	chainedDraftableSaveWithLinks:
		gosub copyBuiltLinksandOpenOriginal
		sleep 2500
		Send ^{PgUp}
		gosub Apdf
		sleep 1450

		Send ^{PgDn}
		sleep 500

		Gosub Bpdf
		sleep 1510

		Gosub draftableComPdf
	Return
	
	chainedDraftableSaveWithLinksPDF:
		gosub copyBuiltLinksandOpenPDF
		sleep 2500
		Send ^{PgUp}
		gosub Apdf
		sleep 1450

		Send ^{PgDn}
		sleep 500

		Gosub Bpdf
		sleep 1510
		Gosub draftableComPdf
	Return
	
	chainedDraftableSaveWithoutLinks:
		gosub makeLinksandOpenOriginal
		sleep 2500
		Send ^{PgUp}
		gosub Apdf
		sleep 1450
		Send ^{PgDn}
		sleep 500
		Gosub Bpdf
		sleep 1510
		Gosub draftableComPdf
	Return
	
	chainedDraftableSaveWithoutLinksPDF:
		gosub makeLinksandOpenPDF
		sleep 2500
		Send ^{PgUp}
		gosub Apdf
		sleep 1450
		Send ^{PgDn}
		sleep 500
		Gosub Bpdf
		sleep 1510
		Gosub draftableComPdf
	Return
	
	chainedDraftableSaveAlreadyOpen:
		gosub Apdf
		sleep 1450
		Send ^{PgDn}
		sleep 500
		Gosub Bpdf
		sleep 1510
		Gosub draftableComPdf
	Return
	
	chainedDiffcheckerTextWithLinks:
		gosub copyBuiltLinksandOpenText
		send ^{PgUp}
		sleep 100
		gosub copyTabCurrentAndNext
		gosub openDiffcheckerCom
		gosub pasteInDiffcheckerCom
		gosub clickAcceptDiffcheckerCom
	return
	
	chainedDiffcheckerTextWithoutLinks:
		gosub makeLinksandOpenText
		send ^{PgUp}
		sleep 100
		gosub copyTabCurrentAndNext
		gosub openDiffcheckerCom
		gosub pasteInDiffcheckerCom
		gosub clickAcceptDiffcheckerCom
	return

	
	chainedDiffcheckerTextAlreadyOpen:
		gosub copyTabCurrentAndNext
		gosub openDiffcheckerCom
		gosub pasteInDiffcheckerCom
		gosub clickAcceptDiffcheckerCom
	return
	
	chainedcountwordsfreeTextWithLinks:
		gosub copyBuiltLinksandOpenText
		send ^{PgUp}
		sleep 100
		gosub copyTabCurrentAndNext
		gosub opencountwordsfree
		gosub pasteIncountwordsfree
		gosub clickAcceptcountwordsfree
	return
	
	chainedcountwordsfreeTextWithoutLinks:
		gosub makeLinksandOpenText
		send ^{PgUp}
		sleep 100
		gosub copyTabCurrentAndNext
		gosub opencountwordsfree
		gosub pasteIncountwordsfree
		gosub clickAcceptcountwordsfree
	return

	
	chainedcountwordsfreeTextAlreadyOpen:
		gosub copyTabCurrentAndNext
		gosub opencountwordsfree
		gosub pasteIncountwordsfree
		gosub clickAcceptcountwordsfree
	return
	
	openMainPages:
		send ^t
		sleep 100
		send https://app.slack.com/client/T04S79KDL{enter}
			send ^t
			sleep 100
			send https://calendar.google.com/calendar/r{enter}
		send ^t
		sleep 100
		send mail.google.com{enter}
	return


;Core Functionality


	greenGoogle:
		send !/
		sleep 50
		sendInput highlight green
		sleep 400
		send {enter}
	Return

	yellowGoogle:
		send !/
		sleep 50
		sendInput highlight yellow
		sleep 400
		send {enter}
	Return
	
	redGoogle:
		send !/
		sleep 150
		sendInput highlight red
		sleep 200
		send {Down}
		sleep 500
		send {enter}
	Return

	greyGoogle:
		send !/
		sleep 50
		sendInput highlight dark gray 2
		sleep 400
		send {enter}
	Return


	redExcel:
		send !h
		sleep 150
		send h
		sleep 150
		loop 6
		{
		send {Down}
		sleep 20
		}
		loop 1
		{
		send {Right}
		sleep 20
		}
		sleep 500
		send {enter}
	Return

	greenExcel:
		send !h
		sleep 150
		send h
		sleep 150
		loop 6
		{
		send {Down}
		sleep 20
		}
		loop 5
		{
		send {Right}
		sleep 20
		}
		sleep 500
		send {enter}
	Return

	yellowExcel:
		send !h
		sleep 150
		send h
		sleep 150
		loop 6
		{
		send {Down}
		sleep 20
		}
		loop 3
		{
		send {Right}
		sleep 20
		}
		sleep 500
		send {enter}
	Return

	greyExcel:
		send !h
		sleep 150
		send h
		sleep 150
		loop 2
		{
		send {Down}
		sleep 20
		}
			sleep 500
		send {enter}
	Return

	clearGoogle:
		send !/
		sleep 50
		sendInput highlight none
		sleep 100
		send {Down}
		sleep 400
		send {enter}
	return

	clearExcel:
		send !h
		sleep 150
		send h
		sleep 150
		loop 1
		{
		send n
		}
	return
	
	
	heading1:
		send !/
		sleep 50
		sendInput Heading 1
		sleep 400
		send {enter}
	Return
	
	heading2:
		send !/
		sleep 50
		sendInput Heading 2
		sleep 400
		send {enter}
	Return
	
	heading3:
		send !/
		sleep 50
		sendInput Heading 3
		sleep 400
		send {enter}
	Return
	
	heading4:
		send !/
		sleep 50
		sendInput Heading 4
		sleep 400
		send {enter}
	Return

	GoToDraftable:
	send ^t
	sleep 100
	sendInput www.draftable.com/compare{enter}
	return
	
	

	
	findIndexPageInLocalServer:
		Send {ctrl down}{c down}{ctrl up}{c up}
		sleep 300
		send ^t 
		sleep 100
		sleep 200
		Send ^l
		Sleep 300 
		SendInput http://192.168.1.19:4000/tables.html
		SendInput {Enter}
		sleep 1500
		Send ^r
		sleep 400
		loop 8
		{ 
		Send {Tab}
		sleep 100
		}
		loop 4
		{
		send {Down}
		sleep 8
		}
		Send {enter}
		sleep 100
		send {Tab}
		Send {ctrl down}{v down}{ctrl up}{v up}
		sleep 100
		loop 24 
		{
		send {Tab}
		sleep 5
		}
		Send, Index Pages
		sleep 100
		Send {Tab}
		sleep 100
		Send, Shown
		sleep 100
		Send {Enter}
	Return


	findURLPageInLocalServer:
		Send {ctrl down}{c down}{ctrl up}{c up}
		sleep 300
		send ^t 
		sleep 100
		sleep 200
		Send ^l
		Sleep 300 
		SendInput http://192.168.1.19:4000/tables.html
		SendInput {Enter}
		sleep 1500
		Send ^r
		sleep 400
		loop 8
		{ 
		Send {Tab}
		sleep 100
		}
		loop 6
		{
		send {Down}
		sleep 8
		}
		Send {enter}
		sleep 100
		send {Tab}
		Send {ctrl down}{v down}{ctrl up}{v up}
		sleep 100
		loop 24 
		{
		send {Tab}
		sleep 5
		}
		Send {Tab}
		sleep 100
		Send, Shown
		sleep 100
		Send {Enter}
	Return
	
	
	openFastPartial1:
			send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
	return 
	
	
	
	
	openOriginalFast:
		sleep 152
		Gosub copy1
		sleep 150
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 150
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		gosub openFastPartial1
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return

	openOriginal2Fast:
		sleep 112 
		Gosub copy1
		sleep 60
		Send {Right}
		sleep 60
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 55
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return
	


	openPdfFast:
		sleep 152
		Gosub copy1
		sleep 150
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 150
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		send ^t 
		sleep 100
		sleep 20
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return

	openPdf2Fast:
		sleep 112 
		Gosub copy1
		sleep 60
		Send {Right}
		sleep 60
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 20
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		Send {Left}
		send ^t 
		sleep 100
		sleep 20
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return
	
	copyTabCurrentAndNext:
		send ^a
		sleep 200
		gosub copy1
		sleep 100
		send ^{PgDn}
		send ^a
		sleep 200
		gosub copy2
		sleep 100
	return


	opentext:
		sleep 52 
		Gosub copy1
		sleep 20
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 52 
		Gosub copy4
		sleep 20
		sleep 52 
		Gosub copy3
		sleep 52 
		Gosub copy5
		Send {ctrl down}{c}{ctrl up}
		sleep 20
		Send {Left}
		Send {Left}
		send ^t 
		sleep 100
		sleep 20
		WinActivate, ahk_exe chrome.exe
		Send ^l
		sleep 52 
		Send, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 20
		Gosub paste1
		sleep 20
		Send, &updateid=
		sleep 52 
		Gosub paste2
		sleep 20
		Send, &type=text
		sleep 20
		SendInput {Enter}
	return

	opentext2:
		sleep 52 
		Gosub copy1
		sleep 20
		Send {Right}
		sleep 52 
		Gosub copy2
		sleep 52 
		Gosub copy4
		sleep 20
		Send {Right}
		sleep 20
		Gosub copy3
		sleep 52 
		Gosub copy5
		Send {ctrl down}{c}{ctrl up}
		sleep 20
		Send {Left}
		Send {Left}
		send ^t 
		sleep 100
		sleep 20
		WinActivate, ahk_exe chrome.exe
		Send ^l
		sleep 52 
		Send, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 20
		Gosub paste1
		sleep 20
		Send, &updateid=
		sleep 52 
		Gosub paste3
		sleep 20
		Send, &type=text
		sleep 20
		SendInput {Enter}
	return
	
	open1Original:
		send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return
	
	open2Original:
	send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste3
		sleep 50
		SendInput, &type=original
		sleep 50
		SendInput {Enter}
	return
	
	open1Text:
		send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=text
		sleep 50
		SendInput {Enter}
	return
	
	open2Text:
	send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste3
		sleep 50
		SendInput, &type=text
		sleep 50
		SendInput {Enter}
	return
	
	open1PDF:
		send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste2
		sleep 50
		SendInput, &type=PDF
		sleep 50
		SendInput {Enter}
	return
	
	open2PDF:
	send ^t 
		sleep 100
		sleep 55
		SendInput, https://portal.policyreporter.com/geturlupdate.php?urlid= 
		sleep 400
		Gosub paste1
		sleep 50
		SendInput, &updateid=
		sleep 50
		Gosub paste3
		sleep 50
		SendInput, &type=PDF
		sleep 50
		SendInput {Enter}
	return
	
	
	
	;save as A.pdf
	Apdf:

		SendInput ^s
		sleep 1250
		Send, a.pdf
		sleep 400
		Send, {Enter}
		Send, `t
		Send, {Enter}
		sleep 500
	Return

	;save as B.pdf
	Bpdf:

		SendInput ^s
		sleep 1000
		send ^a
		sleep 200
		Send, b
		sleep 50
		Send, .pdf
		sleep 400
		Send, {Enter}
		Send, `t
		Send, {Enter}
		sleep 500
	Return
	
	SaveAB:
		Gosub Apdf
		sleep 500
		Send ^{PgDn}
		sleep 500
		Gosub Bpdf
		sleep 500
	Return
	
	saveABandDraft:
		Gosub SaveAB
		gosub GoToDraftable
		gosub draftableComPdf
	Return
	
		makeLinkfromcellsp1:
		send {Enter}
		sleep 100

		send +{Home}
		sleep 100
		send {Backspace}
		sleep 100

		sendInput '=if(isnumber($C2),join("","https://portal.policyreporter.com/geturlupdate.php?urlid=",$A2,"&updateid=",C2,"&type=original"),"n/a")'

		sleep 100
		send {backspace}
		sleep 100
		send {Home}
		sleep 100
		send {Delete}
		sleep 100
		send {Enter}
	return
	
	makeLinkfromcellsp2:
		send {Enter}
		sleep 100
		send +{Home}
		send {Backspace}
		sleep 50
		sendInput '=if(isnumber($B2),join("","https://portal.policyreporter.com/geturlupdate.php?urlid=",$A2,"&updateid=",B2,"&type=original"),"n/a")'
		sleep 100
		send {backspace}
		sleep 100
		send {Home}
		sleep 100
		send {Delete}
		sleep 100
		send {Enter}
	return
	

;Website Navigation



	;diff Clicks - all clicks for diffchecker.com pdfs entry
	diffCheckerComPDF:
		MouseGetPos, StartX, StartY
		Click, 503, 611
		sleep 500
		Send a.pdf
		Send {Enter}
		sleep 1510
		Click, 1388, 611
		sleep 250
		Send b.pdf
		Send, {Enter}
		sleep 3500
		Click, 903, 988
		MouseMove, StartX, StartY
	return
	
	diffCheckerComText:
		gosub copyBuiltLinksandOpenText
		send ^t 
		sleep 100
		Send ^l
		Send, https://www.diffchecker.com/{Enter}

		loop 2	
		Send ^{PgUp}

		Send ^{a}
		sleep 550
		Gosub copy2
		sleep 450
		Send ^{PgDn}
		Send ^{a}
		sleep 550
		Gosub copy3
		sleep 450	

		Send ^{PgDn}
		sleep 75
		Gosub paste1
		Send {Enter}{Enter}
		sleep 75
		Gosub paste4
		Send {Enter}{Enter}
		sleep 75
		Gosub paste2
		sleep 450
		Click, 1327,350
		sleep 75
		Gosub paste1
		Send {Enter}{Enter}
		sleep 75
		Gosub paste5
		Send {Enter}{Enter}
		sleep 75
		Gosub paste3
		sleep 400

		click ,966, 400
		click ,966, 532
		click ,966, 632
		click ,1022, 550
		click ,894, 229
		sleep 400
	return

	draftableComPdf:
		gosub GoToDraftable
		sleep 2000
		loop 15
		{
		send ^{-}
		sleep 100
		}
		sleep 1000
		loop 9
		{
		send ^{+}
		sleep 100
		}
		MouseGetPos, StartX, StartY
		Click, 714, 591
		sleep 1510
		Sendinput a.pdf
		Send {Enter}
		sleep 1200
		Click, 1092, 596
		sleep 1250
		Sendinput b.pdf
		Send, {Enter}
		sleep 1510
		Click, 888, 788
		Click, 888, 701
		Click, 888, 900
		Click, 949, 648
		Click, 888, 949
		MouseMove, StartX, StartY
		send ^l
		gosub paste1
		send, -
		gosub paste2
		send, -
		gosub paste3
	return
	
	clickAcceptDiffcheckerCom:
		sleep 400
		click ,966, 400
		click ,966, 532
		click ,966, 632
		click ,894, 229
		sleep 400
	return
	
	pasteInDiffcheckerCom:
		sleep 2000
		Gosub paste1
		sleep 450
		Click, 1348,338
		sleep 75
		Gosub paste2
	return
	
	openDiffcheckerCom:
		sleep 100
		send ^t	
		sleep 100
		sleep 200
		send, www.diffchecker.com{enter}
	return
	
	clickAcceptcountwordsfree:
		sleep 400
		click , 1044, 798
		click , 1075, 999
		click , 1012, 636
		click , 1044, 700
		sleep 400
	return
	
	pasteIncountwordsfree:
		sleep 2000
		Click, 645,551
		Gosub paste1
		sleep 450
		Click, 1411,338
		sleep 75
		Gosub paste2
		sleep 450
	return
	
	opencountwordsfree:
		sleep 100
		send ^t	
		sleep 100
		sleep 200
		send, https://countwordsfree.com/comparetexts{enter}
		sleep 500
	return
	
	
	
	copyBuiltLinksandOpenOriginal:
		gosub startOfCopyBuiltLinks
		Send {Enter}
		gosub middleOfCopyBuiltLinks
		Send {Enter}
		sleep 2500		
	return
	
	copyBuiltLinksandOpenText:
		gosub startOfCopyBuiltLinks
		gosub backpaceOriginal
		SendInput, text
		sleep 100
		Send {Enter}
		gosub middleOfCopyBuiltLinks
		gosub backpaceOriginal
		SendInput, text
		sleep 100
		Send {Enter}
		sleep 2500
	return


	copyBuiltLinksandOpenPDF:
		gosub startOfCopyBuiltLinks
		gosub backpaceOriginal
		SendInput, pdf
		sleep 100
		Send {Enter}
		send ^t
		gosub middleOfCopyBuiltLinks
		gosub backpaceOriginal
		SendInput, pdf
		sleep 100
		Send {Enter}
		sleep 2500
	return
	
	startOfCopyBuiltLinks:
		loop 10
			{
			Send {Left}
			sleep 50
			}
		loop 4
			{
			Send {Right}
			sleep 50
			}
		gosub copy1
		sleep 100
		send {Right}
		sleep 100
		Gosub copy2
		sleep 100
		loop 8
			{
			Send {Left}
			}
		send ^t
		sleep 400
		gosub paste1
		sleep 300
	return
	
	middleOfCopyBuiltLinks:
		sleep 400
		send ^t
		sleep 500
		gosub paste2
		sleep 300
	return
	
	changeLinksToPDF:
		gosub backpaceOriginal
		SendInput, pdf
		sleep 100
	return
	
	changeLinksToText:
		gosub backpaceOriginal
		SendInput, text
		sleep 100
	return
	
	makeLinksandOpenText:
		gosub GoToCol1
		gosub copyIds
		gosub open1Text
		sleep 200
		gosub open2Text
	return
	
	backpaceOriginal:
		loop 8
			{
		send {Backspace}
		sleep 20
		}
	return

	makeLinksandOpenPDF:
		gosub GoToCol1
		gosub copyIds
		gosub open1PDF
		sleep 200
		gosub open2PDF
	return

	makeLinksandOpenOriginal:
		gosub GoToCol1
		gosub copyIds
		gosub open1Original
		sleep 200
		gosub open2Original
	return


	;markdown
	linebreak:
		sendinput <br>
	return

	spaceNbsp:
		sendinput &nbsp;&nbsp;&nbsp;&nbsp;
	return

	spaceNormal:
		loop 4
		sendinput {space}
	return
	
	markdownTable:
		sendInput | Coverage Criteria | Patient Result | Page # |  {enter}
		sendInput  |----------|----------------|----|  
	return


	endTable:
		send {end}
		sleep 50
		send |||
	return

	startTable:
		send {Del}
		send {Del}
		sleep 50
		send |
	return

	;PMP functions

	goLeftMostCell:
		loop 20
		{
			send {Left}
			Sleep 40
		}
		sleep 100
	return

	GoToCol5:
		loop 4
		{
			send {right}
			Sleep 40
		}
		sleep 100
	return

	GoToCol10:
		loop 9
		{
			send {right}
			Sleep 40
		}
		sleep 200
	return
	
	GoToCol1:
		loop 20
		{
			send {left}
			Sleep 40
		}
		sleep 200
	return
	
	copyIds:
		gosub copy1
		sleep 200
		send {right}
		gosub copy2
		sleep 200
		send {right}
		gosub copy3
		sleep 200
	return
	
	
	resetNumber:
	if not number
			number = 0.0
		number += 1.0
		number = 0.0	
	return
		

	cut1:
		
		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_1 := clipboardall
		sleep 151 
	 Return
	cut2:
		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_2 := clipboardall
		sleep 151 
	 Return
	cut3:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_3 := clipboardall
		sleep 151 
	 Return
	cut4:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_4 := clipboardall
		sleep 151 
	 Return
	cut5:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_5 := clipboardall
		sleep 151 
	 Return
	cut6:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_6 := clipboardall
		sleep 151 
	 Return
	 
	cut7:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_7 := clipboardall
		sleep 151 
	 Return
	cut8:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_8 := clipboardall
		sleep 151 
	 Return
	cut9:

		Clipboard := ""
		Send ^x
		Clipwait, 2
		Clipboard_9 := clipboardall
		sleep 151 
	 Return
	cut0:

		Clipboard := ""
		Send ^x
		Clipwait, 2, 2
		Clipboard_0 := clipboardall
		sleep 151 
	 Return
	copy1:

		Clipboard := ""
		sleep 50
		Send ^c
		sleep 50
		Clipwait, 2
		Clipboard_1 := clipboardall
		Clipboard_1Var := clipboard
		sleep 151 
	 Return
	copy2:

		Clipboard := ""
		sleep 50
		Send ^c
		sleep 50
		Clipwait, 2
		Clipboard_2 := clipboardall
		Clipboard_2Var := clipboard
		sleep 151 
	 Return
	copy3:

		Clipboard := ""
		sleep 50
		Send ^c
		sleep 50
		Clipwait, 2
		Clipboard_3 := clipboardall
		Clipboard_3Var := clipboard
		sleep 151 
	 Return
	copy4:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_4 := clipboardall
		Clipboard_4Var := clipboard
		sleep 151 
	 Return
	copy5:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_5 := clipboardall
		Clipboard_5Var := clipboard
		sleep 151 
	 Return
	copy6:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_6 := clipboardall
		Clipboard_6Var := clipboard
		sleep 151 
	 Return
	copy7:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_7 := clipboardall
		Clipboard_7Var := clipboard
		sleep 151 
	 Return
	copy8:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_8 := clipboardall
		Clipboard_8Var := clipboard
		sleep 151 
	 Return
	copy9:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_9 := clipboardall
		Clipboard_9Var := clipboard
		sleep 151 
	 Return
	copy0:

		Clipboard := ""
		Send ^c
		Clipwait, 2
		Clipboard_0 := clipboardall
		Clipboard_0Var := clipboard
		sleep 151 
	 Return
		 
	paste1:
		sleep 75

		Clipboard := clipboard_1
		sleep 400 
		SendInput {Ctrl Down}v{Ctrl Up}
		Sleep 450
		sleep 151 
	 Return
	 
	paste2:
		sleep 75

		Clipboard := clipboard_2
		sleep 400 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste3:
		sleep 75

		Clipboard := clipboard_3
		sleep 200 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste4:
		sleep 75

		Clipboard := clipboard_4
		sleep 200 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste5:
		sleep 75

		Clipboard := clipboard_5
		sleep 200 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste6:
		sleep 75

		Clipboard := clipboard_6
		sleep 200 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste7:
		sleep 75

		Clipboard := clipboard_7
		sleep 150 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste8:
		sleep 75

		Clipboard := clipboard_8
		sleep 150 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste9:
		sleep 75

		Clipboard := clipboard_9
		sleep 150 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	paste0:
		sleep 75

		Clipboard := clipboard_0
		sleep 150 
		SendInput {Ctrl Down}v{Ctrl Up}
		sleep 150
		sleep 151 
	 Return
	 
	 preventHotkeyStuck:
		send {ctrl up}
		send {alt up}
		send {shift up}
	return
	
		
	lookupUrlID:
		sleep 50
		Send {ctrl down}{c down}{ctrl up}{c up}
		sleep 200
		send ^t 
		sleep 100
		sleep 75
		WinActivate, ahk_exe chrome.exe
		Send ^l
		sleep 75
		SendInput https://portal.policyreporter.com/policy/
		Send {ctrl down}{v down}{ctrl up}{v up}
		Send /review
		SendInput {Enter}
		Send {PgDn}
	return
	
	lookupUrlIDEdit:
		sleep 50
		Send {ctrl down}{c down}{ctrl up}{c up}
		sleep 200
		send ^t 
		sleep 100
		sleep 75
		WinActivate, ahk_exe chrome.exe
		Send ^l
		sleep 75
		SendInput https://portal.policyreporter.com/policy/
		Send {ctrl down}{v down}{ctrl up}{v up}
		Send ?action=edit
		SendInput {Enter}
		Send {PgDn}
	return




	lookupUrlIDinInternalServer:
		send ^t 
		sleep 300
		SendInput http://192.168.1.19:4000/LookupTable.html 
		SendInput {Enter}
		sleep 500
		send ^r
		sleep 1500
		loop 8
		{ 
		Send {Tab}
		sleep 100
		}
		Send {ctrl down}{v down}{ctrl up}{v up}
		sleep 100
		SendInput {Enter}
	return
	
	lookupUrlIDinInternalServerCard:
		send ^t 
		sleep 300
		SendInput http://192.168.1.19:4000/lookup.html 
		SendInput {Enter}
		sleep 500
		send ^r
		sleep 1500
		loop 8
		{ 
		Send {Tab}
		sleep 100
		}
		Send {ctrl down}{v down}{ctrl up}{v up}
		sleep 100
		SendInput {Enter}
	return
	
	
	googleEnterFormula:
		click, left
		sleep 200
		Send {down}
		Send {down}
		sleep 200
		Send {tab}
		Send {down}
		Send {down}
		sleep 200
		Send {tab}	
		sleep 100
		Send {Enter}
		sleep 200
		sleep 100
		send {tab}
		sleep 200
		Send {down}
		Send {up}
		Send {enter}
		send {Tab}
		send ^v
		sleep 200
		loop 6
		{
		send {Tab}
		sleep 50
		}
		send {enter}
	return

	
	removeQuotesFromCell:
		loop 20
			{
			send {PgUp}
			sleep 10
			send {Left}
			send {Home}
			}
		send {delete}
		loop 20
			{
			send {PgDn}
			sleep 10
			}
		send {backspace}
	return 
	
	openNoteURL1:
	sleep 222
	send ^1
	gosub openNote2
	sleep 112
	send ^1
	sleep 444
	loop 10
	{
		sendinput {Left}
		sleep 30
	}
return

openUrlidURL1:
	sleep 222
	loop 6
	{
	Sendinput {Right}
	sleep 100
	}
	sleep 333
	send ^c
	sleep 222
	send ^t
	sleep 222
	send ^v
	sleep 222
	send {enter}
return

findCopyInURLID1:
	send ^1
	sleep 200
	loop 55
	{
		Sendinput {Left}
		sleep 12
	}
	sleep 111
	loop 2
	{
	Sendinput {Right}
	sleep 333
	}
	sleep 333
	send ^c
	sleep 200
	loop 55
	{
		Sendinput {Left}
		sleep 30
	}
	send ^5
	sleep 2000
	send ^f
	sleep 200
	loop 30
	{
	send {backspace}
	sleep 11	
	}
	sleep 100
	send ^v
return


openNote:
	gosub copy1
	loop 1
	{
	sendinput {right}
	sleep 100
	gosub copy2
	sleep 50
	}
	sleep 100
	send ^t
	sleep 100
	sendInput https://portal.policyreporter.com/geturlupdate.php?urlid=
	sleep 100
	gosub paste1
	sleep 100
	sendInput &updateid=
	sleep 100
	gosub paste2
	sleep 100
	sendInput &type=notes
	sleep 100
	Send {Enter}
	return
return 

openNote2:
	gosub copy1
	loop 2
	{
	sendinput {right}
	sleep 100
	gosub copy2
	sleep 50
	}
	sleep 100
	send ^t
	sleep 100
	sendInput https://portal.policyreporter.com/geturlupdate.php?urlid=
	sleep 100
	gosub paste1
	sleep 100
	sendInput &updateid=
	sleep 100
	gosub paste2
	sleep 100
	sendInput &type=notes
	sleep 100
	Send {Enter}
	return
return


	^f1::
		send {f1}
	return
	
	^f2::
		send {f2}
	return
	^f3::
		send {f3}
	return
	^f4::
		send {f4}
	return
	^f5::
		send {f5}
	return
	^f6::
		send {f6}
	return
	^f7::
		send {f7}
	return
	^f8::
		send {f8}
	return
	^f9::
		send {f9}
	return
	^f10::
		send {f10}
	return
	^f11::
		send {f11}
	return
	^f12::
		send {f12}
	return
	
	
	
	
	

;;;;;;play



	f3 & n::
	loop 10
	{

		
		click, right
		sleep 10
		loop 2
		{
		send {Up}
		sleep 30
		}
		loop 2
		{
		send {Enter}
		sleep 30
		}
		loop 30
		{ 
		send {PgDn}
		sleep 5
		}
		
		loop 30
		{ 
		send {Down}
		sleep 5
		}
	}	
	return
	
	sendToBalabolka:
		IfWinExist  ahk_exe balabolka.exe
		winactivate ahk_exe balabolka.exe
	else
		run, "C:\Program Files (x86)\Balabolka\balabolka.exe"
		WinWaitActive ahk_exe balabolka.exe
	
		sleep 100
		send {f7}
		sleep 100
		send ^n
		sleep 100
		send ^v
		sleep 100
		send {f5}
		sleep 100
		send !{tab}
	return    
	
	f13::
		gosub sendToBalabolka
	return
	