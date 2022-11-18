' vibreoffice - Vi Mode for LibreOffice/OpenOffice
'
' The MIT License (MIT)
'
' Copyright (c) 2014 Sean Yeh
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

' 2020 Yamsu 
' Added Support for Librecalc
'
' Added APP() :: Able to check if running Calc, Writer, or Other (Impress, etc.)
' Added simulate_KeyPress_Char() :: Allow generation of key press/release to perform desired motions and operations.
' Added simulate_KeyPress() :: Actually carries out key generation
' Added KEYS() :: Provides key infomation
' Added MODS() :: Provide key modifier information
' TODO Execute as url :: Execute link through web browser
' TODO Borders :: Generate border
' TODO Normal for Intracell movement ::
'
' Incorporated branch from fedorov-ao (Read On)
'
' 2018 axf 
' An attempt to revitalize a useful LibreOffice plugin
' Can be used, but is a work-in-progress so far

' Added Russian keyboard mapping (works in almost all cases)
' Fixed key release processing: releasing of non-char keys is now not consumed
' Fixed char search
' Fixed moving to end of word (e). Moves to the first space after the word.
' Added moving to end of previous word (E). Also moves to the first space after the word. 
' Added search repeating (n and N) 
' Mapped redo to U
' Status bar now displays full vibreoffice state for given frame
' Refactored key translation functions
' Added [c,d][a,i]<char> - will change or delete text fragment surrounded by <char>
' Changed multiplied yank processing: now selects the whole region based on multiplied movement, then yanks. Allows to paste the whole region instead of its last part only.
' Added / and ? searches. String used to search with these is treated as last search and can be used by n and N. 
' Added backspace emulation, called with X
' Replaced string mode id with integer. Should work faster, especially in input.
' Added selection active end swap in visual mode (o).
' Optimized key processing in insert mode.
' Added input logging, log is displayed and cleared by Alt+=
' Added the ability to display the character currently selected by cursor in normal mode
'  If char is printable, shows it as-is, if not - shows its <code>, if empty - shows <>
'  Called by Alt+-
' Added skipping over control and zero-length chars. Downside: cannot join lines easily. Have to go to beggining of next line and press "X"
' Fixed character swapping bug in insert mode by consuming all input and directly using cursor for output
' Fixed dd and cc not working on the last line
' Fixing initialization. Unfortunately, cannot add key handlers and init status bars of all
' currently open windows due to LibreOffice bug (also cannot deinit)
' Thus assuming that vibreoffice is initialized and deinitialized 
' when only one, currently active window is opened
' Added support for multiple windows. Apparently works, but needs cleanup. Vibreoffice is switched into normal mode
' when bringing window in focus
' Fixed s
' Fixed annotation zero-length "char" handling during hjkl movement
' Optimized selection ends swap
' Added polling-based reinitialization on changing windows
' Shift+ESC, as well as "toggle Vbreoffice" menu entry, toggles Vibreoffice

' BUG Annotations break repeated h/l movement
' BUG Anntoations break selection ends swap
' BUG b and e mishadle dashes and periods. Caused by incorrect behaviour of XTextCursor::gotoNext/PreviousWord()
' BUG Searching with f and F behaves incorrectly in visual mode
' TODO Add separate status bar for vibreoffice. Not as easy as it seems.
' TODO Split ProcessMovementKey() into several specialized functions (hjkl movement, word-based movement...).
' Pass number of iterations to these functions to handle repetitions internally.
' 

' Following option allows the use of ActiveCell, but doesn't work'
'Option VBASupport 1 

Option Explicit
' --------
' Globals
' --------
global VIBREOFFICE_STARTED as boolean ' Defaults to False
global VIBREOFFICE_ENABLED as boolean ' Defaults to False

global oXKeyHandler as object
global oListener as object
global oCurrentFrame as object

' Global State (Value are preserved even when the marco is not running)
const M_NORMAL = 0
const M_INSERT = 1
const M_VISUAL = 2
const M_VISUAL_LINE = 3
const M_DISABLED = 254
const M_BAD = 255
global MODE as integer
global OLD_MODE as integer

global VIEW_CURSOR as object
global TEXT_CURSOR as object
global MULTIPLIER as integer
global VISUAL_BASE as object ' Position of line that is first selected when 
                             ' VISUAL_LINE mode is entered
'global ACTIVE_SHEET as object
'global ACTIVE_CELL as object
'global APP as string
global LAST_SEARCH as string

global logged2 as string

' -----------
' Key Generation for Calc 
' -----------
Public Function MODS (key as String)
	Select Case key
		Case "SHIFT":
			MODS = 1
		Case "CTRL": 
			'CMD for Mac OS'
			MODS = 2
		Case "ALT":
			MODS = 4
		Case "MACCTRL":
			'CTRL for Mac OS'
			MODS = 8
	End Select
End Function

Public Function KEYS( key as String, Optional modifier as Integer )
Select Case key
	Case "ESCAPE":
		KEYS = Array(9,	  com.sun.star.awt.Key.ESCAPE,   0)
	Case "RETURN":
		KEYS = Array(13,  com.sun.star.awt.Key.RETURN,   0)
	Case "F":
		KEYS = Array(41,	  com.sun.star.awt.Key.F,   0)
	Case "F2":
		KEYS = Array(68,  com.sun.star.awt.Key.F2,       0)
	Case "DELETE":
		KEYS = Array(91,  com.sun.star.awt.Key.DELETE,   0)
	Case "HOME":
		KEYS = Array(110, com.sun.star.awt.Key.HOME,     0)
	Case "UP":
		KEYS = Array(111, com.sun.star.awt.Key.UP,       0)
	Case "PAGEUP":
		KEYS = Array(112, com.sun.star.awt.Key.PAGEUP,   0)
	Case "LEFT":
		KEYS = Array(113, com.sun.star.awt.Key.LEFT,     0)
	Case "RIGHT":
		KEYS = Array(114, com.sun.star.awt.Key.RIGHT,    0)
	Case "END":
		KEYS = Array(115, com.sun.star.awt.Key.END,      0)
	Case "DOWN":
		KEYS = Array(116, com.sun.star.awt.Key.DOWN,     0)
	Case "PAGEDOWN":
		KEYS = Array(117, com.sun.star.awt.Key.PAGEDOWN, 0)
End Select
If Not IsMissing(modifier) Then KEYS(2)=modifier
End Function

Sub simulate_KeyPress_Char( key as String, Optional modifier as String, Optional modifier2 as String, Optional modifier3 as String)
REM Simulate a RETURN Key press ( and -release ) in the current Window.
REM NB. This can cause the triggering of window elements.
    Dim oKeyEvent As New com.sun.star.awt.KeyEvent
	Dim KeyData(3) As Integer
	Dim finalModfiers As Integer
	finalModfiers = 0
	If Not IsMissing(modifier) Then finalModfiers = finalModfiers + MODS(modifier)
	If Not IsMissing(modifier2) Then finalModfiers = finalModfiers + MODS(modifier2)
	If Not IsMissing(modifier3) Then finalModfiers = finalModfiers + MODS(modifier3)
	KeyData = KEYS(key, finalModfiers)

    oKeyEvent.Modifiers = KeyData(2)     REM A combination of com.sun.star.awt.KeyModifier.
    oKeyEvent.KeyCode   = KeyData(1)               REM 1280.
    oKeyEvent.KeyChar   = chr( KeyData(0) )
    simulate_KeyPress( oKeyEvent )
End Sub

Sub simulate_KeyPress( oKeyEvent As com.sun.star.awt.KeyEvent )
REM Simulate a Key press ( and -release ) in the current Window.
REM NB. This can cause the triggering of window elements.
REM For example if there is a button currently selected in your form, and you call this method
REM while passing the KeyEvent for RETURN, then that button will be activated.
    If Not IsNull( oKeyEvent ) Then
		removeKeyHandlerFromModels()
        Dim oWindow As Object, oToolkit As Object
        oWindow = ThisComponent.CurrentController.Frame.getContainerWindow()
        oKeyEvent.Source = oWindow      
        oToolkit = oWindow.getToolkit()         REM com.sun.star.awt.Toolkit
        oToolkit.keyPress( oKeyEvent )          REM methods of XToolkitRobot.
		oToolkit.keyRelease( oKeyEvent )
		addKeyHandlerToModels()
    End If
End Sub

' -----------
' Singletons
' -----------
Sub setCursor
	VIEW_CURSOR = Nothing
	dim oCurrentController
	oCurrentController = getCurrentController()
	If oCurrentController is Nothing Then
		VIEW_CURSOR = Nothing
	Else
	If APP() <> "CALC" Then
	    VIEW_CURSOR = oCurrentController.getViewCursor()
	Else
		VIEW_CURSOR = Nothing
	End If
	End If
End Sub

Function getCursor
	If APP() <> "CALC" Then
		getCursor = VIEW_CURSOR
	Else
		getCursor = Nothing
	End If
End Function

Sub setTextCursor
	If APP() <> "CALC" Then
		On Error Goto ErrorHandler
		dim oCursor
		oCursor = getCursor()
		dim oText 
		oText = oCursor.getText()
		TEXT_CURSOR = oText.createTextCursorByRange(oCursor)
		Exit Sub
		
	ErrorHandler:
		' Text Cursor does not work in some instances, such as in Annotations
		TEXT_CURSOR = Nothing
	Else
		TEXT_CURSOR = Nothing
	End If
End Sub

Function getTextCursor
	setTextCursor() ' temp
    getTextCursor = TEXT_CURSOR
End Function

Function getCurrentController()
	On Error Goto ErrorHandler
	dim oComponent as object : oComponent = thisComponent
	getCurrentController = oComponent.getCurrentController()
	Exit Function
ErrorHandler:
	getCurrentController = Nothing
End Function	

private function printString(oCursor, s)
	dim l : l = len(s)
	oCursor.setString(s)
	oCursor.goRight(l, False)
end function

' -----------------
' Calc Related Function
' -----------------

Function getSheet()
	getSheet = ThisComponent.CurrentSelection.getSpreadsheet
End Function

Function insertRow(Optional opt As Integer)
	REM 1 Below 0 for Above
	Dim sheet as Object
	Dim myRangeRaw as Object
	Dim Rs, Re
	Dim optL as Integer
	if IsMissing(opt) Then
		optL = 1 
	Else
		optL = opt
	End If

	sheet = getSheet()
	myRangeRaw = ThisComponent.getCurrentSelection.getRangeAddress
	Rs = myRangeRaw.startRow
    Re = myRangeRaw.endRow
    If optL = 1 Then
    	REM Below need to move down one cell
		sheet.Rows.insertByIndex(Rs+1, 1)
	Else
		REM Above
		sheet.Rows.insertByIndex(Re, 1)
	End If
End Function


Function removeRow()
	REM 1 Below 0 for Above
	Dim sheet as Object
	Dim myRangeRaw as Object
	Dim Rs, Re


	sheet = getSheet()
	myRangeRaw = ThisComponent.getCurrentSelection.getRangeAddress
	Rs = myRangeRaw.startRow
    Re = myRangeRaw.endRow

	sheet.Rows.removeByIndex(Re, 1)

End Function
' -----------------
' Helper Functions
' -----------------
' Returns mode name
Function getModeName(m)
	dim sModeName as string
	Select Case m
		Case M_NORMAL:
			sModeName = "NORMAL"
		Case M_INSERT:
			sModeName = "INSERT"
		Case M_VISUAL:
			sModeName = "VISUAL"
		Case M_VISUAL_LINE:
			sModeName = "VISUAL_LINE"
		Case M_BAD:
			sModeName = "BAD"
		Case M_DISABLED:
			sModeName = "DISABLED"
		Case Else:
			sModeName = "BAD"								
	End Select
	getModeName = sModeName
End Function

' Returns key by non-zero code
Function getLatinKeyCharByCode(oEvent)
    dim keyChar
    keyChar = asc(0)
    If (oEvent.modifiers and 1) = 0 Then
        Select Case oEvent.keyCode
            case 1311: 
                keyChar = "`"
            case 257: 
                keyChar = "1"
            case 258: 
                keyChar = "2"
            case 259: 
                keyChar = "3"
            case 260: 
                keyChar = "4"
            case 261: 
                keyChar = "5"
            case 262:
                keyChar = "6"
            case 263:
                keyChar = "7"
            case 264:
                keyChar = "8"
            case 265:
                keyChar = "9"
            case 256:
                keyChar = "0"
            case 1288:
                keyChar = "-"
            case 1295:
                keyChar = "="
            case 0:
                keyChar = "\"
            case 528:
                keyChar = "q"
            case 534:
                keyChar = "w"
            case 516:
                keyChar = "e"
            case 529:
                keyChar = "r"
            case 531:
                keyChar = "t"
            case 536:
                keyChar = "y"
            case 532:
                keyChar = "u"
            case 520:
                keyChar = "i"
            case 526:
                keyChar = "o"
            case 527:
                keyChar = "p"
            case 1315:
                keyChar = "["
            case 1316:
                keyChar = "]"
            case 512:
                keyChar = "a"
            case 530:
                keyChar = "s"
            case 515:
                keyChar = "d"
            case 517:
                keyChar = "f"
            case 518:
                keyChar = "g"
            case 519:
                keyChar = "h"
            case 521:
                keyChar = "j"
            case 522:
                keyChar = "k"
            case 523:
                keyChar = "l"
            case 1317:
                keyChar = ";"
            case 1318:
                keyChar = "'"
            case 537:
                keyChar = "z"
            case 535:
                keyChar = "x"
            case 514:
                keyChar = "c"
            case 533:
                keyChar = "v"
            case 513:
                keyChar = "b"
            case 525:
                keyChar = "n"
            case 524:
                keyChar = "m"
            case 1292:
                keyChar = ","
            case 1291:
                keyChar = "."
            case 1290:
                keyChar = "/"
            case 1284:
                keyChar = " "
        End Select
    Else
        Select Case oEvent.keyCode
            case 1311:
                keyChar = "~"
            case 257:
                keyChar = "!"
            case 258:
                keyChar = "@"
            case 259:
                keyChar = "#"
            case 260:
                keyChar = "$"
            case 261:
                keyChar = "%"
            case 262:
                keyChar = "^"
            case 263:
                keyChar = "&"
            case 264:
                keyChar = "*"
            case 265:
                keyChar = "("
            case 256:
                keyChar = ")"
            case 1288:
                keyChar = "_"
            case 1295:
                keyChar = "+"
            case 528:
                keyChar = "Q"
            case 534:
                keyChar = "W"
            case 516:
                keyChar = "E"
            case 529:
                keyChar = "R"
            case 531:
                keyChar = "T"
            case 536:
                keyChar = "Y"
            case 532:
                keyChar = "U"
            case 520:
                keyChar = "I"
            case 526:
                keyChar = "O"
            case 527:
                keyChar = "P"
            case 1315:
                keyChar = "{"
            case 1316:
                keyChar = "}"
            case 512:
                keyChar = "A"
            case 530:
                keyChar = "S"
            case 515:
                keyChar = "D"
            case 517:
                keyChar = "F"
            case 518:
                keyChar = "G"
            case 519:
                keyChar = "H"
            case 521:
                keyChar = "J"
            case 522:
                keyChar = "K"
            case 523:
                keyChar = "L"
            case 1317:
                keyChar = ":"
            case 1318:
                keyChar = chr(34)
            case 537:
                keyChar = "Z"
            case 535:
                keyChar = "X"
            case 514:
                keyChar = "C"
            case 533:
                keyChar = "V"
            case 513:
                keyChar = "B"
            case 525:
                keyChar = "N"
            case 524:
                keyChar = "M"
            case 1292:
                keyChar = "<"
            case 1291:
                keyChar = ">"
            case 1290:
                keyChar = "?"
            case 1284:
                keyChar = " "
        End Select
    End If
    getLatinKeyCharByCode = keyChar
End Function

Function getLatinKeyCharByRus(oEvent)
    dim keyChar
    keyChar = asc(0)
    If (oEvent.modifiers and 1) = 0 Then
        Select Case oEvent.keyChar
            case "?":
                keyChar = "["
            case "?":
                keyChar = "]"
            case "?":
                keyChar = ";"    
            case "?":
                keyChar = "'"
            case "?":
                keyChar = ","
            case "?":
                keyChar = "."
            case ".":
                keyChar = "."                
        End Select
    Else
        Select Case oEvent.keyChar
            case "?":
                keyChar = "{"
            case "?":
                keyChar = "}"
            case "?":
                keyChar = ":"    
            case "?":
                keyChar = chr(34)
            case "?":
                keyChar = "<"
            case "?":
                keyChar = ">"
        End Select
    End If
    getLatinKeyCharByRus = keyChar    
End Function

Function getLatinKey(oEvent)
    dim keyChar
    keyChar = asc(0)
    If oEvent.keyCode <> 0 Then
        keyChar = getLatinKeyCharByCode(oEvent)
    Else
        keyChar = getLatinKeyCharByRus(oEvent)
    End If
    getLatinKey = keyChar
End Function

Function isControl(c)
	if len(c) = 0 then
		isControl = True
	else
		dim ac as integer : ac = asc(c)
		isControl = ((ac >= 0 and ac <= 31) or (ac = 127))
	end if
End Function

Function isPrintable(c)
	isPrintable = not isControl(c)
End Function

Sub restoreStatus 'restore original statusbar
	On Error Goto ErrorHandler
    dim oCurrentContorller : oCurrentContorller = getCurrentController()
   	dim oFrame : oFrame = oCurrentContorller.Frame
   	dim oLayout : oLayout = oFrame.LayoutManager
  	oLayout.destroyElement("private:resource/statusbar/statusbar")
   	oLayout.createElement("private:resource/statusbar/statusbar")
   	Exit Sub
ErrorHandler:
	MsgBox("restoreStatus() failed!")
End Sub

' Unfortunately, does not work as expected
' Statusbar of the currently active window is restored, because background windows
' return the controller of the currently active window on getCurrentController
' Thus the statusbar of active window is restored several times, and status bars
' of background windows are not restored at all
Sub restoreStatusOfModels()
    dim vComponents
	dim hasUnoI
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
			If APP() <> "CALC" Then
				If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then hasUnoI = True Else hasUnoI = False
			Else
				hasUnoI = True
			End If
    		If hasUnoI Then
				dim oController
				
    			oController = vComponent.getCurrentController()
	    		static oOldController as object
	    		If EqualUnoObjects(oController, oOldController) Then
	    		'	MsgBox("controllers are the same")
	    		End If
	    		oOldController = oController
	    		
    			If not (oController is Nothing) Then
			    	dim oFrame
			    	oFrame = oController.getFrame()
			    	dim oLayout
			    	oLayout = oFrame.LayoutManager
			    	oLayout.destroyElement("private:resource/statusbar/statusbar")
				    oLayout.createElement("private:resource/statusbar/statusbar")
			   	End If
    		End If
    	Loop	
    End If
End Sub

Sub setRawStatus(rawText)
	dim oCurrentController as object
	oCurrentController = getCurrentController()
	If not (oCurrentController is Nothing) Then
	    oCurrentcontroller.StatusIndicator.Start(rawText, 0)
	End If
End Sub

Sub setStatus()
    setRawStatus(getModeName(MODE) & " | " & getMultiplier() & " | special: " & getSpecial() & " | " & "modifier: " & getMovementModifier())
End Sub

' Try to use statusbar to save state of window, allowing for multiple windows to support vibreoffice'
' Seemingly No such funciton is available to read the statusbar from the api?'
Function getStatus() As String
    thisComponent.Currentcontroller.StatusIndicator
    'setRawStatus(MODE & " | " & statusText & " | special: " & getSpecial() & " | " & "modifier: " & getMovementModifier())
End Function

Sub setMode(m)
    MODE = m
    setStatus()
End Sub

' Selects the current line and makes it the Visual base line for use with 
' VISUAL_LINE mode.
Function formatVisualBase()
If APP() <> "CALC" Then
    dim oTextCursor
    oTextCursor = getTextCursor()
    VISUAL_BASE = getCursor().getPosition()

    ' Select the current line by moving cursor to start of the bellow line and 
    ' then back to the start of the current line.
    getCursor().gotoEndOfLine(False)
    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
        getCursor().goRight(1, False)
    End If
    getCursor().goLeft(1, True)
    getCursor().gotoStartOfLine(True)
Else
	simulate_KeyPress_Char("HOME")
	simulate_KeyPress_Char("END","SHIFT")
End If
End Function

Function formatVisual2VisualLine(Optional bKeepLastEOL)
    If IsMissing(bKeepLastEOL) Then bKeepLastEOL = False
    If APP() <> "CALC" Then
        dim oTextCursor, oldEndPos
        oTextCursor = getTextCursor()
        getCursor().gotoRange(oTextCursor.getStart(), False)
        getCursor().gotoStartOfLine(False)
        getCursor().gotoRange(oTextCursor.getEnd(), True)
        oldEndPos = getCursor().getPosition()
        getCursor().gotoEndOfLine(True)
        If Not bKeepLastEOL And getCursor().getPosition().Y() = oldEndPos.Y() Then
            getCursor().goRight(1, True)
        End If
    End If
End Function

Function gotoMode(sMode)
    Select Case sMode
        Case M_NORMAL, M_DISABLED:
            setMode(sMode)
            setMovementModifier("")
		Case M_INSERT:
            setMode(sMode)
		Case M_VISUAL:
            setMode(sMode)
			If APP() <> "CALC" Then
				dim oTextCursor
				oTextCursor = getTextCursor()
				' Deselect TextCursor
				If not (oTextCursor is Nothing) Then
					oTextCursor.gotoRange(oTextCursor.getStart(), False)
					' Show TextCursor selection
					getCurrentController().Select(oTextCursor)
				End If
			End If
		Case M_VISUAL_LINE:
            setMode(sMode)
            formatVisualBase()
		Case Else:
			' Should not happen
			setMode(M_BAD)
    End Select
End Function

Sub cursorPreReset(oTextCursor)
    oTextCursor.gotoRange(oTextCursor.getStart(), False)
    oTextCursor.goRight(1, False)
	oTextCursor.goLeft(1, True)
End Sub

Sub cursorReset(oTextCursor)
	cursorPreReset(oTextCursor)	
    getCurrentController().Select(oTextCursor) ' needed
End Sub

Sub searchAndSet(oTextCursor, sText, bIsBackwards)
	If APP() <> "CALC" Then
    dim oSearchDesc as object
    oSearchDesc = thisComponent.createSearchDescriptor()
    oSearchDesc.setSearchString(sText)
    oSearchDesc.SearchCaseSensitive = True
    oSearchDesc.SearchBackwards = bIsBackwards
    dim oStartRange	    
		If Not bIsBackwards Then
			oStartRange = oTextCursor.getEnd()
		Else
			oStartRange = oTextCursor.getStart()
		End If
    dim oFoundRange	    
    oFoundRange = thisComponent.findNext(oStartRange, oSearchDesc)	    
	If not (oFoundRange is Nothing) Then
		oTextCursor.gotoRange(oFoundRange, False)
		getCurrentController().Select(oTextCursor)
		setMode(M_VISUAL)
	End If
	Else
		simulate_KeyPress_Char("F","CTRL")
	End If
End Sub

' swaps cursor start and end
' have to resort to the crude string-based algorithm because of annotation chars
' BUG minor bug: on swap will exclude annotation chars present on either end of selection
Sub swapCursorEnds(oTextCursor as object)
	dim s, oldLen
	s = oTextCursor.getString()
	oldLen = len(s)
	If oldLen = 0 Then Exit Sub
	oTextCursor.goRight(1, True)
	dim newLen : newLen = len(oTextCursor.getString())
	If newLen > oldLen Then oTextCursor.goLeft(1, False)
	dim pureLen : pureLen = newLen - count(s, chr(10))
	dim t as string
	If newLen >= oldLen Then
		oTextCursor.collapseToEnd()
		oTextCursor.goLeft(pureLen-1, True)
        do while oTextCursor.getString() <> s
        	t = oTextCursor.getString()
			oTextCursor.goLeft(1, True)
		Loop
	Else
		oTextCursor.collapseToStart()
		oTextCursor.goLeft(1, False) 'collapsing to start does not include starting character for some reason
        oTextCursor.goRight(pureLen, True)
        do while oTextCursor.getString() <> s
        	t = oTextCursor.getString()
			oTextCursor.goRight(1, True)
		Loop
	End If
End Sub

Function samePos(oPos1, oPos2)
    samePos = oPos1.X() = oPos2.X() And oPos1.Y() = oPos2.Y()
End Function

Function genString(sChar, iLen)
    dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Counts number of character c in string s
Function count(s as string, c)
	dim cnt
	cnt = 0
	dim length
	length = len(s)
	dim i
	For i = 1 to length
		If asc(Mid(s,i,1)) = asc(c) Then cnt = cnt + 1
	Next
	count = cnt
End Function

' Yanks selection to system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    dim dispatcher As Object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(getCurrentController().Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
		If APP() <> "CALC" Then
			getTextCursor().setString("")
		Else
			simulate_KeyPress_Char("DELETE")
		End If
    End If
End Sub


Sub pasteSelection()
    dim oTextCursor, dispatcher As Object

    ' Deselect if in NORMAL mode to avoid overwriting the character underneath
    ' the cursor
    If MODE = M_NORMAL Then
		If APP() <> "CALC" Then
        oTextCursor = getTextCursor()
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        getCurrentController().Select(oTextCursor)
		End If
    End If

    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(getCurrentController().Frame(), ".uno:Paste", "", 0, Array())
End Sub


' -----------------------------------
' Special Mode (for chained commands)
' -----------------------------------
global SPECIAL_MODE As string
global SPECIAL_COUNT As integer

Sub setSpecial(specialName)
    SPECIAL_MODE = specialName

    If specialName = "" Then
        SPECIAL_COUNT = 0
    Else
        SPECIAL_COUNT = 2
    End If
End Sub

Function getSpecial()
    getSpecial = SPECIAL_MODE
End Function

Sub delaySpecialReset()
    SPECIAL_COUNT = SPECIAL_COUNT + 1
End Sub

Sub resetSpecial(Optional bForce)
    If IsMissing(bForce) Then bForce = False

    SPECIAL_COUNT = SPECIAL_COUNT - 1
    If SPECIAL_COUNT <= 0 Or bForce Then
        setSpecial("")
    End If
End Sub


' -----------------
' Movement Modifier
' -----------------
'f,i,a
global MOVEMENT_MODIFIER As string

Sub setMovementModifier(modifierName)
    MOVEMENT_MODIFIER = modifierName
End Sub

Function getMovementModifier()
    getMovementModifier = MOVEMENT_MODIFIER
End Function


' --------------------
' Multiplier functions
' --------------------
Sub _setMultiplier(n as integer)
    MULTIPLIER = n
End Sub

Sub resetMultiplier()
    _setMultiplier(0)
End Sub

Sub addToMultiplier(n as integer)
    dim sMultiplierStr as String
    dim iMultiplierInt as integer

    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 then
        sMultiplierStr = CStr(MULTIPLIER) & CStr(n)
        _setMultiplier(CInt(sMultiplierStr))
    End If
End Sub

' Should only be used if you need the raw value
Function getRawMultiplier()
    getRawMultiplier = MULTIPLIER
End Function

' Same as getRawMultiplier, but defaults to 1 if it is unset (0)
Function getMultiplier()
    If MULTIPLIER = 0 Then
        getMultiplier = 1
    Else
        getMultiplier = MULTIPLIER
    End If
End Function


' -------------
' Key Handling
' -------------
' Buggy due to the same reason as restoreStatusOfModels() is
Sub addKeyHandlerToModels()
	dim hasUnoI
	If oXKeyHandler is nothing Then
		Exit Sub
	End If
	
    dim vComponents
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
			If APP() <> "CALC" Then
				If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then hasUnoI = True Else hasUnoI = False
			Else
				hasUnoI = True
			End If
    		If hasUnoI Then
				dim oController
    			oController = vComponent.getCurrentController()
    			If not (oController is Nothing) Then
			    	oController().addKeyHandler(oXKeyHandler)
			   	End If
    		End If
    	Loop	
    End If
End Sub


' Buggy due to the same reason as restoreStatusOfModels() is
Sub removeKeyHandlerFromModels()
	dim hasUnoI
    dim vComponents
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
			If APP() <> "CALC" Then
				If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then hasUnoI = True Else hasUnoI = False
			Else
				hasUnoI = True
			End If
    		If hasUnoI Then
				dim oController
    			oController = vComponent.getCurrentController()
    			If not (oController is Nothing) Then
			    	oController().removeKeyHandler(oXKeyHandler)
			    End If
    		End If
    	Loop	
    End If
End Sub


Sub sStartXKeyHandler
	sStopXKeyHandler()
    oXKeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
End Sub

Sub sStopXKeyHandler
End Sub

Sub KeyHandler_Disposing(oEvent)
End Sub


' --------------------
' Main Key Processing
' --------------------
function KeyHandler_KeyPressed(oEvent) as boolean
    If oEvent.KeyCode = 1281 And oEvent.Modifiers = 1 Then
    	toggleVibreoffice()
    	KeyHandler_KeyPressed = True
    	Exit Function
    End If

    ' Exit if plugin is not enabled
    If MODE = M_DISABLED Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If
    
    ' Have to resort to polling because subscribing to theGlobalEventBroadcaster causes crashes
    dim oFrame : oFrame = StarDesktop.getCurrentFrame()
    if not EqualUnoObjects(oFrame, oCurrentFrame) then
    	reinitVibreOffice()
    	oCurrentFrame = oFrame
    end if
    
	if oEvent.keyChar = "=" and oEvent.Modifiers > 1 then
		MsgBox(logged2)
		logged2 = ""
		KeyHandler_KeyPressed = True
		Exit Function
	end if    
	
	if oEvent.keyChar = "-" and oEvent.Modifiers > 1 then
		dim oCur : oCur = getTextCursor()
		dim s : s = oCur.getString()
		
		if len(s) = 0 then 
			s = "<>"
		elseif isControl(s) then 
			s = "<" & asc(s) & ">"
		end if
		
		s = s & chr(13) & TRAP_STATE
		
		MsgBox(s)
		KeyHandler_KeyPressed = True
		Exit Function
	end if	


	'CHECK'
	If APP() <> "CALC" Then
    ' Exit if TextCursor does not work (as in Annotations)
	Dim oTextCursor
    oTextCursor = getTextCursor()
    If oTextCursor Is Nothing Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If
    End If

    dim bConsumeInput : bConsumeInput = True ' Block all inputs by default

        
    ' --------------------------
    ' Process global shortcuts, exit if matched (like ESC)
    If ProcessGlobalKey(oEvent) Then
        ' Pass
    ' If INSERT mode, allow all inputs
    ElseIf MODE = M_INSERT Then
		bConsumeInput = True
		dim c : c = oEvent.keyChar	
		
		if isPrintable(c) then
			logged2 = logged2 & c
			'TODO Revert if characters are still being swapped during input
			'print_string(getTextCursor(), c)
			'bConsumeInput = True
			bConsumeInput = False
		else
			logged2 = logged2 & "<" & asc(c) & ">"
			bConsumeInput = False
		end if
    Else
    	dim bIsMultiplier, bIsModified, bIsControl, bIsSpecial
	    bIsMultiplier = False ' reset multiplier by default
	    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
	    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)
	    bIsSpecial = getSpecial() <> ""        	
	    ' If Change Mode
	    ' ElseIf MODE = M_NORMAL And Not bIsSpecial And getMovementModifier() = "" And ProcessModeKey(oEvent) Then
	    If ProcessModeKey(oEvent) Then
	        ' Pass
	
	    ' Replace Key
	    ElseIf getSpecial() = "r" And Not bIsModified Then
	        dim iLen
	        iLen = Len(getCursor().getString())
	        getCursor().setString(genString(oEvent.KeyChar, iLen))
	
	    ' Normal Key
	    ElseIf ProcessNormalKey(getLatinKey(oEvent), oEvent.Modifiers, oEvent) Then
	        ' Pass
	
	    ' Multiplier Key
	    ElseIf ProcessNumberKey(oEvent) Then
	        bIsMultiplier = True
	        delaySpecialReset()
	
	    ' If is modified but doesn't match a normal command, allow input
	    '   (Useful for built-in shortcuts like Ctrl+a, Ctrl+s, Ctrl+w)
	    ElseIf bIsModified Then
	        ' Ctrl+a (select all) sets mode to VISUAL
	        If bIsControl And getLatinKey(oEvent) = "a" Then
	            gotoMode(M_VISUAL)
	        End If
	        bConsumeInput = False
	
	    ' Movement modifier here?
	    ElseIf ProcessMovementModifierKey(getLatinKey(oEvent)) Then
	        delaySpecialReset()
	
	    ' If standard movement key (in VISUAL mode) like arrow keys, home, end
	    ElseIf (MODE = M_VISUAL Or MODE = M_VISUAL_LINE) And ProcessStandardMovementKey(oEvent) Then
	        ' Pass
	
	    ' If bIsSpecial but nothing matched, return to normal mode
	    ElseIf bIsSpecial Then
	        gotoMode(M_NORMAL)
	
	    ' Allow non-letter keys if unmatched
	    ' TODO Use getLatinKey()
	    ElseIf asc(oEvent.KeyChar) = 0 Then
	        bConsumeInput = False
	    End If
	    ' --------------------------
	
	    ' Reset Special
	    resetSpecial()
	
	    ' Reset multiplier if last input was not number and not in special mode
	    If not bIsMultiplier and getSpecial() = "" and getMovementModifier() = "" Then
	        resetMultiplier()
	    End If
	    setStatus(getMultiplier())
	End If

    KeyHandler_KeyPressed = bConsumeInput
End Function

Function KeyHandler_KeyReleased(oEvent) As boolean
    ' Exit if plugin is not enabled
    If MODE = M_DISABLED Then
        KeyHandler_KeyReleased = False
        Exit Function
    End If
    
    If asc(oEvent.KeyChar) = 0 Then
        KeyHandler_KeyReleased = False
    Else
        dim iModifiers as integer
        iModifiers = oEvent.modifiers
        dim iKeyCode as integer
        iKeyCode = oEvent.keyCode
	    ' Allow Ctrl+c for Copy, so don't change cursor
        If iKeyCode = 514 And (iModifiers = 2 Or iModifiers = 8) Then
        ' Needed to make cursor always select 1 character in NORMAL mode
        ' Constrict to movement keys only?
        ElseIf MODE = M_NORMAL Then
	        ' Show terminal-like cursor
			dim oTextCursor
	        oTextCursor = getTextCursor()
	        If not (oTextCursor Is Nothing) Then
	            ' Do nothing        
	           cursorReset(oTextCursor)
	        End If
        End If

        KeyHandler_KeyReleased = (MODE = M_NORMAL) 'cancel KeyReleased
    End If
End Function


' ----------------
' Processing Keys
' ----------------
Function ProcessGlobalKey(oEvent)
    dim bMatched, bIsControl
    bMatched = False
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)

    ' keycode can be viewed here: http://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1Key.html
    ' PRESSED ESCAPE (or ctrl+[)
    If oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Then
        If APP() <> "CALC" Then
			If getTextCursor() is Nothing Then
				bMatched = False
			Else
				' Move cursor back if was in INSERT (but stay on same line)
				If MODE <> M_NORMAL And Not getCursor().isAtStartOfLine() Then
					getCursor().goLeft(1, False)
				End If
				bMatched = True
			End If
		Else
			If (MODE = M_VISUAL Or MODE = M_VISUAL_LINE) Then
			simulate_KeyPress_Char("ESCAPE")
			simulate_KeyPress_Char("DOWN")
			simulate_KeyPress_Char("UP")
			ElseIf (MODE = M_INSERT) Then
			' Prevents cell entries from being undone'
			simulate_KeyPress_Char("DOWN")
			simulate_KeyPress_Char("UP")
			
			End If
        	bMatched = True
        End If

        resetSpecial(True)
		resetMultiplier()
        gotoMode(M_NORMAL)
    Else
        bMatched = False
    End If
    ProcessGlobalKey = bMatched
End Function


Function ProcessStandardMovementKey(oEvent)
    dim c, bMatched
    c = oEvent.KeyCode

    bMatched = True

    If (MODE <> M_VISUAL And MODE <> M_VISUAL_LINE)Then
        bMatched = False
        'Pass
    ElseIf c = 1024 Then
        ProcessMovementKey("j", True)
    ElseIf c = 1025 Then
        ProcessMovementKey("k", True)
    ElseIf c = 1026 Then
        ProcessMovementKey("h", True)
    ElseIf c = 1027 Then
        ProcessMovementKey("l", True)
    ElseIf c = 1028 Then
        ProcessMovementKey("^", True)
    ElseIf c = "0"   and getRawMultiplier() = 0 Then
    'and getRawMultiplier() = 0
    	' Only if this entry is not part of the multiplier
			ProcessMovementKey("0", True) ' key for zero (0)

     ElseIf c = 1029 Then
        ProcessMovementKey("$", True)
    Else
        bMatched = False
    End If

    ProcessStandardMovementKey = bMatched
End Function


Function ProcessNumberKey(oEvent)
    dim c
    c = CStr(oEvent.KeyChar)

	' Don't treat number keys as multiplier-related 
	' if we are in modified movement mode (like f,t)
	' Otherwise will not be able to search for numbers with f or t at all
    If getMovementModifier() = "" and c >= "0" and c <= "9" Then
        addToMultiplier(CInt(c))
        ProcessNumberKey = True
    Else
        ProcessNumberKey = False
    End If
End Function


Function ProcessModeKey(oEvent)
	dim keyChar
    dim bIsModified
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    ' Don't change modes in these circumstances
    If MODE <> M_NORMAL Or bIsModified Or getSpecial <> "" Or getMovementModifier() <> "" Then
        ProcessModeKey = False
        Exit Function
    End If

    ' Mode matching
    dim bMatched, oTextCursor
    bMatched = True
    keyChar = getLatinKey(oEvent)
	oTextCursor = getTextCursor()
    Select Case oEvent.KeyChar
        ' Insert modes
        Case "i", "a", "I", "A", "o", "O":
			If APP() <> "CALC" Then
				If oEvent.KeyChar = "a" And NOT oTextCursor.isEndOfParagraph() Then getCursor().goRight(1, False)
				If oEvent.KeyChar = "I" Then ProcessMovementKey("^")
				If oEvent.KeyChar = "A" Then ProcessMovementKey("$")
			Else
				If oEvent.KeyChar = "I" Then 
					simulate_KeyPress_Char("F2")
					simulate_KeyPress_Char("HOME")
				End If
				If oEvent.KeyChar = "a" Then simulate_KeyPress_Char("F2")
				If oEvent.KeyChar = "A" Then simulate_KeyPress_Char("F2") 
			End If

            If KeyChar = "o" Then
				If APP() <> "CALC" Then
				    ProcessMovementKey("$")
                	ProcessMovementKey("l")
					getCursor().setString(chr(13))
					If Not getCursor().isAtStartOfLine() Then
						getCursor().setString(chr(13) & chr(13))
						ProcessMovementKey("l")
					End If
				Else
					insertRow(1)
					ProcessMovementKey("j")
				End If
            End If

            If KeyChar = "O" Then
				If APP() <> "CALC" Then
				    ProcessMovementKey("^")
					getCursor().setString(chr(13))
					If Not getCursor().isAtStartOfLine() Then
						ProcessMovementKey("h")
						getCursor().setString(chr(13))
						ProcessMovementKey("l")
					End If
				Else
					insertRow(0)
				End If
            End If

            gotoMode(M_INSERT)
        Case "v":
            gotoMode(M_VISUAL)
        Case "V":
            gotoMode(M_VISUAL_LINE)
        Case Else:
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(keyChar, modifiers, optional oEvent)
    dim i, bMatched, bMatchedMovement, bIsVisual, iIterations, bIsControl, sSpecial
	dim oTextCursor
    bIsControl = (modifiers = 2) or (modifiers = 8)
    bIsVisual = (MODE = M_VISUAL Or MODE = M_VISUAL_LINE) ' is this hardcoding bad? what about visual block?

	If bIsVisual and keyChar = "o" Then
		oTextCursor = getTextCursor()
		swapCursorEnds(oTextCursor)
		getCurrentController().select(oTextCursor)
        ProcessNormalKey = True
        Exit Function		
	End If

    ' ----------------------
    ' 1. Check Movement Key
    ' ----------------------
    iIterations = getMultiplier()
    sSpecial = getSpecial()
    bMatched = False
    bMatchedMovement = False
    ' starting from 0 adds one extra, unneeded movement
    ' FIXME axf Currently, say, <num>f<char> and d<num>f<char> commands differ if cursor is on <char>
    For i = 1 To iIterations 
        ' Movement Key
        ' axf Passing oEvent to make actual key char available for search
        bMatchedMovement = ProcessMovementKey(keyChar, bIsVisual, modifiers, oEvent)
        bMatched = bMatched or bMatchedMovement
    Next i

    ' If Special: d/c + movement
    If bMatched And (sSpecial = "d" Or sSpecial = "c" Or sSpecial = "y") Then
        yankSelection((sSpecial <> "y"))
    End If    

    ' Reset Movement Modifier
    setMovementModifier("")

    ' Exit already if movement key was matched
    If bMatched Then
        ' If Special: d/c : change mode
        If getSpecial() = "d" Or getSpecial() = "y" Then gotoMode(M_NORMAL)
        If getSpecial() = "c" Then gotoMode(M_INSERT)

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 2. Undo/Redo
    ' --------------------
    If keyChar = "u" Or (bIsControl And keyChar = "r") Then
        dim mode
        mode = 0
        If keyChar = "u" Then
            mode = 1
        End If

        For i = 1 To iIterations
            Undo(mode)
        Next i

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 3. Paste
    '   Note: in vim, paste will result in cursor being over the last character
    '   of the pasted content. Here, the cursor will be the next character
    '   after that. Fix?
    ' --------------------
    If keyChar = "p" or keyChar = "P" Then
        ' Move cursor right if "p" to paste after cursor
		If APP() <> "CALC" Then
        Redim oTextCursor
        oTextCursor = getTextCursor()
        ' Move cursor right if "p" to paste after cursor
        If keyChar = "p" And NOT oTextCursor().isEndOfParagraph() Then
            ProcessMovementKey("l", False)
        End If
		End If

        For i = 1 To iIterations
            pasteSelection()
        Next i

        ProcessNormalKey = True
        Exit Function
    End If
    
	'CHECK'
    ' search
    ' HACK Remapping to make compatible with Russian keyboard layout
    If keyChar = "." Then 
    	keyChar = "/" 
    ElseIf keyChar = ">" Then
    	keyChar = "?"
    End If
    
    If keyChar = "/" or keyChar = "?" Then
    	dim sDir
    	If keyChar = "/" Then
    		sDir = "forward"
    	Else
    		sDir = "backward"
    	End If
    	sDir = "Search " & sDir
		If APP() <> "CALC" Then
    	dim sInput
    	sInput = InputBox(sDir, sDir)
    	If sInput <> "" Then
    		dim bIsBackwards
	 	    bIsBackwards = (keyChar = "?")
		    searchAndSet(getTextCursor(), sInput, bIsBackwards)
			LAST_SEARCH = sInput
	   		ProcessNormalKey = True
	        Exit Function	
	     End If	
		 Else
			 simulate_KeyPress_Char("F","CTRL")
		 End If
    End If

    ' --------------------
    ' 4. Check Special/Delete Key
    ' --------------------

    ' There are no special/delete keys with modifier keys, so exit early
    If modifiers > 1 Then
        ProcessNormalKey = False
        Exit Function
    End If

    ' Only 'x' or Special (dd, cc) can be done more than once
    If getSpecial() = "" Then
        iIterations = 1
    End If
    For i = 1 To iIterations
        dim bMatchedSpecial

        ' Special/Delete Key
        bMatchedSpecial = ProcessSpecialKey(keyChar)

        bMatched = bMatched or bMatchedSpecial
    Next i


    ProcessNormalKey = bMatched
End Function


' Function for both undo and redo
Sub Undo(bUndo)
    On Error Goto ErrorHandler

    If bUndo Then
        thisComponent.getUndoManager().undo()
    Else
        thisComponent.getUndoManager().redo()
    End If
    Exit Sub

    ' Ignore errors from no more undos/redos in stack
ErrorHandler:
    Resume Next
End Sub


Function ProcessSpecialKey(keyChar)
    dim i, oCursor, oTextCursor, bMatched, bIsSpecial, bIsDelete, iIterations
    bMatched = True
    bIsSpecial = getSpecial() <> ""
    iIterations = getMultiplier()

    If keyChar = "d" Or keyChar = "c" Or keyChar = "s" Or keyChar = "y" Then
        bIsDelete = (keyChar <> "y")

        ' Special Cases: 'dd' and 'cc'
        If bIsSpecial Then
            dim bIsSpecialCase
            bIsSpecialCase = (keyChar = "d" And getSpecial() = "d") Or (keyChar = "c" And getSpecial() = "c")

            If bIsSpecialCase Then


				If APP() <> "CALC" Then
					 'ProcessMovementKey("^", False)
                     'ProcessMovementKey("j", True)
            	' A bit hacky, but works
					oCursor = getCursor()
					oCursor.gotoStartOfLine(False)
					oCursor.gotoEndOfLine(True)                

					oTextCursor = getTextCursor()
					'oTextCursor.goRight(1, True)
					getCurrentController().Select(oTextCursor)
					yankSelection(bIsDelete)
				Else
					yankSelection(bIsDelete)
					removeRow()
				End If
                
            Elseif (keyChar = "y" And getSpecial() = "y") Then
            	yankSelection(False)
            Else
                bMatched = False
            End If

            ' Go to INSERT mode after 'cc', otherwise NORMAL
            If bIsSpecialCase And keyChar = "c" Then
                gotoMode(M_INSERT)
            Else
                gotoMode(M_NORMAL)
            End If


        ' visual mode: delete selection
        ElseIf MODE = M_VISUAL Or MODE = M_VISUAL_LINE Then
			If APP() <> "CALC" Then
				oTextCursor = getTextCursor()
				getCurrentController().Select(oTextCursor)
			End If

            yankSelection(bIsDelete)

            If keyChar = "c" Or keyChar = "s" Then gotoMode(M_INSERT)
            If keyChar = "d" Or keyChar = "y" Then gotoMode(M_NORMAL)


        ' Enter Special mode: 'd', 'c', 'y' or 's'
        ElseIf MODE = M_NORMAL Then
				If keyChar = "s" Then
					If APP() <> "CALC" Then
						getCursor().gotoRange(getCursor().getStart(), False)
						getCursor().goRight(iIterations, True)
						yankSelection(True)
						gotoMode(M_INSERT)
						resetMultiplier()
						resetSpecial(True)
					Else
						setSpecial("c")
						gotoMode(M_VISUAL)		
						yankSelection(True)
						simulate_KeyPress_Char("DELETE")	
						gotoMode(M_INSERT)	
					End If
				Else
					setSpecial(keyChar)
					gotoMode(M_VISUAL)
				End If

        End If

    ' If is 'r' for replace
    ElseIf keyChar = "r" Then
		If APP() <> "CALC" Then
			setSpecial("r")
		End If
	' gg to go to beginning of text
	ElseIf keyChar = "g" Then
		If bIsSpecial Then
			If getSpecial() = "g" Then
			dim bExpand
			If APP() <> "CALC" Then
                ' If cursor is to left of current visual selection then select 
                ' from right end of the selection to the start of file.
                ' If cursor is to right of current visual selection then select 
                ' from left end of the selection to the start of file.
                If MODE = M_VISUAL Then
                    dim oldPos
                    oldPos = getCursor().getPosition()
                    getCursor().gotoRange(getCursor().getStart(), True)
                    If NOT samePos(getCursor().getPosition(), oldPos) Then
                        getCursor().gotoRange(getCursor().getEnd(), False)
                    End If

                ' If in VISUAL_LINE mode and cursor is bellow the Visual base 
                ' line then move it to the Visual base line, reformat the 
                ' Visual base line, and move cursor to start of file.
                ElseIf MODE = M_VISUAL_LINE Then
                    Do Until getCursor().getPosition().Y() <= VISUAL_BASE.Y()
                        getCursor().goUp(1, False)
                    Loop
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        formatVisualBase()
                    End If
                End If

                bExpand = MODE = M_VISUAL Or MODE = M_VISUAL_LINE
                getCursor().gotoStart(bExpand)
			Else
                bExpand = MODE = M_VISUAL Or MODE = M_VISUAL_LINE
				If bExpand Then simulate_KeyPress_Char("UP", "SHIFT", "CTRL") Else simulate_KeyPress_Char("UP", "CTRL")
			End If
			End If
		ElseIf MODE = M_NORMAL Or MODE = M_VISUAL Or MODE = M_VISUAL_LINE Then
			setSpecial("g")
		End If
			
		
    ' Otherwise, ignore if bIsSpecial
    ElseIf bIsSpecial Then
        bMatched = False

    ElseIf keyChar = "x" Or keyChar = "X" Then
		If APP() <> "CALC" Then
			If MODE = M_NORMAL Then
				getCursor().gotoRange(getCursor().getStart(), False)
				If keyChar = "x" Then
					getCursor().goRight(iIterations, True)
				Else
					getCursor().goLeft(iIterations, True)
				End If
				yankSelection(True)
			ElseIf MODE = M_VISUAL Or MODE = M_VISUAL_LINE Then
				If keyChar = "x" Then
					yankSelection(True)
				Else
					formatVisual2VisualLine()
					yankSelection(True)
				End If
			End If
		Else
			yankSelection(True)
			simulate_KeyPress_Char("DELETE")
		End If

        ' Goto NORMAL mode (in the case of VISUAL mode)
		gotoMode(M_NORMAL)
		resetMultiplier()
		resetSpecial(True)

    ElseIf keyChar = "D" Or keyChar = "C" Then
        If MODE = M_VISUAL Or MODE = M_VISUAL_LINE Then
            ProcessMovementKey("^", False)
            ProcessMovementKey("$", True)
            ProcessMovementKey("l", True)
        Else
            ' Deselect
			If APP() <> "CALC" Then
				oTextCursor = getTextCursor()
				oTextCursor.gotoRange(oTextCursor.getStart(), False)
				getCurrentController().Select(oTextCursor)
			End If
            ProcessMovementKey("$", True)
        End If

        yankSelection(True)

        If keyChar = "D" Then
            gotoMode(M_NORMAL)
        ElseIf keyChar = "C" Then
            gotoMode(M_INSERT)
        End IF

    ElseIf keyChar = "S" Then
		If APP() <> "CALC" Then
			If MODE = M_NORMAL Then
				getCursor().gotoStartOfLine(False)
				getCursor().goDown(iIterations - 1, True)
				getCursor().gotoEndOfLine(True)
				yankSelection(True)
				gotoMode(M_INSERT)
			ElseIf MODE = M_VISUAL Or MODE = M_VISUAL_LINE Then
				formatVisual2VisualLine(True)
				yankSelection(True)
				gotoMode(M_INSERT)
			End If
			resetMultiplier()
			resetSpecial(True)
		End If

    Else
        bMatched = False
    End If

    ProcessSpecialKey = bMatched
End Function


Function ProcessMovementModifierKey(keyChar)
    dim bMatched

    bMatched = True
    Select Case keyChar
        Case "f", "t", "F", "T", "i", "a":
            setMovementModifier(keyChar)
        Case Else:
            bMatched = False
    End Select

    ProcessMovementModifierKey = bMatched
End Function


Function ProcessSearchKey(oTextCursor, searchType, keyChar, bExpand)
    '-----------
    ' Searching
    '-----------
    dim bMatched, oSearchDesc, oFoundRange, bIsBackwards, oStartRange
    bMatched = True
	If APP() <> "CALC" Then
		bIsBackwards = (searchType = "F" Or searchType = "T")

		If Not bIsBackwards Then
			' VISUAL mode will goRight AFTER the selection
			If MODE <> M_VISUAL And MODE <> M_VISUAL_LINE Then
				' Start searching from next character
				oTextCursor.goRight(1, bExpand)
			End If

			oStartRange = oTextCursor.getEnd()
			' Go back one
			oTextCursor.goLeft(1, bExpand)
		Else
			oStartRange = oTextCursor.getStart()
		End If

		oSearchDesc = thisComponent.createSearchDescriptor()
		oSearchDesc.setSearchString(keyChar)
		oSearchDesc.SearchCaseSensitive = True
		oSearchDesc.SearchBackwards = bIsBackwards

		oFoundRange = thisComponent.findNext( oStartRange, oSearchDesc )

		If not IsNull(oFoundRange) Then
			dim oText, foundPos, curPos, bSearching
			oText = oTextCursor.getText()
			foundPos = oFoundRange.getStart()

			' Unfortunately, we must go go to this "found" position one character at
			' a time because I have yet to find a way to consistently move the
			' Start range of the text cursor and leave the End range intact.
			If bIsBackwards Then
				curPos = oTextCursor.getEnd()
			Else
				curPos = oTextCursor.getStart()
			End If
			do until oText.compareRegionStarts(foundPos, curPos) = 0
				If bIsBackwards Then
					bSearching = oTextCursor.goLeft(1, bExpand)
					curPos = oTextCursor.getStart()
				Else
					bSearching = oTextCursor.goRight(1, bExpand)
					curPos = oTextCursor.getEnd()
				End If

				' Prevent infinite if unable to find, but shouldn't ever happen (?)
				If Not bSearching Then
					bMatched = False
					Exit Do
				End If
			Loop

			If searchType = "t" Then
				oTextCursor.goLeft(1, bExpand)
			ElseIf searchType = "T" Then
				oTextCursor.goRight(1, bExpand)
			End If

		Else
			bMatched = False
		End If

		' If matched, then we want to select PAST the character
		' Else, this will counteract some weirdness. hack either way
		If Not bIsBackwards And (MODE = M_VISUAL Or MODE = M_VISUAL_LINE) Then
			oTextCursor.goRight(1, bExpand)
		End If
	Else
    bMatched = False
	End If

    ProcessSearchKey = bMatched

End Function


Function ProcessInnerKey(oTextCursor, movementModifier, keyChar, bExpand)
    dim bMatched, searchType1, searchType2, search1, search2

	If APP() <> "CALC" Then
		' Setting searchType
		If movementModifier = "i" Then
			searchType1 = "T" : searchType2 = "t"
		ElseIf movementModifier = "a" Then
			searchType1 = "F" : searchType2 = "f"
		Else ' Shouldn't happen
			ProcessInnerKey = False
			Exit Function
		End If

		Select Case keyChar
			Case "(", ")":
				search1 = "(" : search2 = ")"
			Case "{", "}":
				search1 = "{" : search2 = "}"
			Case "[", "]":
				search1 = "[" : search2 = "}"
			Case "<", ">":
				search1 = "<" : search2 = ">"
			Case "t":
				search1 = ">" : search2 = "<"
			Case "'":
				search1 = "'" : search2 = "'"
			Case """":
				' Matches "smart" quotes, which is default in libreoffice
				search1 = "?" : search2 = "?"
			Case Else:
				search1 = keyChar : search2 = keyChar
		End Select
		
			dim bMatched1, bMatched2
			bMatched1 = ProcessSearchKey(oTextCursor, searchType1, search1, False)
			bMatched2 = ProcessSearchKey(oTextCursor, searchType2, search2, True)
			' Temp hack - need to search 2nd time to make da<char> work correctly
			If (search1 = search2) and (searchType2 = "f") and bMatched1 Then
				bMatched2 = ProcessSearchKey(oTextCursor, searchType2, search2, True)
			End If
			bMatched = (bMatched1 And bMatched2)
	Else
		bMatched = False
	End If

    ProcessInnerKey = bMatched
End Function


' -----------------------
' Main Movement Function
' -----------------------
'   Default: bExpand = False, keyModifiers = 0
'   axf Need to pass oEvent to make in available to called functions (i.e. search)
Function ProcessMovementKey(keyChar, Optional bExpand, Optional keyModifiers, Optional oEvent)
    dim oTextCursor, bSetCursor, bMatched
    'If APP() <> "CALC" Then
    	oTextCursor = getTextCursor()
    'End If
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        dim bIsControl
        bIsControl = (keyModifiers = 2) or (keyModifiers = 8)

        ' Ctrl+d and Ctrl+u
        If bIsControl and keyChar = "d" Then
			If APP() <> "CALC" Then
				getCursor().ScreenDown(bExpand)
			Else
				If bExpand Then simulate_KeyPress_Char("PAGEDOWN", "SHIFT") Else simulate_KeyPress_Char("PAGEDOWN")
			End If
        ElseIf bIsControl and keyChar = "u" Then
			If APP() <> "CALC" Then
				getCursor().ScreenUp(bExpand)
			Else
				If bExpand Then simulate_KeyPress_Char("PAGEUP", "SHIFT") Else simulate_KeyPress_Char("PAGEUP")
			End If
        Else
            bMatched = False
        End If

        ProcessMovementKey = bMatched
        Exit Function
    End If

    ' Set global cursor to oTextCursor's new position if moved
    bSetCursor = True


    ' ------------------
    ' Movement matching
    ' ------------------

    ' ---------------------------------
    ' Special Case: Modified movements
	If getMovementModifier() <> "" Then
		Select Case getMovementModifier()
			' f,F,t,T searching
		Case "f", "t", "F", "T":
		'If APP() <> "CALC" Then
			bMatched  = ProcessSearchKey(oTextCursor, getMovementModifier(), keyChar, bExpand)
                LAST_SEARCH = oEvent.keyChar
		'End If
		Case "i", "a":
			bMatched = ProcessInnerKey(oTextCursor, getMovementModifier(), keyChar, bExpand)

		Case Else:
			bSetCursor = False
			bMatched = False
		End Select

		If Not bMatched Then
			bSetCursor = False
		End If
	' ---------------------------------

    ' Search repetition
    ElseIf keyChar = "n" or keyChar = "N" Then
        If keyChar = "n" Then
            ' MsgBox("n: " & LAST_SEARCH)
            ' bMatched  = ProcessSearchKey(oTextCursor, "f", LAST_SEARCH_CHAR, bExpand)
            searchAndSet(getTextCursor(), LAST_SEARCH, False) 
        ElseIf keyChar = "N" Then
            ' MsgBox("N: " & LAST_SEARCH)
            ' bMatched  = ProcessSearchKey(oTextCursor, "F", LAST_SEARCH_CHAR, bExpand) 
			searchAndSet(getTextCursor(), LAST_SEARCH, True)
        End If
        bSetCursor = False

    ' Basic movement
    ElseIf keyChar = "l" Then
    	 If APP() = "CALC" Then
	       	'Print "This is Calc"
			If bExpand Then simulate_KeyPress_Char("RIGHT", "SHIFT") Else simulate_KeyPress_Char("RIGHT")
        Else
        oTextCursor.goRight(1, bExpand)
        End If

    ElseIf keyChar = "h" Then
    	 If APP() = "CALC" Then
	       	'Print "This is Calc"
			If bExpand Then simulate_KeyPress_Char("LEFT", "SHIFT") Else simulate_KeyPress_Char("LEFT")
        Else
        oTextCursor.goLeft(1, bExpand)
        End If

    ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
    ' So this is a weird hack
    ElseIf keyChar = "k" Then
        'oTextCursor.goUp(1, False)
        If APP() = "CALC" Then
	       	'Print "This is Calc"
			If bExpand Then simulate_KeyPress_Char("UP", "SHIFT") Else simulate_KeyPress_Char("UP")
        Else
        If MODE = M_VISUAL_LINE Then
            ' This variable represents the line that the user last selected.
            dim lastSelected

            ' If cursor is already on or above the Visual base line.
            If getCursor().getPosition().Y() <= VISUAL_BASE.Y() Then
                lastSelected = getCursor().getPosition().Y()
                ' If on Visual base line then format it for selecting above 
                ' lines.
                If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                    getCursor().gotoEndOfLine(False)
                    ' Make sure that cursor is on the start of the line bellow 
                    ' the Visual base line. This is needed to make sure the 
                    ' new line character will be selected.
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        getCursor().goRight(1, False)
                    End If
                End If

                ' Move cursor to start of the line above last selected line.
                Do Until getCursor().getPosition().Y() < lastSelected
                    If NOT getCursor().goUp(1, bExpand) Then
                        Exit Do
                    End If
                Loop
                getCursor().gotoStartOfLine(bExpand)

            ' If cursor is already bellow the Visual base line.
            ElseIf getCursor().getPosition().Y() > VISUAL_BASE.Y() Then
                ' Cursor will be under the last selected line so it needs to 
                ' be moved up before setting lastSelected.
                getCursor().goUp(1, bExpand)
                lastSelected = getCursor().getPosition().Y()
                ' Move cursor up another line to deselect the last selected
                ' line.
                getCursor().goUp(1, bExpand)

                ' For the case when the last selected line was the line bellow 
                ' the Visual base line, simply reformat the Visual base line.
                If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                    formatVisualBase()

                Else
                    ' Make sure that the current line is fully selected.
                    getCursor().gotoEndOfLine(bExpand)

                    ' Make sure cursor is at the start of the line we 
                    ' deselected. It needs to always be bellow the user's 
                    ' selection when under the Visual base line.
                    If getCursor().getPosition().Y() < lastSelected Then
                        getCursor().goRight(1, bExpand)
                    End If
                End If

            End If

        Else
        ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
        ' So this is a weird hack
            'oTextCursor.goUp(1, False)
            getCursor().goUp(1, bExpand)
        End If
        bSetCursor = False
        End If

    ElseIf keyChar = "j" Then
        If APP() = "CALC" Then
	       	'Print "This is Calc"
			If bExpand Then simulate_KeyPress_Char("DOWN", "SHIFT") Else simulate_KeyPress_Char("DOWN")
        Else
        If MODE = M_VISUAL_LINE Then
            ' If cursor is already on or bellow the Visual base line.
            If getCursor().getPosition().Y() >= VISUAL_BASE.Y() Then
                ' If on Visual base line then format it for selecting bellow 
                ' lines.
                If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                    getCursor().gotoStartOfLine(False)
                    getCursor().gotoEndOfLine(bExpand)
                    ' Move cursor to next line if not already there.
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        getCursor().goRight(1, bExpand)
                    End If

                End If

                If getCursor().goDown(1, bExpand) Then
                    getCursor().gotoStartOfLine(bExpand)

                ' If cursor is on last line then select from current position 
                ' to end of line.
                Else
                    getCursor().gotoEndOfLine(bExpand)
                End If

            ' If cursor is above the Visual base line.
            ElseIf getCursor().getPosition().Y() < VISUAL_BASE.Y() Then
                ' Move cursor to start of bellow line.
                getCursor().goDown(1, bExpand)
                getCursor().gotoStartOfLine(bExpand)
            End If

        Else
        ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
        ' So this is a weird hack
            'oTextCursor.goDown(1, False)
            getCursor().goDown(1, bExpand)
        End If
        bSetCursor = False
        End If
    ' ----------

    ElseIf keyChar = "J" Then
	' Select Previous Sheet'
        If APP() = "CALC" Then
			simulate_KeyPress_Char("PAGEUP", "CTRL")
        End If
    ElseIf keyChar = "K" Then
	' Select Next Sheet'
        If APP() = "CALC" Then
			simulate_KeyPress_Char("PAGEDOWN", "CTRL")
        End If
    ElseIf keyChar = "0"  and getRawMultiplier() = 0 Then
        if keyModifiers = 0 Then
        	If APP() <> "CALC" Then
				getCursor().gotoStartOfLine(bExpand)
				bSetCursor = False
			Else
				If bExpand Then simulate_KeyPress_Char("HOME", "SHIFT") Else simulate_KeyPress_Char("HOME")
        	End If
        End If
    ElseIf keyChar = "^" Then
        If APP() <> "CALC" Then
        ' This variable represents the original line the cursor was on before 
        ' any of the following changes.
        dim oldLine
        oldLine = getCursor().getPosition().Y()

        ' Select all of the current line and put it into a string.
        getCursor().gotoEndOfLine(False)
        If getCursor().getPosition.Y() > oldLine Then
            ' If gotoEndOfLine moved cursor to next line then move it back.
            getCursor().goLeft(1, False)
        End If
        getCursor().gotoStartOfLine(True)
        dim s as String
        s = getCursor().String

        ' Undo any changes made to the view cursor, then move to start of 
        ' line. This way any previous selction made by the user will remain.
        getCursor().gotoRange(oTextCursor, False)
        getCursor().gotoStartOfLine(bExpand)

        ' This integer will be used to determine the position of the first 
        ' character in the line that is not a space or a tab.
        dim i as Integer
        i = 1

        ' Iterate through the characters in the string until a character that 
        ' is not a space or a tab is found.
        Do While i <= Len(s)
            dim c
            c = Mid(s,i,1)
            If c <> " " And c <> Chr(9) Then
                Exit Do
            End If
            i = i + 1
        Loop

        ' Move the cursor to the first non space/tab character.
        getCursor().goRight(i - 1, bExpand)
        bSetCursor = False
        Else
			If bExpand Then simulate_KeyPress_Char("HOME", "SHIFT") Else simulate_KeyPress_Char("HOME")
		End If

    ElseIf keyChar = "$" Then
        If APP() <> "CALC" Then
			dim oldPos, newPos
			oldPos = getCursor().getPosition()
			getCursor().gotoEndOfLine(bExpand)
			newPos = getCursor().getPosition()

			' If the result is at the start of the line, then it must have
			' jumped down a line; goLeft to return to the previous line.
			'   Except for: Empty lines (check for oldPos = newPos)
			If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
				getCursor().goLeft(1, bExpand)
			End If

			' maybe eventually cursorGoto... should return True/False for bsetCursor
			bSetCursor = False

		Else
			If bExpand Then simulate_KeyPress_Char("END", "SHIFT") Else simulate_KeyPress_Char("END")
        End If
    ElseIf keyChar = "G" Then
		If APP() <> "CALC" Then
        If MODE = M_VISUAL_LINE Then
            ' If cursor is above Visual base line then move cursor down to it. 
            Do Until getCursor().getPosition.Y() >= VISUAL_BASE.Y()
                getCursor().goDown(1, False)
            Loop
            ' If cursor is on Visual base line then move it to start of line.
            If getCursor().getPosition.Y() = VISUAL_BASE.Y() Then
                getCursor().gotoStartOfLine(False)
            End If
        End If
        getCursor().gotoEnd(bExpand)
        bSetCursor = False
		Else
			If bExpand Then simulate_KeyPress_Char("DOWN", "SHIFT", "CTRL") Else simulate_KeyPress_Char("DOWN", "CTRL")
		End If

    ElseIf keyChar = "w" or keyChar = "W" Then
        If APP() <> "CALC" Then
        ' For the case when the user enters "cw":
        If getSpecial() = "c" Then
            ' If the cursor is on a word then delete from the current position to 
            ' the end of the word.
            ' If the cursor is not on a word then delete from the current position 
            ' to the start of the next word or until the end of the paragraph.

            If NOT oTextCursor.isEndOfParagraph() Then
               ' Move cursor to right in case it is already at start or end of 
               ' word.
               oTextCursor.goRight(1, bExpand)
            End If

            Do Until oTextCursor.isEndOfWord() Or oTextCursor.isStartOfWord() Or oTextCursor.isEndOfParagraph()
                oTextCursor.goRight(1, bExpand)
            Loop

        ' For the case when the user enters "w" or "dw":
        Else
            ' Note: For "w", using gotoNextWord would mean that the cursor 
            ' would not be moved to the next word when it involved moving down 
            ' a line and that line happened to begin with whitespace. It would 
            ' also mean that the cursor would not skip over lines that only 
            ' contain whitespace.

            If NOT (getSpecial() = "d" And oTextCursor.isEndOfParagraph()) Then
                ' Move cursor to right in case cursor is already at the start 
                ' of a word. 
                ' Additionally for "w", move right in case already on an empty 
                ' line.
                oTextCursor.goRight(1, bExpand)
            End If

            ' Stop looping when the cursor reaches the start of a word, an empty 
            ' line, or cannot be moved further (reaches end of file).
            ' Additionally, if "dw" then stop looping if end of paragraph is reached.
            Do Until oTextCursor.isStartOfWord() Or (oTextCursor.isStartOfParagraph() And oTextCursor.isEndOfParagraph())
                ' If "dw" then do not delete past the end of the line
                If getSpecial() = "d" And oTextCursor.isEndOfParagraph() Then
                    Exit Do
                ' If "w" then stop advancing cursor if cursor can no longer 
                ' move right
                ElseIf NOT oTextCursor.goRight(1, bExpand) Then
                    Exit Do
                End If
            Loop
        End If
		Else
			If bExpand Then simulate_KeyPress_Char("RIGHT", "SHIFT", "CTRL") Else simulate_KeyPress_Char("RIGHT", "CTRL")
        End If
    ElseIf keyChar = "b" or keyChar = "B" Then
        If APP() <> "CALC" Then
        ' When the user enters "b", "cb", or "db":

        ' Note: The function gotoPreviousWord causes a lot of problems when 
        ' trying to emulate vim behavior. The following method doesn't have to 
        ' account for as many special cases.

        ' "b": Moves the cursor to the start of the previous word or until an empty 
        ' line is reached.

        ' "db": Does same thing as "b" only it deletes everything between the 
        ' orginal cursor position and the new cursor position. The exception to 
        ' this is that if the original cursor position was at the start of a 
        ' paragraph and the new cursor position is on a separate paragraph with 
        ' at least two words then don't delete the new line char to the "left" 
        ' of the original paragraph.

        ' "dc": Does the same as "db" only the new line char described in "db" 
        ' above is never deleted.


        ' This variable is used to tell whether or not we need to make a 
        ' distinction between "b", "cb", and "db".
        dim dc_db as boolean

        ' Move cursor to left in case cursor is already at the start of a word 
        ' or on on an empty line. If cursor can move left and user enterd "dc" 
        ' or "db" and the cursor was originally on the start of a paragraph 
        ' then set dc_db to true and unselect the new line character separating 
        ' the paragraphs. If cursor can't move left then there is no line above 
        ' the current one and no need to make a distinction between "b", "cb", 
        ' and "db".
        dc_db = False
        If oTextCursor.isStartOfParagraph() And oTextCursor.goLeft(1, bExpand) Then
            If getSpecial() = "c" Or getSpecial() = "d" Then
                dc_db = True
                ' If all conditions above are met then unselect the \n char.
                oTextCursor.collapseToStart()
            End If
        End If

        ' Stop looping when the cursor reaches the start of a word, an empty 
        ' line, or cannot be moved further (reaches start of file).
        Do Until oTextCursor.isStartOfWord() Or (oTextCursor.isStartOfParagraph() And oTextCursor.isEndOfParagraph())
            ' Stop moving cursor if cursor can no longer move left
            If NOT oTextCursor.goLeft(1, bExpand) Then
                Exit Do
            End If
        Loop

        If dc_db Then
            ' Make a clone of oTextCursor called oTextCursor2 and use it to 
            ' check if there are at least two words in the "new" paragraph. 
            ' If there are <2 words then the loop will stop when the cursor 
            ' cursor reaches the start of a paragraph. If there >=2 words then 
            ' then the loop will stop when the cursor reaches the end of a word.
            dim oTextCursor2
            oTextCursor2 = getCursor().getText.createTextCursorByRange(oTextCursor)
            Do Until oTextCursor2.isEndOfWord() Or oTextCursor2.isStartOfParagraph()
                oTextCursor2.goLeft(1, bExpand)
            Loop
            ' If there are less than 2 words on the "new" paragraph then set 
            ' oTextCursor to oTextCursor 2. This is because vim's behavior is 
            ' to clear the "new" paragraph under these conditions.
            If oTextCursor2.isStartOfParagraph() Then
                oTextCursor = oTextCursor2
                oTextCursor.gotoRange(oTextCursor.getStart(), bExpand)
                ' If user entered "db" then reselect the \n char from before.
                If getSpecial() = "d" Then
                    oTextCursor.goRight(1, bExpand)
                End If
            End If
        End If
		Else
			If bExpand Then simulate_KeyPress_Char("LEFT", "SHIFT", "CTRL") Else simulate_KeyPress_Char("LEFT", "CTRL")
        End If
    ElseIf keyChar = "e" Then
        If APP() <> "CALC" Then
        ' When the user enters "e", "ce", or "de":

        ' Note: The function gotoNextWord causes a lot of problems when trying 
        ' to emulate vim behavior. The following method doesn't have to account 
        ' for as many special cases.

        ' Moves the cursor to the end of the next word or end of file if there 
        ' are no more words.

        ' Move cursor to right by two in case cursor is already at vim's 
        ' definition of endOfWord.
        oTextCursor.goRight(2, bExpand)

        ' If moving cursor to right by 2 places cursor just to the right of a 
        ' "." then move cursor right again. This is needed to ensure that the 
        ' cursor does not get stuck.
        getCursor().gotoRange(oTextCursor.getEnd(), False)
        getCursor().goLeft(1, True)
        If getCursor().String = "." Then
            oTextCursor.goRight(1, bExpand)
        End If

        ' gotoEndOfWord gets stuck sometimes so manually moving the cursor 
        ' right is necessary in these cases.
        Do Until oTextCursor.gotoEndOfWord(bExpand)
            ' If cursor can no longer move right then break loop
            If NOT oTextCursor.goRight(1, bExpand) Then
                Exit Do
            End If
        Loop

        If oTextCursor.isEndOfWord() Then
            ' LibreOffice defines a "." directly following a word to be the 
            ' endOfWord and vim does not. So in this case we need to move the 
            ' the cursor to the left.
            getCursor().gotoRange(oTextCursor.getEnd(), False)
            getCursor().goLeft(1, True)
            If getCursor().String = "." Then
                oTextCursor.goLeft(1, bExpand)
            End If

            ' gotoEndOfWord moves the cursor one character further than vim 
            ' does so move it back one if end of word is reached and not 
            ' expanding selection.
            If NOT bExpand Then
                oTextCursor.goLeft(1, bExpand)
            End If
        End If
        End If

    ElseIf keyChar = "E" Then
        If APP() <> "CALC" Then
        oTextCursor.gotoPreviousWord(bExpand)
        oTextCursor.gotoPreviousWord(bExpand)
        oTextCursor.gotoEndOfWord(bExpand)
        End If
    ElseIf keyChar = ")" Then
        If APP() <> "CALC" Then
			oTextCursor.gotoNextSentence(bExpand)
        End If
    ElseIf keyChar = "(" Then
        If APP() <> "CALC" Then
			oTextCursor.gotoPreviousSentence(bExpand)
        End If
    ElseIf keyChar = "}" Then
        If APP() <> "CALC" Then
			oTextCursor.gotoNextParagraph(bExpand)
        End If
    ElseIf keyChar = "{" Then
        If APP() <> "CALC" Then
			oTextCursor.gotoPreviousParagraph(bExpand)
        End If

    Else
        bSetCursor = False
        bMatched = False
    End If

    ' If oTextCursor was moved, set global cursor to its position
    If APP() <> "CALC" Then

	If bSetCursor Then
		getCursor().gotoRange(oTextCursor.getStart(), False)

		' ---- REALLY BAD HACK
		' I can't seem to get the View Cursor (getCursor()) to update its
		' position without calling its own movement functions.
		' Theoretically, the above call to gotoRange should work, but I don't
		' know why it doesn't. Visually it works, but its X position is reset
		' when you move lines. Bug??

		' dim oTempPos
		' oTempPos = getCursor().getPosition()
		' ' Move left 1 and then right 1 to stay in same position
		' getCursor().goLeft(1, False)
		' If Not samePos(oTempPos, getCursor().getPosition()) Then
		' 	getCursor().goRight(1, False)
		' End If
	End If


	' If oTextCursor was moved and is in VISUAL mode, update selection
	If bSetCursor and bExpand then
		thisComponent.getCurrentController.Select(oTextCursor)
	End If

    End If

    ProcessMovementKey = bMatched
End Function

Function APP() as String
	If thisComponent.VBAGlobalConstantName = "ThisExcelDoc" Then
		APP = "CALC"
	ElseIf thisComponent.VBAGlobalConstantName = "ThisWordDoc" Then
		APP = "WRITER"
	Else
		' Drawing, Presentations, etc
		APP = "UNK"
	End If
End Function

Sub sStartViewEventListener
	sStopViewEventListener()
	oListener = CreateUnoListener("VEListener_", "com.sun.star.document.XEventListener")
	dim oGlobalEventBroadcaster
	oGlobalEventBroadcaster = GetDefaultContext().getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
	oGlobalEventBroadcaster.addEventListener(oListener)
End Sub


Sub sStopViewEventListener
	dim oGlobalEventBroadcaster
	oGlobalEventBroadcaster = GetDefaultContext().getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
	oGlobalEventBroadcaster.removeEventListener(oListener)
End Sub


Sub VEListener_notifyEvent(o)
    ' Exit if plugin is not enabled
    If MODE = M_DISABLED Then
        KeyHandler_KeyPressed = False
        Exit Sub
    End If
    
	If o is Nothing Then
		Exit Sub
	End If
	dim oSource as object
	oSource = o.Source
	If oSource is Nothing Then
		Exit Sub
	End If	
    If APP() = "CALC" Then
		dim bHasInterface
		bHasInterface = HasUnoInterfaces(oSource, "com.sun.star.text.XTextDocument")
		If not bHasInterface Then	
			Exit Sub
		End If
	End If
	dim oController as object
	If o.EventName = "OnFocus" Then
		reinitVibreoffice()
	ElseIf o.EventName = "OnViewCreated" Then
		oController = oSource.getCurrentController()
		If not (oController is Nothing) Then
			oController.addKeyHandler(oXKeyHandler)
		End If
	ElseIf o.EventName = "OnViewClosed" Then
		oController = oSource.getCurrentController()
		If not (oController is Nothing) Then
			oController.removeKeyHandler(oXKeyHandler)
		End If
	End If
End Sub


sub VEListener_disposing()
end sub


Sub reinitVibreoffice
    dim oTextCursor, oCurrentController
    oCurrentController = getCurrentController()
    If oCurrentController is Nothing Then
    	Exit Sub
    End If

    resetMultiplier()
    setCursor()
    setTextCursor()
    gotoMode(M_NORMAL)

    ' Show terminal cursor
    oTextCursor = getTextCursor()
    If not (oTextCursor Is Nothing) Then
        cursorReset(oTextCursor)
    End If
End Sub

Sub startVibreoffice()
	If not VIBREOFFICE_STARTED Then
    	sStartXKeyHandler()
    
    	VIBREOFFICE_STARTED = True
		gotoMode(M_NORMAL)
		
		oCurrentFrame = StarDesktop().getCurrentFrame()
    End If
    
    reinitVibreoffice()
    getCurrentController().addKeyHandler(oXKeyHandler)
End Sub


Sub stopVibreoffice()
    restoreStatus()
    getCurrentController().removeKeyHandler(oXKeyHandler)
End Sub


Sub toggleVibreoffice()
    if MODE = M_DISABLED then
    	gotoMode(OLD_MODE)
    else
	    OLD_MODE = MODE
	    gotoMode(M_DISABLED)
    end if
End Sub
