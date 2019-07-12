#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=greatestUnitTestRunner.ico
#AutoIt3Wrapper_Outfile=greatestUnitTestRunner.exe
#AutoIt3Wrapper_Outfile_x64=greatestUnitTestRunner_x64.exe
#AutoIt3Wrapper_Compile_Both=n
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=Tool to manage and execute SUITE/TEST of a given test-runner, created by the Greatest Unit-Test Framework
#AutoIt3Wrapper_Res_Fileversion=0.3.0.0
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ***************************************************************************************************************************
Description:
	Tool to manage and execute SUITE of a given test-runner, created by the Greatest Unit-Test Framework.

Parameter:

Problems:

Refactor:
	- Refactor complete code to minimize
	- Reporting with CSS|HTML templates

Ideas/TODO:
	- Menu items (File-Open|Recent|Exit, Help-About|Help, ...)
	- Command line arguments for automated execution (Build server)
	- Allow selection of TEST per SUITE?
	- TEST|SUITE Shuffle (Seed by Framework)
	- TEST|SUITE Execution order definition
	- Order HTML report SUITE results by FAIL/SKIP/PASS
	- CSS Styling by command line style delivery? (Default style shall remain as currently exists, command overwrites)
	- Log file(s) (usable to see progress, debugging, ...)
		- Level: Debug (mainly interesting for developer, all necessary information, technical information)
		- Level: Execution (interesting for user, general information)
	- IDE integration --> configuration how to execute tests, for example:
		- onsave 	Compile unit-test runner and execute it after file save
		- onexec	Compile unit-test runner and execute it after user command (from IDE, command line, ...)
		... Related to commands above:
			- mysuite		Only the SUITE of the current module, where the file is executed
			- all			All SUITEs ...
				- sre		... in SRE mode
				- alone		... All SUITEs in standalone execution one-by-one

History (current on top):
	v0.4.0.0	- Add: Command line arguments to allow execution automation (WIP)
	v0.3.0.0	- Change: Comment-out of "_WinAPI_ChangeWindowMessageFilterEx()" because of Avira AntiVir warning
				- Change: Removed external "TreeStateViewLib.au3) because of no tristate usage at the moment
				- Change: Replaced string "All" with "SRE" when checkbox for SRE is checked
				- Change: "-" will be shown instead of "0" when no suite and no SRE checkbox are checked
				- Add: x64 compile options
				- Add: Single click on treeview item toggles its checked state
	v0.2.0.0	- Change: Refactoring to use "TristateTreeViewLib" UDF instead of "GuiTreeViewEx" UDF --> Resolves previous
					problem with GuiTreeViewEx (Checkbox always iterates over TriState, when just shall be checked/unchecked
					(TriState only for parent checkboxes when not all child checkboxes are selected))
				- Add/Change: Small GUI re-design
					- Add: Progress bar to show execution progress
					- Add: Checkbox "Open Report when finished"
					- Add: Checkbox "Single Runner Execution"
					- Add: Progress Bar
					- Add: Status Bar
				- Add: "Single Runner Execution" or selection of checked SUITEs allowed (not both)
				- Add: Check whether the dropped file is a GREATEST test runner (by evaluating -h result)
				- Add: Allow additional Drag/Drop after initial one done (move drop area to botton)
				- Change: Exec button only enabled when at least one SUITE or "Single Runner Execution" checked
				- Change: StatusBar as replacemenent of TrayTip
	v0.1.0.0	- Add: Listing of all SUITE definitions
				- Add: TreeView to select SUITE to execute
				- Add: HTML Report creation after SUITE execution
				- Add: Console mode (GUI created but no shown) to run all SUITEs separately
				- Add: Unify results -> Greatest 1.4.0 checks for Suite name by "starts-with" instead of "equals", which
					   results in execution of all SUITEs starting with "str1" also when name is "str1blabla"
#ce ***************************************************************************************************************************
#include <File.au3>
#include <Array.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <WinAPISysWin.au3>
#include <GuiStatusBar.au3>
OnAutoItExitRegister('OnAutoItExit')

#cs**************************************************************************************************************
* Declarations
#ce**************************************************************************************************************
; Application
Local $sTitle = "GREATEST Unit-Test Runner (v" & FileGetVersion(@ScriptFullPath) & ")"

; Working with paths
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = "", $sWorkingDir = "", $sPathToTestRunner = ""
Local $sFileGreatestTestRunner = "", $sFileTestRunnerOutput = "testRunnerResult.txt", $sFileGreatestTestReport = ""
Local $arrCmdLine[2]

; GUI
Local $hForm = 0
Local $tvSuiteTests = 0;, $sStateIconFile = "..\inc\TriStateStuff\modern.bmp"
Local $guiSuiteTestWidth = 500, $guiSuiteTestHeight = 600
Local $guiDropWidth = 400, $guiDropHeight = 100, $lblDropArea = 0, $lblDropIdText = 0, $lblDropWidth = 112, _
	$lblDropHeight = 12, $lblDropText = "Drop Test-Runnter here"
Local $btnWidth = 50, $btnHeight = 30
Local $btnExec = 0, $btnCheckAll = 0, $btnCheckNone = 0
Local $btnExecCaption = "Exec", $btnCheckAllCaption = "All", $btnCheckNoneCaption = "None"
Local $cbSingleRunnerExec = 0, $cbSingleRunnerExecCaption = "Single Runner Execution", $cbSingleRunnerExecWidth = 135
Local $cbOpenReportWhenFinished = 0, $cbOpenReportWhenFinishedCaption = "Show Report", $cbOpenReportWhenFinishedWidth = 80
Local $prgProgress = 0
Local $lblProgress = 0
Local $iSpaceMainFrame = 10, $iSpaceElements = 2
Local $g_hLabel = 0
Local Enum $eSB_PART_STATE, $eSB_PART_CURRENT_STEP, $eSB_PART_TOTAL_STEP, $eSB_PART_COUNT
Local $sbStatusbar = 0, $arrStatusbarParts[$eSB_PART_COUNT] = [400, 450, 500]
Local Enum $eGUI_STATE_READY, $eGUI_STATE_EXEC
Local $eGuiState = $eGUI_STATE_READY

; Working with SUITE/TEST
Local $arrTvSuiteItems[0]

; Report
Local $bSuiteStart = False
Local $bAnySuiteChecked = False

#cs**************************************************************************************************************
* Processing
#ce**************************************************************************************************************
; At the beginning, because GUI elements are used at command line processing too (checkboxes for SUITEs)
createGuiToDropRunner()

; Processing of command line arguments (tbd)
If 1 == $CmdLine[0] Then
	$arrCmdLine[0] = 1
;~ 	$arrCmdLine[1] = "C:\Users\Taneeda\Google Drive\dev\Au3\GreatestUnitTestRunner\mwccm_unittesting.exe" ; Just for testing
	$arrCmdLine[1] = $CmdLine[1]
	dropFiles($arrCmdLine)
	checkAll()
	runnerExec()
	openReport()
Else
	; Regular working code
	GUISetState(@SW_SHOW, $hForm)
	While True
		$idGuiMsg = GUIGetMsg()
		$bCbAdj = False
		$bAnySuiteChecked = False

		; In case of checked SUITE, uncheck "single runner" checkbox
		$iTotalStep = 0
		For $i In $arrTvSuiteItems
			; Toggle state on click
			If $idGuiMsg == $i Then
				_GUICtrlTreeView_SetChecked($tvSuiteTests, $i, Not _GUICtrlTreeView_GetChecked($tvSuiteTests, $i))
			EndIf

			; Check for checked items
			If _GUICtrlTreeView_GetChecked($tvSuiteTests, $i) Then
				$iTotalStep += 1
				$bAnySuiteChecked = True
				If $idGuiMsg == $i Then
					$bCbAdj = True
				EndIf
			EndIf
		Next
		If 0 == $iTotalStep Then
			$iTotalStep = "-" ; To show char in statusbar instead of 0
		EndIf
		If $bCbAdj And $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) Then
			GUICtrlSetState($cbSingleRunnerExec, $GUI_UNCHECKED)
		EndIf
		setStatusBarInfoBySingleRunnerExec(bIsCheckBoxSingleRunnerExecChecked())

		; GUI Elements
		Switch $idGuiMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $btnExec
				$eGuiState = $eGUI_STATE_EXEC
			Case $btnCheckAll
				checkAll()
			Case $btnCheckNone
				checkNone()
			Case $cbSingleRunnerExec
				; When "single runner" checkbox checked, disable all SUITEs
				If $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) Then
					checkNone()
				EndIf
		EndSwitch

		; Exec-Button only available when any SUITE or "Single Runner" checked
		If(		($bAnySuiteChecked) _
			Or 	($GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec)) _
		) Then
			If $GUI_DISABLE == BitAND(GUICtrlGetState($btnExec), $GUI_DISABLE) Then
				GUICtrlSetState($btnExec, $GUI_ENABLE)
			EndIf
		Else
			If $GUI_ENABLE == BitAND(GUICtrlGetState($btnExec), $GUI_ENABLE) Then
				GUICtrlSetState($btnExec, $GUI_DISABLE)
			EndIf
		EndIf

		; Update StatusBar / GUI-State
		Switch $eGuiState
			Case $eGUI_STATE_READY
				GUICtrlSetData($prgProgress, 0)

				$sText = "Ready"
				If 0 <> StringCompare(_GUICtrlStatusBar_GetText($sbStatusbar, $eSB_PART_STATE), $sText) Then
					_GUICtrlStatusBar_SetText($sbStatusbar, $sText, $eSB_PART_STATE)
				EndIf
				$sText = "---"
				If 0 <> StringCompare(_GUICtrlStatusBar_GetText($sbStatusbar, $eSB_PART_CURRENT_STEP), $sText) Then
					_GUICtrlStatusBar_SetText($sbStatusbar, $sText, $eSB_PART_CURRENT_STEP)
				EndIf
			Case $eGUI_STATE_EXEC
				$sText = "Processing..."
				If 0 <> StringCompare(_GUICtrlStatusBar_GetText($sbStatusbar, $eSB_PART_STATE), $sText) Then
					_GUICtrlStatusBar_SetText($sbStatusbar, $sText, $eSB_PART_STATE)
				EndIf
				runnerExec()
				$eGuiState = $eGUI_STATE_READY
		EndSwitch
	WEnd
EndIf

#cs**************************************************************************************************************
* Functions
#ce**************************************************************************************************************
Func bIsGreatestTestRunner($sFile)
	$bIsGreatestTestRunner = False

	If 		0 == StringCompare(StringRight($sFile, 4), ".exe") _
		And FileExists($sFile) _
	Then
		_PathSplit($sFile, $sDrive, $sDir, $sFileName, $sExtension)
		$sShortPath = FileGetShortName($sDrive & $sDir)
		$sFileRunner = $sShortPath & $sFileName & $sExtension
		$sHelpFile = $sShortPath & $sFileName & "_help.txt"
		$sRunCmd = @ComSpec & " /c " & $sFileRunner & " -h > " & $sHelpFile

		ConsoleWrite("$sHelpFile = " & $sHelpFile & @CRLF)
		ConsoleWrite("bIsGreatestTestRunner($sFile = " & $sFile & ")")
		ConsoleWrite("bIsGreatestTestRunner() @ $sRunCmd = " & $sRunCmd & @CRLF)

 		RunWait($sRunCmd, $sWorkingDir, @SW_HIDE)

		$fh = FileOpen($sHelpFile)
		If Not @error Then
			$content = FileRead($fh)
			ConsoleWrite("@error = " & @error & @CRLF)
			ConsoleWrite($sFileRunner & " = " & $content & @CRLF)
			If 0 < StringInStr($content, $sFileRunner) Then $bIsGreatestTestRunner = True
		Else
			MsgBox(16,"bIsGreatestTestRunner()","Can't open file: """ & $sFile & """")
		EndIf
		FileClose($fh)
		FileDelete($sHelpFile)
	EndIf

	Return $bIsGreatestTestRunner
EndFunc

Func bIsCheckBoxSingleRunnerExecChecked()
	Return $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec)
EndFunc

Func setStatusBarInfoBySingleRunnerExec($bIsCheckBoxSREChecked)
	If $bIsCheckBoxSREChecked Then
		$sText = "SRE"
		If 0 <> StringCompare(_GUICtrlStatusBar_GetText($sbStatusbar, $eSB_PART_TOTAL_STEP), $sText) Then
			_GUICtrlStatusBar_SetText($sbStatusbar, $sText, $eSB_PART_TOTAL_STEP)
		EndIf
	Else
		If 0 <> StringCompare(_GUICtrlStatusBar_GetText($sbStatusbar, $eSB_PART_TOTAL_STEP), $iTotalStep) Then
			_GUICtrlStatusBar_SetText($sbStatusbar, $iTotalStep, $eSB_PART_TOTAL_STEP)
		EndIf
	EndIf
EndFunc

Func checkAll()
	GUICtrlSetState($cbSingleRunnerExec, $GUI_UNCHECKED)
	For $i In $arrTvSuiteItems
		_GUICtrlTreeView_SetChecked($tvSuiteTests, $i, True)
	Next
EndFunc

Func checkNone()
	GUICtrlSetState($cbSingleRunnerExec, $GUI_CHECKED)
	For $i In $arrTvSuiteItems
		_GUICtrlTreeView_SetChecked($tvSuiteTests, $i, False)
	Next
EndFunc

Func extractSuite($sFile, ByRef $arrTestSuites, ByRef $sSubReport, $sSuite = "")
	$fh = FileOpen($sFile)
	If Not @error Then
		$sSubReport = ""
		$line = FileReadLine($fh)
		While Not @error
			If 0 < StringLen($line) Then
				If $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) Then
					If 0 == StringCompare(StringLeft($line, 7), "* Suite") Then
						If 0 < StringLen($sSubReport) Then
							_ArrayAdd($arrTestSuites, StringStripWS($sSubReport, 1+2))
						EndIf
						$sSubReport = ""
					EndIf
					$sSubReport = $sSubReport & $line & @CRLF
				Else
					If 0 == StringCompare(StringLeft($line, 7), "* Suite") Then
						If		$GUI_UNCHECKED == GUICtrlRead($cbSingleRunnerExec) _
							And 0 == StringCompare(StringStripWS(StringReplace(StringReplace($line, "* Suite", ""), ":", ""), 1+2), StringStripWS($sSuite, 1+2), True) _
						Then
							$bSuiteStart = True
						Else
							If $bSuiteStart Then
								_ArrayAdd($arrTestSuites, StringStripWS($sSubReport, 1+2))
								$sSubReport = ""
								$bSuiteStart = False
							EndIf
						EndIf
					EndIf

					If $bSuiteStart Then
						$sSubReport = $sSubReport & $line & @CRLF
					EndIf
				EndIf
			EndIf

			$line = FileReadLine($fh)
		WEnd
		If $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) Then
			If 0 < StringLen($sSubReport) Then
				_ArrayAdd($arrTestSuites, StringStripWS($sSubReport, 1+2))
			EndIf
		Else
			If $bSuiteStart Then
				_ArrayAdd($arrTestSuites, StringStripWS($sSubReport, 1+2))
				$sSubReport = ""
				$bSuiteStart = False
			EndIf
		EndIf
		FileClose($fh)
		FileDelete($sFile)
	Else
		MsgBox(16, "Can't open file...", $sFile)
	EndIf
EndFunc

Func runnerExec()
	Local $sReport = "", $sSubReport = "", $sTestReport = "", $sStatsReport = ""
	Local $iNumSuites = 0, $iNumTests = 0, $iNumPass = 0, $iNumSkip = 0, $iNumFail = 0
	Enum $STAT_NUM_TESTS, $STAT_NUM_PASS, $STAT_NUM_SKIP, $STAT_NUM_FAIL, $STAT_COUNT
	Local $arrSuiteStats[0][$STAT_COUNT]
	Local $arrTestSuites[0]
	Local $arrSuitesExec[0]

	GUICtrlSetData($prgProgress, 0)

	; Get all marked SUITEs
	For $i In $arrTvSuiteItems
		If 		_GUICtrlTreeView_GetChecked($tvSuiteTests, $i) _
			Or	$GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) _
		Then
			_ArrayAdd($arrSuitesExec, GUICtrlRead($i, True)) ; Returns text of the treeviewitem
		EndIf
	Next

	; Reset statistics
	$iNumSuites = 0
	$iNumTests = 0
	$iNumPass = 0
	$iNumSkip = 0
	$iNumFail = 0
	$sReport = ""

	If $GUI_CHECKED == GUICtrlRead($cbSingleRunnerExec) Then
		_PathSplit($sFileGreatestTestRunner, $sDrive, $sDir, $sFileName, $sExtension)
		$sFile = $sWorkingDir & "SRE" & "_" & $sFileName & ".txt"

		$sCmd = @ComSpec & " /c " & FileGetShortName($sFileGreatestTestRunner) & " -v > " & $sFile
		ConsoleWrite($sCmd & @CRLF)
		RunWait($sCmd, $sWorkingDir, @SW_HIDE)

		extractSuite($sFile, $arrTestSuites, $sSubReport)

		GUICtrlSetData($prgProgress, 100)
		_GUICtrlStatusBar_SetText($sbStatusbar, UBound($arrSuitesExec), $eSB_PART_CURRENT_STEP)
	Else
		; Execute all marked SUITEs
		For $i=0 To UBound($arrSuitesExec)-1
			$sFile = $sWorkingDir & "SUITE" & "_" & $arrSuitesExec[$i] & ".txt"

			$sCmd = @ComSpec & " /c " & FileGetShortName($sFileGreatestTestRunner) & " -v -s " & $arrSuitesExec[$i] & " > " & $sFile
			ConsoleWrite($sCmd & @CRLF)
			RunWait($sCmd, $sWorkingDir, @SW_HIDE)

			extractSuite($sFile, $arrTestSuites, $sSubReport, $arrSuitesExec[$i])

			GUICtrlSetData($prgProgress, (100 * ($i+1) / UBound($arrSuitesExec)))
			_GUICtrlStatusBar_SetText($sbStatusbar, $i+1, $eSB_PART_CURRENT_STEP)
		Next
	EndIf

	; Check the results and create the report
	For $i=0 To UBound($arrTestSuites)-1
		$sInit = ""
		For $num = 0 To $STAT_COUNT-1
			If 0 == StringLen($sInit) Then
				$sInit &= "0"
			Else
				$sInit &= "|0"
			EndIf
		Next
		_ArrayAdd($arrSuiteStats, $sInit)

		; Extract PASS/FAIL info
		$fh = FileOpen($sFile)
		Local $arrSuiteData = StringSplit($arrTestSuites[$i], @CRLF)
		If Not @error Then
			$sSubReport = ""
			For $j=1 To $arrSuiteData[0]
				$line = $arrSuiteData[$j]
				If 0 < StringLen($line) Then
					If 0 == StringCompare(StringLeft($line, 7), "* Suite") Then
						; Start
						$sTestReport = ""
						$iNumSuites = $iNumSuites + 1
					ElseIf 	0 == StringCompare(StringLeft($line, 4), "PASS", True) _
						Or	0 == StringCompare(StringLeft($line, 4), "SKIP", True) _
						Or  0 == StringCompare(StringLeft($line, 4), "FAIL", True) _
						Then
						If 0 == StringCompare(StringLeft($line, 4), "PASS", True) Then
							$iNumPass = $iNumPass + 1
							$arrSuiteStats[$i][$STAT_NUM_PASS] += 1
						EndIf
						If 0 == StringCompare(StringLeft($line, 4), "SKIP", True) Then
							$iNumSkip = $iNumSkip + 1
							$arrSuiteStats[$i][$STAT_NUM_SKIP] += 1
						EndIf
						If 0 == StringCompare(StringLeft($line, 4), "FAIL", True) Then
							$iNumFail = $iNumFail + 1
							$arrSuiteStats[$i][$STAT_NUM_FAIL] += 1
						EndIf
						$iNumTests = $iNumTests + 1
						$arrSuiteStats[$i][$STAT_NUM_TESTS] += 1

						; TEST output ended and new started
						$sTestReport = $sTestReport & $line & @CRLF
						$sSubReport = $sSubReport _
							& "<div class=""testReport_" & StringLeft($line, 4) & """>" _
								& $sTestReport _
							& "</div>"
						$sTestReport = ""
					ElseIf 	0 < StringInStr($line, " tests - ") _
						And 0 < StringInStr($line, " passed") _
						And 0 < StringInStr($line, " failed") _
						And 0 < StringInStr($line, " skipped") _
						And 0 < StringInStr($line, " ticks") _
						And 0 < StringInStr($line, " sec") _
						Then
						; Summary line
					ElseIf  0 == StringCompare(StringLeft($line, 5), "Total: ") _
						And 0 < StringInStr($line, " tests") _
						And 0 < StringInStr($line, " ticks, ") _
						And 0 < StringInStr($line, " sec), ") _
						And 0 < StringInStr($line, " assertions") _
						Then
						; Total line
					ElseIf  0 == StringCompare(StringLeft($line, 5), "Pass: ") _
						And 0 < StringInStr($line, ", fail: ") _
						And 0 < StringInStr($line, ", skip: ") _
						And 0 < StringInStr($line, ".") _
						Then
						; Pass/Fail/Skip line (at last)
					Else
						; Content of FAIL
						$sTestReport = $sTestReport & $line & @CRLF
					EndIf
				EndIf
			Next

			$sStats = ""
			$sPass = (0 < $arrSuiteStats[$i][$STAT_NUM_PASS]) _
				? "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuMWMqnEsAAAB+SURBVDhPvZMxDoAwDAOzwENY+f/K1wI2auWWAI2QsHSINvXRBXN3YtPiGWqPj3l12yzH0WGXxujACOj+LuC9yzor0HAvI9DU/V7QDAVNM1OBRg9pdJ9ENyiJ1hd6AYii84ZIADT9rOFOAF7L4EkwBLr40KefCQ++wJbg7LntI3P1bEGvdhYAAAAASUVORK5CYII=""> " & $arrSuiteStats[$i][$STAT_NUM_PASS] _
				: ""
			$sSkip = (0 < $arrSuiteStats[$i][$STAT_NUM_SKIP]) _
				? "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xNkRpr/UAAABPSURBVDhPY/j//z8YM7Aq/ScFw/WBCTbt/wcPHiQJg/SA9YJMw6aAGAx2zagBg9EAYMSCMT42XC1NDCAFDxIDKM5MIALMAJlGAobo+88AAG5htjderMxYAAAAAElFTkSuQmCC""> " & $arrSuiteStats[$i][$STAT_NUM_SKIP] _
				: ""
			$sFail = (0 < $arrSuiteStats[$i][$STAT_NUM_FAIL]) _
				? "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xNkRpr/UAAABvSURBVDhPtZNBCsAwCAS9tA/ptf9/n81uaBFRqTRdmCCLGXKJqCqR7dAOzz0e+znHBryDibZg4Q18zXrBHdslfS1Aqm4QC0AWt5cLgE+w86Mgi9uLBTZVN6gFtkv6WNBgjYCP+vKZcHCArcG8p3IBTw0uCxD8kpUAAAAASUVORK5CYII=""> " & $arrSuiteStats[$i][$STAT_NUM_FAIL] _
				: ""
			$sLvl = _
				(0 < $arrSuiteStats[$i][$STAT_NUM_FAIL]) _
				? "FAIL" _
				: 	(0 < $arrSuiteStats[$i][$STAT_NUM_SKIP]) _
					? "SKIP" _
					: "PASS"
			$sReport = $sReport _
				& "<button class=""accordion"">" _
					& "<div class=""statBorder"">" _
					& 	"<div class=""suiteStat_PASS""> " & $sPass & "</div>" _
					& 	"<div class=""suiteStat_SKIP""> " & $sSkip & "</div>" _
					& 	"<div class=""suiteStat_FAIL""> " & $sFail & "</div>" _
					& "</div>" _
					& "<div class=""suiteReport"">" & $arrSuitesExec[$i] & "</div>" _
				& "</button>" _
				& "<div class=""subReport"">" _
					& $sSubReport _
				& "</div>" & @CRLF
		Else
			MsgBox(16, "Can't open file...", $sFile)
		EndIf
	Next

	; Create statistics
	$sStatsReport = _
			"<table>" _
		&		"<tr>" _
		&			"<td>Date</td>" _
		&			"<td>Suites</td>" _
		&			"<td>Tests</td>" _
		&			"<td class=""pass"">PASS</td>" _
		&			"<td class=""skip"">SKIP</td>" _
		&			"<td class=""fail"">FAIL</td>" _
		&		"</tr>" _
		&		"<tr>" _
		&			"<td>" & @MON & "/" & @MDAY & "/" & @YEAR & "</td>" _
		&			"<td>" & $iNumSuites & "</td>" _
		&			"<td>" & $iNumTests & "</td>" _
		&			"<td>" & $iNumPass & "</td>" _
		&			"<td>" & $iNumSkip & "</td>" _
		&			"<td>" & $iNumFail & "</td>" _
		&		"</tr>" _
		&	"</table>"

	; HTML formatting
	$sReport = _
			"<html>" _
		&	"<style>" _
		&		"body {" _
		&			"font-family: ""Verdana"", ""Arial"", Sans-serif;" _
		&			"font-size: 10pt;" _
		&			"padding: 5px; " _
		&			"margin: 5px; " _
		&		"}" _
		&		"table {" _
		&			"border-collapse: collapse;" _
		&			"margin-bottom: 20px;" _
		&		"}" _
		&		"table, th, td {" _
		&			"border: 1px black solid;" _
		&		"}" _
		&		"td {" _
		&			"padding: 5px;" _
		&			"text-align: center;" _
		&		"}" _
		&		"td.pass {" _
		&			"height: 24px;" _
		&			"background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwQAADsEBuJFr7QAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS4xYyqcSwAAAKZJREFUSEvtlUEKwlAMRLNQwVuIW++/9WrRifQztZO0atx14H1KkuZB/6Lm7oEdr97J2BvH6eZ2t16eO2N3GNVAB9i9C0q6BXGrXOsUcEa9S8CZ9ToEnEX/VwFH9VNB+gLBUf0Au8+Hy6zI4TrDUf2BEgDOJ70F2ScCnKpWUglAFjUrWROA96iZlC0CMEX1SrYKvmYXrILduLu//vRxxANsjbz2uj0A3GEwtrHVB1UAAAAASUVORK5CYII=') no-repeat;" _
		&			"background-position: 2px center;" _
		&			"padding-left: 28px;" _
		&		"}" _
		&		"td.skip {" _
		&			"height: 24px;" _
		&			"background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xNkRpr/UAAABkSURBVEhLY/j//z8YM7Aq/acmhpsLJti0/x88eJCqGGQm2GyQbdgUUAODfTNqAT48agFBPGoBQTxMLQCWTiiYXHGw3KgF+MTBctgsoCYetYAgBltA80ofRIAZINuoiCHm/mcAAODzNAd8KFLMAAAAAElFTkSuQmCC') no-repeat;" _
		&			"background-position: 2px center;" _
		&			"padding-left: 28px;" _
		&		"}" _
		&		"td.fail {" _
		&			"height: 24px;" _
		&			"background: url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4xNkRpr/UAAADASURBVEhLvZVLCsMwDAW9aAu9Rem29z+fi5QPg/qsKOBmYILRE34km7Teu9vu7z7T/V5/PD7LcaJ+p528TSzM0N/m+oINzipuYJYXGJxnEsx1gUliFiUhGxeYROUmEXleYJIz2epxgUmymbBWYI5Qu7BeYEbUTtALnreXDH+MqJ1gvWCE2oW1T0SymfC4gJzJVvMConKTiHxcQGIWJSHTBYTzTIJ5XsBZxQ3MdMFErynwl/rnT98efrC2iS739vYFlFbwVSp73QwAAAAASUVORK5CYII=') no-repeat;" _
		&			"background-position: 2px center;" _
		&			"padding-left: 28px;" _
		&		"}" _
		&		"" _
		&		"/* Style the buttons that are used to open and close the accordion subReport */" _
		&		".accordion {" _
		&		"  font-size: 18px;" _
		&		"  background-color: #eee;" _
		&		"  color: #444;" _
		&		"  cursor: pointer;" _
		&		"  padding: 18px;" _
		&		"  width: 100%;" _
		&		"  text-align: left;" _
		&		"  vertical-align: middle;" _
		&		"  border: none;" _
		&		"  outline: none;" _
		&		"  transition: 0.4s;" _
		&		"  overflow: hidden;" _
		&		"}" _
		&		"" _
		&		"/* Add a background color to the button if it is clicked on (add the .active class with JS), and when you move the mouse over it (hover) */" _
		&		".active, .accordion:hover {" _
		&		"  background-color: #ccc;" _
		&		"}" _
		&		"" _
		&		"/* Style the accordion subReport. Note: hidden by default */" _
		&		".subReport {" _
		&		"  display: none;" _
		&		"  overflow: hidden;" _
		&		"}" _
		&		".testReport_PASS {" _
		&			"padding: 5px;" _
		&			"background-color: #00BE00;" _
		&			"color: black;" _
		&		"}" _
		&		".testReport_SKIP {" _
		&			"padding: 5px;" _
		&			"background-color: lightgray;" _
		&			"color: black;" _
		&		"}" _
		&		".testReport_FAIL {" _
		&			"padding: 5px;" _
		&			"background-color: red;" _
		&			"color: white;" _
		&		"}" _
		&		".suiteStat_PASS {" _
		&			"float: left;" _
		&			"padding-right: 10px;" _
		&		"}" _
		&		".suiteStat_SKIP {" _
		&			"float: left;" _
		&			"padding-right: 10px;" _
		&		"}" _
		&		".suiteStat_FAIL {" _
		&			"float: left;" _
		&			"padding-right: 10px;" _
		&		"}" _
		&		".statBorder {" _
		&			"float: left;" _
		&			"width: 175px;" _
		&		"}" _
		&		".suiteReport {" _
		&			"float: left;" _
		&		"}" _
		&	"</style>" & @CRLF _
		&	"<body>" & @CRLF _
		&		"<h1>Test-Runner Report for " & $sFileGreatestTestRunner & "</h1>" & @CRLF _
		&		$sStatsReport & @CRLF _
		&		$sReport & @CRLF _
		&	"<script>" & @CRLF _
		&	"var acc = document.getElementsByClassName(""accordion"");" & @CRLF _
		&	"var i;" & @CRLF _
		&	"" & @CRLF _
		&	"for (i = 0; i < acc.length; i++) {" & @CRLF _
		&	"  acc[i].addEventListener(""click"", function() {" & @CRLF _
		&	"    /* Toggle between adding and removing the ""active"" class," & @CRLF _
		&	"    to highlight the button that controls the subReport */" & @CRLF _
		&	"    this.classList.toggle(""active"");" & @CRLF _
		&	"" & @CRLF _
		&	"    /* Toggle between hiding and showing the active subReport */" & @CRLF _
		&	"    var subReport = this.nextElementSibling;" & @CRLF _
		&	"    if (subReport.style.display === ""block"") {" & @CRLF _
		&	"      subReport.style.display = ""none"";" & @CRLF _
		&	"    } else {" & @CRLF _
		&	"      subReport.style.display = ""block"";" & @CRLF _
		&	"    }" & @CRLF _
		&	"  });" & @CRLF _
		&	"}" & @CRLF _
		&	"</script>" _
		&	"</body>" _
		&	"</html>"
	$fh = FileOpen($sFileGreatestTestReport, $FO_OVERWRITE+$FO_CREATEPATH)
	FileWrite($fh, $sReport)
	FileClose($fh)
	If WinActive($sTitle) And $GUI_CHECKED == GUICtrlRead($cbOpenReportWhenFinished) Then
		openReport()
	EndIf
EndFunc

Func openReport()
	ShellExecute($sFileGreatestTestReport)
EndFunc

Func getSuiteTestTree($sFileRunner)
	Local $hId = 0

	$sSuiteTestTree = ""
	$sFileGreatestTestRunner = $sFileRunner

	If FileExists($sFileRunner) Then
		ConsoleWrite("getSuiteTestTree($sFileRunner = """ & $sFileRunner & """)" & @CRLF)

		_PathSplit($sFileRunner, $sDrive, $sDir, $sFileName, $sExtension)

		; Shorten path for CMD execution
		$sWorkingDir = $sDrive & $sDir
		$sWorkingDir = FileGetShortName($sWorkingDir)

		; Execution paths
		$sFileRunner = $sWorkingDir & $sFileName & $sExtension
		$sFileGreatestTestReport = $sWorkingDir & $sFileName & "_report.html"
		$sResultOutput = $sWorkingDir & $sFileTestRunnerOutput
		$sRunCmd = @ComSpec & " /c " & $sFileRunner & " -l > " & $sResultOutput & ""

		; debug
		ConsoleWrite("$sWorkingDir = " & $sWorkingDir & @CRLF)
		ConsoleWrite("$sFileGreatestTestReport = " & $sFileGreatestTestReport & @CRLF)
		ConsoleWrite("$sResultOutput = " & $sResultOutput & @CRLF)
		ConsoleWrite("$sRunCmd = " & $sRunCmd & @CRLF)

 		RunWait($sRunCmd, $sWorkingDir, @SW_HIDE)

		_ArrayDelete($arrTvSuiteItems, "0-" & UBound($arrTvSuiteItems)-1)

		$fh = FileOpen($sResultOutput)
		If Not @error Then
			$line = FileReadLine($fh)
			While Not @error
				$sSearchSuite = "* Suite"
				$sSearchTest = "  "

				If 0 < StringLen($line) And 0 == StringCompare(StringLeft($line, StringLen($sSearchSuite)), $sSearchSuite) Then
					$sSuite = StringStripCR(StringStripWS(StringReplace(StringReplace($line, $sSearchSuite, ""), ":", ""), 8))
					$hId = GUICtrlCreateTreeViewItem($sSuite, $tvSuiteTests)
					If Not @error Then
						ConsoleWrite($sSuite & @CRLF)
						_ArrayAdd($arrTvSuiteItems, $hId)
					Else
						MsgBox(16,"getSuiteTestTree()","Can't add SUITE: """ & $sSuite & """")
					EndIf
;~ 				ElseIf 0 < StringLen($line) And 0 == StringCompare(StringLeft($line, StringLen($sSearchTest)), $sSearchTest) Then
;~ 					$sTest = StringStripCR(StringStripWS(StringReplace(StringReplace($line, $sSearchTest, ""), ":", ""), 8))
;~ 					If 0 == StringLen($sSuiteTestTree) Then
;~ 						$sSuiteTestTree = $sTest
;~ 					Else
;~ 						$sSuiteTestTree = $sSuiteTestTree & "|~" & $sTest
;~ 					EndIf
				EndIf

				$line = FileReadLine($fh)
			WEnd
		Else
			MsgBox(16,"getSuiteTestTree()","Can't open file: """ & $sResultOutput & """")
		EndIf
		FileClose($fh)
		FileDelete($sResultOutput)

;~ 		LoadStateImage($tvSuiteTests, $sStateIconFile)
	Else
		MsgBox(16,"Test-Runner doesn't exists...","Test-Runner doesn't exists: """ & $sFileRunner & """")
	EndIf
EndFunc

Func createGuiToDropRunner()
	$hForm = GUICreate($sTitle, $guiDropWidth, $guiDropHeight, @DesktopWidth/2-$guiDropWidth/2, @DesktopHeight/2-$guiDropHeight/2)

	$lblDropArea = GUICtrlCreateLabel("", $iSpaceMainFrame, $iSpaceMainFrame, $guiDropWidth-($iSpaceMainFrame*2), $guiDropHeight-($iSpaceMainFrame*2))
	$g_hLabel = GUICtrlGetHandle($lblDropArea)
	GUICtrlSetBkColor(-1, 0xD3D8EF)
	$lblDropIdText = GUICtrlCreateLabel($lblDropText, $guiDropWidth/2-$lblDropWidth/2, $guiDropHeight/2-$lblDropHeight/2, $lblDropWidth, $lblDropHeight)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	; The following statements cause an Avira AntiVir warning, don't know why...
;~ 	If IsAdmin() Then
;~ 		_WinAPI_ChangeWindowMessageFilterEx($g_hLabel, $WM_COPYGLOBALDATA, $MSGFLT_ALLOW)
;~ 		_WinAPI_ChangeWindowMessageFilterEx($g_hLabel, $WM_DROPFILES, $MSGFLT_ALLOW)
;~ 	EndIf

	; Register label window proc
	Global $g_hDll = DllCallbackRegister('_WinProc', 'ptr', 'hwnd;uint;wparam;lparam')
	Global $g_pDll = DllCallbackGetPtr($g_hDll)
	Global $g_hProc = _WinAPI_SetWindowLong($g_hLabel, $GWL_WNDPROC, $g_pDll)

	_WinAPI_DragAcceptFiles($g_hLabel)
EndFunc

Func dropFiles($sFileList)
	If 1 == $sFileList[0] Then
		ConsoleWrite("Dropped path: """ & $sFileList[1] & """" & @CRLF)

		If bIsGreatestTestRunner($sFileList[1]) Then
			If 0 == $tvSuiteTests Then
				WinMove($sTitle, "", @DesktopWidth/2-$guiSuiteTestWidth/2, @DesktopHeight/2-$guiSuiteTestHeight/2, $guiSuiteTestWidth, $guiSuiteTestHeight+23)

				$btnCheckAll = GUICtrlCreateButton("All", $iSpaceElements, $iSpaceElements, $btnWidth, $btnHeight)
				$btnCheckNone = GUICtrlCreateButton("None", $iSpaceElements*2+$btnWidth, $iSpaceElements, $btnWidth, $btnHeight)
				$cbSingleRunnerExec = GUICtrlCreateCheckbox($cbSingleRunnerExecCaption, $iSpaceElements*3+$btnWidth*2, $iSpaceElements, $cbSingleRunnerExecWidth, $btnHeight)
				GUICtrlSetState($cbSingleRunnerExec, $GUI_CHECKED)
				$cbOpenReportWhenFinished = GUICtrlCreateCheckbox($cbOpenReportWhenFinishedCaption, $iSpaceElements*4+$btnWidth*2+$cbSingleRunnerExecWidth, $iSpaceElements, $cbOpenReportWhenFinishedWidth, $btnHeight)
				GUICtrlSetState($cbOpenReportWhenFinished, $GUI_CHECKED)
				$prgProgress = GUICtrlCreateProgress($iSpaceElements, $guiSuiteTestHeight-59, $guiSuiteTestWidth-$btnWidth-$iSpaceElements*5, $btnHeight-2)
				$tvSuiteTests = GUICtrlCreateTreeView(0, $iSpaceElements*2+$btnHeight, $guiSuiteTestWidth-5, $guiSuiteTestHeight-$btnHeight-$iSpaceElements*2-61, BitOR($GUI_SS_DEFAULT_TREEVIEW, $TVS_CHECKBOXES))
				$btnExec = GUICtrlCreateButton("Exec", $guiSuiteTestWidth-5-$btnWidth-$iSpaceElements, $guiSuiteTestHeight-60, $btnWidth, $btnHeight)

				GUICtrlSetPos($lblDropArea, $guiSuiteTestWidth-5-$btnWidth-$iSpaceElements, $iSpaceElements, $btnWidth, $btnHeight)
				GUICtrlSetPos($lblDropIdText, $guiSuiteTestWidth-5-$btnWidth-$iSpaceElements+15, $iSpaceElements+8, $btnWidth, $btnHeight)
				GUICtrlSetData($lblDropIdText, "Drop")

				$sbStatusbar = _GUICtrlStatusBar_Create($hForm)
				_GUICtrlStatusBar_SetParts($sbStatusbar, $arrStatusbarParts)
			Else
				_GUICtrlTreeView_DeleteAll($tvSuiteTests)
			EndIf
			_GUICtrlStatusBar_SetText($sbStatusbar, "---", $eSB_PART_CURRENT_STEP)

			getSuiteTestTree($sFileList[1])
			checkNone() ;Reset to default state after extracting SUITE info
			$eGuiState = $eGUI_STATE_READY
		Else
			ConsoleWrite("No GREATEST Unit-Test Runner detected: """ & $sFileList[1] & """" & @CRLF)
		EndIf
	EndIf
EndFunc

Func _WinProc($hWnd, $iMsg, $wParam, $lParam)
    Switch $iMsg
		Case $WM_DROPFILES
			Local $sFileList = _WinAPI_DragQueryFileEx($wParam)
			If Not @error Then
				dropFiles($sFileList)
			EndIf
			_WinAPI_DragFinish($wParam)
			Return 0
    EndSwitch
    Return _WinAPI_CallWindowProc($g_hProc, $hWnd, $iMsg, $wParam, $lParam)
EndFunc

Func OnAutoItExit()
    _WinAPI_SetWindowLong($g_hLabel, $GWL_WNDPROC, $g_hProc)
    DllCallbackFree($g_hDll)
EndFunc
