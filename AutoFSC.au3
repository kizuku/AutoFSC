#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Zachary Zhao

 Script Function:
   Automates FSC Documentation Process as per WI0876
   Currently only covers the following operations:
	  Extracts FS Test Pages
	  Combines Protocol and Test Pages into a single document
	  Formats header and footer on document
   In WI0876 v4.0, these steps are:
	  13.3
	  13.8 - 13.11.1

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; Includes
#include <GUIConstantsEx.au3>
#include <FontConstants.au3>
#include <FileConstants.au3>
#include <File.au3>

Opt('MustDeclareVars', 1)

Global $fileChosen1, $fileChosen2, $baseFileChosen, $headerFileChosen, $fileExtractChosen

; Execution
main()

Func main()

   ; Possible feature: complete form fields in child popup window
   ;Local $childGUI
   ;$childGUI = GUICreate("Complete form", 600, 400)
   ;Local $testButton = GUICtrlCreateButton("Test", 550, 350)
   ;GUISetState(@SW_HIDE)

   Local $GUI, $msg
   Local $shrinkScale = ""
   Local $pageSize = ""

   $GUI = GUICreate("AutoFSC Formatting", 800, 600)
   GUISetState(@SW_SHOW)

   ; README
   Local $readmeButton = GUICtrlCreateButton("Open README", 700, 550)

   ; Operations
   Local $extractTestPagesButton = GUICtrlCreateButton("Begin extraction", 585, 10, 150, 25)
   Local $combineDocumentsButton = GUICtrlCreateButton("Begin combination", 585, 180, 150, 25)

   ; Notes
   Local $extractNoteLabel = GUICtrlCreateLabel("Note: Requires FS in PDF form and header image in JPG form.", 585, 95, 200, 80)
   GUICtrlSetFont($extractNoteLabel, 14)
   Local $combineDocumentsLabel = GUICtrlCreateLabel("Note: Protocol must end in 'protocol' and be a pdf. Test pagse must end in 'test-pages' and be a pdf.", 585, 265, 200, 100)
   GUICtrlSetFont($combineDocumentsLabel, 14)

   ; Extract FS test pages section BEGIN
   Local $wantToExtractFS = GUICtrlCreateLabel("Extract FS Test Pages and Customize Header", 5, 10, 400)
   GUICtrlSetFont($wantToExtractFS, 14, $FW_BOLD)
   Local $notApplicableExtractBox = GUICtrlCreateCheckbox("Not applicable", 410, 11)

   Local $chooseFS = GUICtrlCreateButton("Choose FS", 20, 40, 85, 25)
   Local $fsLabel = GUICtrlCreateLabel("FS file: ", 110, 42, 70)
   GUICtrlSetFont($fsLabel, 14)
   Local $fsPicked = GUICtrlCreateLabel("None", 170, 44, 600)
   GUICtrlSetFont($fsPicked, 12)
   Local $chooseHeader = GUICtrlCreateButton("Choose header image file", 20, 70, 127, 25)
   Local $headerFile = GUICtrlCreateLabel("Header Image: ", 152, 72, 120)
   GUICtrlSetFont($headerFile, 14)
   Local $chosenHeaderFile = GUICtrlCreateLabel("None", 272, 74, 500)
   GUICtrlSetFont($chosenHeaderFile, 12)
   Local $pageSizes = GUICtrlCreateLabel("Choose page size:", 20, 100, 170, 21)
   GUICtrlSetFont($pageSizes, 14, $FW_BOLD)
   Local $smallPage = GUICtrlCreateRadio("8.5 x 11", 190, 102)
   Local $largePage = GUICtrlCreateRadio("11 x 17", 260, 102)
   GUICtrlSetState($smallPage, $GUI_CHECKED)
   Local $pagesLabel = GUICtrlCreateLabel("Pages to extract: ", 20, 125, 160)
   GUICtrlSetFont($pagesLabel, 14, $FW_BOLD)
   Local $pagesInput = GUICtrlCreateInput("Relevant page numbers", 180, 126, 120)
   GUICtrlSetTip($pagesInput, "Separate pages with commas, using a dash for contiguous pages. Always include FS DCA Approval Page (Page 1).")
   ; Extract FS test pages section END

   ; Combine protocol and fs section BEGIN
   Local $wantToCombineDocs = GUICtrlCreateLabel("Combine Documents and Customize Footer", 5, 180, 390)
   GUICtrlSetFont($wantToCombineDocs, 14, $FW_BOLD)
   Local $notApplicableCombineBox = GUICtrlCreateCheckbox("Not applicable", 400, 181)

   Local $chooseFile1 = GUICtrlCreateButton("Choose protocol", 20, 210, 85, 25)
   Local $chooseFile2 = GUICtrlCreateButton("Choose FS", 20, 240, 85, 25)
   Local $protocol = GUICtrlCreateLabel("Protocol: ", 110, 212, 80)
   Local $fs = GUICtrlCreateLabel("FS Test Pages: ", 110, 242, 130)
   GUICtrlSetFont($protocol, 14)
   GUICtrlSetFont($fs, 14)
   Local $chosenProtocol = GUICtrlCreateLabel("None", 190, 214, 600)
   Local $chosenFS = GUICtrlCreateLabel("None", 240, 244, 600)
   GUICtrlSetFont($chosenProtocol, 12)
   GUICtrlSetFont($chosenFS, 12)
   ; Combine protocol and fs section END

   ; TEST begin
   ;Local $exampleButton = GUICtrlCreateButton("Open child", 700, 550)
   ; TEST end

   Local $tempDrive, $tempDir, $filename, $tempExtension
   Local $tempDrive1, $tempDir1, $file1, $tempExtension1
   Local $tempDrive2, $tempDir2, $file2, $tempExtension2
   Local $extractDrive, $extractDir, $extractFile, $extractExtension

   $pageSize = "8.5x11"
   $shrinkScale = "75%"

   While 1
	  $msg = GUIGetMsg()
	  Switch $msg
		 Case $GUI_EVENT_CLOSE
			ExitLoop

		 Case $readmeButton
			Run("explorer.exe C:\Users\zhaoz29\Documents\AutoFSC\README.txt")

		 ; Child window for filling out form
		 ;Case $exampleButton
			;GUISetState($GUI_DISABLE, $GUI)
            ;GUISetState(@SW_SHOW, $childGUI)
            ;While 1
			   ;Switch GUIGetMsg()
				  ;Case $GUI_EVENT_CLOSE
					 ;GUISetState(@SW_HIDE, $childGUI)
					 ;GUISetState($GUI_ENABLE, $GUI)
					 ;ExitLoop
				  ;Case $testButton
					 ;MsgBox("", "", "test")
                ;EndSwitch
            ;WEnd

		 ; Extract cases
		 Case $smallPage
			$pageSize = "8.5x11"
		 Case $largePage
			$pageSize = "11x17"

		 Case $notApplicableExtractBox
			If GUICtrlRead($notApplicableExtractBox) = $GUI_UNCHECKED Then
			   GUICtrlSetState($chooseFS, $GUI_SHOW)
			   GUICtrlSetState($fsLabel, $GUI_SHOW)
			   GUICtrlSetState($fsPicked, $GUI_SHOW)
			   GUICtrlSetState($chooseHeader, $GUI_SHOW)
			   GUICtrlSetState($headerFile, $GUI_SHOW)
			   GUICtrlSetState($chosenHeaderFile, $GUI_SHOW)
			   GUICtrlSetState($pagesLabel, $GUI_SHOW)
			   GUICtrlSetState($pagesInput, $GUI_SHOW)
			   GUICtrlSetState($pageSizes, $GUI_SHOW)
			   GUICtrlSetState($smallPage, $GUI_SHOW)
			   GUICtrlSetState($largePage, $GUI_SHOW)
			Else
			   GUICtrlSetState($chooseFS, $GUI_HIDE)
			   GUICtrlSetState($fsLabel, $GUI_HIDE)
			   GUICtrlSetState($fsPicked, $GUI_HIDE)
			   GUICtrlSetState($chooseHeader, $GUI_HIDE)
			   GUICtrlSetState($headerFile, $GUI_HIDE)
			   GUICtrlSetState($chosenHeaderFile, $GUI_HIDE)
			   GUICtrlSetState($pagesLabel, $GUI_HIDE)
			   GUICtrlSetState($pagesInput, $GUI_HIDE)
			   GUICtrlSetState($pageSizes, $GUI_HIDE)
			   GUICtrlSetState($smallPage, $GUI_HIDE)
			   GUICtrlSetState($largePage, $GUI_HIDE)
			EndIf

		 Case $chooseFS
			ChooseFileToExtract()
			If $fileExtractChosen = "" Then
			   GUICtrlSetData($fsPicked, "None")
			Else
			   _PathSplit($fileExtractChosen, $extractDrive, $extractDir, $extractFile, $extractExtension)
			   GUICtrlSetData($fsPicked, $extractFile & $extractExtension)
			EndIf

		 Case $chooseHeader
			ChooseHeaderFile()
			If $headerFileChosen = "" Then
			   GUICtrlSetData($chosenHeaderFile, "None")
			Else
			   _PathSplit($headerFileChosen, $tempDrive, $tempDir, $filename, $tempExtension)
			   GUICtrlSetData($chosenHeaderFile, $filename & $tempExtension)
			EndIf

		 Case $extractTestPagesButton
			Local $pages = GUICtrlRead($pagesInput)
			If NOT ($pageSize = "8.5x11") AND NOT ($pageSize = "11x17") Then
			   MsgBox("", "AutoFSC", "Please choose a valid page size")
			ElseIf ($fileExtractChosen = "" OR $headerFileChosen = "") Then
			   MsgBox("", "AutoFSC", "Please choose a valid base and header file")
			ElseIf (NOT StringRight($headerFileChosen, 10) = "header.jpg") AND (NOT StringRight($headerfileChosen, 11) = "header.jpeg") Then
			   MsgBox("", "AutoFSC", "Please choose a valid header file")
			ElseIf ($pages = "Relevant page numbers" OR $pages = "") Then
			   MsgBox("", "AutoFSC", "Please enter valid page numbers to extract")
			Else
			   extractTestPages($fileExtractChosen, $headerFileChosen, $pages, $pageSize)
			EndIf

		 ; Combine cases
		 Case $notApplicableCombineBox
			If GUICtrlRead($notApplicableCombineBox) = $GUI_UNCHECKED Then
			   GUICtrlSetState($chooseFile1, $GUI_SHOW)
			   GUICtrlSetState($chooseFile2, $GUI_SHOW)
			   GUICtrlSetState($protocol, $GUI_SHOW)
			   GUICtrlSetState($fs, $GUI_SHOW)
			   GUICtrlSetState($chosenProtocol, $GUI_SHOW)
			   GUICtrlSetState($chosenFS, $GUI_SHOW)
			Else
			   GUICtrlSetState($chooseFile1, $GUI_HIDE)
			   GUICtrlSetState($chooseFile2, $GUI_HIDE)
			   GUICtrlSetState($protocol, $GUI_HIDE)
			   GUICtrlSetState($fs, $GUI_HIDE)
			   GUICtrlSetState($chosenProtocol, $GUI_HIDE)
			   GUICtrlSetState($chosenFS, $GUI_HIDE)
			EndIf

		 Case $chooseFile1
			ChooseProtocol()
			If $fileChosen1 = "" Then
			   GUICtrlSetData($chosenProtocol, "None")
			Else
			   _PathSplit($fileChosen1, $tempDrive1, $tempDir1, $file1, $tempExtension1)
			   GUICtrlSetData($chosenProtocol, $file1 & $tempExtension1)
			EndIf

		 Case $chooseFile2
			ChooseFS()
			If $fileChosen2 = "" Then
			   GUICtrlSetData($chosenFS, "None")
			Else
			   _PathSplit($fileChosen2, $tempDrive2, $tempDir2, $file2, $tempExtension2)
			   GUICtrlSetData($chosenFS, $file2 & $tempExtension2)
			EndIf

		 Case $combineDocumentsButton
			If $fileChosen1 == "" OR $fileChosen2 == "" Then
			   MsgBox("", "AutoFSC", "Please choose a protocol and FS file")
			ElseIf NOT(StringRight($fileChosen1, 12) = "PROTOCOL.pdf") Then
			   MsgBox("", "AutoFSC", "Please choose a valid protocol file")
			Else
			   combineDocuments($fileChosen1, $fileChosen2)
			EndIf
	  EndSwitch
   WEnd
EndFunc

Func ChooseFileToExtract()

   $fileExtractChosen = ""
   $fileExtractChosen = FileOpenDialog("Select a file", ".", "PDF files (*.pdf)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, $FD_MULTISELECT))
   If @error Then
	  FileChangeDir(@ScriptDir)
   EndIf

EndFunc

Func ChooseHeaderFile()

   $headerFileChosen = ""
   $headerFileChosen = FileOpenDialog("Select a file", ".", "JPG files (*.jpg;*.jpeg)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, $FD_MULTISELECT))
   If @error Then
	  FileChangeDir(@ScriptDir)
   EndIf

EndFunc

; Extract FS test pages and configure header
Func extractTestPages($file, $header, $pages, $size)

   MsgBox("", "AutoFSC", "Opening up Adobe Acrobat. Your mouse will move automatically.")
   MouseMove(@DesktopWidth * 3 / 4, @DesktopHeight - 25, 10)
   Sleep(500)

   ShellExecute("Acrobat.exe", $file)

   Local $drive, $dir, $filename, $extension
   _PathSplit($file, $drive, $dir, $filename, $extension)

   Local $hDrive, $hDir, $hFilename, $hExtension
   _PathSplit($header, $hDrive, $hDir, $hFilename, $hExtension)

   WinActivate($filename & $extension & " - Adobe Acrobat")
   WinWaitActive($filename & $extension & " - Adobe Acrobat", "", 10)

   If NOT WinActive($filename & $extension & " - Adobe Acrobat") Then
	  MsgBox("", "AutoFSC", "Something went wrong while opening the file.")
	  Return
   EndIf

; 13.9 Workflow Step 7: Add JPG header to combined FSC Protocol
; 13.9.1 Adjust header margin
   ; Open "Add Header and Footer" window
   Sleep(500)
   Send("!")
   Sleep(500)
   Send("v")
   Sleep(500)
   Send("t")
   Sleep(500)
   Send("p")
   Sleep(1000)
   ControlClick($filename & $extension & " - Adobe Acrobat", "", "[TEXT:AVScrollView]", "primary", 1, 29, 497)
   Sleep(500)
   Send("a")

   ; Set header
   WinWaitActive("Add Header and Footer", "", 10)
   If Not WinActive("Add Header and Footer") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the header window.")
	  Return
   EndIf
   Sleep(500)
   Send("{TAB 6}")
   ;ControlFocus("Add Header and Footer", "", "[CLASS:RICHEDIT50W; INSTANCE:1]")
   Send("1.5")
   Send("+{TAB}")
   ;ControlFocus("Add Header and Footer", "", "[CLASS:Button; INSTANCE:6]")
   Send("{SPACE}")
   Send("s")
   Send("{ENTER}")
   Send("{TAB 15}")
   ;ControlFocus("Add Header and Footer", "", "[CLASS:Button; INSTANCE:11]")
   Send("{SPACE}")
   Send("{ENTER}")
   Sleep(1500)


   Send("!o")  ; Exit Add Header and Footer menu

   MsgBox("", "AutoFSC", "Hit enter when the header space is made.")


; 13.9.2 Insert JPG bg image file into FSC Protocol
   WinWaitActive($filename & $extension & " - Adobe Acrobat", "", 10)
   WinActivate($filename & $extension & " - Adobe Acrobat")
   Sleep(1500)
   If Not WinActive($filename & $extension & " - Adobe Acrobat") Then
	  MsgBox("", "AutoFSC", "Something went wrong when adding the header.")
	  Return
   EndIf
   Sleep(500)
   ControlClick($filename & $extension & " - Adobe Acrobat", "", "[TEXT:AVScrollView]", "primary", 1, 30, 529)
   Sleep(500)
   Send("a")

   WinWaitActive("Add Background", "", 10)
   If Not WinActive("Add Background") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the background window.")
	  Return
   EndIf
   Sleep(500)
   Send("{TAB 2}")

   ; Opens File Browse for background
   Send("{DOWN}")
   Send("{TAB 2}")
   Send("{SPACE}")
   Sleep(500)
   WinWaitActive("Open", "", 10)
   If Not WinActive("Open") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the file dialog."
	  Return
   EndIf
   Sleep(1000)

   ; Select header jpg file
   ControlFocus("Open", "", "[CLASS:Edit; INSTANCE:1]")
   Sleep(500)
   Send($hDrive & $hDir)
   Send("{ENTER}")
   Sleep(1500)
   Send($hFilename & $hExtension)
   Send("{ENTER}")
   Sleep(1000)

   ; 13.9.2.5
   Send("+{TAB 3}")
   ;ControlFocus("Add Background", "", "[CLASS:Button; INSTANCE:3]")
   Send("{SPACE}")
   Send("{ENTER}")
   Sleep(500)

   ; 13.9.2.6
   Send("+{TAB 13}")
   ;ControlFocus("Add Background", "", "[CLASS:RICHEDIT50W; INSTANCE:7]")
   Send("0.2")
   Send("{TAB 2}")
   Send("{UP}")
   Send("+{TAB 4}")
   ;ControlFocus("Add Background", "", "[CLASS:RICHEDIT50W; INSTANCE:6]")
   Sleep(500)

   If $size = "8.5x11" Then
	  Send("75")
	  Send("+5")
   Else
	  Send("40")
	  Send("+5")
   EndIf

   Send("!o") ; Press OK

   ; Test background menu
   MsgBox("", "AutoFSC", "Hit enter when the header image is added.")
   ;Return

   WinWaitActive($filename & $extension & " - Adobe Acrobat", "", 10)
   WinActivate($filename & $extension & " - Adobe Acrobat")
   Sleep(1500)
   If Not WinActive($filename & $extension & " - Adobe Acrobat") Then
	  MsgBox("", "AutoFSC", "Something went wrong when closing the background menu.")
	  Return
   EndIf
   Sleep(500)
   ;Send("^s")  ; SAVE

   Sleep(500)
   Send("!")
   Sleep(500)
   Send("f")
   Sleep(500)
   Send("p")
   WinWaitActive("Print", "", 10)
   If Not WinActive("Print") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the print menu.")
	  Return
   EndIf
   Send("a")
   Send("{TAB 2}")
   Send("{SPACE}")

   WinWaitActive("Adobe PDF Document Properties", "", 10)
   If Not WinActive("Adobe PDF Document Properties") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the properties window.")
	  Return
   EndIf
   Sleep(500)
   If ($size = "8.5x11") Then
	  Send("{ENTER}")
   Else
	  Send("{TAB 5}")
	  Send("{DOWN 6}")
	  Send("{ENTER}")
   EndIf

   WinWaitActive("Print", "", 10)
   WinActivate("Print")
   Sleep(1500)
   If Not WinActive("Print") Then
	  MsgBox("", "AutoFSC", "Something went wrong when closing the properties window.")
	  Return
   EndIf
   Sleep(500)
   Send("{TAB 4}")
   Send("{DOWN 2}")
   Send("{TAB}")
   Send($pages)
   MsgBox("", "AutoFSC", "Done. The new file will be saved in the same folder with -TEST-PAGES appended to the end.")
   Send("{ENTER}") ; Print
   WinWaitActive("Save PDF File As", "", 10)
   If Not WinActive("Save PDF File As") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the print menu.")
	  Return
   EndIf
   ControlFocus("Save PDF File As", "", "[CLASS:Edit; INSTANCE:1]")
   Send($drive & $dir)
   Send("{ENTER}")
   Sleep(1500)
   Send($filename & "-TEST-PAGES")
   Send("{ENTER}")
   Sleep(1500)
   MsgBox("", "AutoFSC", "Done.")

EndFunc

Func ChooseProtocol()

   $fileChosen1 = ""
   $fileChosen1 = FileOpenDialog("Select a file", ".", "PDF files (*.pdf)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, $FD_MULTISELECT))
   If @error Then
	  ; Display error message
	  FileChangeDir(@ScriptDir)
   EndIf

EndFunc

Func ChooseFS()

   $fileChosen2 = ""
   $fileChosen2 = FileOpenDialog("Select a file", ".", "PDF files (*.pdf)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, $FD_MULTISELECT))
   If @error Then
	  ; Display error message
	  FileChangeDir(@ScriptDir)
   EndIf

EndFunc

; Combine protocol and FS and configure footer
Func combineDocuments($protocol, $fs)

   MsgBox("", "AutoFSC", "Opening up Adobe Acrobat. Your mouse will move automatically.")
   MouseMove(@DesktopWidth * 3 / 4, @DesktopHeight - 25, 10)
   Sleep(500)

   Local $tempDrive, $tempDir, $filename, $tempExtension
   Local $fsDrive, $fsDir, $fsName, $fsExtension
   _PathSplit($protocol, $tempDrive, $tempDir, $filename, $tempExtension)
   _PathSplit($fs, $fsDrive, $fsDir, $fsName, $fsExtension)
   Local $name = $filename & $tempExtension

   ShellExecute("Acrobat.exe", $protocol)
   WinActivate($name & " - Adobe Acrobat")
   WinWaitActive($name & " - Adobe Acrobat", "", 10)
   If Not WinActive($name & " - Adobe Acrobat") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the file.")
	  Return
   EndIf

   Sleep(500)
   Send("!")
   Sleep(500)
   Send("v")
   Sleep(500)
   Send("t")
   Sleep(500)
   Send("p")
   Sleep(1000)
   Local $pos = ControlGetPos($name & " - Adobe Acrobat", "", "[TEXT:AVScrollView]")
   MouseMove($pos[0] + 50, $pos[1] + 200)
   MouseWheel("up", 5)
   Sleep(500)
   ControlClick($name & " - Adobe Acrobat", "", "[TEXT:AVScrollView]", "primary", 1, 28, 365)
   WinWaitActive("Select File To Insert", "", 10)
   Sleep(500)

   If NOT WinActive("Select File To Insert") Then
	  MsgBox("", "AutoFSC", "Something went wrong.")
	  Return;
   EndIf

   ControlFocus("Select File To Insert", "", "[CLASS:Edit; INSTANCE:1]")
   Sleep(500)
   Send($fsDrive & $fsDir)
   Send("{ENTER}")
   Sleep(1500)
   ;Send("{ENTER}")
   Send($fsName & $fsExtension)
   Send("{ENTER}")
   Sleep(1000)

   WinWaitClose("Select File To Insert")
   WinActivate("Insert Pages")
   Sleep(500)
   If NOT WinActive("Insert Pages") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the insert window.")
	  Return;
   EndIf
   Send("{TAB}")
   Send("{UP}")
   Send("{ENTER}")
   ; Send("^s") ; Save result

; 13.10 Workflow Step 8: Add footer to FSC Protocol
   Sleep(1000)
   Send("!")
   Sleep(500)
   Send("v")
   Sleep(500)
   Send("t")
   Sleep(500)
   Send("p")
   Sleep(1000)
   ControlClick($name & " - Adobe Acrobat", "", "[TEXT:AVScrollView]", "primary", 1, 29, 497)
   Sleep(500)
   Send("a")

   ; Set footer
   WinWaitActive("Add Header and Footer", "", 15)
   Sleep(500)
   If NOT WinActive("Add Header and Footer") Then
	  MsgBox("", "AutoFSC", "Something went wrong when opening the header window.")
	  Return
   EndIf
   Sleep(500)
   ; Select Page 1 of n page num format
   Send("{TAB 18}")
   Send("{SPACE}")
   Send("{TAB}")
   Send("{DOWN}")
   Send("{ENTER}")
   Sleep(500)

   ;Send("+{TAB 4}")
   ControlFocus("Add Header and Footer", "", "[CLASS:RICHEDIT50W; INSTANCE:9]")
   Sleep(500)
   Send("!i")
   Sleep(500)
   Send("+{TAB 8}")
   Send("0.25")
   Sleep(1000)


   Send("!o") ; Press OK


   ; Footer test
   ;MsgBox("", "", "Done")
   ;Return

   WinWaitActive($name & " - Adobe Acrobat", "", 10)
   WinActivate($name & " - Adobe Acrobat")
   Sleep(1500)
   If Not WinActive($name & " - Adobe Acrobat") Then
	  MsgBox("", "AutoFSC", "Something went wrong when closing the header window.")
	  Return
   EndIf
   Sleep(500)

   ; 13.11 Workflow Step 9: Save completed FSC Protocol
   ; Send("^s")
   MsgBox("", "AutoFSC", "Done. Please save the file to the proper folder with the proper name format.")

   Send("^+s")

EndFunc
