;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Windows XP, Windows Vista, Windows 7
; Thunderbird:	  1.x, 8.x, 9.x, 10.x, 11.x, 12.x
; Author:         Gavin Quast (code@gavinquast.de)
; Version:	  0.1
;
; Script Function:
;   Export Thunderbird Addressbooks under Windows.
;
; Copyright (C) 2012  Gavin Quast
;
; This AutoIt source file is free software: you can redistribute it
; and/or modify ; it under the terms of the GNU General Public License
; as published by the Free Software Foundation, either version 3 of
; the License, or (at your option) any later version.
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU General Public License for more details.
;
; You should have received a copy of the GNU General Public License
; along with this program.  If not, see <http://www.gnu.org/licenses/>.
;
; AutoIt is no longer under GPL. See more Information here:
; http://www.autoitscript.com/autoit3/docs/license.htm

#include <File.au3>
#include <Array.au3>

; =================================================================================================
; Start setting up a lot of variables. If you want to use it in full auto mode the
; most can be deleted.

; Set Version
Local $VERSION = "0.1"

; Set Export done variable
Local $EXPORTDONE = "FALSE"

; Set message box text and title
Local $MSGTITLE = "Thunderbird Addressbook Export " & $VERSION
Local $MSGIT = "Please call the IT-Hotline or run arround with your hair on fire!"
Local $MSGNOTB = "Thunderbird is not installed on your system."
Local $MSGERROR = "Error during addressbook export."
Local $MSGTIMEOUT = "Timeout error during export."
Local $MSGNOPROFILE = "No Thunderbird profile was found on your system."
Local $MSGNOBOOK = "No addressbooks found on your system."
Local $MSGSTART = "Thunderbird addressbook export starts right now."
Local $MSGNOINPUT = "Please no keyboard or mouse interaction during the export."
Local $MSGEXPORT1 = "1 Addressbook exported."
Local $MSGEXPORTX = " Addressbooks exported."

; Set traytip text
Local $TTTEXT = "Export addressbook "

; Set time info for suicide function
Local $MSGTIME = "Export can take up to 1 minute."

; Set locale specific window titles
Local $TBAPPWINDOW = "Adressbuch"
Local $TBAPPEXPORT = "Adressbuch exportieren"

; Set Thunderbird profile path and executable
Local $TBPATH = @AppDataDir &"\Thunderbird\Profiles\"
Local $TBPROFILE = _FileListToArray($TBPATH, "*.default", 2)
Local $TBEXE = @ProgramFilesDir & "\Mozilla Thunderbird\Thunderbird.exe"
Local $TBSWITCH = " -addressbook"

; Set path for export directory and filename
Local $EXPORTDIR = @DesktopDir & "\Thunderbird-Export\"
Local $EXPORTFILE = "TBAddressbook-"

; Get screen resolution
Local $SCREENRES = _GetTotalScreenResolution()

; =================================================================================================
; Now after we set up everything we start checking some conditions.

; Check if Thunderbird is installed
If Not FileExists($TBEXE) Then
   MsgBox (48,$MSGTITLE, $MSGNOTB & @CRLF & $MSGERROR & @CRLF & @CRLF & $MSGIT)
   Exit
EndIf

; Check if there is a default profile
If Not IsArray($TBPROFILE) Or $TBPROFILE[0] = 0 Then
   MsgBox (48,$MSGTITLE, $MSGNOPROFILE & @CRLF & $MSGERROR & @CRLF & @CRLF & $MSGIT)
   Exit
EndIf

; Count number of addressbooks 
Local $COUNT = _FileListToArray($TBPATH & $TBPROFILE[1], "*ab*.mab", 1)
If Not IsArray($COUNT) Or $COUNT[0] = 0 Then
   MsgBox (48,$MSGTITLE, $MSGNOBOOK & @CRLF & $MSGERROR & @CRLF & @CRLF & $MSGIT)
   Exit
EndIf

; =================================================================================================
; Just start anything

; Some error handling
If @error then 
  Msgbox (48,$MSGTITLE,$MSGERROR & @CRLF & @CRLF & $MSGIT & @CRLF & "Error Code: " & @error)
  Exit
EndIf

; Create export dir if not exists
DirCreate($EXPORTDIR)

; Show information befor export starts
If Not Msgbox (1,$MSGTITLE,$MSGSTART & @CRLF & $MSGNOINPUT & @CRLF & @CRLF & $MSGTIME) = "1" Then
   Exit
EndIf

; Start Thunderbird in addressbook mode
Run($TBEXE & $TBSWITCH)

; Wait for the window and move it away
WinWaitActive($TBAPPWINDOW)
WinMove($TBAPPWINDOW,"",$SCREENRES[0],$SCREENRES[1])

; Register suicide function
AdlibRegister("_Suicide", 45*1000)

; Loop for suicide function
While 1

; =================================================================================================
; Lets start to export something.

; Export addressbooks
TrayTip($MSGTITLE, $TTTEXT & "1 of " & $count[0], 20, 1)
$ADDRESSBOOK = $EXPORTDIR & $EXPORTFILE & "1"
Send("!x")
Send("x")
WinWaitActive($TBAPPEXPORT)
WinMove($TBAPPEXPORT,"",$SCREENRES[0],$SCREENRES[1])
Sleep(250)
ControlSetText($TBAPPEXPORT,"","[CLASS:Edit; INSTANCE:1]",$ADDRESSBOOK)
Sleep(100)
Send("{ENTER}")
;ControlClick($TBAPPEXPORT,"","[CLASS:Button; INSTANCE:2]")
WinWaitActive($TBAPPWINDOW)

If $count[0] = 1 Then
   WinKill($TBAPPWINDOW)
   Sleep(250)
   Local $EXPORTDONE = "TRUE"
   ; Abschlussmeldung
   Msgbox (64, $MSGTITLE, $MSGEXPORT1)
   Exit
Else
   ; For Thunderbird 1.x you need just two TAB actions!
   Send("{TAB}{TAB}{TAB}")
   For $i = 2 to ($count[0])
	  TrayTip($MSGTITLE, $TTTEXT & $i & " of " & $count[0], 20, 1)
	  $ADDRESSBOOK = $EXPORTDIR & $EXPORTFILE & $i
	  Sleep(2000)
	  Send("{LEFT}")
	  Sleep(100)
	  Send("{DOWN}")
	  Send("!x")
	  Send("x")
	  WinWaitActive($TBAPPEXPORT)
	  Sleep(250)
	  ControlSetText($TBAPPEXPORT,"","[CLASS:Edit; INSTANCE:1]",$ADDRESSBOOK)
	  Sleep(100)
	  Send("{ENTER}")
	  ;ControlClick($TBAPPEXPORT,"","[CLASS:Button; INSTANCE:2]")
	  WinWaitActive($TBAPPWINDOW)
	  Sleep(250)
   Next
   WinKill($TBAPPWINDOW)
   Sleep(250)
   Local $EXPORTDONE = "TRUE"
   ; Show finish message dialog
   Msgbox (64,$MSGTITLE,$count[0] & $MSGEXPORTX)
   Exit
EndIf

WEnd

; =================================================================================================
; Functions used in the script.

; Function for suicide
Func _Suicide()
	  WinKill($TBAPPWINDOW)
	  If $EXPORTDONE = "TRUE" Then
		 Exit
	  Else
		 MsgBox (48,$MSGTITLE,$MSGTIMEOUT & @CRLF & @CRLF & $MSGIT)
		 Exit
	  EndIf
EndFunc

;Function for screen resolution
Func _GetTotalScreenResolution()
	Local $aRet[2]
	Global Const $SM_VIRTUALWIDTH = 78
	Global Const $SM_VIRTUALHEIGHT = 79
	$VirtualDesktopWidth = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_VIRTUALWIDTH)
	$aRet[0] = $VirtualDesktopWidth[0] - 1
	$VirtualDesktopHeight = DllCall("user32.dll", "int", "GetSystemMetrics", "int", $SM_VIRTUALHEIGHT)
	$aRet[1] = $VirtualDesktopHeight[0] - 1
	Return $aRet
EndFunc
