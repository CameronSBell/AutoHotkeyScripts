#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; Script to obtain meeting state and mute state from mutesync application (www.mutesync.com)
; This script will create 2 binary_sensor entities in Home Assistant; one for mute status and one for in_meeting status
;
; To obtain the status from Mutesync, a token must be obtained from the Mutesync application.
;   1. Choose mutesync preferences, authentication tab, and check "allow external app"
;   2. Open a browser and navigate to http://127.0.0.1:8249/authenticate
;   3. Copy the 16-character token and paste it into this next line between the quotes.
msToken := "NGIQYMWDLYQLIEZB"

; These settings should not change
global apiVersion := 1
global msURL := "http://127.0.0.1:8249/state"
global msTokenText := "Token " msToken
StringCaseSense Off

; End of configuration settings
; ------------------------------------------------------------------------------


SetTitleMatchMode, 2 ; 2 = a partial match on the title

; ------------------------------ MICROSOFT TEAMS - Unmute ------------------------------
+^!N::

if (IsMuted())
{
	ToggleTeamsMute()
}
return


; ------------------------------ MICROSOFT TEAMS - mute -----------------------------
+^!M::

if (!IsMuted())
{
	ToggleTeamsMute()
}
return


; ------------------------------ Functions ------------------------------

ToggleTeamsMute()
{
	WinGet, winid, ID, A	; Save the current window ID
	if !WinExist("Microsoft Teams") ;Yes, every Teams meeting has that in the title bar - even if it's not visible to you
		return
	WinActivate ; Without any parameters this activates the previously retrieved window - in this case your meeting
    Sleep, 10 ; wait a bit
	Send, ^+M   ; Teams' native Mute shortcut
    Sleep, 10 ; wait a bit
    WinActivate ahk_id %winid% ; Restore previous window focus
	return
}


IsMuted()
{
; Create oHttp object
oHttp := ComObjCreate("WinHttp.Winhttprequest.5.1")
; GET request, synchronous mode
oHttp.Open("GET", msURL, false)
; Add token header
oHttp.SetRequestHeader("Authorization", msTokenText)
; Add API version header
oHttp.SetRequestHeader("x-mutesync-api-version", apiVersion)
; send Request
try
{
	oHttp.send()
	; Wait for the response for 5 seconds
	oHttp.WaitForResponse(5)
	responseText := oHttp.responseText
	;MsgBox % "initial response: " responseText

	;Parse response for Meeting Status
	inMeetingStatusLoc := Instr(responseText, """in_meeting"":") + 13
	inMeetingStatusRaw := SubStr(responseText, inMeetingStatusLoc, 4)
	;MsgBox % "raw extract from inMeetingStatusRaw: " inMeetingStatusRaw
	if (inMeetingStatusRaw = "true")
	{
		;MsgBox In Meeting
		meetingState := "on"
	}
	else
	{
		;MsgBox Not In Meeting
		meetingState := "off"
	}

	; Parse response for Mute Status
	muteStatusLoc := Instr(responseText, """muted"":") + 8
	muteStatusRaw := SubStr(responseText, muteStatusLoc, 4)
	;MsgBox % "raw extract from muteStatusRaw: " muteStatusRaw
	if (muteStatusRaw = "true")
	{
		 ;MsgBox Mute is On
		 isMuted := True
	}
	else
	{
		 ;MsgBox Mute is Off
		 isMuted := False
	}
}
catch e
{
	return
}

return isMuted
}