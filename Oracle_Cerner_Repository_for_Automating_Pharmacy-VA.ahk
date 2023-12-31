﻿/*
Autohotkey v2 must be installed on user's computer.  TRM approved:  https://trm.oit.va.gov/ToolPage.aspx?tid=6458#  Download available:  https://www.autohotkey.com/
Stores scripted hotkeys to be used for automating processes in Oracle Cerner pharmacy applications. All concepts must consider significant variability in system lag and responsiveness and scripts should only be created that can be reliably run without consequence.

Contact: Lewis DeJaegher (lewis.dejaegher@va.gov)
Additional documentation: https://dvagov.sharepoint.com/:u:/r/sites/PharmacyCernerEHRCommunityofPractice/Shared%20Documents/Outpatient%20Prescribing%20and%20Dispensing/AutoHotKey%20Automating%20Pharmacy/How%20to%20Guide%20for%20OCRAP-VA.url?csf=1&web=1&e=QdVDxB (alt+shift+h)
Change Log:
    2023/11/01 - Added fix for inactive meds (KeyWait); adjusted CheckHP to differentiate Bill Only (MMR) v. "Bill Only" (resuspend/WQM); created function PharmacyPatientLookup(ID,Term) and incorporated into alt+LC
    2023/10/24 - Added alt+left click to OK button on new orders + refills; modified MMR window handling logic
    2023/10/19 - Added GUI to Alt+Shift+F to allow user to limit script to specific items and receive input. Added CheckHP function to evaluate and remove DoD HPs; added Alt+O hotkey to use that function and incorporated into Alt+Shift+F
    2023/10/10 - Added IP/Med Manager lookup from PC patient (click encounter) or patient list (click row). Added logic to look for "open" (no patient) MMR and handling for when it appears ctrl+p failed to clear prior patient (e.g. open History window, unsubmitted actions). Started replacing unnecessary variables (e.g. used once) with hard coded values to simplify code.
        - Ronnie Major developed launcher/updater file and iconography
    2023/10/04 - Added CMOP product line search; modified Outlook message to not restrict to HTML messages; updated MMR lookup for PC and DM to use WinClose instead of alt+c to avoid conflict with CPRSBooster; added MRN lookup from PBI report; applied fix to Alt+LClick to properly remove inactive meds if previously added
    2023/09/28 - Added package tracking search and MMR search from message center in PC. Updated help to load SPO reference page and exit hotkey to initiate email for reporting. Added block release where previously omitted. Updated message when MMR lookup fails to include copied value and clarify message.
    2023/09/15 - Applied fixes to infinite loop and failure to sort with Alt+LC (inactive meds), routing option change in Alt+Shift+F. Added alt+LC MMR lookup from V20 OP Dashboard.
    2023/08/31 - Added MMR search from Teams/Excel/Outlook, other Cerner window maximization to alt+left click; updated input block on unclaim; updated WinMatch to release blocks prior to exit; modified Exit hotkey to exit better; updated comments/descriptions
	2023/08/09 - Added MMR patient search/inactive meds hotkey; modified failure hotkey to allow user input; modified WinMatch to include time parameter (delay) and inserted into unclaim script
    2023/07/28 - Added failures and help hotkeys; WinMatch function; tooltip during script execution
    2023/07/06 - Initial release with unclaim hotkey
*/

#Requires AutoHotkey v2.0
#SingleInstance force

/************************************************************ Icon *********************************************************************************
RM added 10/6/23
Checks for icon presence in script directory, changes icon to it if exists.
****************************************************************************************************************************************************/

IconFileName := "ocrap.ico"

if (FileExist(IconFileName))
{
    IconPath := A_ScriptDir "\" IconFileName 
    TraySetIcon IconPath,,true
}


/************************************************************ WinMatch function *********************************************************************
; This is a function that may be used in hotkey to simplify code written for window handling. 
; System lag is a significant barrier and not all possible windows are encountered during development.
; This function provides a method for addressing those pitfalls.
; Parameter "Expected" is WinTitle and will match based on SetTitleMatchMode; delay is timeout period in seconds
****************************************************************************************************************************************************/

WinMatch(Expected, Delay)
{
	if !WinActive(Expected)
		{
		if !WinWaitActive(Expected, ,Delay)
			{
            BlockInput "Off"
            BlockInput "MouseMove" "Off"
			OtherWindow := MsgBox("Waiting for " Expected " window to load. Please hit OK to proceed (window is now loaded) or Cancel to exit.", "Expected Window Not Found", "OC")
			if OtherWindow = "Cancel"													    ; If user selects cancel then exit
				{
				ExitReason := ": exit requested"
                ToolTip "", , ,1                                                            ; Disable Tooltip and blocks if exist
                BlockInput "Off"    
                BlockInput "MouseMove" "Off"
				;MsgBox "Exiting script" ExitReason
				exit
				}
			Else if !WinExist(Expected)
				{
				ExitReason := ": unable to activate expected window"
                ToolTip "", , ,1                                                            ; Disable Tooltip and blocks if exist
				BlockInput "Off"
                BlockInput "MouseMove" "Off"
				MsgBox "Exiting script" ExitReason
				Exit
				}
			Else
				{
				OtherWindow := ""														    ; Reset MsgBox input
				WinActivate(Expected)
                BlockInput "On"
                BlockInput "MouseMove"
				}
			}
		}
}

/************************************************************ CheckHP function *********************************************************************
; This is a function that may be used in hotkey to remove DoD health plans during script processing to avoid unnecessary claim workload
; Script checks initial HP, then uses down key to check next HP listed. If only DoD HP is available, the DoD HP will be removed.
; Function expected to only be called when a [New Order/Refill/Bill Only/Modify] for Patient: [Patient Name]  [DOB: ] is open
;   NOTE: "Bill Only" window varies in size depending on if true bill only (MMR) or pseudo-bill only (resuspend in WQM) and also varies from other action windows
;   May consider adding a parameter or resorting to checking both locations but for now will assume Alt+Shift+F is how this will be used in WQM, and independent call will be from MMR
;   Similarly titled windows may exist (e.g. Inquire for Patient Name:) so attention to window title/context is necessary when using this function
****************************************************************************************************************************************************/

CheckHP()
{
WinActivate("for Patient:")
HP := "empty"
If WinActive("Bill Only") and PixelGetColor(27,480) = 0xFFFFFF							    ; Bill only window size varies between resuspend (WQM) and bill only (MMR) but cannot distinguish by title
	{
	MouseClick "left", 27, 480, 1, 0                                                        ; Health Plan is at 480 via MMR; 480 is F0F0F0 via WQM (background)                                                    
	}
Else if WinActive("Bill Only")											                
	{
	MouseClick "left", 27, 540, 1, 0
	}
Else
	{	
	MouseClick "left", 27, 500, 1, 0
	}
Send "{F4}"																					; This expands drop down and improves copy performance
Loop 5																						; This is to prevent an infinite loop should logic fail
	{
	A_Clipboard := ""
	Send "^c"
	Clipwait(2)
	If A_Clipboard = "(None)" OR RegExMatch(SubStr(A_Clipboard,1,1),"\D") = 1				; If selected HP is (None) or starts with alpha character, accept and move on
		{
		Send "{Enter}"
		Break
		}
	Else if RegExMatch(SubStr(A_Clipboard,1,1),"\d") = 1 AND HP = A_Clipboard				; If selected HP begins with number, it is DoD. If it is the same HP as last run, delete and move on
		{
		Send "{Backspace}"
		Send "{Enter}"
		Break
		}
	Else if RegExMatch(SubStr(A_Clipboard,1,1),"\d") = 1									; If selected HP begins with number, store it as variable and check next HP
		{
		HP := A_Clipboard
		Send "{Down}"
		}
	}
Return
}

/************************************************************ PharmacyPatientLookup function ********************************************************
; This is a function that may be used in hotkey to search for a patient in MMR or Med Manager using identifier passed to function
; This is used in Alt+Left Click to load patient profile
; Potential other uses exists that would build on the initial search - e.g. load patient by MRN, find rx for drug, do something with rx
; Find rx on amb profile may be a more straightforward approach in most cases
****************************************************************************************************************************************************/

PharmacyPatientLookup(ID,Term){
If RegExMatch(StrReplace(ID,"-"),"\D") = 1 OR ID = ""                                       ; Check that ID has only numeric characters (after removing potential hyphen) and is not empty
    {
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    MsgBox "Copied value is invalid or copy failed. Please try again.",,"T5"
    Exit
    }

If Term = "FIN" AND WinExist("PharmNet: Pharmacy Med")                                      ; Term is FIN and an open MM is available
    {
    WinActivate("PharmNet: Pharmacy Med")                                                   ; Activate MMR
    Sleep 250
    Send "!i"                                                                               ; Navigate to change the search term
    Send "{Enter}"                                                                          ; Trigger drop down
    Send "{Down}"
    Send "{Enter}"                                                                          ; Select term
    Sleep 250                                                                               
    MouseClick "left", 67,173, 1                                                            ; If user clicks in side bar, cursor won't exist in search field
    Send ID                                                                                 ; Send value
    Send "{Enter}"                                                                          ; Initiate search
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    If !WinWaitNotActive("Pharmacy Medication", ,5)											; Wait up to 5 seconds to check if window doesn't change (i.e. patient loading)
        {
        A_Clipboard := ID
        MsgBox Format("Hotkey appears to have failed, but FIN: {1} is copied. Try to paste into Med Manager search field or running hotkey again. Sorry about that.", ID),,"T5"
        }
    Exit
    }
Else if Term = "FIN"
    {
    BlockInput "Off"
    BlockInput "MouseMove" "Off"    
    MsgBox Format("Unable to find open Med Manager without a patient loaded. FIN: {1} is copied, please search manually. Thanks!", ID),,"T5"
    Exit
    }

If !WinExist("PharmNet: Retail Med")                                                        ; If no MMR is open, remove blocks and exit; window name is possibly truncated when a patient is already loaded
    {
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    Exit
    }
Else if WinExist("PharmNet: Retail Med", ,",")                                              ; Check for MMR w/o loaded patient (assume patient name has comma; could move STMM = 1 up)
    {
    WinActivate("PharmNet: Retail Med", ,",")                                               ; Activate MMR
    }
Else WinActivateBottom("PharmNet: Retail Med")                                              ; If user has multiple windows open, grab the window idle longest

WinWaitActive("PharmNet: Retail Med")
SetTitleMatchMode 1																		    ; Title must begin with- changed to differentiate MMR that has patient loaded v. not
Sleep 250
MouseClick "left", 63, 172, 1                                                               ; Placing cursor because ctrl+p doesn't work if no patient is selected and user has clicked out of search box
Send "^p"                                                                                   ; And sending ctrl+p because it seems to work most reliably to do both

Loop 8                                                                                      ; After ctrl+p, WinTitle will change, but when is variable so check a few times
    {
    Sleep 250
    } until WinActive("PharmNet: Retail Med")

If !WinActive("PharmNet: Retail Med")                                                       ; If ctrl+p did not result in WinTitle("A") beginning with PharmNet, a window may be open in MMR
    {
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    MsgBox Format("MMR may not be available for search. Please check for an open window or warning. {1}: {2} is copied, please complete search manually when ready. Thanks!", Term, ID), ,"4096 T5"
    Exit
    }
Send "{Tab}"                                                                                ; Navigate to change the search term
Send "{Enter}"                                                                              ; Trigger drop down
If Term = "MRN"                                                                             
    {
    Send "{Down 3}"                                                                         ; 3rd value in list
    Send "{Enter}" 
    Sleep 500
    }
Else if Term = "Rx"                                                                         
    {
    Send "{Down 4}"                                                                         ; 4th value in list
    Send "{Enter}"                                                                          ; Select term
    Sleep 500                                                                               
    Send "{Backspace 2}"                                                                    ; Clear prefix
    }
Send ID                                                                                     ; Send value
Send "{Enter}"                                                                              ; Initiate search
Return
}

/************************************************************ Alt+O (OK) ***************************************************************************
; This hotkey will check for presence of a DoD HP and attempt to remove or replace using CheckHP() function above. 
; NOTE: Scripts that include "!o" at these two windows may be adversely impacted. Additional controls and/or modification of (#)HotIf logic may be necessary. https://www.autohotkey.com/docs/v2/lib/_HotIf.htm
;   Expected outcome would be that initial hotkey would interrupt, this hotkey will run, then prior hotkey will resume. This may be acceptable in this instance and would just need to ensure action is not duplicated.
;   Initial testing of failures hotkey 
;   Consider using A_Priorhotkey: https://www.autohotkey.com/docs/v2/Variables.htm#PriorHotkey
; NOTE: "!o" will not trigger itself "By default, a given hotkey or hotstring subroutine cannot be run a second time if it is already running." https://www.autohotkey.com/docs/v2/misc/Threads.htm
****************************************************************************************************************************************************/

GroupAdd "HP", "Refill for Patient:"
GroupAdd "HP", "New Order for Patient:"
GroupAdd "HP", "Bill Only for Patient:"

#HotIf WinActive("ahk_group HP")															; #HotIf does not prevent other hotkeys from calling this hotkey
!o::
{
CheckHP()
Send "!o"																					; See threads documentation in header
Exit
}
#HotIf

/************************************************************ Alt + Shift + U (Unclaim) ***********************************************************************
; This hotkey is used to process unclaim actions in a batches of up to 40 rows
; Potential enhancement:
    ; Use page down to increase batch size (see eRx requests script for a way to achieve this)
    ; Utility of that change may be limited considering potential reduction in dummy claims, potential increase in "unclaim" fails, list dynamism, etc. and it may be prudent to keep max 40
**************************************************************************************************************************************************************/

!+U::                                                                               	    ; alt+shift+u (Unclaim)
{
CoordMode "Mouse", "Window"
SetTitleMatchMode 2 	                                                                    ; A window's title can contain WinTitle anywhere inside it to be a match.

Loop
{
iter := InputBox("How many unclaim actions to perform? (Max: 40)", "Batch Unclaim", "w300 h100")	; Limited by window size. May be able to increase using pg down/pg up but 40 is probably sufficient.
if iter.Result = "Cancel"											                        ; If user selects cancel then exit
    exit
if (iter.Value < 1 or iter.Value > 40) and A_Index = 3								        ; If user enters value that is out of bounds on 3rd try then exit
    exit
if iter.Value < 1 or iter.Value > 40
    continue													                            ; Restart loop if entered value is out of bounds
    {														                                ; User has entered value between 1-40, script will proceed
    BlockInput "On"												                            ; Prevent accidental input that could throw off script
    BlockInput "MouseMove"
    ToolTip "Please wait, script is executing", 881, 32, 1                                  ; Display tool tip so user knows script is running
    if WinExist("PharmNet: Claims Monitor - \\Remote")								        ; Check for open Claims Monitor
        WinActivate("PharmNet: Claims Monitor - \\Remote")							        ; Activate Claims Monitor window
	    WinMaximize("PharmNet: Claims Monitor - \\Remote")							        ; To standardize amount of rows available
        xPos := 160												                            ; Expected to be column 2 of CM since CoordMode is set to Window this should be reliable across users and computers, etc.
        yPos := 210												                            ; Expected to be row 1 of CM since CoordMode is set to Window this should be reliable across users and computers, etc.	
        i := 0													                            ; initialize counter variable
        while i < iter.Value and WinActive("PharmNet: Claims Monitor - \\Remote")		    ; check counter and active window before entering loop
        {
            MouseClick "left", xPos, yPos, 2								                ; double click on row to open claim
            WinMatch("View Claim", 4)
		    Send "!s"											                            ; alt+s to select claim status
		    Send "u"    										                            ; u to update status to unclaimed
		    Send "!o"											                            ; alt+o to select OK
            WinMatch("PharmNet: Claims Monitor - \\Remote", 4)
            yPos += 18											                            ; adjust vertical position
            i += 1												                            ; update counter
        }
    ToolTip "", , ,1                                                                        ; Disable tool tip
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    MsgBox "Please review claim status prior to submitting.",,"T4"							; alert user that process is complete
    }
exit
}
}

/************************************************************ Alt + Shift + O (Open) *********************************************************************
; This hotkey is used to open the Citrix Storefront

*************************************************************************************************************************************************************/

!+o::
{
Run ("https://ssoiaccess.valehr.cernerworks.ehr.gov/Citrix/USVADCweb/#")
MsgBox("Good luck today!","Welcome","4096 T4")
Exit
}


/************************************************************ Alt + Shift + F (Failures) *********************************************************************
; This hotkey is used to process specific failure types in the Work Queue Monitor

; Row details are compared to stored variables (e.g. failure types) and, if matched, script attempts to process, if no match, script will skip
; NOTE: an array may offer more precision (string allows script to proceed without an exact match) but strings are acceptable for the initial cohort of failure details
; Currently, a user must have semi-standard view defined. As written, first 3 columns should be Action, RxNumber, Failure Detail
; Planned enhancements:
    ; Copy row rather than single cell to allow more flexibility and allow for include/exclude criteria to be passed by user, for instance: drug name. Example scenario: user wants to process all stock kickbacks except semaglutide, so exclude semaglutide and vice versa.
*************************************************************************************************************************************************************/

#HotIf WinActive("Work Queue Monitor")
!+f::                                                                                       ; Alt+Shift+F (Failures)
{
CoordMode "Mouse", "Window"
SetTitleMatchMode 2 	                                                                    ; A window's title can contain WinTitle anywhere inside it to be a match.

; Create strings that hold failure details that this script will work on

LocalResuspend := ("Contact Support: Unable to transmit to CMOP, 015:REFRIG ITEM TO PO BOX FILL LOCALLY, 016:FLAMMABLE ITEM SHIPPING RESTRICTION, Patient mailing preference is not eligible for CMOP, Product is not available at CMOP, 007:MANUFACTURER'S BACKORDER, 001:TEMP OUT OF STOCK-OPEN MARKET, 013:TEMP OUT OF STOCK-PRIME VENDOR")

CMOPResuspend := ("039:PLEASE RESUBMIT FOR REROUTING, 020:DATABASE ERROR-PLEASE RESUBMIT")

HPIneffective := ("Health plan either ineffective or card holder identity changed")         ; This isn't necessary but think it helps with transparency and code legibility

ProdInactive := ("The product has been inactivated")

; Create variables
yPosFailureDetail := 160                                                                    ; Define y-position of Failure Detail in Work Queue Monitor (initially row 1, may increment in script)
ExpWin := ""																				; Variable used when checking active window prior to progressing
ExitReason := ""																			; Variable to populate message box triggered prior to exit
IncludeFlag := 0																			; Flag value = 1, user checked box to limit to specific items; flag value = 0, all items OK
RowCount := 0																				; Variable to allow user to define iterations
IncludeList := []																			; Array allows multiple items to be entered so building for that

; Create GUI for failure resolution routine to prompt and receive variables, wrapping as function to ensure script waits for GUI submit
GUIFail(*)
{
	FileMenu := Menu()
	FileMenu.Add "E&xit", call_reload														; Function call will reload script, hiding GUI
	HelpMenu := Menu()
	HelpMenu.Add "&About", (*) => Run("https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&e=n4xVU9&siteid=%7B0D555D22-714A-473B-B971-CDE4D9CE013F%7D&webid=%7B643266C2-271B-46E0-A96F-2FA1068C485C%7D&uniqueid=%7B2B6C0596-0A5D-41BF-BEC7-A1EE0293F2D3%7D#use-a-specific-view")
	Menus := MenuBar()
	Menus.Add "&File", FileMenu  ; Attach the two submenus that were created above.
	Menus.Add "&Help", HelpMenu
	MyGui := Gui()
	MyGui.Opt("+AlwaysOnTop")
	MyGui.MenuBar := Menus
	MyGui.Title := "Alt+Shift+F (Failures) Hotkey"
	MyGui.AddText("", "This script will attempt to resolve specific failures in Work Queue `rMonitor. Please make sure you have the appropriate view selected.`n`nHow many rows would you like to attempt?")
	MyGui.AddEdit("r1 vRowCount w24 x240 y42")
	MyGui.AddText("Section x10","`nCheck this box to limit script to a specific item. `rItem name will be entered at next screen.`r")
	MyGui.AddCheckBox("vInclusion x246 y88")
	MyGui.AddButton("Section w80 x130", "&OK").OnEvent("Click", Submit)  			; Button to call Submit function
	MyGui.OnEvent("Close", call_reload)														; Function call will reload script, hiding GUI
	MyGui.Show "w340"
	WinWaitClose("Alt+Shift+F (Failures) Hotkey")											; Wait for GUI window to close before returning call
	Return

Submit(*)																					; Function to save GUI input
{
    Saved := MyGui.Submit(1)
	IncludeFlag := Saved.Inclusion
	If IsInteger(Saved.RowCount) 															; If RowCount value is non-integer, leave it as zero so script will exit
		{
		RowCount := Saved.RowCount
		}
	Return	
}

call_reload(*)																				; Function to remove GUI and reload script
{
	MyGui.Destroy()
	Reload 																					; Reload is giving me less headache than Exit and it keeps script available
	Return
}
}

GUIFail()																					; Call GUI

If RowCount < 1																				; Exit if entered value is invalid
	{
	Exit
	}

If IncludeFlag = 1
; String version
	{
	Include := InputBox("If you would like to limit to specific items, enter an item name here as displayed in your view. `n`nIf no item is entered, the script will attempt to process all prescriptions with a qualifying failure detail that are displayed in view.", "Inclusion List")

	If include.Result = "OK" AND include.Value != ""			
		{
		IncludeList.Push(StrUpper(include.Value))											; Array isn't necessary at this point but could be developed to put multiple items into list (worth it?)
		}
	Else if include.Result = "Cancel"
		{
		Exit
		}	
	}

ToolTip "Please wait, script is executing", 875, 32, 1                                      ; Display tool tip so user knows script is running
Loop RowCount                                                                           	; Loop as many times as user defined
    {
    WinActivate("Work Queue Monitor")
    Loop 3                                                                                  ; If copy fails 3 times, stop looping
        {
        Sleep 250
        FailureDetail := ""																	; Empty FailureDetail variable
        MouseClick "left", 330, yPosFailureDetail, 1
        A_Clipboard := ""																	; Start off empty to allow ClipWait to detect when the text has arrived
        Send "^c"																			; Copy failure detail
        Clipwait(2)
        FailureDetail := StrReplace(A_Clipboard,"`r`n")                                     ; Remove hidden characters from copied value
        } until FailureDetail != ""                                                         ; Break copy loop once variable is loaded
    If StrLen(FailureDetail) = 0
        {
		ExitReason := ": copy from WQM failed"
        Break                                                                               ; Break action loop if copy unsuccessful
        }
	Else if SubStr(FailureDetail,1,1) = "`t" 												; If copy action is taken in empty list space the headers are copied and first character will be a tab
	    {
		ExitReason := ": end of list or list empty"
        Break                                                                               ; Break action loop if copy unsuccessful
        }

	If IncludeList.Length > 0																; In case a user checks the box but then enters no value
		{
		MouseClick "left", 33, yPosFailureDetail, 1											; Click on row to 
		A_Clipboard := ""																	; Start off empty to allow ClipWait to detect when the text has arrived
		Send "^c"																			; Copy failure detail
		Clipwait(2)

/* Since currently limited to single entry, we only need value at index position = 1 */						
		If InStr(StrUpper(A_Clipboard),"`t" IncludeList.Get(1)) = 0							; If clipboard (row) does not include item from include list, increment position and try next row							
			{
			yPosFailureDetail += 18
			Continue
			}
		}
		
; Once copied and initial validitation checks performed, compare to string values to determine which resolution pathway to take

; CMOP or Local Resuspend sequence; pre-transmission failures may not always trigger initial Bill-Only window (e.g. external refill requests will not have a fill to cancel) so may need to revise this section for those to consistently work

    If InStr(LocalResuspend, FailureDetail) > 0 OR InStr(CMOPResuspend, FailureDetail) >0	; Resuspend to local or CMOP sequence (most are "post-transmission" failures so sequence is very similar)
        {
        Send "{Home}"                                                                       ; Navigate to action column dropdown
        Send "R"                                                		                    ; r to select "resuspend" action
        Send "!s"                                               		                    ; alt+s to submit action
		WinWaitNotActive("PharmNet: Work Queue Monitor")                                    ; Wait for WQM window to become inactive
		if WinActive("No Actions")											                ; "No actions submitted" window indicates either patient locked or list is empty
			{
			Send "{Enter}"													                ; Acknowledge alert
			ExitReason := ": no actions selected"
			Break															                ; Break loop so user can assess and restart if appropriate; alternatively could adjust position and try next row but exiting is probably good bet for now
			}
		WinMatch("Bill-Only", 15)															; "Bill-Only" window loads (cancel fill); HYPHEN distinguishes from later Bill Only for Patient window.
		Send "{Tab 2}"                                          		                	; Navigate to reason code
		Send "Pharmacy Out of Stock"			                		                	; Primary use case is probably CMOP stock issue, could map more precisely (e.g. refrig/haz as patient request or other)
		Send "!o"                                               		                	; alt+o for OK
		WinWaitNotActive("Bill-Only")                           		                	; Wait for window to disappear
		if WinActive("Cancel Fill Warning - \\Remote")              		                ; Cancel fill warning may appear if dispense is > 14 days old
			{
			Send "!o" 					                            		                ; If it does, alt+o for OK
			WinWaitNotActive("Cancel Fill Warning")											; Wait for window to close before next sequence to avoid sending duplicate "!o"
			}
		if WinActive("Warning")                                         					; Health plan unavailable for this encounter
			{
			Send "!o"
			WinWaitNotActive("Warning")
			}
		WinMatch("Bill Only for Patient", 15)
		if InStr(LocalResuspend, FailureDetail) > 0                                         ; If resuspending to local, need to update routing option
			{
			MouseClick "left", 81, 441, 1					                				; Navigate to routing option
            Send "^a^a"                                                                     ; Ctrl+A x2 seems to clear the value (rather than select all)
            Send "{Backspace}"                                                              ; But if it does select all, then backspace to clear
			Send "Local Regular Mail"                               		                ; Update to local regular mail
			WinMatch("Routing Option Override", 15)
			Send "!u" 					                            		                ; Continue to accept default routing option override reason
			}
		WinWaitActive("Bill Only for Patient")                  		                	; Rx ready to submit - could pause here for pharmacist to review
		CheckHP()                                                                           ; Calling CheckHP() as !o on next line won't trigger due to #HotIf
		Send "!o"					                            		                    ; OK to accept and move to next or exit if done
        ;Send "!c"
		WinWaitNotActive("Bill Only for Patient")
		if WinActive("Cannot Select Health Plan")                                           ; Cannot Select Health Plan window may appear
			{
			Send "{Enter}"                                                                  ; Acknowledge and move on
			}
        WinWaitActive("Work Queue Monitor")                                                 ; Wait for WQM to become active
        Continue                                                                            ; This rx is done, restart loop
        }

; Inactive product sequence

    else if InStr(FailureDetail, ProdInactive) > 0					                        ; Inactive product sequence
        {
        Send "{Home}"
		Send "M"                                                			                ; m to select "modify" action
		Send "!s"                                               			                ; alt+s to submit action
		WinWaitNotActive("PharmNet: Work Queue Monitor")
		if WinActive("Unable to Refill")								                    ; Product has been inactivated, order must be replaced
			{
			Send "{Enter}"													                ; Acknowledge alert
			yPosFailureDetail += 18					                                        ; This row will need to be skipped, so update position
			Continue														                ; Restart loop
			}
		if WinActive("No Actions")											                ; Either patient locked or list is empty
			{
			Send "{Enter}"													                ; Acknowledge alert
			ExitReason := ": no actions selected"
			Break															                ; Break loop so user can assess and restart if appropriate
			}
		WinMatch("Warning: Inactive Product", 15)
		Send "!o"                                               		                	; alt+o for OK
		WinWaitNotActive("Warning: Inactive Product")                                   	; Wait for window to disappear
		WinMatch("Generic Substitution", 15)												; Generic Substitution window loads
		Send "{Tab}"													                	; Navigate to NDC list
		Send "{End}"													                	; P NDC is expected to be at end of list (based on experience, may need to confirm validity)
		Send "!o"														                	; OK to accept new NDC
		WinMatch("Refill for Patient", 15)													; Refill for Patient window loads (header = "Refill" for Modify action), this is rx screen
		if WinActive("Cannot Select Health Plan")
			{
			Send "{Enter}"
			}
		WinWaitActive("Refill for Patient")                  		                    	; Rx ready to submit - could pause here for pharmacist to review
		CheckHP()                                                                           ; Calling CheckHP() here does not appear to lead to duplication with !o
		Send "!o"					                            		                	; OK to accept and move to next or exit if done. Consider if adding pause for input would be better.
		WinWaitNotActive("Refill for Patient")
        WinWaitActive("Work Queue Monitor")
        Continue                                                                            ; This rx is done, restart loop
        }

; Health plan is ineffective sequence
    else if InStr(FailureDetail, HPIneffective) > 0											; Health plan change sequence - Need to identify and test various alternative windows.
        {
        Send "{Home}"
        Send "R"
        Send "!s"
		WinWaitNotActive("PharmNet: Work Queue Monitor")
		if WinActive("No Actions")											                ; Either patient locked or list is empty
			{
			Send "{Enter}"													                ; Acknowledge alert
			ExitReason := ": no actions selected"
			Break															                ; Break loop so user can assess and restart if appropriate
			}
		WinMatch("Refill for Patient", 15)
		Sleep 250
		CheckHP()                                                                           ; Calling CheckHP() here does not appear to lead to duplication with !o
        Send "!o"
        WinWaitActive("Work Queue Monitor")
        Continue                                                            				; This rx is done, restart loop
        }
    yPosFailureDetail += 18    
    }
If ExitReason = ""
    {
	ExitReason := ": process complete"
    }
    
BlockInput "Off"
BlockInput "MouseMove" "Off"    
ToolTip "", , ,1                                                                            ; Disable Tooltip if exists
MsgBox "Exiting script" ExitReason
Exit
}
#HotIf

/************************************************************ Alt+Left Mouse Button  ***************************************************************
; This hotkey does a lot of different things in various Cerner or Microsoft applications.  
; Hotkey is limited to windows with "\\Remote" in the WinTitle, Microsoft Teams/Excel/Outlook or a PowerBI report viewed in a browser. It is unexpected to act in other applications.
; Redundant logic to check windows should also help avoid unintended actions. 
; Path is dependent upon the active window when the hotkey is called.

; PowerChart:
;   - Call hotkey on patient name in banner to copy MRN to search patient in MMR
;   - Script to initiate BVE via communicate and PM conversation were created but not released due to concerns with system lag/script consistency and possible unintended data modifications (e.g. patient name edits)
;       - Will continue to try and develop but unclear if this will be successful

; Claims Monitor, Dispense Monitor, Work Queue Monitor:
;   - Call hotkey on row in DM or CM to extract rx number, or on rx number cell in WQM; rx number is used to search for patient in MMR

; Medication Manager Retail:
;   - Call hotkey with a patient loaded to add inactive meds to amb profile and sort by order sentence
;   - Call hotkey with tracking number highlighted to track package (MMR or WQM activity log; MMR dispense details)

; Other Cerner Windows
;   - Call hotkey in an otherwise undefined window  (e.g. History, Suspended Rx Activity Log) to maximize that window

; Teams, Outlook
;   - Highlight rx number (12 digit, with or without -), then call hotkey to search in MMR

; Excel , PowerBI report (Rx or MRN)
;   - Call hotkey with mouse over cell containing rx number value (12 digit, with or without -) to search in MMR
****************************************************************************************************************************************************/

; AHK documentation suggests using groups to optimize #HotIf, so creating group to define context for hotkey. Also grouping MSFT suite.
GroupAdd "MSFT", "Microsoft Teams"
GroupAdd "MSFT", "- Excel"
GroupAdd "MSFT", "- Outlook"                                                                ; Outlook via reading pane
GroupAdd "MSFT", "- Message"                                                                ; Open outlook message
;GroupAdd "MSFT", "SQL Server Management Studio"                                             ; SSMS. Data users can uncomment this if they like but Alt+LClick is useful in ADS, SSMS.

GroupAdd "ALC", "ahk_group MSFT"                                                            ; Add above group to overarching group
GroupAdd "ALC", "\\Remote"                                                                  ; Cerner applications (possibly other remote desktop things too)
GroupAdd "ALC", "ahk_class Chrome_WidgetWin_1",,"Azure Data Studio"                         ; Edge or Chrome browser, exclude ADS

#HotIf WinActive("ahk_group ALC")
!LButton::                                                                                  ; Alt + left mouse click - Lots of different things

{
CoordMode "Mouse", "Window"
SetTitleMatchMode 2 	                                                                    ; A window's title can contain WinTitle anywhere inside it to be a match.

/* Script logic variables */
ID := ""                                                                                    ; Variable to store ID value (PC = MRN or FIN; DM, CM, WQM = Rx#)
Term := ""                                                                                  ; Variable to store the search term type

MouseGetPos &xPosMouse, &yPosMouse                                                          ; Record current mouse position

BlockInput "On"                                                                             ; Block keyboard while executing
BlockInput "MouseMove"                                                                      ; Block mouse movement while executing

If WinActive("Opened by") or WinActive("PowerChart Organizer for")                          ; Window title for open chart in PC or Message Center    
    {
    MouseClick "left", , , 1                                                                ; Click where mouse is
    If WinActive("PowerChart Organizer for") AND PixelGetColor(255,340) != 0x007BBD         ; Message center expected to have blue banner here, otherwise assume viewing Patient List in PC Organizer
        {
        A_Clipboard := ""																	; Start off empty to allow ClipWait to detect when the text has arrived
        Send "^c"																			; Copies row
        Clipwait(2)
        ID := StrReplace(SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\t\d\d\d\d\d\d\d\d\t"),10),"`t")    ; FIN values won't always be 8 characters long but this works for now
        Term := "FIN"
        }
    Else if WinWaitActive("Custom Information:", ,2)                                        ; Wait for Custom Information to load
        {
        WinGetPos , , &W, , "Custom Information:"                                           ; Get window width
        If W < 800                                                                          ; Assess if we triggered by clicking Patient (narrower) or Encounter
            {
            MouseClick "left", 126, 246, 2                                                  ; Double click to highlight MRN
            A_Clipboard := ""																; Start off empty to allow ClipWait to detect when the text has arrived
            Send "^c"																		; Copy MRN
            Clipwait(2)
            ID := StrReplace(A_Clipboard,"`r`n")											; Remove line feed/carriage return and assign copied text to ID variable
            WinClose("Custom Information:")                                                 ; Close window
            Term := "MRN"                                                                   ; Will use MRN as search term in MMR
            }
        Else
            {                                
            MouseClick "left", 470, 397, 2                                                  ; Double click to highlight FIN
            A_Clipboard := ""																; Start off empty to allow ClipWait to detect when the text has arrived
            Send "^c"																		; Copy FIN
            Clipwait(2)
            ID := StrReplace(A_Clipboard,"`r`n")											; Remove line feed/carriage return and assign copied text to ID variable
            WinClose("Custom Information:")                                                 ; Close window
            Term := "FIN"                                                                   ; Will use FIN as search term in Med Manager
            }
        }
    Else                                                                                    ; Window did not load in 2 seconds so exit (e.g. clicked somewhere else in PC) 
        {
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit
        }
    }

Else if WinActive("PharmNet: Dispense Monitor")                                             ; Called from dispense monitor
    {
    MouseClick "left", , ,2                                                                 ; Double click on highlighted row to open details
    WinWaitActive("View Details", ,1)                                                       ; Wait up to 1 second for window to load 
    If !WinActive("View Details")                                                           ; If clicking in empty space, window never launches so release blocks and exit
        {
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit      
        }
    MouseClickDrag "left", 20, 68, 115, 68                                                  ; Need to click and drag to grab rx number
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copy rx number
    Clipwait(2)
    ID := StrReplace(A_Clipboard,"`r`n")												    ; Remove line feed/carriage return and assign copied text to ID variable
    WinClose("View Details")                                                                ; Close window
    Term := "Rx"                                                                            ; Will use Rx Number as search term
    }

Else if WinActive("PharmNet: Work Queue Monitor")                                           ; Called from WQM
    {
    MouseClick "left", , ,1                                                                 ; Click at mouse cursor - cell is highlighted not row as in other apps
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copy cell
    Clipwait(2)
    ID := StrReplace(A_Clipboard,"`r`n")												    ; Remove line feed/carriage return and assign copied text to ID variable
    Term := "Rx"
    }

Else if WinActive("PharmNet: Claims Monitor")                                               ; Called from Claims Monitor
    {
    MouseClick "left", , ,1                                                                 ; Click wherever cursor is
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copies row
    Clipwait(2)
    ID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d"),13)        ; Strip rx number from copied row - pattern matching for "[0-9][0-9][0-9][0-9]-..." (e.g. "3001-")                           
    Term := "Rx"
    }

Else if WinActive("- PharmNet: Retail Med")                                                 ; Called in MMR with patient loaded so we're going to add inactive orders
    {
    WinActivate("- PharmNet: Retail Med")
    KeyWait "Alt"                                                                           ; Alt button from hotkey trigger must be released for this to work
    Send "!v"                                                                               ; Alt+V to open View menu
    Send "{Enter}"                                                                          ; Enter to select first option (Inactive Orders)
    If WinWaitActive("View Profile by Status",,0.5)                                         ; If adding, View Profile by Status should appear
        {                                                                          
        Send "!o"                                                                           ; OK to close window
        WinWaitActive("- PharmNet: Retail Med",,1)
        MouseClick "left", 1100, 220, 1                                                     ; Click where we think the Order Sentence column header may exist (will vary)
        }
    MouseMove xPosMouse, yPosMouse, 0                                                       ; Put the mouse back where it was
    BlockInput "Off"                                                                        ; Release blocks and exit
    BlockInput "MouseMove" "Off"
    Exit
    }

/* Bill Only (MMR) and "Bill Only" (resuspend in WQM) are not the same size; do Bill Only window first since Bill Only is in ahk_group HP */

Else if WinActive("Bill Only for Patient") and xPosMouse > 657 and xPosMouse < 748 
    and (yPosMouse > 716 and yPosMouse < 740
    or yPosMouse > 779 and yPosMouse < 804)
    {
    CheckHP()
    MouseClick "left", xPosMouse, yPosMouse, 1
    BlockInput "Off"                                                                        ; Release blocks and exit
    BlockInput "MouseMove" "Off"
    Exit
    }

Else if WinActive("ahk_group HP") and xPosMouse > 646 and xPosMouse < 748 and yPosMouse > 740 and yPosMouse < 765
    {
    CheckHP()
    MouseClick "left", xPosMouse, yPosMouse, 1
    BlockInput "Off"                                                                        ; Release blocks and exit
    BlockInput "MouseMove" "Off"
    Exit
    }

Else if WinActive("Product Selection") OR WinActive("for Patient:")                         ; Use NDC or CMOP ID to search CMOP Product Line
    {
    If WinActive("for Patient:")                                                            ; [Inquire/New Order/Refill/Bill Only/Modify] for Patient:" is Rx screen
        {
        MouseClick "left", , ,1                                                             ; User to call while hovering over "Product" window
        WinWaitActive("Product",,1)                                                         ; Wait to see if window loads
        If !WinActive("Product")                                                            ; If not, release blocks and exit
            {
            BlockInput "Off"                                                                
            BlockInput "MouseMove" "Off"
            Exit                
            }
        }

    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copies row
    Clipwait(2)

    If WinActive("Product")                                                                 ; If window did open, close now that we have copied NDC
        {
        WinClose("Product")
        }

    CMOPID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"[A-Z]\w\d\d\d"),5)                 ; CMOP ID is alpha, alphanumeric, numeric, numeric, numeric (e.g. S0020, XS515)
    NDC := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d\d-\d\d\d\d-\d\d"),13)        ; NDC in 5-4-2 format

    If CMOPID != ""
        {
        Term := CMOPID
        }
    Else if NDC != ""
        {
        Term := NDC
        }
    Else
        {
        BlockInput "Off"                                                                    ; Release blocks and exit
        BlockInput "MouseMove" "Off"
        Exit
        }
    Run Format('msedge.exe "https://vaww.pbi.cdw.va.gov/PBI_RS/report/GPE/CMOP_Analytics/DART/ProductLineSearch?Product={1}"', Term)

    BlockInput "Off"                                                                        ; Release blocks and exit
    BlockInput "MouseMove" "Off"
    Exit
    }

Else if WinActive("ahk_group MSFT")                                                         ; Called from Teams, Outlook, or Excel
    {
    If WinActive("- Excel")
        {
        MouseClick "left", , ,1                                                             ; Click wherever cursor is; for Teams or Outlook, user must have value highlighted
        }
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copies highlighted text
    Clipwait(2)
    ID := StrReplace(A_Clipboard,"-")
    ID := SubStr(ID,RegExMatch(ID,"\d\d\d\d\d\d\d\d\d\d\d\d"),12)                           ; Assume Excel or Teams values +/- hyphen so adjusting SubStr (12) to pass INT test below
    Term := "Rx"
    }

Else if WinActive("Power BI")                                                               ; Hotkey will run in browser, this limits it to a Power BI report on browser
    {
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    MouseClick "left", , , 1                                                                ; Left-click on cell   
    MouseGetPos &xPosMouse, &yPosMouse                                                      ; Check that we are in a cell and not whitespace
    If PixelGetColor(xPosMouse,yPosMouse) = "0xFFFFFF"                                      ; Exit if in whitespace
        {
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit            
        }
    MouseClick "right", , , 1                                                               ; Right-click on cell to open context menu
    Send "{Up}"                                                                             ; Navigate context menu
    Send "{Right}"
    Send "{Enter}"
    Clipwait(2)
    If RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d") = 1                              ; Will equal 1 if rx number with hyphen is beginning of clipboard value
        {
        ID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d"),13)    ; If so, extract and remove any trailing hidden characters
        Term := "Rx"
        }
    Else if RegExMatch(A_Clipboard,"\D") = 0                                                ; If copied value is strictly numeric, treat it like MRN
        {
        ID := StrReplace(A_Clipboard,"`n`r")                                                ; MRN length can vary so strip trailing characters if exist
        Term := "MRN"
        }
    }

Else if WinActive("Suspended Orders Activity Log") OR WinActive("Dispense Details")         ; Tracking info visible here
    {
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    ; Copies row
    Clipwait(0.5)
    Tracking := StrReplace(SubStr(A_Clipboard,InStr(A_Clipboard,"`t",,,-1)),"`t")           ; Extract tracking number from copied row, carrier not always available

    If InStr(StrUpper(Tracking),"Tracking") > 0                                             ; "Tracking info unavailable" should not trigger search
        {
        WinMaximize("A")                                                                    ; So instead make it big!
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit
        }
    Else if InStr(A_Clipboard,"Acknowledged") > 0                                           ; If called on Ack row, load CMOP Cerner Rx Lookup report
        {
        ALWin := WinGetTitle("Suspended Orders Activity Log - <")
        RxNumber := SubStr(ALWin,InStr(ALWin,"<")+1,13)
        Run Format('msedge.exe "https://vaww.pbi.cdw.va.gov/PBI_RS/report/GPE/CMOP_Analytics/DART/Cerner%20Rx%20Lookup?SearchTerm={1}"', RxNumber)
        }
    Else if InStr(A_Clipboard,"USPS") > 0 OR SubStr(Tracking,1,1) = "9"                     ; USPS starts with 9
        {
        Run Format('msedge.exe "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1={1}"', Tracking)
        }
    Else if InStr(A_Clipboard,"UPS") > 0 OR SubStr(Tracking,1,2) = "1Z"                     ; UPS starts with 1Z
        {
        Run Format('msedge.exe "https://www.ups.com/track?track=yes&trackNums={1}&loc=en_US&requester=ST/trackdetails"', Tracking)
        }
    Else if InStr(A_Clipboard,"FEDEX") > 0                                                  ; Fedex is only other carrier so far, seems to typically start with 7
        {
        Run Format('msedge.exe "https://www.fedex.com/fedextrack/?cntry_code=us&trknbr={1}"', Tracking)
        }
    Else
        {
        WinMaximize("A")                                                                    ; Make it big!
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit
        }
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    A_Clipboard := Tracking                                                                 ; Load tracking number to clipboard in case website fails or wrong carrier
    WinActivate("Edge")
    Exit
    }

Else if WinActive("Suspended Order Details")                                                ; This is where we can get TRN, so we can search CMOP application/web portal
    {
    MouseClickDrag "left", 267, 420, 360, 420                                               ; Drag to highlight, double-click fails 2/2 hyphens
    A_Clipboard := ""																	    ; Start off empty to allow ClipWait to detect when the text has arrived
    Send "^c"																			    
    Clipwait(2)
    If A_Clipboard = ""                                                                     ; Make sure value exists in clipboard
        {
        WinMaximize("A")                                                                    
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        Exit            
        }
    Run("msedge.exe https://vaww.cmop.med.va.gov/CMOPNationalWebApplication/", ,"Max")
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    WinWaitActive("Consolidated Mail Outpatient Pharmacy")                                  ; This may not always work due to log in, but will be in clipboard anyway
    Send "^0"                                                                               ; Reset zoom so we know where to click
    BlockInput "MouseMove"
    Loop                                                                                    ; Check sidebar color as proof of completed page load
        {
        Sleep 250
        } until PixelGetColor(75,800) = 0x1E2D54
    MouseClick "left", 360, 440, 1                                                          ; Click in TRN box
    Send A_Clipboard
    Send "{Enter}"
    BlockInput "MouseMove" "Off"
    Exit    
    }

Else                                                                                        ; If user does this in other Cerner window
    {
    WinMaximize("A")                                                                        ; Make it big!
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    Exit
    }

PharmacyPatientLookup(ID,Term)                                                              ; Extracted code as standalone function, replaced here 11/01/2023

BlockInput "Off"
BlockInput "MouseMove" "Off"
If !WinWaitNotActive("PharmNet: Retail", ,3)											    ; Wait up to 3 seconds to check if window doesn't change (i.e. patient loading)
	{
    A_Clipboard := ID
    MsgBox Format("Hotkey appears to have failed, but {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID),,"T5"
	}
Exit
}
#HotIf                                                                                      ; Turn off HotIf

/************************************************************ Alt + Shift + X (Exit) *********************************************************************
; Hotkey will interrupt executing script, release any input blocks that may exist, and exit the script
; Script must be reloaded before hotkey will run again
; Panic button that will hopefully never be used
*********************************************************************************************************************************************************/

!+X::
{
    BlockInput "Off"
    BlockInput "MouseMove" "Off"
    Report := MsgBox("Exit command called. If you experienced a failure, please click OK to send a report via email.`n`nIf everything is fine, please click cancel and have a nice day.","Catastrophic Exit","OC")
    If Report = "OK"
        {
        Run "mailto:lewis.dejaegher@va.gov?subject=OCRAP, it failed!&body=Here's what I was doing and what went wrong."
        }
    ExitApp
}

/************************************************************ Alt + Shift + H (Help) *********************************************************************
; Hotkey that opens help page in Edge browser
*********************************************************************************************************************************************************/

!+H::			; Alt+Shift+H (Help)
{
/* Open PBM OCP COP SP reference page */
Run "msedge.exe https://dvagov.sharepoint.com/:u:/r/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo"
}