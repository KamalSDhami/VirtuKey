; ===============================
; Virtual Desktop Manager v2 (Instant Switch)
; Uses Ctrl+Win+D to create desktops if needed
; Then jumps directly to target using DLL
; ===============================

dllPath := A_ScriptDir "\VirtualDesktopAccessor.dll"

; Load DLL
hModule := DllCall("LoadLibrary", "Str", dllPath, "Ptr")
if (!hModule) {
    MsgBox "Error: Could not load VirtualDesktopAccessor.dll"
    ExitApp
}

; --- Functions ---
GetDesktopCount() {
    return DllCall("VirtualDesktopAccessor\GetDesktopCount", "Int")
}

GetCurrentDesktop() {
    return DllCall("VirtualDesktopAccessor\GetCurrentDesktopNumber", "Int")
}

GoToDesktop(num) {
    return DllCall("VirtualDesktopAccessor\GoToDesktopNumber", "Int", num-1)
}

; Get the topmost window on the current desktop after switch
GetTopWindowOnCurrentDesktop() {
    ; Wait a bit more for desktop switch to complete
    Sleep 100
    
    ; Try multiple approaches to find and focus a window
    
    ; Method 1: Try to get the current foreground window
    hwnd := DllCall("GetForegroundWindow", "Ptr")
    if (hwnd && WinExist("ahk_id " hwnd)) {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if (title != "" && !InStr(title, "Program Manager") && !InStr(title, "Task View")) {
                return hwnd
            }
        }
    }
    
    ; Method 2: Find the topmost visible window
    hwnd := DllCall("GetTopWindow", "Ptr", 0, "Ptr")
    while (hwnd) {
        if (DllCall("IsWindowVisible", "Ptr", hwnd)) {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                ; Skip windows without titles or system windows
                if (title != "" && !InStr(title, "Program Manager") && !InStr(title, "Task View") && !InStr(title, "Windows Input Experience")) {
                    ; Check if it's not minimized
                    if (WinGetMinMax("ahk_id " hwnd) != -1) {
                        return hwnd
                    }
                }
            }
        }
        hwnd := DllCall("GetWindow", "Ptr", hwnd, "UInt", 2) ; GW_HWNDNEXT
    }
    
    return 0
}

; Clear focus from current desktop before switching
ClearCurrentFocus() {
    ; Get the current foreground window and minimize it briefly to clear focus
    hwnd := DllCall("GetForegroundWindow", "Ptr")
    if (hwnd) {
        ; Set focus to desktop/shell to clear app focus
        try {
            ; Find the desktop window
            desktopHwnd := DllCall("GetShellWindow", "Ptr")
            if (desktopHwnd) {
                DllCall("SetForegroundWindow", "Ptr", desktopHwnd)
            }
        }
    }
    ; Clear any lingering focus
    DllCall("SetFocus", "Ptr", 0)
}

; Ensure desktop exists, then switch instantly with proper focus
SwitchOrCreateDesktop(num) {
    count := GetDesktopCount()
    if (num > count) {
        ; Create missing desktops quickly
        Loop (num - count) {
            Send("^#d")   ; Ctrl+Win+D = new desktop
            Sleep 50     ; minimal delay for Windows to register
        }
    }
    
    ; Clear focus from current desktop BEFORE switching
    ClearCurrentFocus()
    Sleep 50
    
    ; Switch instantly to target desktop
    GoToDesktop(num)
    
    ; Wait longer for the desktop switch to complete
    Sleep 200
    
    ; Try to find and focus a window on this desktop
    hwnd := GetTopWindowOnCurrentDesktop()
    
    if (hwnd != 0) {
        ; Force activate the window with multiple attempts
        try {
            ; First attempt - activate window
            WinActivate("ahk_id " hwnd)
            Sleep 50
            
            ; Second attempt - bring to foreground
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            Sleep 50
            
            ; Third attempt - ensure focus
            DllCall("SetFocus", "Ptr", hwnd)
            Sleep 50
            
            ; Fourth attempt - make sure it's really active
            WinActivate("ahk_id " hwnd)
            
        } catch {
            ; Ultimate fallback: Alt+Tab to cycle through windows
            Send("{Alt down}{Tab}{Alt up}")
        }
    } else {
        ; If no window found, try Alt+Tab to activate any available window
        Send("{Alt down}{Tab}{Alt up}")
        Sleep 100
        
        ; Check if Alt+Tab worked
        hwnd := DllCall("GetForegroundWindow", "Ptr")
        if (!hwnd || !WinExist("ahk_id " hwnd)) {
            ; Final fallback: click on desktop center to ensure focus
            MouseGetPos(&mouseX, &mouseY)
            Click(A_ScreenWidth//2, A_ScreenHeight//2)
            MouseMove(mouseX, mouseY, 0)
            ; Send Windows key to refresh desktop focus
            Send("{LWin down}{LWin up}")
        }
    }
}


; --- HOTKEYS: Win+1..9,0 ---
#1::SwitchOrCreateDesktop(1)
#2::SwitchOrCreateDesktop(2)
#3::SwitchOrCreateDesktop(3)
#4::SwitchOrCreateDesktop(4)
#5::SwitchOrCreateDesktop(5)
#6::SwitchOrCreateDesktop(6)
#7::SwitchOrCreateDesktop(7)
#8::SwitchOrCreateDesktop(8)
#9::SwitchOrCreateDesktop(9)
#0::SwitchOrCreateDesktop(10)

; Optional: exit script
#Esc::ExitApp

; TrayTip on load
TrayTip("✅ Virtual Desktop Manager Loaded", "Use Win+1..9,0 to instantly switch/create desktops")
