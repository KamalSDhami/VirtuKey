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

; Ensure desktop exists, then switch instantly
SwitchOrCreateDesktop(num) {
    count := GetDesktopCount()
    if (num > count) {
        ; Create missing desktops quickly
        Loop (num - count) {
            Send("^#d")   ; Ctrl+Win+D = new desktop
            Sleep 50     ; minimal delay for Windows to register
        }
    }
    ; Switch instantly to target desktop
    GoToDesktop(num)
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
