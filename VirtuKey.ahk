#Requires AutoHotkey v2.0
#SingleInstance Force
#UseHook
Persistent

; ===============================
; VirtuKey - instant virtual desktop switching for Windows 11
; Win+1..9,0 switches (creating desktops if needed), Win+Shift+N moves the
; active window. Powered by VirtualDesktopAccessor.dll.
; ===============================

KeyHistory 0
ListLines False
; AHK treats windows on other virtual desktops (DWM-cloaked) as hidden.
; IsUserWindow() checks WS_VISIBLE itself.
DetectHiddenWindows True
SetWinDelay -1
SetKeyDelay -1, -1
CoordMode "Mouse", "Screen"
ProcessSetPriority "AboveNormal"
A_MaxHotkeysPerInterval := 200

APP_NAME := "VirtuKey"
APP_VERSION := "2.0.0"
SETTINGS_FILE := A_AppData "\VirtuKey\settings.ini"
RUN_KEY := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
WM_DESKTOP_CHANGED := 0x1400 + 30
MAX_PILLS := 20

; Menu label -> setting name
OPTION_ITEMS := Map(
    "Auto-create missing desktops", "AutoCreate",
    "Show desktop indicator", "ShowOsd",
    "Follow moved windows", "FollowMove"
)

Settings := LoadSettings()
VDA := LoadVDA()
Theme := LoadTheme()
IconCache := Map()
Osd := ""
OsdAlpha := 0
SuppressHookOsdUntil := 0
SwitchMenu := Menu()

if !EnsureVDA()
    Fatal("VirtualDesktopAccessor.dll does not support this Windows build (" A_OSVersion ").`n`n"
        . "Download the latest DLL from github.com/Ciantic/VirtualDesktopAccessor/releases "
        . "and replace the one next to " A_ScriptName ".")

; --- Hotkeys ---
Loop 10 {
    key := Mod(A_Index, 10)
    Hotkey("#" key, SwitchHotkey.Bind(A_Index))
    Hotkey("#+" key, MoveHotkey.Bind(A_Index))
}
Hotkey("#SC029", (*) => RemoveEmptyDesktops())   ; Win+`
Hotkey("#Esc", (*) => ExitApp())

; --- Events ---
OnMessage(WM_DESKTOP_CHANGED, OnDesktopChanged)
OnMessage(0x404, OnTrayNotify)                      ; AHK_NOTIFYICON
OnMessage(0x1A, OnSettingChange)                    ; WM_SETTINGCHANGE
OnMessage(DllCall("RegisterWindowMessage", "Str", "TaskbarCreated", "UInt"), OnExplorerRestart)
OnExit(OnAppExit)
RegisterDesktopHook()

BuildTrayMenu()
RefreshTray()

if Settings.FirstRun {
    TrayTip("Win+1..9,0 switches desktops. Right-click the tray icon for options.", APP_NAME " is running")
    Settings.FirstRun := false
    SaveSetting("FirstRun")
}

SwitchHotkey(n, *) => SwitchToDesktop(n)
MoveHotkey(n, *) => MoveActiveWindowTo(n)

; ===============================
; DLL access
; ===============================

LoadVDA() {
    h := 0
    for path in [A_ScriptDir "\VirtualDesktopAccessor.dll", "VirtualDesktopAccessor.dll"] {
        if (h := DllCall("LoadLibrary", "Str", path, "Ptr"))
            break
    }
    if !h
        Fatal("Could not load VirtualDesktopAccessor.dll.`n`nPlace it next to " A_ScriptName " and try again.")

    ; Resolve every export once so hot paths skip the name lookup
    vda := {}
    for name in ["GetDesktopCount", "GetCurrentDesktopNumber", "GoToDesktopNumber",
        "IsWindowOnCurrentVirtualDesktop", "GetWindowDesktopNumber", "MoveWindowToDesktopNumber",
        "RegisterPostMessageHook", "UnregisterPostMessageHook", "RestartVirtualDesktopAccessor"] {
        if !(proc := DllCall("GetProcAddress", "Ptr", h, "AStr", name, "Ptr"))
            Fatal("VirtualDesktopAccessor.dll is missing " name "(). Please update the DLL.")
        vda.%name% := proc
    }
    return vda
}

DesktopCount() => DllCall(VDA.GetDesktopCount, "Int")
CurrentDesktop() => DllCall(VDA.GetCurrentDesktopNumber, "Int")   ; 0-based
GoToDesktop(index) => DllCall(VDA.GoToDesktopNumber, "Int", index, "Int")
; Windows can ignore a switch while a desktop-create animation is running, so retry
GoToAndWait(index) {
    Loop 3 {
        GoToDesktop(index)
        if WaitUntil(() => CurrentDesktop() = index, 400)
            return true
    }
    return false
}

IsOnCurrentDesktop(hwnd) => DllCall(VDA.IsWindowOnCurrentVirtualDesktop, "Ptr", hwnd, "Int") = 1

EnsureVDA() {
    if (DesktopCount() > 0)
        return true
    ; The DLL loses its COM connection when explorer.exe restarts
    DllCall(VDA.RestartVirtualDesktopAccessor)
    Sleep 250
    return DesktopCount() > 0
}

OnAppExit(*) {
    DllCall(VDA.UnregisterPostMessageHook, "Ptr", A_ScriptHwnd, "Int")
}

RegisterDesktopHook() {
    DllCall(VDA.UnregisterPostMessageHook, "Ptr", A_ScriptHwnd, "Int")
    DllCall(VDA.RegisterPostMessageHook, "Ptr", A_ScriptHwnd, "Int", WM_DESKTOP_CHANGED, "Int")
}

; ===============================
; Desktop actions
; ===============================

SwitchToDesktop(n) {
    Critical
    global SuppressHookOsdUntil := A_TickCount + 600
    if !EnsureDesktops(n) {
        ShowOsd(CurrentDesktop() + 1, DesktopCount(), "Desktop " n " doesn't exist")
        return
    }
    if (CurrentDesktop() != n - 1) {
        ; Park focus on the taskbar first so the old window doesn't keep
        ; keyboard focus (or flash) after the switch
        try WinActivate("ahk_class Shell_TrayWnd")
        GoToAndWait(n - 1)
        FocusTopWindow()
    }
    RefreshTray(true)
}

MoveActiveWindowTo(n) {
    Critical
    global SuppressHookOsdUntil := A_TickCount + 600
    hwnd := WinExist("A")
    if !hwnd || IsShellWindow(hwnd)
        return
    orig := CurrentDesktop()
    if !EnsureDesktops(n) {
        ShowOsd(orig + 1, DesktopCount(), "Desktop " n " doesn't exist")
        return
    }
    DllCall(VDA.MoveWindowToDesktopNumber, "Ptr", hwnd, "Int", n - 1, "Int")

    if Settings.FollowMove {
        SwitchToDesktop(n)
        try WinActivate(hwnd)
        return
    }
    ; Creating desktops with Ctrl+Win+D jumps to them, so come back
    if (CurrentDesktop() != orig) {
        GoToAndWait(orig)
    }
    FocusTopWindow()
    RefreshTray()
    ShowOsd(n, DesktopCount(), "Moved to Desktop " n)
}

; Creates desktops up to n with the native shortcut. Returns false if n is
; out of range and auto-create is off or Windows didn't respond.
EnsureDesktops(n) {
    count := DesktopCount()
    if (n <= count)
        return true
    if !Settings.AutoCreate
        return false
    while (count < n) {
        Send "^#d"
        if !WaitUntil(() => DesktopCount() > count, 1000)
            return false
        count := DesktopCount()
    }
    return true
}

; Closes empty desktops from the end until one in use (or the current one) is found
RemoveEmptyDesktops() {
    Critical
    global SuppressHookOsdUntil := A_TickCount + 5000
    orig := CurrentDesktop()
    used := UsedDesktops()
    count := DesktopCount()
    removed := 0
    while (count > 1) {
        last := count - 1
        if (used.Has(last) || last = orig)
            break
        if !GoToAndWait(last) || HasWindowsOnCurrentDesktop()
            break
        Send "^#{F4}"
        if !WaitUntil(() => DesktopCount() < count, 1000)
            break
        count -= 1
        removed += 1
    }
    if (CurrentDesktop() != orig) {
        GoToAndWait(orig)
        FocusTopWindow()
    }
    RefreshTray()
    ShowOsd(orig + 1, DesktopCount(), removed ? "Removed " removed " empty desktop" (removed > 1 ? "s" : "") : "No empty desktops")
    SuppressHookOsdUntil := A_TickCount + 600
}

NewDesktop() {
    Send "^#d"
}

; ===============================
; Window helpers
; ===============================

; Activates the topmost real window on the current desktop (WinGetList is in z-order)
FocusTopWindow() {
    for hwnd in WinGetList() {
        if IsUserWindow(hwnd) && IsOnCurrentDesktop(hwnd) {
            try {
                WinActivate(hwnd)
                return true
            }
        }
    }
    try WinActivate("ahk_class Progman")
    return false
}

HasWindowsOnCurrentDesktop() {
    for hwnd in WinGetList() {
        if IsUserWindow(hwnd, true) && IsOnCurrentDesktop(hwnd)
            return true
    }
    return false
}

UsedDesktops() {
    used := Map()
    for hwnd in WinGetList() {
        if IsUserWindow(hwnd, true) {
            d := DllCall(VDA.GetWindowDesktopNumber, "Ptr", hwnd, "Int")
            if (d >= 0)
                used[d] := true
        }
    }
    return used
}

IsUserWindow(hwnd, includeMinimized := false) {
    try {
        style := WinGetStyle(hwnd)
        exStyle := WinGetExStyle(hwnd)
        if !(style & 0x10000000)                          ; WS_VISIBLE
            return false
        if (style & 0x20000000) && !includeMinimized      ; WS_MINIMIZE
            return false
        if (exStyle & 0x80) && !(exStyle & 0x40000)       ; tool window without WS_EX_APPWINDOW
            return false
        if (exStyle & 0x08000000)                         ; WS_EX_NOACTIVATE
            return false
        if (WinGetTitle(hwnd) = "") || IsShellWindow(hwnd)
            return false
    } catch {
        return false
    }
    ; Suspended UWP apps leave cloaked ghost frames behind
    cloaked := 0
    DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "UInt*", &cloaked, "UInt", 4)
    return !(cloaked & 0x5)   ; DWM_CLOAKED_APP | DWM_CLOAKED_INHERITED
}

IsShellWindow(hwnd) {
    try cls := WinGetClass(hwnd)
    catch
        return true
    return cls ~= "^(Shell_TrayWnd|Shell_SecondaryTrayWnd|Progman|WorkerW|Windows\.UI\.Core\.CoreWindow|XamlExplorerHostIslandWindow|TopLevelWindowForOverflowXamlIsland)$"
}

WaitUntil(condition, timeoutMs) {
    deadline := A_TickCount + timeoutMs
    while !condition() {
        if (A_TickCount > deadline)
            return false
        Sleep 10
    }
    return true
}

; ===============================
; Events
; ===============================

OnDesktopChanged(wParam, lParam, *) {
    RefreshTray(A_TickCount > SuppressHookOsdUntil)
}

OnTrayNotify(wParam, lParam, *) {
    if (lParam = 0x204)   ; WM_RBUTTONDOWN: menu is about to open
        RebuildSwitchMenu()
}

OnSettingChange(wParam, lParam, *) {
    global Theme, Osd
    if !lParam || StrGet(lParam) != "ImmersiveColorSet"
        return
    Theme := LoadTheme()
    if IsObject(Osd) {
        Osd.Destroy()
        Osd := ""
    }
    old := []
    for , hIcon in IconCache
        old.Push(hIcon)
    IconCache.Clear()
    RefreshTray()
    for hIcon in old
        DllCall("DestroyIcon", "Ptr", hIcon)
}

OnExplorerRestart(*) {
    DllCall(VDA.RestartVirtualDesktopAccessor)
    SetTimer(() => (RegisterDesktopHook(), RefreshTray()), -1000)
}

; ===============================
; Tray
; ===============================

BuildTrayMenu() {
    m := A_TrayMenu
    m.Delete()
    m.Add("Keyboard shortcuts", (*) => ShowHelp())
    m.Default := "Keyboard shortcuts"
    m.Add()
    m.Add("Switch to", SwitchMenu)
    m.Add("New desktop", (*) => NewDesktop())
    m.Add("Remove empty desktops", (*) => RemoveEmptyDesktops())
    m.Add()
    for label in OPTION_ITEMS
        m.Add(label, ToggleOption)
    m.Add("Start with Windows", ToggleStartup)
    m.Add()
    m.Add("Pause hotkeys", TogglePause)
    m.Add("Exit", (*) => ExitApp())
    UpdateMenuChecks()
    RebuildSwitchMenu()
}

RebuildSwitchMenu() {
    SwitchMenu.Delete()
    cur := CurrentDesktop() + 1
    Loop DesktopCount() {
        label := "Desktop " A_Index
        SwitchMenu.Add(label, (name, pos, *) => SwitchToDesktop(pos))
        if (A_Index = cur)
            SwitchMenu.Check(label)
    }
}

UpdateMenuChecks() {
    m := A_TrayMenu
    for label, name in OPTION_ITEMS
        Settings.%name% ? m.Check(label) : m.Uncheck(label)
    IsStartupEnabled() ? m.Check("Start with Windows") : m.Uncheck("Start with Windows")
    A_IsSuspended ? m.Check("Pause hotkeys") : m.Uncheck("Pause hotkeys")
}

ToggleOption(label, *) {
    name := OPTION_ITEMS[label]
    Settings.%name% := !Settings.%name%
    SaveSetting(name)
    UpdateMenuChecks()
}

ToggleStartup(*) {
    if IsStartupEnabled()
        RegDelete(RUN_KEY, APP_NAME)
    else
        RegWrite(A_IsCompiled ? '"' A_ScriptFullPath '"' : '"' A_AhkPath '" "' A_ScriptFullPath '"', "REG_SZ", RUN_KEY, APP_NAME)
    UpdateMenuChecks()
}

IsStartupEnabled() => RegRead(RUN_KEY, APP_NAME, "") != ""

TogglePause(*) {
    Suspend(-1)
    UpdateMenuChecks()
    RefreshTray()
}

RefreshTray(withOsd := false) {
    cur := CurrentDesktop() + 1
    count := DesktopCount()
    if (cur < 1)
        return
    TraySetIcon("HICON:*" NumberIcon(cur, A_IsSuspended), , true)
    A_IconTip := APP_NAME " - Desktop " cur " of " count (A_IsSuspended ? " (paused)" : "")
    if withOsd
        ShowOsd(cur, count)
}

; Draws the desktop number on an accent-colored tile. Icons are cached per number.
NumberIcon(n, paused := false) {
    key := n "|" paused
    if IconCache.Has(key)
        return IconCache[key]

    size := SysGet(49)   ; SM_CXSMICON
    hdc := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
    bi := Buffer(40, 0)
    NumPut("UInt", 40, "Int", size, "Int", -size, "UShort", 1, "UShort", 32, bi)
    bits := 0
    hbm := DllCall("CreateDIBSection", "Ptr", hdc, "Ptr", bi, "UInt", 0, "Ptr*", &bits, "Ptr", 0, "UInt", 0, "Ptr")
    oldBm := DllCall("SelectObject", "Ptr", hdc, "Ptr", hbm, "Ptr")

    fill := paused ? 0x808080 : (Theme.AccentRef || 0x010101)
    brush := DllCall("CreateSolidBrush", "UInt", fill, "Ptr")
    oldBrush := DllCall("SelectObject", "Ptr", hdc, "Ptr", brush, "Ptr")
    oldPen := DllCall("SelectObject", "Ptr", hdc, "Ptr", DllCall("GetStockObject", "Int", 8, "Ptr"), "Ptr")   ; NULL_PEN
    radius := Round(size * 0.35)
    DllCall("RoundRect", "Ptr", hdc, "Int", 0, "Int", 0, "Int", size + 1, "Int", size + 1, "Int", radius, "Int", radius)

    ; GDI leaves alpha at 0: remember the transparent corners before text is drawn
    holes := []
    Loop size * size {
        if !NumGet(bits, (A_Index - 1) * 4, "UInt")
            holes.Push((A_Index - 1) * 4)
    }

    fontHeight := -Round(size * (n >= 10 ? 0.62 : 0.8))
    font := DllCall("CreateFont", "Int", fontHeight, "Int", 0, "Int", 0, "Int", 0, "Int", 700,
        "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 0, "UInt", 4, "UInt", 0, "Str", "Segoe UI", "Ptr")
    oldFont := DllCall("SelectObject", "Ptr", hdc, "Ptr", font, "Ptr")
    DllCall("SetBkMode", "Ptr", hdc, "Int", 1)   ; TRANSPARENT
    DllCall("SetTextColor", "Ptr", hdc, "UInt", paused ? 0xFFFFFF : Theme.OnAccentRef)
    rc := Buffer(16, 0)
    NumPut("Int", size, "Int", size, rc, 8)
    DllCall("DrawText", "Ptr", hdc, "Str", String(n), "Int", -1, "Ptr", rc, "UInt", 0x25)   ; DT_CENTER|DT_VCENTER|DT_SINGLELINE

    Loop size * size {
        off := (A_Index - 1) * 4
        NumPut("UInt", NumGet(bits, off, "UInt") | 0xFF000000, bits, off)
    }
    for off in holes
        NumPut("UInt", 0, bits, off)

    DllCall("SelectObject", "Ptr", hdc, "Ptr", oldFont)
    DllCall("SelectObject", "Ptr", hdc, "Ptr", oldPen)
    DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBrush)
    DllCall("SelectObject", "Ptr", hdc, "Ptr", oldBm)
    DllCall("DeleteObject", "Ptr", font)
    DllCall("DeleteObject", "Ptr", brush)

    hMask := DllCall("CreateBitmap", "Int", size, "Int", size, "UInt", 1, "UInt", 1, "Ptr", 0, "Ptr")
    ii := Buffer(8 + 3 * A_PtrSize, 0)   ; ICONINFO
    NumPut("Int", 1, ii, 0)
    NumPut("Ptr", hMask, ii, 8 + A_PtrSize)
    NumPut("Ptr", hbm, ii, 8 + 2 * A_PtrSize)
    hIcon := DllCall("CreateIconIndirect", "Ptr", ii, "Ptr")
    DllCall("DeleteObject", "Ptr", hMask)
    DllCall("DeleteObject", "Ptr", hbm)
    DllCall("DeleteDC", "Ptr", hdc)
    return IconCache[key] := hIcon
}

; ===============================
; On-screen indicator
; ===============================

BuildOsd() {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000020")   ; WS_EX_NOACTIVATE | WS_EX_TRANSPARENT
    g.BackColor := Theme.Bg
    g.MarginX := 0, g.MarginY := 0
    g.SetFont("s13 bold c" Theme.Text, "Segoe UI")
    g.SetFont(, "Segoe UI Variable Display")
    g.Heading := g.Add("Text", "x0 y0 w10 h10 Center BackgroundTrans")
    g.SetFont("s10")
    g.Pills := []
    Loop MAX_PILLS
        g.Pills.Push(g.Add("Text", "x0 y0 w10 h10 Center Hidden 0x200", A_Index))   ; SS_CENTERIMAGE
    corner := 2   ; DWMWCP_ROUND
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g.Hwnd, "UInt", 33, "Int*", &corner, "UInt", 4)
    return g
}

ShowOsd(highlight, count, title := "") {
    global Osd, OsdAlpha
    if !Settings.ShowOsd
        return
    if !IsObject(Osd)
        Osd := BuildOsd()

    s := A_ScreenDPI / 96
    pill := Round(30 * s), gap := Round(6 * s), pad := Round(18 * s)
    top := Round(14 * s), titleH := Round(30 * s)
    shown := count <= MAX_PILLS ? count : 0
    rowW := shown ? shown * pill + (shown - 1) * gap : 0
    w := Max(rowW + 2 * pad, Round(220 * s))

    Osd.Heading.Value := title != "" ? title : "Desktop " highlight
    Osd.Heading.Move(0, top, w, titleH)
    y := top + titleH + Round(8 * s)
    x := (w - rowW) // 2
    for i, p in Osd.Pills {
        if (i > shown) {
            p.Visible := false
            continue
        }
        active := (i = highlight)
        p.Opt("+Background" (active ? Theme.Accent : Theme.Surface))
        p.SetFont((active ? "bold c" Theme.OnAccent : "norm c" Theme.SubText))
        p.Move(x, y, pill, pill)
        p.Visible := true
        p.Redraw()
        x += pill + gap
    }
    h := shown ? y + pill + Round(16 * s) : top + titleH + Round(14 * s)

    MouseGetPos(&mx, &my)
    MonitorGetWorkArea(MonitorAt(mx, my), &l, &t, &r, &b)
    SetTimer(FadeOsd, 0)
    ; Top-center: Windows 11 shows its own desktop-name flyout at the bottom
    Osd.Show("NoActivate x" (l + (r - l - w) // 2) " y" (t + Round(56 * s)) " w" w " h" h)
    OsdAlpha := 245
    WinSetTransparent(OsdAlpha, Osd)
    SetTimer(StartFade, -900)
}

StartFade() {
    SetTimer(FadeOsd, 15)
}

FadeOsd() {
    global OsdAlpha
    OsdAlpha -= 35
    if (OsdAlpha <= 0) {
        SetTimer(FadeOsd, 0)
        Osd.Hide()
        return
    }
    WinSetTransparent(OsdAlpha, Osd)
}

MonitorAt(x, y) {
    Loop MonitorGetCount() {
        MonitorGet(A_Index, &l, &t, &r, &b)
        if (x >= l && x < r && y >= t && y < b)
            return A_Index
    }
    return MonitorGetPrimary()
}

; ===============================
; Help window
; ===============================

ShowHelp() {
    static g := ""
    if IsObject(g) {
        g.Show()
        return
    }
    g := Gui("-MinimizeBox", APP_NAME " - Keyboard shortcuts")
    g.BackColor := Theme.Bg
    g.MarginX := 22, g.MarginY := 18
    g.SetFont("s16 bold c" Theme.Text, "Segoe UI")
    g.SetFont(, "Segoe UI Variable Display")
    g.Add("Text", , APP_NAME)
    g.SetFont("s9 norm c" Theme.SubText, "Segoe UI")
    g.SetFont(, "Segoe UI Variable Text")
    g.Add("Text", "y+0", "Version " APP_VERSION "  ·  Virtual desktop shortcuts for Windows 11")
    g.SetFont("s10 c" Theme.Text)

    lv := g.Add("ListView", "xm y+16 w480 r9 -Hdr -Multi NoSort ReadOnly -E0x200 Background" Theme.Surface, ["Shortcut", "Action"])
    for row in [
        ["Win + 1 … 9, 0", "Switch to desktop 1–10 (creates it if missing)"],
        ["Win + Shift + 1 … 9, 0", "Move the active window to desktop 1–10"],
        ["Win + ``", "Remove empty desktops at the end"],
        ["Win + Esc", "Exit VirtuKey"],
        ["", ""],
        ["Ctrl + Win + ← / →", "Previous / next desktop (Windows)"],
        ["Ctrl + Win + D", "New desktop (Windows)"],
        ["Ctrl + Win + F4", "Close current desktop (Windows)"],
        ["Win + Tab", "Task View (Windows)"]] {
        lv.Add(, row*)
    }
    lv.ModifyCol(1, 170)
    lv.ModifyCol(2, 300)

    g.SetFont("s9 c" Theme.SubText)
    g.Add("Text", "xm y+12", "Right-click the tray icon for options. Settings: " SETTINGS_FILE)
    g.SetFont("s10 c" Theme.Text)
    btn := g.Add("Button", "xm+390 y+14 w90 h30 Default", "Close")
    btn.OnEvent("Click", (*) => g.Hide())
    g.OnEvent("Close", (*) => g.Hide())
    g.OnEvent("Escape", (*) => g.Hide())

    if Theme.Dark {
        dark := 1
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g.Hwnd, "UInt", 20, "Int*", &dark, "UInt", 4)   ; dark title bar
        for ctrl in [lv, btn]
            DllCall("uxtheme\SetWindowTheme", "Ptr", ctrl.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    }
    g.Show()
}

; ===============================
; Settings & theme
; ===============================

LoadSettings() {
    s := {}
    s.AutoCreate := IniRead(SETTINGS_FILE, "General", "AutoCreate", 1) = 1
    s.ShowOsd := IniRead(SETTINGS_FILE, "General", "ShowOsd", 1) = 1
    s.FollowMove := IniRead(SETTINGS_FILE, "General", "FollowMove", 0) = 1
    s.FirstRun := IniRead(SETTINGS_FILE, "General", "FirstRun", 1) = 1
    return s
}

SaveSetting(name) {
    try {
        DirCreate(A_AppData "\VirtuKey")
        IniWrite(Settings.%name% ? 1 : 0, SETTINGS_FILE, "General", name)
    }
}

LoadTheme() {
    light := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", 0)
    accent := RegRead("HKCU\Software\Microsoft\Windows\DWM", "AccentColor", 0xFFD47800)   ; ABGR
    r := accent & 0xFF, g := (accent >> 8) & 0xFF, b := (accent >> 16) & 0xFF
    lightAccent := (r * 299 + g * 587 + b * 114) / 1000 > 160
    return {
        Dark: !light,
        Accent: Format("{:02X}{:02X}{:02X}", r, g, b),
        AccentRef: accent & 0xFFFFFF,                     ; COLORREF (0x00BBGGRR)
        OnAccent: lightAccent ? "000000" : "FFFFFF",
        OnAccentRef: lightAccent ? 0x000000 : 0xFFFFFF,
        Bg: light ? "F3F3F3" : "202020",
        Surface: light ? "E3E3E3" : "2E2E2E",
        Text: light ? "1A1A1A" : "FFFFFF",
        SubText: light ? "5C5C5C" : "C5C5C5"
    }
}

Fatal(message) {
    MsgBox(message, APP_NAME, "Iconx")
    ExitApp
}
