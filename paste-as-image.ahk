#Requires AutoHotkey v2.0
#SingleInstance force
#Include libs/WinClipAPI.ahk
#Include libs/WinClip.ahk

wc := WinClip()
$~^v:: {
    if !WinActive("ahk_exe explorer.exe") {
        return
    }
    ; check if clipboard is an snipped image

    if !wc.GetBitmap() {
        SendInput "^v" ; send default pasting behavior
        return
    }
    savePath := explorerGetPath()
    Label:
    fileName := InputBox(Format("Saving to {1}`n Please enter the filename:", savePath) , , , "image")
    if fileName.Result == "Cancel" {
        MsgBox("Canceled save!")
        return
    }
    fullPath := savePath . "\" . RTrim(fileName.Value) . ".png"

    if FileExist(fullPath) {
        MsgBox("File already exists!")
        goto('Label')
    }
    wc.SaveBitmap(fullPath, "png")
}

; get current explorer path
; https://www.reddit.com/r/AutoHotkey/comments/10fmk4h/get_path_of_active_explorer_tab/
explorerGetPath(hwnd:=WinExist("A")) {
    activeTab := 0
    try activeTab := ControlGetHwnd("ShellTabWindowClass1", hwnd)
    for w in ComObject("Shell.Application").Windows {
        if (w.hwnd != hwnd)
            continue
        if activeTab {
            static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
            shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
            ComCall(3, shellBrowser, "uint*", &thisTab:=0)
            if (thisTab != activeTab)
                continue
        }
        return w.Document.Folder.Self.Path
    }
    return false
}
