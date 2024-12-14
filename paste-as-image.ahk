#Requires AutoHotkey v2.0
#SingleInstance force
#Include libs/WinClipAPI.ahk
#Include libs/WinClip.ahk

wc := WinClip()
~^v:: {
    if !WinActive("ahk_exe explorer.exe") {
        return
    }
    ; check if clipboard is an snipped image
    if !wc.GetBitmap() {
        return
    }
    fileName := InputBox("Enter Filename", "Please enter the filename:",,"image")
    if !fileName {
        MsgBox("Canceled save!")
        return
    }
    savePath := explorerGetPath()
    fullPath := savePath . "\" . fileName.Value . ".jpeg"
    if FileExist(fullPath) {
        MsgBox("File already exists!")
        return
    }
    wc.SaveBitmap(fullPath, "jpeg")
}

explorerGetPath(hwnd := 0) { ; https://www.autohotkey.com/boards/viewtopic.php?p=387113#p387113
    static winTitle := 'ahk_class CabinetWClass'
    hWnd ? explorerHwnd := WinExist(winTitle ' ahk_id ' hwnd)
        : ((!explorerHwnd := WinActive(winTitle)) && explorerHwnd := WinExist(winTitle))
    if explorerHwnd
        for window in ComObject('Shell.Application').Windows
            try if window && window.hwnd && window.hwnd = explorerHwnd
                return window.Document.Folder.Self.Path
    return False
}
