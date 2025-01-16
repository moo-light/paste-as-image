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
    bitmap := wc.GetBitmap()
    if !bitmap {
        SendInput "^v" ; send default pasting behavior
        return
    }
    savePath := explorerGetPath()
    saveImage(bitmap, savePath)
}

; get current explorer path
; https://www.reddit.com/r/AutoHotkey/comments/10fmk4h/get_path_of_active_explorer_tab/
explorerGetPath(hwnd := WinExist("A")) {
    activeTab := 0
    try activeTab := ControlGetHwnd("ShellTabWindowClass1", hwnd)
    for w in ComObject("Shell.Application").Windows {
        if (w.hwnd != hwnd)
            continue
        if activeTab {
            static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
            shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
            ComCall(3, shellBrowser, "uint*", &thisTab := 0)
            if (thisTab != activeTab)
                continue
        }
        return w.Document.Folder.Self.Path
    }
    return false
}

saveImage(bitmap, savePath) {
    label := Format("Saving image to: {1}", savePath)
    wc.SaveBitmap("./temp.png", "png")
    ;; creating save image gui
    saveImageGui := Gui()
    saveImageGui.Add("Text", "w270", label)
    picture := saveImageGui.Add("Picture", "x55 w200 h-1", "temp.png")
    saveImageGui.Add("Text", "x5", "Please enter file name:")
    saveImageGui.Add("Text", "xp+210", "image type")

    inpName := saveImageGui.Add("Edit", "x5 w200")
    ; inpName.AddText("image name")
    cboType := saveImageGui.Add("DropDownList", "xp+210 w90 Choose1", ["png", "jpeg", "bmp"])
    inpName.Value := "image"
    btnSave := saveImageGui.Add("Button", "x20 w90", "Save")
    btnCancel := saveImageGui.Add("Button", "xp180 w90", "Cancel")
    saveImageGui.Show("w310 Center")

    ; setting up events
    picture.OnEvent("Click", (*) => PictureClick())
    btnSave.OnEvent("Click", (*) => SaveClick())
    btnCancel.OnEvent("Click", (*) => CancelClick())
    saveImageGui.OnEvent("Close", (*) => CancelClick())
    ; save image function
    SaveClick() {
        fileName := RTrim(inpName.Value)
        imageType := cboType.Value
        switch (cboType.Value) {
            case 1: ; png
                imageType := "png"
            case 2: ; jpeg
                imageType := "jpeg"
            case 3: ; webp
                imageType := "bmp"
        }
        fullPath := savePath . "\" . fileName . "." . imageType

        if FileExist(fullPath) {
            MsgBox("File already exists! Try Again")
            return
        }
        wc.SaveBitmap(fullPath, imageType)
        saveImageGui.Destroy()
        FileDelete("temp.png")
    }
    ; cancel function
    CancelClick() {
        saveImageGui.Destroy()
        FileDelete("temp.png")
    }
    ; open image in default image viewer
    PictureClick() {
        ;Todo: open image in default image viewer
        Run("explorer.exe " . "temp.png")
        return
    }

}
