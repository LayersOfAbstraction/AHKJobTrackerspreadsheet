#Requires AutoHotkey v2
logFile := "E:\Work\ProgrammingExperiments\AutoHotKey\SkipCompanyFolder_logFile.txt"
targetDir := "C:\Users\Jordan Nash\OneDrive\Job Tracking Docs"
checkInterval := 1000 ; Time in milliseconds between checks (1 second)

try {
    currentDateTime := FormatTime(A_Now, ' yyyy/MM/dd hh:mmtt')
    
    ; This ensures that result will always be in English even if user's locale is not.
    currentDateTime := FormatTime(A_Now . ' L0x809', ' yyyy/MM/dd hh:mmtt')
    FileAppend("Script started on " currentDateTime "`n", logFile)
    SetTimer CheckDirectory, checkInterval
} 
catch Error as err { 
    FileAppend("Error: " err.Message "`n", logFile) 
}

CheckDirectory() {
    static hwnd := 0
    static navigated := false ; Declare static variable within the function

    try {
        ; Find the File Explorer window with the specified title
        hwnd := WinExist("ahk_class CabinetWClass ahk_exe explorer.exe")
        if !hwnd {
            return
        }

        for window in ComObject("Shell.Application").Windows {
            if (window.HWND == hwnd) {
                currentDir := StrReplace(window.LocationURL, "file:///", "")
                break
            }
        }             

        currentDir := StrReplace(currentDir, "%20", " ") ; Decode URL-encoded spaces
        currentDir := StrReplace(currentDir, "/", "\") ; Convert to backslashes for consistency

        ; Check if the current directory starts with the target directory
        if (InStr(currentDir, targetDir) = 1) {
            if (!navigated) {
                FileAppend("Current directory starts with the target directory" "`n", logFile)
                subfolders := []
                Loop Files, currentDir "\*.*", "D"  ; D = directories only
                {
                    subfolders.Push(A_LoopFileFullPath)
                }
                if (subfolders.Length = 1) {  ; Check if only one subfolder is present
                    folder := subfolders[1]
                    ; Navigate to the subfolder within the same window
                    for window in ComObject("Shell.Application").Windows {
                        if (window.HWND == hwnd) {
                            window.Navigate(folder)
                            break
                        }
                    }
                    FileAppend("Navigated to: " folder "`n", logFile)
                } else {
                    FileAppend("Manual navigation required: multiple subfolders found" "`n", logFile)
                }
            }
        } else {
            ; Reset the navigated flag if the current directory is not within the target directory
            navigated := false
            FileAppend("Current directory does not start with the target directory" "`n", logFile)
        }
    } catch Error as err { 
        FileAppend("Error: " err.Message "`n", logFile) 
    }
}