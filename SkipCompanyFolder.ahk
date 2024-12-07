#Requires AutoHotkey v2
_logFile := "SkipCompanyFolder_logFile.txt"
targetDir := "C:\Users\Jordan Nash\OneDrive\Job Tracking Docs"
checkInterval := 1000 ; Time in milliseconds between checks (1 second)

try {    
    ; This ensures that result will always be in English even if user's locale is not.
    _currentDateTime := FormatTime(A_Now . ' L0x809', ' yyyy/MM/dd hh:mmtt')
    FileAppend("Script started" _currentDateTime "`n", _logFile)
    SetTimer CheckDirectory, checkInterval
} 
catch Error as err { 
    FileAppend("Error: " err.Message "`n", _logFile) 
}

CheckDirectory() {
    static hwnd := 0

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
        if (InStr(currentDir, targetDir) = 1) 
        {
            FileAppend("Current directory starts with the target directory" _currentDateTime "`n", _logFile)                
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
                FileAppend("Navigated to: " folder " " _currentDateTime "`n", _logFile)
            } 
            else {
                FileAppend("Manual navigation required: multiple subfolders found" _currentDateTime "`n", _logFile)                              
            }            
        } 
        else
        {
            FileAppend("Current directory does not start with target directory" _currentDateTime "`n", _logFile)
        }

    } catch Error as err { 
        FileAppend("Error: " err.Message " " _currentDateTime "`n", _logFile) 
    }
}