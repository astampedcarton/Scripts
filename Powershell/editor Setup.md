# Setup of editor

The following is a step through on setting up of running a powershell script in either VSCodium/VScode/Notepad++

## VScodium/Vscode
Create a new task or if tasks already exist add to the exiting file.

### Example:
This setup will create a task that will create a new terminal for the Todo. And will contain the name of the Label in the task below.

Note that the Terminal will display the following:

Terminal will be reused by tasks, press any key to close it. This is not something that can be fixed. Once you press any key the terminal will reset and close.
```
{
    "version": "2.0.0",
    "tasks": [
        {
        "label": "TODOs in active file",
        "type": "shell",
        "command": "powershell",
        "args": [
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File",
            "C:/Users/userid/OneDrive/Programming/Powershell/listtodo.ps1",
            "-File",
            "${file}"
        ],
        "problemMatcher": [],
        "presentation": {
            "reveal": "always",
            "panel": "dedicated",
            "close": false,
            "focus": true            
        }
        }
    ]
}
```

## Notepad++
For Notepad++ there are two ways, but NppExec was found to tbe the best as the setup via Run launch an external terminal and this wasn't what was desired.

### NppExec Setup
Enter the following into NppExec updating this portion:

```
-File "C:\Users\userid\OneDrive\Programming\Powershell\listTodo.ps1"
```

as needed.
```
 CLS 
 cmd /c pwsh.exe -NoProfile -ExecutionPolicy Bypass  -File "C:\Users\userid\OneDrive\Programming\Powershell\listTodo.ps1" "$(FULL_CURRENT_PATH)"
```
CLS is to clear the console. It was found to have everything in one line to work the best.

### Run command
Open the Run option
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\userid\OneDrive\Programming\Powershell\listTodo.ps1" "$(FULL_CURRENT_PATH)"
```

Note that the script will need to be updated by placing this piece back

```powershell
Read-Host "Press Enter to close"
```

To Execute the script on any file being opened
 - Make sure that NppEventExec is installed.
 - Make sure that there is no active scripts be executed. Otherwise it will seem like NPP is not responding.
 - Add a new Rule that executes when a new file is opened i.e. NPPN_FILEOPENED. The command must be the same as the name you gave in NppExec.
 - To limit the script to just example py files. Update the regex portion.

**IMPORTANT:**
- Make sure that the Event is enabled.
- Make sure that the Event runs in the background.

  <img width="902" height="180" alt="image" src="https://github.com/user-attachments/assets/9a9665fd-99ee-441c-a618-5b96e7b22afd" />
