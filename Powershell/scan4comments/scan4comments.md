# scan4comments

A simple powershell script that scans a SAS File for comments and prints those comments containing potential SAS code.

Due to the complexity of what could be Code there is the possibility of false positives.This is expected and the final decision on if they should be present will be left to the user.

The purpose of this script is to flag as much potential cases as possible. 

## Editor setup
This script works the best when automated in VScodium/VScode/Notepad++:
For editor setup see: [https://github.com/astampedcarton/GeneralTools/tree/main/Powershell]

## Command line execution

The following is an example of executing the script in the command line.

Change directory to the location of the scripts.
```
cd "C:\Users\userid\OneDrive\Programming\Powershell\scan4comments\"
```

Execute the script as per example below.
```
C:\Users\userid\OneDrive\Programming\Powershell\scan4comments> clear | .\scan4comments.ps1 "G:\SAS\all_data_for_subject.sas"
```
