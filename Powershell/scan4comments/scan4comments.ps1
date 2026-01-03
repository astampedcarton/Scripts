#Small script to check for the existence of comments in code and if these comments 
#Contains SAS code or not. Due to the complexity of what could be Code there is the 
#possibility of false positives.

#To allow for the script to be ran from the command line

param([Parameter(Mandatory)]
      [string]$file
) 

#Count the nr of data/proc
$data_pat = "^\\s*DATA\\s+[^;]+;"        # DATA step
$data_pat = "^\\s*DATA\\s+[^;]+;"        # DATA step
$proc_pat = "^\\s*PROC\\s+[^;]+;"        # PROC step
$blkmacr_pat = "^\\%macro\\s+[^;]+"       # Macro declaration 

$code_patterns = '(?i)\b(?:data|proc|set|merge|if|then|else|run;|quit;|input|put|retain|
                  keep|drop|rename|by|output|do|end|%macro|%mend|%let|%if|%do|while|
                  unitl|%sysfunc|tranwrd|%bquote|substr|%str|%nrstr|%quote|superq|scan|
                  cat|catt|cats|translate|findw|findc|propcase|lowercase|format|informat|
                  &SYSADDRBITS|&SYSBUFFR|&SYSCC|&SYSCHARWIDTH|&SYSCMD|&SYSDATASTEPPHASE|
                  &SYSDATE|&SYSDATE9|&SYSDAY|&SYSDEVIC|&SYSDMG|&SYSDSN|&SYSENCODING|&SYSENDIAN|
                  &SYSENV|&SYSERR|&SYSERRORTEXT|&SYSFILRC|&SYSHOSTINFOLONG|&SYSHOSTNAME|
                  &SYSINDEX|&SYSINFO|&SYSJOBID|&SYSLAST|&SYSLCKRC|&SYSLIBRC|&SYSLOGAPPLNAME|
                  &SYSMACRONAME|&SYSMENV|&SYSMSG|&SYSNCPU|&SYSNOBS|&SYSODSESCAPECHAR|&SYSODSPATH|
                  &SYSPARM|&SYSPBUFF|&SYSPRINTTOLIST|&SYSPRINTTOLOG|&SYSPROCESSID|&SYSPROCESSMODE|
                  &SYSPROCESSNAME|&SYSPROCNAME|&SYSRC|&SYSSCP|&SYSSCPL|&SYSSITE|&SYSSIZEOFLONG|
                  &SYSSIZEOFPTR|&SYSSIZEOFUNICODE|&SYSSTARTID|&SYSSTARTNAME|&SYSTCPIPHOSTNAME|
                  &SYSTIME|&SYSTIMEZONE|&SYSTIMEZONEIDENT|&SYSTIMEZONEOFFSET|&SYSUSERID|&SYSVER|
                  &SYSVLONG|&SYSVLONG4|&SYSWARNINGTEXT)\b\s*[\w;]'

#Different SAS comment patterns that must be searched
$snglpat = "(?m)^\\s*\\*[^;]*?;"
$multpat = "(?m)^\\s*\\*.*?;"
$blkpat = "/\\*.*?\\*/"

#Scan the file for potential comments with SAS Code in it.
$blkcoms = Get-Content $file | Select-String -pattern $blkpat
$sngcoms = Get-Content $file | Select-String -pattern $snglpat 
$mulcoms = Get-Content $file | Select-String -pattern $multpat

#Combine all results into one
$allcoms= $blkcoms + $sngcoms + $mulcoms

#Scan through all the comments looking for what might be SAS Code in comments 
#Which should not be present. Only print if the line contains a potential SAS Code entry
ForEach ($blkm in $allcoms){  
  $comtxt = $blkm.Line -replace  '^/\*','' -replace '\*/$', ''

  if ($comtxt -match $code_patterns){
    Write-Host "$($blkm.LineNumber) $($comtxt)"
  }
}  

#Ensure that the terminal window is kept active. 
#this can be removed without an effect to the script

# Read-Host "Press Enter to close"

