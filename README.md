# powershell
(Soon to be) collection of some (personally) useful powershell snippets/functions that I don't mind being public.

# profile.ps1
What I usually have in my personal powershell profile. Add it to the file located in $profile (usually %userprofile%\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1).

Contains:
- Path to a functions.ps1-file that is loaded automatically. Path is saved in a global variable called CUSTOMFUNCTIONS so that the functions can be reloaded by simply typing '. $CUSTOMFUNCTIONS'
- Auto-connection to Exchange if $env:ExchangeInstallPath is present.
- Logging of every completed command to %userprofile%\Documents\WindowsPowerShell\posh_history_COMPUTERNAME.txt
- Display of current time and duration of last completed command in the prompt.
- Display of certain options like executionPolicy, current user, current IP, etc. when starting a new session.