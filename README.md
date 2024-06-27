# Useful COM & WMI snippets

This README contains some Nim and Powershell snippets for (ab)using COM and WMI for various useful purposes.
Powershell snippets should work as is, for Nim import [winim/com](https://github.com/khchen/winim).
I may add further examples in the future.

### Speak thru the microphone

```
$SpeakObject = new-object -com SAPI.SpVoice
$SpeakObject.volume = 100
$SpeakObject.voice = $SpeakObject.GetVoices().item(1) # <- You can even change to a woman voice
$speak_object.Speak("hello")
```
```
var obj = CreateObject("SAPI.SpVoice")
obj.volume = 100
obj.speak("hello")
```

### Get username, computer name, and domain
```
$WN = new-object -com wscript.network
Write-Host $WN.UserDomain
Write-Host $WN.UserName
Write-Host $WN.ComputerName
```
```
var obj = CreateObject("WScript.Network")
echo obj.userDomain 
echo obj.computerName
echo obj.userName
```

### Get OS Info
```
$OsInfo = Get-CimInstance Win32_OperatingSystem
Write-Host $OsInfo.Caption
Write-Host $OsInfo.Version
Write-Host $OsInfo.SerialNumber
```
```
var wmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
var query = "SELECT * FROM Win32_OperatingSystem"
for osInfo in wmi.execQuery(query):
    echo "Operating System Name: ", osInfo.Caption
    echo "Version: ", osInfo.Version
```

### Enumerate processes
```
Get-CimInstance win32_process
```
```
var wmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
for i in wmi.execQuery("select * from win32_process"):
    echo i.handle, " | ", i.name
```

### Enumerate installed software
```
Get-CimInstance Win32_Product | Select-Object Name, Vendor, Version
```
```
var wmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
for i in wmi.execQuery("SELECT * FROM Win32_Product"):
    echo "Name: ", i.Name
    echo "Version: ", i.Version
    echo "Vendor: ", i.Vendor
```

### Enumerate AVs
```
Get-CimInstance -Namespace root\securitycenter2 AntiVirusProduct | select displayName
```
```
var wmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\securitycenter2")
for i in wmi.execQuery("SELECT displayName FROM AntiVirusProduct"):
    echo "AntiVirusProduct: ", i.displayName
```

### Spawn process
> This technique spawns the child process under WmiPrvSE.exe, hence breaks causality
```
Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="path\to\malware.exe arg1 arg2"}
```
```
var wmi = GetObject(r"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:Win32_Process")
wmi.Create("notepad.exe")
```