$ErrorActionPreference = "silentlycontinue"
$tasks = Get-Process -Name EXCEL
foreach($task in $tasks) {
    Stop-Process $task
}
# taskkill $task
cscript vbac.wsf combine
$path = "./bin/" + $Args[0]
Invoke-Item $path