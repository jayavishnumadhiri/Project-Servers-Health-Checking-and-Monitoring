$time=New-ScheduledTaskTrigger -Once -At 01:00 -RepetitionInterval (New-TimeSpan -hours 12) -RepetitionDuration (New-TimeSpan -Days 9999)

$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "C:\Users\Venkatram\Desktop\Servers Health Checking and Monitoring.ps1" #here give path of the script for which wants to create scheduled task
 
Register-ScheduledTask -TaskName "Healthcheck" -Trigger $time -Action $Action
