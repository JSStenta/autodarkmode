Add-Type -AssemblyName System.Device #Required to access System.Device.Location namespace
$GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher #Create the required object
$GeoWatcher.Start() #Begin resolving current locaton
$wait = 5
while (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied')) {
    'Wait...'
    Start-Sleep -Seconds $wait #Wait for discovery.
    $wait += 5
}

if ($GeoWatcher.Permission -eq 'Denied') {
    Write-Error 'Access Denied for Location Information'
}
else {
    #Get location
    while ($GeoWatcher.Status -ne 'Ready') {
        Start-Sleep -Milliseconds 100 #Wait for discovery.
    }
    $lat = $GeoWatcher.Position.Location.Latitude
    $lng = $GeoWatcher.Position.Location.Longitude

    #Get daylight
    $url = "https://api.sunrise-sunset.org/json?lat=" + $lat + "&lng=" + $lng
    $Daylight = (Invoke-RestMethod $url).results
    $Sunrise = ($Daylight.Sunrise | Get-Date).ToLocalTime()
    $Sunset = ($Daylight.Sunset | Get-Date).ToLocalTime()
    #$Tomorrow = (((Invoke-RestMethod ($url + '&date=tomorrow')).results).Sunrise | Get-Date).AddDays(1).ToLocalTime()
    $Now = Get-Date
    $Modo = $Sunrise.CompareTo($Now) * $Sunset.CompareTo($Now)

    #Change the mode
    if ($Modo -lt 0) {
        $Name = "LightMode"
        Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value 1
        $Time = $Sunset
    }
    else {
        $Name = "DarkMode"
        Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value 0
    }
    #Schedule next time
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy bypass -noprofile -file $PSCommandPath"
    $trigger = New-ScheduledTaskTrigger -Once -At $Time
    $principal = New-ScheduledTaskPrincipal -UserId 'Administrator' -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet #-StartWhenAvailable
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal

    Unregister-ScheduledTask -TaskName $Name -Confirm:$false
    Register-ScheduledTask $Name -InputObject $task
        
    #Log data
    Add-Content logAutoDark.txt ($Now.ToString() + $line + ($Daylight -split ('; ') -join $ln) + $line + 'ScheduledTask: Name=' + $Name + ", trigger=" + $Time)
}

    

$settings = New-ScheduledTaskSettingsSet #-StartWhenAvailable
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings

Unregister-ScheduledTask -TaskName $Name -Confirm:$false
Register-ScheduledTask $Name -InputObject $task    
