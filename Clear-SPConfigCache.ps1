$ErrorActionPreference = "Stop";

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent());
$isAdminUser = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

if (-not $isAdminUser)
{
   Write-Host "You need elevated administrator privileges to run this script." -ForegroundColor Red;
   Exit;
}

if ((Get-PSSnapin | ? { $_.Name -eq "Microsoft.SharePoint.PowerShell" }) -eq $null)
{
    Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

Write-Host "Stopping SharPoint Timer Service ..." -ForegroundColor Yellow;
try
{
    Stop-Service "SPTimerV4";
}
catch
{
    Write-Host "Failed to stop SharePoint Timer Service." -ForegroundColor Red;
    Exit;
}
Write-Host "Succeeded to stop SharPoint Timer Service." -ForegroundColor Green;

try
{
    $farm = Get-SPFarm;
    $configDB = Get-SPDatabase | ? { $_.Name -eq $farm.Name };
    if ($configDB -eq $null)
    {
        throw "Failed to get information of the SharePoint configuration database.";
    }

    Write-Host ("Clearing configuration cache files of '{0}' ({1}) ..." -F $configDB.Name, $configDB.Id) -ForegroundColor Yellow;

    $path = ("{0}\Microsoft\SharePoint\Config\{1}" -F $env:ALLUSERSPROFILE, $configDB.Id);
    $colXml = Get-ChildItem $path -Filter *.xml;

    Write-Host ("Deleting {0} XML files ..." -F $colXml.Count) -ForegroundColor Yellow;
    $colXml | Remove-Item -Force;
    Write-Host ("Deleted." -F $colXml.Count) -ForegroundColor Green;

    Write-Host "Modifying cache.ini file ..." -ForegroundColor Yellow;
    $modified = $False;
    for ($numRetryAttempt = 0; $numRetryAttempt -lt 10; $numRetryAttempt++)
    {
        try
        {
            "1" | Set-Content -Path (Join-Path -Path $path -ChildPath "cache.ini");
            $modified = $True;
        }
        catch [System.IOException]
        {
            Write-Host ("cache.ini file is locked, retry modifying in 5 seconds. ({0} time(s), the maximum number of retry attempts: 10)" -F ($numRetryAttempt + 1)) -ForegroundColor Yellow;
            Start-Sleep 5;
        }
    }

    if (-not $modified)
    {
        throw "Failed to modify cache.ini file.";
    }
    else
    {
        Write-Host "Succeeded to modify cache.ini file." -ForegroundColor Green;
    }

    Write-Host ("Succeeded to clear configuration cache files of '{0}' ({1})." -F $configDB.Name, $configDB.Id) -ForegroundColor Green;
}
catch
{
    Write-Host $error[0] -ForegroundColor Red;
}
finally
{
    Write-Host "Restarting SharPoint Timer Service ..." -ForegroundColor Yellow;

    try
    {
        Start-Service "SPTimerV4";
        Write-Host "Succeeded to restart SharPoint Timer Service." -ForegroundColor Green;
    }
    catch
    {
        Write-Host "Failed to restart SharePoint Timer Service. Please restart the service manually at 'Services' console." -ForegroundColor Red;
    }
    finally
    {
        if ($farm.BuildVersion.Major -eq 16 -and $farm.BuildVersion.Build -lt 10000)
        {
            Write-Host "Run 'iisreset /noforce' manually." -ForegroundColor Yellow;
        }
    }
}
