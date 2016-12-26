Function Get-OfficeInstall
{
<#
.Synopsis
Gets office install information on a machine

.DESCRIPTION
Gets office install information on a machine

.NOTES   
Name: Get-OfficeInstall
Author: Jacob Boykin

.PARAMETER ComputerName
The computer in which the office install info will be fetched from

.EXAMPLE
Get-OfficeInstall CLIENT01

Description:
Will fetch office install info on CLIENT01

.EXAMPLE
Get-OfficeInstall CLIENT01, CLIENT02, CLIENT03

Description:
Will fetch office install info on CLIENT01, CLIENT02 and CLIENT03

.EXAMPLE
'CLIENT01' | Get-OfficeInstall

Description:
Will fetch office install info on CLIENT01 using the piped input

#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [string[]] $ComputerName = $env:COMPUTERNAME
    )

    begin
    {
        $results = @()
        $error.clear()

        $Baseline = @{
            "Update_Channel" = "Current"
            "Installed_Channel" = "Current"
            "Installed_Version" = "16.0.7571.2075"
            "Automatic_Updates" = "True"
            "SCCM_Updates" = "False"
            "Update_Path" = "\\Server\Share"
        }
    }
    
    process
    {
        foreach ($computer in $ComputerName)
        {
            $result = New-Object System.Object

            try
            {
                # FETCH OFFICE INSTALL INFO =================================================================
                # ===========================================================================================

                if (!(Test-Connection -Computername $computer -BufferSize 16 -Count 2 -Quiet))
                {
                    throw "computerOffline"
                }

                $ErrorActionPreference = "Stop"

                Write-Host -ForegroundColor Green "`nConnected to " $Computer "..."
                Write-Host -ForegroundColor Green "`nFetching info from registry..."

                $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Computer)
                
                # Check if Office is installed
                $Key = $Reg.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us")
                if (!$Key)
                {
                    throw "officeNotInstalled"
                }

                # Get info from Configuration key
                $Key = $Reg.OpenSubKey("SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
                $VersionToReport = $Key.GetValue("VersionToReport")
                $SharedComputerLicensing = [boolean]$Key.GetValue("SharedComputerLicensing")

                # If Office is managed by GP
                $Key = $Reg.OpenSubKey("SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate")
                if ($Key)
                {
                    # Get info from Policies key
                    $UpdateBranch = $Key.GetValue("updatebranch")
                    $EnableAutomaticUpdates = [boolean]$key.GetValue("enableautomaticupdates")
                    $OfficeMGMTCOM = [boolean]$key.GetValue("officemgmtcom")
                    $UpdatePath = $key.GetValue("updatepath")
                }
                else
                {
                    Write-Host -ForegroundColor Yellow "`nOffice 365 is not managed by Group Policy`n"
                    $UpdateBranch=$EnableAutomaticUpdates=$OfficeMGMTCOM=$UpdatePath = 'not set'
                }



                # ATTEMPT TO CONVERT VERSION NUM TO CHANNEL ===================================================
                # ===========================================================================================

                $ErrorActionPreference = "SilentlyContinue"

                Write-Host -ForegroundColor Green "`nAttempting to fetch latest build info from MS..."

                # Fetch latest build numbers from MS
                $Request = Invoke-WebRequest "https://technet.microsoft.com/en-us/library/mt592918.aspx"
                if ($Request.StatusCode -eq '200')
                {
                    Write-Host -ForegroundColor Green "`nParsing build info..."

                    $VersionTable = $Request.AllElements | Where {$_.tagname -eq "td"}

                    # Trim version number to parse easily
                    $VersionToReportTrimmed = [string]$VersionToReport.Substring(5)

                    # Look for version number
                    $VersionNumberInfo = $VersionTable -match $VersionToReportTrimmed

                    # Check what channel is linked to version number
                    if (($VersionNumberInfo -match 'Current') -and ($VersionNumberInfo -notmatch 'Deferred'))
                    {
                        $InstalledChannel = 'Current'
                    }
                    elseif (($VersionNumberInfo -match 'First Release Deferred') -and ($VersionNumberInfo -notmatch 'Current'))
                    {
                        $InstalledChannel = 'First Release Deferred'
                    }
                    elseif (($VersionNumberInfo -match 'Deferred') -and ($VersionNumberInfo -notmatch 'First Release'))
                    {
                        $InstalledChannel = 'Deferred'
                    }
                    else 
                    {
                        $InstalledChannel = 'Could not determine channel'
                    }
                }
                
                # ANALYZE SETTINGS ==========================================================================
                # ===========================================================================================

                $BaslineCheck = $true

                Write-Host -ForegroundColor Green "`nComparing to basline settings...`n"

                if ($Baseline.Update_Channel -ne $UpdateBranch)
                {
                    Write-Host -ForegroundColor Yellow "Updated Channel is "$UpdateBranch "| Expected " $Baseline.Update_Channel
                    $BaslineCheck = $false
                }

                if ($Baseline.Installed_Channel -ne $InstalledChannel)
                {
                    Write-Host -ForegroundColor Yellow "Installed Channel is "$InstalledChannel "| Expected " $Baseline.Installed_Channel
                    $BaslineCheck = $false
                }

                if ($Baseline.Installed_Version -ne $VersionToReport)
                {
                    Write-Host -ForegroundColor Yellow "Installed Version is "$VersionToReport "| Expected " $Baseline.Installed_Version
                    $BaslineCheck = $false
                }

                if ($Baseline.Automatic_Updates -ne $EnableAutomaticUpdates)
                {
                    Write-Host -ForegroundColor Yellow "Automatic Updates enabled is "$EnableAutomaticUpdates "| Expected " $Baseline.Automatic_Updates
                    $BaslineCheck = $false
                }

                if ($Baseline.SCCM_Updates -ne $OfficeMGMTCOM)
                {
                    Write-Host -ForegroundColor Yellow "SCCM Updates Enabled is "$OfficeMGMTCOM "| Expected " $Baseline.SCCM_Updates
                    $BaslineCheck = $false
                }

                if ($Baseline.Update_Path -ne $UpdatePath)
                {
                    Write-Host -ForegroundColor Yellow "Update Path is "$UpdatePath "| Expected " $Baseline.Update_Path
                    $BaslineCheck = $false
                }

            }
            catch
            {
                Switch -Wildcard ($Error[0].Exception)
                {
                    "*computerOffline*"
                    {
                        Write-Host -BackgroundColor Black -ForegroundColor Red "`n$computer appears to be offline!`n"      
                    }
                    "*System.Net.WebException*"
                    {
                        Write-Host -ForegroundColor Yellow "`nFailed to fetch build info!"
                        Write-Host -ForegroundColor Yellow "HTTP Status Code: " $Request.StatusCode
                        $InstalledChannel = "Failed to fetch"
                    }
                    "*officeNotInstalled*"
                    {
                        Write-Host -ForegroundColor Yellow "`nOffice 365 is not installed!`n"
                    }
                    Default
                    {
                        Write-Host -BackgroundColor Black -ForegroundColor Red "`nUnable to fetch office install!`n"
                        Write-Host -BackgroundColor black -ForegroundColor Red $Error[0].Exception
                    }
                }

            }
            finally
            {
                if (!$error)
                {
                    Write-Host -ForegroundColor Green "`nGathering results..."
                    $result | Add-Member -MemberType NoteProperty -Name "Computer" -Value "$Computer"
                    $result | Add-Member -MemberType NoteProperty -Name "Baseline Check Passed" -Value "$BaslineCheck"
                    $result | Add-Member -MemberType NoteProperty -Name "Update Channel" -Value $UpdateBranch
                    $result | Add-Member -MemberType NoteProperty -Name "Installed Channel" -Value $InstalledChannel
                    $result | Add-Member -MemberType NoteProperty -Name "Installed Version" -Value $VersionToReport
                    $result | Add-Member -MemberType NoteProperty -Name "Shared Computer Licensing" -Value $SharedComputerLicensing
                    $result | Add-Member -MemberType NoteProperty -Name "Automatic Updates Enabled" -Value $EnableAutomaticUpdates
                    $result | Add-Member -MemberType NoteProperty -Name "SCCM Updates Enabled" -Value $OfficeMGMTCOM
                    $result | Add-Member -MemberType NoteProperty -Name "Update Path" -Value $UpdatePath
                    $results += $result
                }
                
            }
        }
    }
    
    end
    {
        $results
    }
}