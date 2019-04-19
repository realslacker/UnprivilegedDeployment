$ModuleName = 'UnprivilegedDeployment'

$Deployment = $null

if ( -not( Get-Variable -Name 'PSScriptRoot' -ErrorAction SilentlyContinue ) ) {

    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent

}

# local language data
$LocalizedDataSplat = @{
    BindingVariable = 'Messages'
    FileName        = 'Messages'
    BaseDirectory   = Join-Path $PSScriptRoot 'lang'
}
Import-LocalizedData @LocalizedDataSplat


function Initialize-UnprivilegedDeployment {

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        throw (Get-UnprivilegedDeploymentMessage DeploymentVariableUndefinedException)

    }

    # create the registry key for the deployment
    $RegistryPath = 'HKLM:\Software\{0}\{1}' -f $ModuleName, $Deployment
    New-Item -Path $RegistryPath -Force > $null

    # get the BUILTIN\Users group
    # we convert the SID to the group name in case we are
    # running in another language
    [System.Security.Principal.SecurityIdentifier]$UsersSID = 'S-1-5-32-545'
    $UsersGroup = $UsersSID.Translate([System.Security.Principal.NTAccount]).Value

    # add an access rule to the registry key
    $Acl = Get-Acl -Path $RegistryPath
    $AclRule = New-Object System.Security.AccessControl.RegistryAccessRule $UsersGroup, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
    $Acl.AddAccessRule($AclRule)
    $Acl | Set-Acl $RegistryPath

    Set-UnprivilegedDeploymentStatus -Installer -Status 'Waiting'
    Set-UnprivilegedDeploymentStatus -Client -Status 'Waiting'

}


function Set-UnprivilegedDeploymentVariable {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory)]
        [string]
        $Name,

        $Value

    )

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        throw (Get-UnprivilegedDeploymentMessage DeploymentVariableUndefinedException)

    }

    $RegistryPath = 'HKLM:\Software\{0}\{1}' -f $ModuleName, $Deployment

    $SetItemPropertySplat = @{
        Path         = $RegistryPath
        Name         = $Name
        Value        = $Value
        Force        = $true
    }
    Set-ItemProperty @SetItemPropertySplat

}


function Get-UnprivilegedDeploymentVariable {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory)]
        [string]
        $Name

    )

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        return $null

    }

    $RegistryPath = 'HKLM:\Software\{0}\{1}' -f $ModuleName, $Deployment

    $GetItemPropertySplat = @{
        Path        = $RegistryPath
        Name        = $Name
        ErrorAction = 'SilentlyContinue'
    }
    (Get-ItemProperty @GetItemPropertySplat).$Name

}


function Set-UnprivilegedDeploymentStatus {

    param(
    
        [Parameter(Mandatory, ParameterSetName='InstallerStatus')]
        [switch]
        $Installer,

        [Parameter(Mandatory, ParameterSetName='ClientStatus')]
        [switch]
        $Client,

        [Parameter(Mandatory)]
        [ValidateSet('Waiting','Ready','Running','Complete','Failed')]
        [string]
        $Status
        
    )

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        throw (Get-UnprivilegedDeploymentMessage DeploymentVariableUndefinedException)

    }

    $Splat = @{
        Name = $PSCmdlet.ParameterSetName
        Value = $Status
    }
    Set-UnprivilegedDeploymentVariable @Splat

}


function Get-UnprivilegedDeploymentStatus {

    param(
    
        [Parameter(Mandatory, ParameterSetName='InstallerStatus')]
        [switch]
        $Installer,

        [Parameter(Mandatory, ParameterSetName='ClientStatus')]
        [switch]
        $Client
        
    )

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        return $null

    }

    Get-UnprivilegedDeploymentVariable -Name $PSCmdlet.ParameterSetName

}

function Test-UnprivilegedDeploymentReady {
    
    $InstallerStatus = Get-UnprivilegedDeploymentStatus -Installer
    $ClientStatus = Get-UnprivilegedDeploymentStatus -Client

    return ( $InstallerStatus -eq 'Ready' -and $ClientStatus -eq 'Ready' )

}

function Get-UnprivilegedDeploymentMessage {

    param(

        [string]
        $Key

    )

    if ( $Messages.$Key ) {
    
        return $Messages.$Key
        
    } else {
    
        Write-Warning ( "Language key '{0}' is undefined." -f $Key )
        
        return $Key
        
    }

}

function Initialize-UnprivilegedDeploymentLog {

    [CmdletBinding()]
    param(

        [ValidateScript({ $_ = Resolve-Path $_; Test-Path -Path $_ -PathType Container })]
        [string]
        $Path = $env:TEMP,

        [ValidateNotNullOrEmpty()]
        [string]
        $FileName
        
    )

    $ErrorActionPreference = 'Stop'

    if ( $null -eq $Deployment ) {

        throw (Get-UnprivilegedDeploymentMessage DeploymentVariableUndefinedException)

    }

    # if LogPath already set this function has been called
    if ( Get-UnprivilegedDeploymentVariable -Name 'LogPath' ) { return }

    $Path = Resolve-Path $Path

    if ( -not $FileName ) { $FileName = '{0}_{1}.log' -f $Deployment, (Get-Date -Format 'yyyyMMddHHmm') }

    $LogPath = Join-Path $Path $FileName

    # store the LogPath for later use
    Set-UnprivilegedDeploymentVariable -Name 'LogPath' -Value $LogPath

    if ( -not(Test-Path -Path $LogPath) ) {

        New-Item -Path $LogPath -ItemType File -ErrorAction Stop > $null

    }

    Add-Content -Path $LogPath -Value ( (Get-UnprivilegedDeploymentMessage InitStatusLogMessage) -f $Deployment, $env:USERDOMAIN, $env:USERNAME )

}

function Get-UnprivilegedDeploymentLogContent {

    if ( $LogPath = Get-UnprivilegedDeploymentVariable -Name 'LogPath' ) {

        Get-Content -Path $LogPath -ErrorAction Stop

    } else {

        [array](Get-UnprivilegedDeploymentMessage LogPathNotSetMessage)

    }

}

function Write-UnprivilegedDeploymentLog {

    [CmdletBinding()]
    param(

        [Parameter(Position=1)]
        [string]
        $Message,

        [ValidateSet('Information', 'Warning', 'Error')]
        [string]
        $Type = 'Information'

    )

    if ( -not $PSBoundParameters.ErrorAction ) {

        $ErrorActionPreference = 'Stop'

    }

    $TimeStamp = Get-Date -f 'yyyy-MM-dd HH:mm:ss'

    $LogMessage = '[{0}] {1}: {2}' -f $TimeStamp, $Type.ToUpper(), $Message

    if ( $LogPath = Get-UnprivilegedDeploymentVariable -Name 'LogPath' ) {

        Add-Content -Value $LogMessage -Path $LogPath -ErrorAction Stop

    }

    switch ( $Type ) {

        'Information' {

            if ( Get-Command 'Write-Information' -ErrorAction SilentlyContinue ) {

                Write-Information $Message -InformationAction Continue

            } else {

                Write-Host $Message

            }

        }

        'Warning' {

            Write-Warning $Message

        }

        'Error' {

            Write-Error $Message

        }

    }
    
}

function Start-UnprivilegedDeploymentInstaller {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory)]
        [string]
        $Deployment,

        [Parameter(Mandatory, ParameterSetName='Installer')]
        [ValidateScript({ $_ = Resolve-Path $_; Test-Path -Path $_ -PathType Leaf })]
        [string]
        $FilePath,

        [Parameter(ParameterSetName='Installer')]
        [string]
        $ArgumentList,

        [Parameter(ParameterSetName='Installer')]
        [ValidateScript({ $_ = Resolve-Path $_; Test-Path -Path $_ -PathType Container })]
        [string]
        $WorkingDirectory,

        [Parameter(Mandatory, ParameterSetName='ScriptFile')]
        [ValidateScript({ $_ = Resolve-Path $_; Test-Path -Path $_ -PathType Leaf })]
        [string]
        $PS1Script,

        [Parameter(Mandatory, ParameterSetName='ScriptBlock')]
        [ScriptBlock]
        $ScriptBlock,

        [ScriptBlock]
        $SkipInstallCheck = { $false },

        [ScriptBlock]
        $PreCheck = { $true },

        [ScriptBlock]
        $PostCheck = { $true },

        [switch]
        $Force,

        [switch]
        $Reset

    )

    $ErrorActionPreference = 'Stop'

    # set the active deployment
    $Script:Deployment = $Deployment

    # if reset is supplied we change the status back to Waiting
    if ( $Reset ) {

        Set-UnprivilegedDeploymentStatus -Installer -Status Waiting
        Set-UnprivilegedDeploymentVariable -Name UserReady -Value 0

    }

    # check for a previously failed deployment
    if ( (Get-UnprivilegedDeploymentStatus -Installer) -eq 'Failed' ) {

        throw (Get-UnprivilegedDeploymentMessage 'InstallationPreviouslyFailedException')
        exit 99

    }

    # check for a previously complete deployment
    if ( (Get-UnprivilegedDeploymentStatus -Installer) -eq 'Complete' ) {

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallationAlreadyCompletedWarning) -Type Warning
        return

    }

    # on our first run we initialize the installer
    if ( -not(Get-UnprivilegedDeploymentStatus -Installer) ) {

        Initialize-UnprivilegedDeployment -Name $Deployment

    }

    # check if we can skip the installation
    if ( $SkipInstallCheck.Invoke() ) {
    
        Set-UnprivilegedDeploymentStatus -Installer -Status Complete

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallationSkippedMessage) -Type Information

        return
    
    }

    # if -Force is present we set the Client status to Ready
    # this basically allows for a headless install
    if ( $Force ) {

        Set-UnprivilegedDeploymentStatus -Client -Status Ready

    }

    # set the status to Ready once the Client is detected
    Set-UnprivilegedDeploymentStatus -Installer -Status Ready

    # wait for the client to be ready
    while ( (Get-UnprivilegedDeploymentStatus -Client) -ne 'Running' ) {
        
        for ( $i = 10; $i -gt 0; $i -- ) {

            $ProgressSplat = @{
                Activity        = Get-UnprivilegedDeploymentMessage InstallationWaitingForClientActivityMessage
                Status          = (Get-UnprivilegedDeploymentMessage InstallationWaitingForClientStatusMessage) -f $i
                PercentComplete = ( 10 - $i ) / 10 * 100
            }
            Write-Progress @ProgressSplat

            Start-Sleep -Seconds 1

        }
        
    }

    Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallationStartingMessage) -Type Information

    Write-UnprivilegedDeploymentLog -Message ((Get-UnprivilegedDeploymentMessage InstallationTypeMessage) -f $PSCmdlet.ParameterSetName) -Type Information

    try {

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage PerformingInstallerPreCheckMessage) -Type Information

        # verify that the pre-check passes
        if ( -not $PreCheck.Invoke() ) {

            Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallerFailedPreCheckError) -Type Error -ErrorAction Stop

        }

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage PerformingInstallationMessage) -Type Information
        
        # run the installer
        switch ( $PSCmdlet.ParameterSetName ) {

            'Installer' {

                $FilePath = Resolve-Path $FilePath

                $ArgumentListSplat = @{}
                if ( $ArgumentList ) { $ArgumentListSplat.ArgumentList = $ArgumentList }
    
                $WorkingDirectorySplat = @{}
                if ( $WorkingDirectory ) { $WorkingDirectorySplat.WorkingDirectory = Resolve-Path $WorkingDirectory }
            
                Start-Process -FilePath $FilePath @ArgumentListSplat @WorkingDirectorySplat -Wait
            
            }

            'ScriptFile' {

                if ( -not $PowerShellScript.Contains('\') -and (Test-Path -Path (Join-Path $PSScriptRoot $PS1Script) -PathType Leaf) ) {

                    $PS1Script = Join-Path $PSScriptRoot $PS1Script

                }
            
                $PS1Script = Resolve-Path $PS1Script

                . "$PS1Script"
            
            }

            'ScriptBlock' {

                $ScriptBlock.Invoke()

            }

        }

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage PerformingInstallerPostCheckMessage) -Type Information

        # verify that the post-check passes
        if ( -not $PostCheck.Invoke() ) {

            Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallerFailedPostCheckException) -Type Error -ErrorAction Stop

        }

        Set-UnprivilegedDeploymentStatus -Installer -Status Complete

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallationCompleteMessage) -Type Information

    } catch {

        Set-UnprivilegedDeploymentStatus -Installer -Status Failed

        Write-UnprivilegedDeploymentLog -Message (Get-UnprivilegedDeploymentMessage InstallationFailedError) -Type Error

        throw $_.Exception

    }

}


function Start-UnprivilegedDeploymentClient {

    [CmdletBinding()]
    param(

        [Parameter(Mandatory)]
        [string]
        $Deployment,
    
        [int]
        $WaitMinutes = 15,

        [ScriptBlock]
        $PreCheck = { $true },

        [switch]
        $Force

    )

    $ErrorActionPreference = 'Stop'

    # set the active deployment
    $Script:Deployment = $Deployment

    # wait for installer to be anything other than Waiting or $null
    while ( -not(Get-UnprivilegedDeploymentStatus -Installer) -or (Get-UnprivilegedDeploymentStatus -Installer) -eq 'Waiting' ) {
    
        for ( $i = 10; $i -gt 0; $i -- ) {

            $WriteProgressSplat = @{
                Activity         = Get-UnprivilegedDeploymentMessage 'InstallationWaitingActivity'
                Status           = (Get-UnprivilegedDeploymentMessage 'InstallationWaitingStatus') -f $i
                PercentComplete  = ( 10 - $i ) / 10 * 100
            }
            Write-Progress @WriteProgressSplat

            Start-Sleep -Seconds 1
    
        }
    
    }

    # check for a previously complete deployment
    if ( (Get-UnprivilegedDeploymentStatus -Installer) -eq 'Complete' ) { return }

    # tell the installer we're ready
    Set-UnprivilegedDeploymentStatus -Client -Status Ready

    # wait for user to be ready
    if ( -not [bool]( Get-UnprivilegedDeploymentVariable -Name 'UserReady' ) ) {

        while ( (New-Object -ComObject WScript.Shell).Popup(((Get-UnprivilegedDeploymentMessage 'ReadyToInstallPrompt') -f $Deployment), 0, (Get-UnprivilegedDeploymentMessage 'InstallationStartingTitle'), 36) -eq 7 ) {
    
            $SecondsToWait = $WaitMinutes * 60
            for ( $i = $SecondsToWait; $i -gt 0; $i-- ) {

                $MinutesRemaining = [math]::Ceiling(($SecondsToWait/60))
                $SecondsRemaining = $SecondsToWait - ( $MinutesRemaining * 60 )

                $WriteProgressSplat = @{
                    Activity         = Get-UnprivilegedDeploymentMessage 'InstallationWaitingForUserActivity'
                    Status           = ( Get-UnprivilegedDeploymentMessage 'InstallationWaitingForUserStatus' ) -f $WaitMinutes
                    SecondsRemaining = $i
                    PercentComplete  = ($SecondsToWait - $i) / $SecondsToWait * 100
                }
                Write-Progress @WriteProgressSplat

                Start-Sleep -Seconds 1
    
            }
    
        }

        # verify that the pre-check passes
        if ( -not $PreCheck.Invoke() ) {

            throw (Get-UnprivilegedDeploymentMessage ClientFailedPrecheckException)

        }
    
        # tell the user what's up
        $null = (New-Object -ComObject WScript.Shell).Popup((Get-UnprivilegedDeploymentMessage 'InstallationStartingMessage'), 5, (Get-UnprivilegedDeploymentMessage 'InstallationStartingTitle'), 64)

        # set UserReady for next run, if any
        Set-UnprivilegedDeploymentVariable -Name 'UserReady' -Value 1

    }

    # once the user is ready we change our status to running
    Set-UnprivilegedDeploymentStatus -Client -Status Running

    # initialize the log
    Initialize-UnprivilegedDeploymentLog

    # show the client GUI
    Show-UnprivilegedDeploymentClientGUI

    # once the client closes we change the status back to waiting
    Set-UnprivilegedDeploymentStatus -Client -Status Waiting

}

function Show-UnprivilegedDeploymentClientGUI {

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $StatusForm                      = New-Object system.Windows.Forms.Form
    $StatusForm.ClientSize           = '800,400'
    $StatusForm.text                 = Get-UnprivilegedDeploymentMessage ClientWindowTitle
    $StatusForm.TopMost              = $true

    $StatusListBox                   = New-Object system.Windows.Forms.ListBox
    $StatusListBox.text              = "listBox"
    $StatusListBox.width             = 784
    $StatusListBox.height            = 384
    $StatusListBox.location          = New-Object System.Drawing.Point(8,8)
    $StatusListBox.Font              = New-Object System.Drawing.Font('Lucida Console', 12, [System.Drawing.FontStyle]::Regular)
    $StatusListBox.Anchor            = [System.Windows.Forms.AnchorStyles]::Top, [System.Windows.Forms.AnchorStyles]::Bottom, [System.Windows.Forms.AnchorStyles]::Left, [System.Windows.Forms.AnchorStyles]::Right

    $StatusForm.controls.AddRange(@($StatusListBox))

    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 1000
    $Timer.add_tick({

        $StatusListBox.Items.Clear()
        $StatusListBox.Items.AddRange((Get-UnprivilegedDeploymentLogContent | ForEach-Object { ($_ -replace '^\[[^\]]+\]', '').Replace('INFORMATION:','').Trim() }))
        $StatusListBox.SelectedIndex = $StatusListBox.Items.Count - 1
        $StatusListBox.Update()

    })

    $StatusForm.Add_Shown({

        $StatusForm.Activate()
        $Timer.Start()
    
    })

    $StatusForm.Add_Closing({

        #Write-DeploymentStatus $StatusListBox.Items[($StatusListBox.Items.Count - 1)]

        if ( ($StatusListBox.Items[($StatusListBox.Items.Count - 1)] -notmatch (Get-UnprivilegedDeploymentMessage InstallationCompleteMessage)) -and ($StatusListBox.Items[($StatusListBox.Items.Count - 1)] -notmatch (Get-UnprivilegedDeploymentMessage InstallationFailedError)) -and ((New-Object -ComObject WScript.Shell).Popup((Get-UnprivilegedDeploymentMessage ClientWindowCloseWarning), 0, (Get-UnprivilegedDeploymentMessage ClientWindowCloseTitle), 20) -eq 7) ) { $_.Cancel = $true }
        
    })

    $StatusForm.Add_Closed({

        $Timer.Stop()

    })

    [void]$StatusForm.ShowDialog()

}