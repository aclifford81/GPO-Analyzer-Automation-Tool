<#
.SYNOPSIS

    Enumerates GPOs in use in on premises AD, converts them to XML and then
    uploads them to GPO Analitics in MEMC


.DESCRIPTION

    This script automates the conversion and upload of GPOs to the GPO Analitics
    tool in MEMC for conversion to MEM policies

    To use this script, you will need:

    1. An Internet connection
    2. Line of site to a domain controller
    3. MEMC administrator account, or sufficest privlidges to upload via Graph
    4. The ability to install MS-Graph and AD PS modules if not already installed.
   

.EXAMPLE
    .\Upload-GPOReportsToMEM.ps1

    Requires and checks for MS graph and AD PS modules
    
    Will prompt for MEMC admin credentials to ingest GPOs.

.NOTES

    The purge feature removes all uploaded GPOs from GPO Analitics but will not
    delete any policies created from the migration of them.  

#>

<#
Revision history
    1.0.0  - Initial release  ActiveDirectory
#>
Function Get-CurrentLineNumber { 
        # Simply Displays the Line number within the script.
        [string]$line = $MyInvocation.ScriptLineNumber
        $line.PadLeft(4,'0')
    }

Function Update-Log {
    
        [CmdletBinding(DefaultParameterSetName = 'DefParamSet',
                       SupportsShouldProcess = $true,
                       PositionalBinding = $false,
                       ConfirmImpact = 'Medium')]
        Param (
            # This is the Message that will be logged into the Log File
            [Parameter(Mandatory = $False,
                       ValueFromPipeline = $true,
                       ValueFromPipelineByPropertyName = $true,
                       Position = 0)]
            $Message,
            # If you also want to send it to a different Error Log, include this switch.

            [Parameter(ParameterSetName = 'ParamSet02')]
            [switch]$IncludeErrorLog,
            # To include a section break in your log file, include this switch.

            # A section break is three blank lines in the log file.

            [Parameter(ParameterSetName = 'ParamSet03')]
            [switch]$SectionBreak,
            # To Add a minor break in your log file, include this switch.

            # A minor break is ************* in the log file.

            [Parameter(ParameterSetName = 'ParamSet04')]
            [switch]$MinorBreak
        )

        Begin { }
        Process {
            if (-not $PSBoundParameters.ContainsKey('Verbose')) {
                $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
            }
            Try {
                If ($SectionBreak) {
                    Add-Content -Value "`n`n`n" -Path $LogFile -Confirm:$false
                } ElseIf ($MinorBreak) {
                    If ($IncludeErrorLog) {
                        Add-Content -Value "$(Get-Date -Format o) *************" -Path $LogFile, $ErrorFile -ErrorAction Stop -Confirm:$false
                    } else {
                        Add-Content -Value "$(Get-Date -Format o) *************" -Path $LogFile -ErrorAction Stop -Confirm:$false
                    }
                } Else {
                    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
                        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
                    }
                    Write-Verbose -Message "$(Get-Date -Format o) $Message"

                    If ($IncludeErrorLog) {
                        Add-Content -Value "$(Get-Date -Format o) $Message" -Path $LogFile, $ErrorFile -ErrorAction Stop -Confirm:$false
                    } else {
                        Add-Content -Value "$(Get-Date -Format o) $Message" -Path $LogFile -ErrorAction Stop -Confirm:$false
                    }
                }
            } Catch {
                Write-Error -Message "Cannot Update log because: $_" -ErrorAction Continue
            }
        }
        End { }
    }

Function Create-ProgressBar {

    Param
    (
        [Parameter(Mandatory = $true)] [string] $Status

    )

 #title for the winform
$Title = "Processing GPO Reports"
#winform dimensions
$height=100
$width=400
#winform background color
$color = "gray"
$i=0

#create the form

$form1.Text = $title
$form1.Height = $height
$form1.Width = $width
$form1.BackColor = $color
$form1.ForeColor = "white"
$form1.Font.Bold

$form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
#display center screen
$form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# create label

$label1.Text = "not started"
$label1.Left=5
$label1.Top= 10
$label1.Width= $width - 20
#adjusted height to accommodate progress bar
$label1.Height=15
$label1.Font= "Verdana"
#optional to show border
#$label1.BorderStyle=1

#add the label to the form
$form1.controls.add($label1)


$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Style="Continuous"
$progressBar1.ForeColor = "blue"




$System_Drawing_Size.Width = $width - 40
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size

$progressBar1.Left = 5
$progressBar1.Top = 40
$form1.Controls.Add($progressBar1)
$form1.FormBorderStyle = 'None'
$form1.Show()| out-null
#give the form focus
$form1.Focus() | out-null
#update the form
$label1.text=$Status
$form1.Refresh()





}

function Ingest-XMLGPOReports {
    
    $files = Get-ChildItem $WorkingPath
    Update-Log -Message "Ingesting Reports" -SectionBreak
    Update-Log -Message $files

    Create-ProgressBar -Status "Uploading GPO Reports to MEM"
    $i=1
    
  
     
    
    foreach ($file in $files) {
        [int]$pct = ($i/$files.count)*100
        $progressbar1.Value = $pct
        $form1.Refresh()
        Update-Log -Message "Processing file $($file.fullname) $($pct)% completed."

        $File_Path = $file.fullname
        $get_GPO_XML_Content = [xml](get-content $File_Path)
        $Get_GPO_Name = $get_GPO_XML_Content.GPO.name
        $GPO_XML_Content = [convert]::ToBase64String((Get-Content $File_Path -Encoding byte))

$MyProfile = @"
        {
        "groupPolicyObjectFile": {
        "ouDistinguishedName": "$Get_GPO_Name",
        "content": "$GPO_XML_Content"
            }
        }
"@
 

    Update-Log -Message $MyProfile
    New-MgBetaDeviceManagementGroupPolicyMigrationReport -BodyParameter $MyProfile
    $i++

    }
    $form1.hide()
    Update-Log -Message "Cleaning up..."
    Update-WorkingDirectory
}

function Check-forNeededModules {
$Modules = Get-Module -listavailable
   If (!($Modules | where {$_.name -like "Microsoft.Graph.Beta.DeviceManagement.Administration"})) 
        { 
            Update-Log -Message "Graph PS modules missing, trying to install"            
            try {
                 Install-Module Microsoft.Graph.Intune -ErrorAction SilentlyContinue
                 } catch {
                 Update-Log -Message "Error installing Graph PS module, please try re-ruinning the tool as admin" -IncludeErrorLog
                 Write-host "Error installing Graph PS module, please try re-ruinning the tool as admin"
                 exit
                 }
        } 

     If (!($Modules | where {$_.name -like "*ActiveDirectory*"})) 
        { 
             Update-Log -Message "AD PS modules missing, trying to install"            
            try {
                 Install-Module Microsoft.Graph.Intune -ErrorAction SilentlyContinue 

                    If ((Get-WMIObject win32_operatingsystem).name -like "*Server*")
                    {
                    Import-Module ServerManager
                    Add-WindowsFeature -Name "RSAT-AD-PowerShell" –IncludeAllSubFeature
                    } else {
                    Add-WindowsCapability -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -online
                    } 
                 } catch {
                 Update-Log -Message "Error installing Active Directory PS module, Please install the Active Directory RSAT Powershell modeule manually and then re-run this script." -IncludeErrorLog
                 Write-host "Error installing Active Directory PS module, Please install the Active Directory RSAT Powershell modeule manually and then re-run this script."
                 exit
                 }
        }

}

Function Draw-GUI {
    
    # Import Windows Forms Assembly
    Add-Type -AssemblyName System.Windows.Forms;
    # Create a Form
    $Form = New-Object -TypeName System.Windows.Forms.Form;
    $form.Text = 'Select GPOs to import';
    # Create a CheckedListBox
    $CheckedListBox = New-Object -TypeName System.Windows.Forms.CheckedListBox;
    # Add the CheckedListBox to the Form
    $Form.Controls.Add($CheckedListBox);
    # Widen the CheckedListBox
    $CheckedListBox.Width = 500;
    $CheckedListBox.Height = 600;
    $form.AutoSize = $true
    $GPOs = Get-GPO -all


    # Load the GPOs to the CheckedListBox
    $GPOs = Get-GPO -all
    Foreach ($GPO in $GPOs) 
    {
        [void] $CheckedListBox.Items.Add($GPO.DisplayName)
        $CheckedListBox.SetItemChecked($CheckedListBox.Items.IndexOf($GPO.DisplayName), $true)
    }

    # Define Buttons
    $AllButton = New-Object System.Windows.Forms.Button
    $AllButton.Location = New-Object System.Drawing.Size(10,600)
    $AllButton.Size = New-Object System.Drawing.Size(120,23)
    $AllButton.Text = "Select All"
    $Form.Controls.Add($AllButton)


    $NoneButton = New-Object System.Windows.Forms.Button
    $NoneButton.Location = New-Object System.Drawing.Size(140,600)
    $NoneButton.Size = New-Object System.Drawing.Size(120,23)
    $NoneButton.Text = "Select None"
    $Form.Controls.Add($NoneButton)

    $ImportButton = New-Object System.Windows.Forms.Button
    $ImportButton.Location = New-Object System.Drawing.Size(270,600)
    $ImportButton.Size = New-Object System.Drawing.Size(120,23)
    $ImportButton.Text = "Import to MEM"
    $form.Controls.Add($ImportButton)

    $PurgeButton = New-Object System.Windows.Forms.Button
    $PurgeButton.Location = New-Object System.Drawing.Size(10,630)
    $PurgeButton.Size = New-Object System.Drawing.Size(300,23)
    $PurgeButton.Text = "Purge All GPOs from MEM Group Policy analytics"
    $form.Controls.Add($PurgeButton)

    #Event handlers
    
    $PurgeButton.Add_Click(
    {
    
    Create-ProgressBar -Status "Purging ALL GPO Reports from MEM"
    $i=1    
    
    $GPOReports = Get-MgBetaDeviceManagementGroupPolicyMigrationReport -All
        
        Foreach ($GPOReport in $GPOReports) 
            {
            [int]$pct = ($i/$files.count)*100
            $progressbar1.Value = $pct
            $form1.Refresh()
            
            $GPOReportID = $GPOReport.ID
            Update-Log -Message "Deleting $($GPOReport.DisplayName)"
            Remove-MgBetaDeviceManagementGroupPolicyMigrationReport -GroupPolicyMigrationReportId $GPOReportID
            

            }
    $form1.hide()
    }
    )
    
    
    
    $AllButton.Add_Click(
    {
        
        Foreach ($GPO in $GPOs) 
            {
            $CheckedListBox.SetItemChecked($CheckedListBox.Items.IndexOf($GPO.displayname), $true);
            }
    }
    )

    $NoneButton.Add_Click(
    {
        Foreach ($GPO in $GPOs) 
            {
            $CheckedListBox.SetItemChecked($CheckedListBox.Items.IndexOf($GPO.displayname), $false);
            }

    }
    )

    $ImportButton.Add_Click(
    {
            
            Foreach ($GPO in $GPOs) 
            {
            
            $check = $false
                try 
                {
                Update-Log -Message "Processing $($GPO.displayname)"
                $check = $CheckedListBox.GetItemChecked($CheckedListBox.Items.IndexOf($GPO.displayname));
                } catch {
                
                }
                
                Update-Log -Message "GPO $($GPO.DisplayName) Is Selected: $($check)"
                if ($check) 
                {
                $ReportPath = (Get-Location).Path + "\MEMExportReports\" + ($GPO.displayname -replace " ","_")  + ".xml"
                Update-Log -Message "Saving GPO report as $($ReportPath)"
                $XML = Get-GPOReport -Guid $GPO.ID -ReportType Xml
                $XML | Out-File -FilePath $ReportPath
                }


            } 
            Update-Log -Message "Ingesting GPO report files."      
            Ingest-XMLGPOReports
    }
    )



    $Form.ShowDialog();

}

Function Update-WorkingDirectory
{
Update-Log -Message "Checking for temporary directory. "

if (Test-path $WorkingPath) 
    {
    Update-Log -Message "Removing previous reports "
    [void](Remove-item "$($WorkingPath)\*.*" -Confirm:$false -force)
    }




}

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.ProgressBar") 

    $MyRootPath = (Get-Location).Path + "\"
    $logFile = "$($MyRootPath)\Log Files\Log_Upload_GPO_Report_to_MEM_$(get-date -Format yyyyMMddTHHmmss).txt"
    $errorFile = "$($MyRootPath)\Log Files\Error_Log_Upload_GPO_Report_to_MEM_$(get-date -Format yyyyMMddTHHmmss).txt"
    

    #Check if 'Log Files' exists and if not, create it
    if (Test-Path -Path "$($MyRootPath)\Log Files") {
        Write-Verbose -Message 'Log Files folder already existed.'
    } else {
        Write-Verbose -Message 'Need to create Log Files folder.'
        Try {
            $null = New-Item -Path "$($MyRootPath)\Log Files" -ItemType Directory -ErrorAction Stop
            Write-Verbose -Message "Created Folder '$($MyRootPath)\Log Files' because it did not exist."
        } Catch {
            Write-Error -Message "Could not create '$($MyRootPath)\Log Files' because: $_" -ErrorAction Stop
        }
    }
    


$form1 = New-Object System.Windows.Forms.Form
$WorkingPath = (Get-Location).Path + "\" + "MEMExportReports\"
$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$label1 = New-Object system.Windows.Forms.Label
$System_Drawing_Size = New-Object System.Drawing.Size

Update-WorkingDirectory
Write-Host "Checking for required PS Modules."
Update-Log -Message "Checking for required PS Modules." -SectionBreak
Check-forNeededModules
Write-Host "Logging into GRAPH."
Update-Log -Message "Logging into GRAPH." -SectionBreak
Import-Module Microsoft.Graph.Beta.DeviceManagement.Administration
#Select-MgProfile -Name Beta
$connectParams = @{
    Scopes = "DeviceManagementConfiguration.ReadWrite.All"
}
Connect-MgGraph @connectParams
Write-Host "Building UI."
Update-Log -Message "Logging into GRAPH." -SectionBreak
Draw-GUI
