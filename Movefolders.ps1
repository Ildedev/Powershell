#######################################################################################
#                                                                                     #
#   Prerequisites                                                                     #
#   Download and Install Sharepoint Client Components SDK                             #
#   Sharepoint 2013                                                                   #
#   https://www.microsoft.com/en-us/download/details.aspx?id=35585                    #
#   Sharepoint 2016                                                                   #
#   https://www.microsoft.com/en-us/download/details.aspx?id=51679                    #
#                                                                                     #
#   Author: Sebastian Ilde                                                            #
#                                                                                     #
#######################################################################################

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

if((Test-Path -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI")){
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
}
elseif((Test-Path -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI")){
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
}
else{
    throw [System.IO.FileNotFoundException]"No Sharepoint Client Components is installed."
    #Install-Prerequisites -Version 16
}
[switch]$isSupported | Out-Null
$global:folderPath
$global:newSite
$global:SharepointRoot
$UnSupportedFileExtensions  =   ".ade", ".adp", ".asa", ".ashx", ".asmx", ".asp", ".bas", ".bat", ".cdx", ".cer", ".chm",
                                ".class", ".cmd", ".com", ".config", ".cnt", ".cpl", ".crt", ".csh", ".der", ".dll", ".exe",
                                ".fxp", ".gadget", ".grp", ".hlp", ".hpj", ".hta", ".htr", ".htw", ".ida", ".idc", ".idq", ".ins",
                                ".isp", ".its", ".json", ".ksh", ".lnk", ".mad", ".maf", ".mag", ".mam", ".maq", ".mar",
                                ".mas", ".mat", ".mau", ".mav", ".maw", ".mcf", ".mda", ".mdb", ".mde", ".mdt", ".mdw", ".mdz",
                                ".ms-one-stub", ".msc", ".msh", ".msh1", ".msh1xml", ".msh2", ".msh2xml", ".mshxml", ".msi", ".msp",
                                ".mst", ".ops", ".pcd", ".pif", ".pl", ".prf", ".prg", ".printer", ".ps1", ".ps1xml", ".ps2", ".ps2xml",
                                ".psc1", ".psc2", ".pst", ".reg", ".rem", ".scf", ".scr", ".sct", ".shb", ".shs", ".shtm", ".shtml",
                                ".soap", ".stm", ".svc", ".url", ".vb", ".vbe", ".vbs", ".vsix", ".ws", ".wsc", ".wsf", ".wsh", ".xamlx"
$UnSupportedChar            =   '"', '#', '%', '*', ':', '<', '>', '?', '/', '\', '|'

function Select-Folder{
    param([string]$Description="Select Folder",
          [string]$RootFolder="Desktop")
     
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = $RootFolder
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    if($Show -eq "OK"){
        $global:folderPath =  $objForm.SelectedPath
    }else{
        Write-Error "Operation cancelled by user."
    }
}
function Out-Logfile($filename){
    if((Test-Path -Path "$PSScriptRoot\logfile.log") -eq $false){
        New-Item -Path "$PSScriptRoot\logfile.log" | Out-Null
    }     
    Add-Content -Path "$PSScriptRoot\logfile.log" -Value $filename.fullname
}

function Install-Prerequisites{
    [CmdletBinding()]
    param(
        [parameter( Mandatory = $true )]
        [ValidateSet(13,16)][int]$Version
    )
    BEGIN{
        if($Version -eq 13){
            $msifile = "$PSScriptRoot\file"
        }
        if($Version -eq 16){
            $msifile = "$PSScriptRoot\Prerequisites\Sharepoint16.msi"
        }
        $arguments = @(
            "/i"
            "`"$msifile`""
            "/qn"
            "/norestart"
        )
    }
    PROCESS{
        Write-Verbose "Installing....."
        $process = Start-Process -FilePath "$PSScriptRoot\Prerequisites\Sharepoint16.msi" /qn -Wait -PassThru
        if($process.ExitCode -eq 0){
            Write-Verbose "Client Components was successfully installed!"
        }
        else{
            Write-Verbose "Couldnt install, Exit code $($process.ExitCode)"
        }
    }
    END{
    }
}

function Select-Site{
    $webURL = [Microsoft.VisualBasic.interaction]::InputBox("Enter full URL for site`n`n`n"+
                                                            "Example: https://Companyname.sharepoint.com/wffd_it/", "URL")
    if($webURL -eq $null -or $webURL -eq ""){
        Throw "URL cannot be null"
        exit
    }
    $global:context = New-Object Microsoft.Sharepoint.Client.ClientContext($webURL)
    $credentials = (Get-Credential)
    $folderPaths = $global:folderPath.Split('\')
    $global:newSite = $folderPaths[$folderPaths.Count-1]
    $cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)
    $global:context.Credentials = $cred
    $LibInput = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Document library where you want move to.`n`n`n"+
                                                              "Example: Projects/ProjectName/folder", "Document Library")
    $global:SharepointRoot = "Shared Documents/"+$LibInput
    return $global:SharepointRoot
}

function Start-Move{
    try{
        Clear-Host
        Write-Host "Connection to Sharepoint Site...`n`n`n`n`n`n" -ForegroundColor Yellow
        $web = $global:context.Web
        $global:context.Load($web)
        $global:context.ExecuteQuery()
        $folder = $web.Folders.Add($global:SharepointRoot + '/' + $global:newSite)
        $global:context.Load($folder)
        $global:context.ExecuteQuery()
        $DocLib = $web.GetFolderByServerRelativeUrl($web.ServerRelativeUrl + '/' + $global:SharepointRoot)
        $global:context.Load($DocLib)
        $global:context.ExecuteQuery()
    }catch{
        throw "Could not connect to this sharepoint site"
        exit
    }
    $i = 1
    #progress bar
    $Form.Close()
    $form1.Show() | Out-Null
    $form1.Focus()| Out-Null
    $Files = Get-ChildItem -Path $global:folderPath -Recurse
    foreach ($File in $Files) {
        $progressBar1.Value = ($i/$Files.Count)*100
        $label1.Text = "Moving $file"
        $form1.Refresh()

        $extension = [System.IO.Path]::GetExtension($File)
        Foreach($UnSupportedFileExtension in $UnSupportedFileExtensions){
            if($extension -eq $UnSupportedFileExtension){
                $isSupported = $false
                break
            }else{
                $isSupported = $true
            }
        }
        if($isSupported){
            $string = $File.FullName
            $string = $string.Split(':')
            $Relativefolder = $string[1].Replace('\', '/')
            if((Get-Item $File.fullname) -is [System.IO.DirectoryInfo]){
                #Write-Host "Creating Directory $File" -ForegroundColor Green
                $folder = $web.Folders.Add($global:SharepointRoot + $Relativefolder)
                $global:context.Load($folder)
                $global:context.ExecuteQuery()                
            }else{
                $Relativefolder1 = $Relativefolder.Replace('/'+$File.Name, '')
                #Write-Host "Uploading $File" -ForegroundColor Green
                $newPath = $web.GetFolderByServerRelativeUrl($web.ServerRelativeUrl + '/' + $global:SharepointRoot + $Relativefolder1)
                $global:context.Load($newPath)
                $FileFullName = $File.FullName
                $FileStream = New-Object IO.FileStream($FileFullName, [System.IO.FileMode]::Open)
                $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                $FileCreationInfo.Overwrite = $true
                $FileCreationInfo.ContentStream = $FileStream
                $FileCreationInfo.URL = $File.Name
                $FileUpload = $newPath.Files.Add($FileCreationInfo)
                $global:context.Load($FileUpload)
                $global:context.ExecuteQuery()
            }
        }else{
            #Write-Host "Unsupported fileformat $File" -ForegroundColor Red
            Out-Logfile -filename $File
        }
        $i++
    }
    Write-Host "Completed!" -ForegroundColor Green
    $form1.Close()
}

# GUI
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Move folders to Sharepoint"
$Form.TopMost = $true
$Form.Width = 574
$Form.Height = 405

$button2 = New-Object system.windows.Forms.Button
$button2.Text = "Select Folder"
$button2.Width = 95
$button2.Height = 30
$button2.Add_Click({
    Select-Folder
    $label11.Text = $global:folderPath
})
$button2.location = new-object system.drawing.point(22,161)
$button2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button2)

$button3 = New-Object system.windows.Forms.Button
$button3.Text = "Select site"
$button3.Width = 95
$button3.Height = 30
$button3.Add_Click({
    $shareSite = Select-Site
    $label10.Text = $shareSite
})
$button3.location = new-object system.drawing.point(22,220)
$button3.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button3)

$button4 = New-Object system.windows.Forms.Button
$button4.Text = "Start"
$button4.Width = 82
$button4.Height = 30
$button4.Add_Click({
    Start-Move
    #$Form.Close()
})
$button4.location = new-object system.drawing.point(443,302)
$button4.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button4)

$label10 = New-Object system.windows.Forms.Label
$label10.Text = ""
$label10.AutoSize = $true
$label10.Width = 25
$label10.Height = 10
$label10.location = new-object system.drawing.point(144,224)
$label10.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label10)

$label11 = New-Object system.windows.Forms.Label
$label11.Text = ""
$label11.AutoSize = $true
$label11.Width = 25
$label11.Height = 10
$label11.location = new-object system.drawing.point(144,167)
$label11.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label11)

$form1 = New-Object System.Windows.Forms.Form
$form1.Text = "Moving Files"
$form1.Height = 100
$form1.Width = 400
$form1.BackColor = "white"
$form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
$form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

$label1 = New-Object system.Windows.Forms.Label
$label1.Text = "not started"
$label1.Left=5
$label1.Top= 10
$label1.Width= (400 - 20)
$label1.Height=15
$label1.Font= "Verdana"

$form1.controls.add($label1)
$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Top = $true
$progressBar1.Value = 0
$progressBar1.Style="Continuous"
    
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = (400 - 40)
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size
$progressBar1.Left = 5
$progressBar1.Top = 40
$form1.Controls.Add($progressBar1)

[void]$Form.ShowDialog()
$Form.Dispose()
