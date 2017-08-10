if(Test-Path C:\script -eq $false){
    mkdir 'C:\script\'
    mkdir 'C:\script\mail\'
}
[string]$mail
Set-Location C:\script\
$name = $env:COMPUTERNAME
$date = Get-date | select -ExpandProperty date
if(Test-Path C:\script\mail\Auto.txt -eq $false){
    Read-Host -assecurestring | Convertfrom-SecureString | Out-file 'C:\script\mail\Auto.txt'
}
$password = Get-Content .\mail\Auto.txt | ConvertTo-SecureString
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mail,$password


function Out-HTMLWhitered{
    <#
        .SYNOPSIS
            You can pipe any .NET object into this function and it will out the object in a HTML output and HTML file
        .DESCRIPTION
            Will print out information in a nicely formatted (Whitered) HTML output.
            First parameter it takes in is its Piped in Object which is mandatory
            You can add a title and subtitle by adding the -Title and the -SubTitle as Arguments after.
        .EXAMPLE
            PSObject | Out-HTMLWhitered -Title 'A Title' -SubTitle 'Another lesser title'
            Get-Service | Out-HTMLWhitered -Title 'Service on Computer' -SubTitle 'Services:'
    #>
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   Position = 0)][System.Management.Automation.PSObject]$Object,
        [Parameter(Mandatory = $true)][String]$Title = 'The Big Title',
                                      [String]$SubTitle = ''
    )
    BEGIN{
        Write-Verbose -Message "Setting up Prevariables"
        $pipelineArray = @()
        $ReportTitle="Rapport"
        $footer="Report run {0} by {1}" -f (Get-Date),$env:USERNAME
        $head =     "<style>"
        $head +=    "body { background-color:#FFFFFF;font-family:Helvetica;font-size:12pt; }"
        $head +=    "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
        $head +=    "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;color:white;background-color:#FF0000;}"
        $head +=    "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
        $head +=    "table, tr, td, th { padding: 2px; margin: 0px }"
        $head +=    "table { margin-left:50px; }"
        $head +=    "</style>"
        $head +=    "<Title>$ReportTitle</Title>"
        Write-Verbose -Message "Putting object $($Object) into array"
    }
    PROCESS{
        $pipelineArray += $Object
    }
    END{
        Write-Verbose -Message "Creating and converting objects into HTML code"
        $fragments = @()
        $fragments += "<Img src='http://www.whitered.se/wp-content/uploads/2017/01/logo_110x68.png' style='float:left' width='110' height='68' hspace=10><H1><center>$($Title)</center></H1><br><br>"
        $fragments += $pipelineArray | ConvertTo-HTML -Fragment -PreContent "<H2>$($SubTitle)</H2>"
        $html = ConvertTo-Html -Head $head -Title $ReportTitle -PreContent "$($fragments)" -PostContent "<br><I>$footer</I>"
        Write-Verbose -Message "Complete!"
        return $html
    }
}


try{
    $data = Get-VM -ComputerName $name | Select-Object name,state,replicationmode,replicationhealth |
    Sort-Object replicationmode | Out-HTMLWhitered -Title "VM / Replica Health"
    $data | Out-File VMHealth.html
    $From = "Senders@Email.com"
    $To = "someemail@mail.com"
    $cc = "thecc@mail.com"
    $Attachment = "C:\script\VMHealth.html"
    $Subject = "$($env:COMPUTERNAME), $date"
    $Body = "$data"
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    Send-MailMessage -From $From -to $To -cc $cc -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    -Credential $credential -Attachments $Attachment -BodyAsHtml
}catch{
    $Error | Out-File C:\script\log.log
}