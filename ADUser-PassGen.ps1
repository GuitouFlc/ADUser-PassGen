#global variables
$Myuser = Read-Host 'A quel compte souhaitez vous apporter des modifications ? (SamAccountName ex: gufloch)'
function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [ValidateRange(4,[int]::MaxValue)]
        [int] $length,
        [int] $upper = 1,
        [int] $lower = 1,
        [int] $numeric = 1,
        [int] $special = 1
    )
    if($upper + $lower + $numeric + $special -gt $length) {
        throw "number of upper/lower/numeric/special char must be lower or equal to length"
    }
    $uCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lCharSet = "abcdefghijklmnopqrstuvwxyz"
    $nCharSet = "0123456789"
    $sCharSet = "/*-+,!?=()@;:._"
    $charSet = ""
    if($upper -gt 0) { $charSet += $uCharSet }
    if($lower -gt 0) { $charSet += $lCharSet }
    if($numeric -gt 0) { $charSet += $nCharSet }
    if($special -gt 0) { $charSet += $sCharSet }
    
    $charSet = $charSet.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
    $rng.GetBytes($bytes)
 
    $result = New-Object char[]($length)
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
    }
    $password = (-join $result)
    $valid = $true
    if($upper   -gt ($password.ToCharArray() | Where-Object {$_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
    if($lower   -gt ($password.ToCharArray() | Where-Object {$_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
    if($numeric -gt ($password.ToCharArray() | Where-Object {$_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
    if($special -gt ($password.ToCharArray() | Where-Object {$_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }
 
    if(!$valid) {
         $password = Get-RandomPassword $length $upper $lower $numeric $special
    }
    return $password
}
function choice{
    Write-Host "0 - Quitter" -BackgroundColor Black -ForeGroundColor White
    Write-Host "1 - Set Maintenance Password (Manuel)" -BackgroundColor Gray -ForeGroundColor Black
    Write-Host "2 - Generer un mot de passe Utilisateur 'jetable' " -BackgroundColor Gray -ForeGroundColor Black
    Write-Host ""
    $num = Read-Host "Faites votre choix"
        Switch ($num)
    {
        0 {Write-Host "See You Soon" Exit}
        1 {MaintenancePass}
        2 {SessionReset}
    }
}
function SessionReset{
   
    $NewPwd = ConvertTo-SecureString $password -AsPlainText -Force
    Set-ADAccountPassword -Identity $Myuser -NewPassword $NewPwd -Reset
    Set-ADUser -Identity $Myuser -ChangePasswordAtLogon $true
    Unlock-ADAccount $Myuser

    Write-Host -BackgroundColor DarkGreen "New pass is set and $MyUser account is unlocked" 
    Write-Host -BackgroundColor DarkGreen "New pass for first login is $password" 

    $choice = Read-Host "Voulez vous Générer une fiche pour l' Utilisateur (Y pour Oui)"
    if($choice -eq "Y"){
    FicheAccueil
    }
    else{
        Exit
    }
}
function FicheAccueil{
    $mail = Get-ADUser $Myuser -Properties mail | select -ExpandProperty UserPrincipalName
    # unzip function
    Add-Type -AssemblyName "System.IO.Compression.FileSystem"
    function Unzip {
        param([string]$zipfile, [string]$outpath)
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }
    function Zip {
        param([string]$folderInclude, [string]$outZip)
        [System.IO.Compression.CompressionLevel]$compression = "Optimal"
        $ziparchive = [System.IO.Compression.ZipFile]::Open( $outZip, "Update" )
    
        # loop all child files
        $realtiveTempFolder = (Resolve-Path $config.tempFolder -Relative).TrimStart(".\")
        foreach ($file in (Get-ChildItem $folderInclude -Recurse)) {
            # skip directories
            if ($file.GetType().ToString() -ne "System.IO.DirectoryInfo") {
                # relative path
                $relpath = ""
                if ($file.FullName) {
                    $relpath = (Resolve-Path $file.FullName -Relative)
                }
                if (!$relpath) {
                    $relpath = $file.Name
                } else {
                    $relpath = $relpath.Replace($realtiveTempFolder, "")
                    $relpath = $relpath.TrimStart(".\").TrimStart("\\")
                }
    
                # debug line for *.doc creation 
                # Write-Host $relpath -Fore Green
                # Write-Host $file.FullName -Fore Yellow
    
                # add file
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ziparchive, $file.FullName, $relpath, $compression) | Out-Null
            }
        }
        $ziparchive.Dispose()
    }
    
    # prepare folder
    Remove-Item $config.tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
    mkdir $config.tempFolder | Out-Null
    
    # unzip DOCX
    Unzip $config.template $config.tempFolder
    
    # replace text
    $bodyFile = $config.tempFolder + $config.bodyFileXML
    $body = Get-Content $bodyFile
    $body = $body.Replace($config.field1, $Myuser)
    $body = $body.Replace($config.field2, $password)
    $body = $body.Replace($config.field3, $mail)
    $body | Out-File $bodyFile -Force -Encoding ascii
    
    # zip DOCX
    $destfile = $config.template -Replace(".docx", "-$user2000.docx")
    Remove-Item $destfile -Force -ErrorAction SilentlyContinue
    Zip $config.tempFolder $destfile
    
    # clean folder
    Remove-Item $config.tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
    $MypcID = Read-Host $config.PCNumber
    Copy-Item "$destfile" $config.EditFileDestination
    
    Write-Host "User creation complete" -BackgroundColor DarkGreen
}
function MaintenancePass{
    $choice = Read-Host "Voulez vous utiliser le mot de passe généré ? Y pour Oui"

    if($choice -eq "Y"){
        Write-host "Mot de passe validé : $password"
    }
    else{
        $password = Read-Host "saisissez le mot de passe que vous souhaitez utiliser"
    }

    $NewPwd = ConvertTo-SecureString $password -AsPlainText -Force
    Set-ADAccountPassword -Identity $Myuser -NewPassword $NewPwd -Reset
    Set-ADUser -Identity $Myuser -ChangePasswordAtLogon $true
    Unlock-ADAccount $Myuser

    Write-Host -BackgroundColor DarkGreen "Maintenance pass for $MyUser is $password" 
}
function main () {
    # config Import
        $configfile = $args[0]
    
        if (Test-Path -Path $configfile) {
        } 
        else {
            Write-Host ("File " + $configfile + " does not exists")
            exit 1
        }
        try {
            $config = Import-PowerShellDataFile -Path $configfile
            Write-Host $configfile "import success" -BackgroundColor Yellow -ForegroundColor Black
        }
        catch {
            Write-Host $_.Exception
        }
    # Call function  
    # Get-Password >> [0]>Lengt [1]>upper [2]>lower [3]number [4]special
    $password = Get-RandomPassword 12 1 1 1 1
    choice
    
    }
    #Test
    if ($args.Length -gt 0) {
        main $args[0]
    } 
    else {
        Write-Host ("Arguments not fully specified")
        Write-Host ("Specify config file as argument like: DomainDeploymentTool.ps1 <config file name>")
        exit 1
    }
