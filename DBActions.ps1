using namespace System.Management.Automation.Host
#=================================================================================
# Designed to deploy a database from a dacpac
#
# Usage:
# .\sqlPackageDeploymentCMD.ps1  -targetServer "LOCALHOST" -targetDB "IamADatabase" -dacpacFile "C:\ProjectDirectory\bin\Debug\IamADatabase.dacpac" -SQLCMDVariable1 "IamASQLCMDVariableValue"
# 
#
#=================================================================================

Param(
    #Action
    [Parameter(Mandatory=$false)]
    [string]$action = "",

    #Database connection
    [Parameter(Mandatory=$false)]
    [string]$targetServerName = "localhost\SQLEXPRESS",
    [Parameter(Mandatory=$false)]
    [string]$targetDBname = "SiomaxDB",

    #DacPac source
    #Note PSScriptRoot is the location where this script is called from. Good idea to keep it in the root of 
    # your solution then the absolute path is easy to reconstruct
    [Parameter(Mandatory=$false)]
    [string]$dacpacFile = "",

    [Parameter(Mandatory=$false)]
    [string]$bacpacFile = ""

    #"""$PSScriptRoot\ProjectDirectory\bin\Debug\IamADatabase.dacpac""" #Quotes in case your path has spaces

    ##SQLCMD variables
    #[Parameter(Mandatory=$false)]
    #[string]$SQLCMDVariable1 = "IamASQLCMDVariableValue",
    #
    #[Parameter(Mandatory=$false)]
    #[string]$SQLCMDVariable2 = "IamSomeOtherSQLCMDVariableValue"

)
#SQLPackage
$sqlpackage ="$env:ProgramData\DAC\sqlpackage\sqlpackage.exe"

filter timestamp {"$(Get-Date -Format G): $_"}

Function Sleep-Progress($seconds) {
    $s = 0;
    Do {
        $p = [math]::Round(100 - (($seconds - $s) / $seconds * 100));
        Write-Progress -Activity "Waiting..." -Status "$p% Complete:" -SecondsRemaining ($seconds - $s) -PercentComplete $p;
        [System.Threading.Thread]::Sleep(1000)
        $s++;
    }
    While($s -lt $seconds);
    
}
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}
function CheckExistSqlpackage {
    [OutputType([bool])]
    $URL = "https://go.microsoft.com/fwlink/?linkid=2215400"
    $Path = "$env:TEMP\sqlpackage.zip"
    $result=$true
    if (-not(Test-Path -Path $sqlpackage -PathType Leaf))
    {
        try {
            Write-Host "Obteniendo SqlPackage..."
            Start-BitsTransfer -Source $URL -Destination $Path
            Write-Host "Descomprimiendo archivos..."
            Expand-Archive -Path $Path -DestinationPath "$env:ProgramData\DAC\sqlpackage" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $Path -Force
        }
        catch {
        Write-Warning "Algo paso al obtener SqlPackage: $message"
        $result=$false
        }
        
    }
    return $result
}
function Create-Credential {
    [OutputType([pscredential])]
       param(
        [Parameter(Mandatory)]
        [string]$User,

        [Parameter(Mandatory)]
        [string]$Password
    )

    [securestring]$secStringPassword = ConvertTo-SecureString $Password -AsPlainText -Force
    [pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($User, $secStringPassword)
    return $credObject
}
function Test-SqlConnection {
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$ServerName,

        [Parameter(Mandatory)]
        [string]$DatabaseName,

        [Parameter(Mandatory)]
        [pscredential]$Credential
    )

    $ErrorActionPreference = 'Stop'
    $result=$true
    try {
        $userName = $Credential.UserName
        $password = $Credential.GetNetworkCredential().Password
        $connectionString = 'Data Source={0};database={1};User ID={2};Password={3}' -f $ServerName,$DatabaseName,$userName,$password
        $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
        $sqlConnection.Open()        
    } catch {
        $message = $_
        Write-Warning "Falló la prueba: $message"
        $result= $false
    } finally {
        ## Close the connection when we're done
        $sqlConnection.Close()
    }
    return $result
}
function UpdateDataBase {
    [OutputType([bool])]
    Param(
        #Action
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$fileDapac
        )
    
    Write-Host "Actualizando base de datos $targetDBname..."
    #%SQLPCK% /a:Publish /sf:"%DBFILE%" /Profile:".\SiomaxDB.publish.xml" /drp:".\SiomaxDB.Report.xml"
    & "$sqlpackage" /a:"Publish" /sf:"$fileDapac"  /Profile:"$PSScriptRoot\SiomaxDB.publish.xml" /drp:"$targetDBname.$timeStamp.xml"
    if ( $LASTEXITCODE -ne 0 ) {
        Write-Host "No se pudo actualizar la base de datos $targetDBname..."
        return $false
    }
    else { return $true }
    
}
function UpdateSiomaxDBArtifact{
    $latest = Get-ChildItem $PSScriptRoot -Attributes !Directory *SiomaxDB*.zip | Sort-Object -Descending -Property LastWriteTime | select -First 1
    if ($latest.count -eq 1) {
         Write-Host "Leyendo paquete de actualización de SiomaxDB..."
         $tmpFolder=New-TemporaryDirectory
         Expand-Archive -Path $latest.FullName -DestinationPath $tmpFolder -Force -ErrorAction SilentlyContinue
         
         $result=UpdateDataBase -fileDapac "$tmpFolder\SiomaxDB.dacpac"
         Remove-Item -LiteralPath $tmpFolder -Force -Recurse
         if ($result -eq $true){
            $BkpFolder=(Join-Path $PSScriptRoot "UpdatesBackup")
            $fileName= [System.IO.Path]::GetFileNameWithoutExtension($latest.FullName)  
            
            if(-not ([System.IO.Directory]::Exists($BkpFolder))) {
                $null = New-Item -ItemType Directory -Path $BkpFolder}
            
            Move-Item –Path $latest.FullName -Destination "$BkpFolder\$fileName$timeStamp.old"
         }
    }
    else {
        Write-Host "No se encontró ningin paquete de actualización de SiomaxDB..."
    }
}
function RestoreDataBase {
   
    Write-Host "Restaurando base de datos $targetDBname..."
    $RestoreContinue=$true
     if ([string]::IsNullOrEmpty($bacpacFile)) {
        $latest = Get-ChildItem $PSScriptRoot -Attributes !Directory *.bacpac | Sort-Object -Descending -Property LastWriteTime | select -First 1
        $bacpacFile=$latest.FullName
        if (-not [string]::IsNullOrEmpty($bacpacFile)) { Write-Host "Se usará el archivo mas reciente encontrado: $bacpacFile" }

         if ([string]::IsNullOrEmpty($bacpacFile)) {
            Write-Host "No se estableció un archivo bacpac para importar la base de datos: $targetDBname"
            $RestoreContinue=$false
            }
     }

    
     if ($RestoreContinue -eq $true){
    Write-Host "Haciendo DROP a la base de datos $targetDBname..."
    $qcd = "USE [master];IF DB_ID('$targetDBname') IS NOT NULL BEGIN ALTER DATABASE [$targetDBname] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$targetDBname]; END"
    try {
        Invoke-Sqlcmd -ServerInstance $targetServerName  -Query $qcd -Verbose -ErrorAction 'Stop'
    } catch{
        $RestoreContinue=$false
        Write-Host($_)
    }
    }
    if ($RestoreContinue -eq $true){
        & "$sqlpackage" /a:"Import" /sf:"$bacpacFile" /tsn:"$targetServerName" /tdn:"$targetDBname" /q:True /TargetTrustServerCertificate:True
        if ( $LASTEXITCODE -ne 0 ) { $RestoreContinue=$false }
    }
    if ($RestoreContinue -eq $false){ Write-Host "No fue restaurada la base de datos $targetDBname..."}
}
function BackupDataBase {

    Write-Host "Exportando la base de datos $targetDBname..."
    
    #User Id=sa;Password=Admin2908
    
    & "$sqlpackage" /a:"Export" /SourceServerName:"$targetServerName" /SourceDatabaseName:"$targetDBname" /SourceTrustServerCertificate:True /q:True /tf:"$PSScriptRoot\$targetDBname.$timeStamp.bacpac"
    if ( $LASTEXITCODE -ne 0 ) {
        Write-Host "No se pudo exportar la base de datos $targetDBname..."
        return $false
    }
    else { return $true }
}
function InteractiveOptions{

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Question
    )

    $update =  [ChoiceDescription]::new('&Update', 'Actualizar base de datos')
    $restore = [ChoiceDescription]::new('&Restore', 'Restaurar base de datos')
    $backup =  [ChoiceDescription]::new('&Backup', 'Respaldar base de datos')
    $MkToke =  [ChoiceDescription]::new('&Crear Token', 'Crea un toque para conectarce a la base de datos')
    
    $options = [ChoiceDescription[]]($update, $restore, $backup,$MkToke)

    $result = $host.ui.PromptForChoice($Title, $Question, $options, 0)

     return $result
        
}
Function Create-SQLToken() {
    $keyfile = ".\Sql.token"
    $serverName = Read-Host "Enter Server name"
    $DBName = Read-Host "Enter Database name"
    $user = Read-Host "Enter Username"
    $pass = Read-Host "Enter Password" -AsSecureString
    $strPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))

    if ([string]::IsNullOrWhiteSpace($serverName)) { $serverName = 'ZEROCOOL\SQLEXPRESS' }
    if ([string]::IsNullOrWhiteSpace($DBName)) { $DBName = 'SiomaxDB' }

    $json = @{ 'Server'= "$serverName"
           'Database' = "$DBName"
           'UseCredential' = $true
           'UserID'= "$user"
           'Password'= "$strPwd" } | ConvertTo-Json -Compress
    $secureJson=$json | ConvertTo-SecureString -AsPlainText -Force
    $encryptedJson = $secureJson | ConvertFrom-SecureString
    $encryptedJson | Set-Content -Encoding UTF8 -Path $keyfile
}
Function Load-Token() {
    $keyfile = ".\Sql.token"
    $token = Get-Content -Path $keyfile -Encoding UTF8
    $secureToken = ConvertTo-SecureString -String $token
    
    $Marshal = [System.Runtime.InteropServices.Marshal]
    $Bstr = $Marshal::SecureStringToBSTR($secureToken)
    $decryptedJSON = $Marshal::PtrToStringAuto($Bstr)
    $Marshal::ZeroFreeBSTR($Bstr)

    #write-host  $jsonObj
    #$j= $decryptedJSON | ConvertFrom-Json
    return $decryptedJSON
}

$now = Get-Date
$logfile = "$PSScriptRoot\logfiles\file-" + $now.ToString("yyyy-MM-dd") + ".log"
Start-Transcript -path $logfile -force


Get-WMIObject Win32_ComputerSystem
Get-NetIPAddress

$timeStamp=$(((get-date).ToUniversalTime()).ToString("yyyyMMddTHHmmss"))

#UpdateSiomaxDBArtifact

CheckExistSqlpackage

if ([string]::IsNullOrEmpty($action)) {
    $result=InteractiveOptions -Title 'Menu' -Question '¿Que deseas hacer con la base de datos?'
    switch ($result)
        {
            0 { $action="update"}
            1 { $action="restore"}
            2 { $action="backup"}
            2 { $action="token"}
        }
}

switch ($action)
{
    "update" {UpdateDataBase;Break}
    "restore" {RestoreDataBase;Break}
    "backup" {BackupDataBase;Break}
    "token" {BackupDataBase;Break}
    Default {
        Write-Host "comando de acción no válido '$action'"
    }
}
Stop-Transcript
 Sleep-Progress 5