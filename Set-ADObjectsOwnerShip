[CmdLetBinding(DefaultParameterSetName = 'Default')]
Param(
    [Parameter(ParameterSetName='Default',Mandatory=$False)]
    [Switch]
    $DoRemediation,
    [Parameter(ParameterSetName='Default',Mandatory=$False)]
    [Switch]
    $Mail,
    [Parameter(ParameterSetName='Default',Mandatory=$False)]
    [Switch]
    $Html
)
Remove-Module *
Clear-Host
$savedLogs = @()
$PSDefaultParameterValues = @{ '*:Encoding' = 'utf8' }
#region Enums
    enum Environ
    {
        PRD
        UAT
        HOMOL
        LAB
    }
    enum Category
    {
        ASSESS
        CHECK
        CONFIGURE
        CREATE
        DELETE
        DISABLE
        ENABLE
        MAINTAIN
        REMOVE
        UPDATE
    }
    enum Classes
    {
        ADAdmin
        ADComputer
        ADGeneric
        ADGroup
        ADObject
        ADService
        ADUser
    }
#endregion
#region Class
    class UserConfig
    {
        [String[]]$EmailTo
        [String]$ConfigFileName
        [String]$LogFileName
        [String]$CsvFileName
        [String]$ZipFileName
        [String]$HtmlFileName
    }
    class Files
    {
        [String]$Config
        [String]$Log
        [String]$Csv
        [String]$Zip
        [String]$Html
    }
    class Paths
    {
        [String]$Logs
        [String]$Results
        [String]$Script
        [String]$grouped
    }
    class OutSettings
    {
        [DateTime]$Ref
        [String]$Env
        [String]$Category
        [String]$Class
        [String]$ShortTitle
        [Boolean]$hasConfigFile
        [String]$configFile
        [String]$EmailTo
        [String]$ConfigFileName
        [String]$LogFileName
        [String]$CsvFileName
        [String]$ZipFileName
        [String]$HtmlFileName
        [String]$LogPath
        [String]$ResultPath
        [String]$ScriptPath
        [String]$groupedPath
        [String]$logFile
        [String]$csvFile
        [String]$zipFile
        [String]$htmlFile
    }
    class ScriptData
    {
        static [String]$ShortTitle
        static [String]$Env
        static [String]$Category
        static [String]$Class
        static [Boolean]$hasConfigFile
        Static [DateTime]$Ref          = [datetime]::Now
        static [Files]$Files           = [Files]::new()
        static [UserConfig]$ConfigData = [UserConfig]::new()
        static [Paths]$Paths           = [Paths]::new()
        static [Object] GetSettings()
        {
            $output = [OutSettings]::new()
            $output.Ref            = [ScriptData]::Ref
            $output.ShortTitle     = [ScriptData]::ShortTitle
            $output.Env            = [ScriptData]::Env
            $output.Category       = [ScriptData]::Category
            $output.Class          = [ScriptData]::Class
            $output.hasConfigFile  = [ScriptData]::hasConfigFile
            $output.configFile     = [ScriptData]::Files.Config
            $output.EmailTo        = [ScriptData]::ConfigData.EmailTo
            $output.ConfigFileName = [ScriptData]::ConfigData.ConfigFileName
            $output.LogFileName    = [ScriptData]::ConfigData.LogFileName
            $output.CsvFileName    = [ScriptData]::ConfigData.CsvFileName
            $output.ZipFileName    = [ScriptData]::ConfigData.ZipFileName
            $output.HtmlFileName   = [ScriptData]::ConfigData.HtmlFileName
            $output.LogPath        = [ScriptData]::Paths.Logs
            $output.ResultPath     = [ScriptData]::Paths.Results
            $output.ScriptPath     = [ScriptData]::Paths.Script
            $output.groupedPath    = [ScriptData]::Paths.grouped
            $output.logFile        = [ScriptData]::Files.Log
            $output.csvFile        = [ScriptData]::Files.Csv
            $output.zipFile        = [ScriptData]::Files.Zip
            $output.htmlFile       = [ScriptData]::Files.Html
            return $output
        }
    }
#endregion
#region User Variables
    $ConfigData = ([ScriptData]::configData)
    # Fichier de Config
    $ConfigData.ConfigFileName = ".json"
    # Fichier de Logs
    $ConfigData.LogFileName = ".log"
    # Fichier CSV
    $ConfigData.CsvFileName = ".csv"
    # Fichier ZIP
    $ConfigData.ZipFileName = ".zip"
    # Fichier HTML
    if ($html.IsPresent)
    {
        $ConfigData.HtmlFileName = ".html"
    }
    # Creates a Title for the Script
    [ScriptData]::Env        = [Environ]::PRD
    [ScriptData]::Category   = [Category]::CHECK
    [ScriptData]::Class      = [Classes]::ADObject
    [ScriptData]::ShortTitle = "OwnerShip"
#endregion
#region Preparation
    #region Functions
        Function Set-Log {
            [CmdLetBinding()]
            param (
                [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
                [ValidateNotNullOrEmpty()]
                [String]
                $message,
                [Parameter(Mandatory=$false)]
                [ValidateSet("Success","Info","Warning","Error","No")]
                [String]
                $severity="Info",
                [Parameter(Mandatory=$false)]
                [Switch]
                $hold
            )
            if ($severity -ne "No") {
                $message = "`t" + $severity.toupper() + ":`t"+ $message
            }
            else {
                $message = "`t"+ $message
            }
            if ($hold.IsPresent) {
                return ($(([datetime]::Now).toString("HH:mm:ss") + " " + $message))
            }
            else
            {
                Write-Host $(([datetime]::Now).toString("HH:mm:ss") + " " + $message)
            }
        }
        Function Set-ErrorToLog {
            [CmdLetBinding()]
            param (
                [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
                [ValidateNotNullOrEmpty()]
                [object]
                $theError
            )
            ("Message: `t" + $theError.Exception.Message)      | Set-Log -severity No
            ("Activity: `t" + $theError.CategoryInfo.Activity) | Set-Log -severity No
            ("Category: `t" + $theError.CategoryInfo.Category) | Set-Log -severity No
            ("Reason: `t" + $theError.CategoryInfo.Reason)     | Set-Log -severity No
            ("Target: `t" + $theError.CategoryInfo.TargetName) | Set-Log -severity No
        }
        function Get-ScriptName {
            return $MyInvocation.ScriptName | Split-Path -Leaf
        }
    #endregion
    #region Time / Date Info
        $DateRef         = [ScriptData]::Ref
        $Timer           = [system.diagnostics.stopwatch]::StartNew()
        $dateSimple      = $DateRef.toString("yyyyMMdd")
        $dateStamped     = $DateRef.ToString("yyyyMMdd.HHmmss")
    #endregion
    #region Locations
        #region Script Location
            # Gets Script Name
            $scriptFullName = Get-ScriptName
            $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($scriptFullName)
            # Determine script location for PowerShell
            [ScriptData]::Paths.Script = split-path -parent $MyInvocation.MyCommand.Definition
            # Determine current location
            Push-Location
            $SAVWorkDir = Get-Location
            # Force WorkDir to be the same as Script location
            if ($SAVWorkDir.Path -ne [ScriptData]::Paths.Script) {
                Set-Location ([ScriptData]::Paths).Script
            }
            $savedLogs += ("Script name is: $($scriptName)"  | Set-Log -hold)
        #endregion
        #region Location Grouping
            if(([ScriptData]::Paths).grouped -ne $null) {
                $rep = [ScriptData]::Paths.grouped
                $savedLogs += ("Folder(s) will be with: $(([ScriptData]::Paths).grouped)"  | Set-Log -hold)
            }
            else {
                $rep = $scriptName
                $savedLogs += ("Folder(s) will be with: $($rep)"  | Set-Log -hold)
            }
        #endregion
        #region Logs Location
            # Looks where the logs folder should be
            [ScriptData]::Paths.Logs =  "D:\_logs\" + $rep
            # Check for existing Log location or creates it
            if (-not(Test-Path "D:\_logs")) {
                New-Item -Path "D:\" -Name "_logs" -ItemType "directory" | out-Null
                $savedLogs += ("Creation of directory D:\_logs"  | Set-Log -hold)
            }
            if (-not(Test-Path ([ScriptData]::Paths).Logs)) {
                New-Item -Path "D:\_logs\" -Name $rep -ItemType "directory" | out-Null
                $savedLogs += ("Creation of directory $([ScriptData]::Paths.Logs)"  | Set-Log -hold)
            }
        #endregion
        #region Result Location
            # Looks where the Result folder should be
            [ScriptData]::Paths.Results =  "D:\_results\" + $rep
            # Check for existing Result location or creates it
            if (-not(Test-Path "D:\_results" )) {
                New-Item -Path "D:\" -Name "_results" -ItemType "directory" | out-Null
                $savedLogs += ("Creation of directory D:\_results"  | Set-Log -hold)
            }
            if (-not(Test-Path ([ScriptData]::Paths).Results)) {
                New-Item -Path "D:\_results\" -Name $rep -ItemType "directory" | out-Null
                $savedLogs += ("Creation of directory $([ScriptData]::Paths.Results)"  | Set-Log -hold)
            }
        #endregion
    #endregion
    #region Load Models
        #. (([ScriptData]::Paths).Script)\models\DCsModel.ps1
        class ADObjAcl
        {
            static [Boolean]$DoRemediation
            hidden [object]$Acl
            hidden [String]$CompoundPath
            [String]$ObjectDN
            [String]$Name
            [Boolean]$isCriticalSystemObject
            [String]$ObjectCategory
            [String]$ObjectClass
            [String]$whenCreated
            [String]$CurrentOwner
            [Boolean]$ToChange
            [String]$NewOwner
            ADObjAcl([Object]$obj)
            {
                $this.ObjectDN = $obj.distinguishedname.tostring()
                $this.Name = $obj.Name
                $this.isCriticalSystemObject = $obj.isCriticalSystemObject
                $this.ObjectCategory = $obj.ObjectCategory
                $this.ObjectClass = $obj.ObjectClass
                $this.whenCreated = $obj.whenCreated
                $this.GetAcl($false)
                $this.CheckOwner()
                if($this.ToChange -eq $true -and [ADObjAcl]::DoRemediation -eq $true)
                {
                    $this.UpdateOwner()
                    $this.GetAcl($true)
                }
            }
            [void] GetAcl([Boolean]$newAcl)
            {
                $this.CompoundPath = ("AD:" + ($this.ObjectDN))
                $this.Acl = get-acl -Path ($this.CompoundPath)
                if ($newAcl -eq $true)
                {
                    $this.NewOwner = ($this.Acl).Owner
                }
                else
                {
                    $this.CurrentOwner = ($this.Acl).Owner
                }
            }
            [void] CheckOwner()
            {
                If ($this.CurrentOwner -ne 'AMEDMZ\Domain Admins' -and
                $this.CurrentOwner -ne 'BUILTIN\Administrators' -and
                $this.CurrentOwner -ne 'NT AUTHORITY\SYSTEM')
                {
                    $this.ToChange = $true
                }
            }
            [void] UpdateOwner()
            {
                if ($this.ToChange -eq $true)
                {
                    ($this.Acl).SetOwner([Security.Principal.NTaccount]('AMEDMZ\Domain Admins'))
                    set-acl -path ($this.CompoundPath) -AclObject ($this.Acl)
                }
            }
        }
    #endregion
    #region Load AD Module
        try{
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch{
            $savedLogs += ("Unable to load Active Directory PowerShell Module"  | Set-Log -severity Warning -hold)
        }
    #endregion
    #region Gather Active Directory Data
        $domainInfo = Get-ADDomain
        $savedLogs += ("Working in domain: $($domainInfo.NetBIOSName)"  | Set-Log -hold)
        $PDC = $domainInfo.PDCEmulator
    #endregion
    #region Load SendMail Module
        try{
            Import-Module SendMail -ErrorAction Stop
        }
        catch{
            $savedLogs += ("Unable to load SendMail Module"  | Set-Log -severity Warning -hold)
        }
    #endregion
    #region Load CreateHtml Module
        try{
            Import-Module CreateHtml -ErrorAction Stop
        }
        catch{
            $savedLogs += ("Unable to load CreateHtml Module"  | Set-Log -severity Warning -hold)
        }
    #endregion
    #region System Settings
        #region Parameters for Import and Export CSV
            $Param4CSV = @{
                Delimiter         = ';'
                Encoding          = "UTF8"
                NoTypeInformation = $true
                Path              = ""
            }
        #endregion
        #region Config File
            if ($ConfigData.ConfigFileName -ne ".json")
            {
                [ScriptData]::Files.Config = "$(([ScriptData]::Paths).Script)\configs\" + $ConfigData.ConfigFileName
            }
            else
            {
                [ScriptData]::Files.Config = "$(([ScriptData]::Paths).Script)\configs\" + $domainInfo.NetBIOSName + ".json"
            }
            if (-not(test-Path ([ScriptData]::Files.Config)))
            {
                $savedLogs += ("No Config File." | Set-Log -severity Warning -hold)
                if ($Html.IsPresent)
                                                                        {
            # list of mails to send the results TO
            $ConfigData.EmailTo = @(
                list.amer-gts-dws-cps@sgcib.com
                #philippe.grivel-ext@socgen.com
            )
            $savedLogs += ("EmailTo set to: $($ConfigData.EmailTo)." | Set-Log -severity Info -hold)
        }
            }
            else
            {
                $savedLogs += ("Config File : $([ScriptData]::Files.Config)" | Set-Log -severity Success -hold)
                [ScriptData]::hasConfigFile = $true
                # Load Config
                $loadedConfigFull        = ConvertFrom-Json -InputObject ([ScriptData]::Files.Config)
                $loadedConfig            = $loadedConfigFull.Settings
                $ConfigData.LogFileName  = $loadedConfig.LogFileName
                $ConfigData.CsvFileName  = $loadedConfig.CsvFileName
                $ConfigData.ZipFileName  = $loadedConfig.ZipFileName
                $ConfigData.HtmlFileName = $loadedConfig.HtmlFileName
                $ConfigData.EmailTo      = $loadedConfig.EmailTo
            }
        #endregion
        #region Set all File Names and Paths
            # Fichier de Logs
            if ($ConfigData.LogFileName -ne ".log" -and [ScriptData]::hasConfigFile -eq $true)
            {
                [ScriptData]::Files.Log = "$(([ScriptData]::Paths).Logs)\" + $ConfigData.LogFileName
            }
            else
            {
                [ScriptData]::Files.Log = "$(([ScriptData]::Paths).Logs)\" + $dateSimple + "." + $scriptName + ".log"
            }
            $savedLogs += ("LOG File : $([ScriptData]::Files.Log)" | Set-Log -hold)
            # Fichier CSV
            if ($ConfigData.CsvFileName -ne ".csv")
            {
                [ScriptData]::Files.Csv = "$(([ScriptData]::Paths).Results)\" + $ConfigData.CsvFileName
            }
            else
            {
                [ScriptData]::Files.Csv    = "$(([ScriptData]::Paths).Results)\" + $dateStamped + "." + $domainInfo.DNSRoot + "." + $scriptName + ".csv"
            }
            $savedLogs += ("CSV File : $([ScriptData]::Files.Csv)" | Set-Log -hold)
            # Fichier ZIP
            if ($ConfigData.ZipFileName -ne ".zip")
            {
                [ScriptData]::Files.Zip = "$(([ScriptData]::Paths).Results)\" + $ConfigData.ZipFileName
            }
            else
            {
                [ScriptData]::Files.Zip    = "$(([ScriptData]::Paths).Results)\" + $dateStamped + "." + $domainInfo.DNSRoot + "." + $scriptName + ".zip"
            }
            $savedLogs += ("ZIP File : $([ScriptData]::Files.Zip)" | Set-Log -hold)
            # Fichier HTML
            if ($ConfigData.HtmlFileName -ne ".html" -and $html.IsPresent)
            {
                [ScriptData]::Files.Html = "$(([ScriptData]::Paths).Results)\" + $ConfigData.HtmlFileName
            }
            elseif ($html.IsPresent)
            {
                [ScriptData]::Files.Html = "$(([ScriptData]::Paths).Results)\" + $dateStamped + "." + $domainInfo.DNSRoot + "." + $scriptName + ".html"
            }
            if ($html.IsPresent)
            {
                $savedLogs += ("HTML File : $([ScriptData]::Files.Html)" | Set-Log -hold)
            }
        #endregion
    #endregion System Settings
    #region Cleaning old logs
        $savedLogs += ("Cleaning old log files" | Set-Log -hold)
        ([ScriptData]::Paths).Logs | Get-ChildItem  -File |
        Where-object {$_.Name -like "*$($scriptName)*" -and $_.CreationTime -lt $((get-date).adddays(-1))} |
            Remove-ItemProperty -Force -ErrorAction SilentlyContinue
    #endregion
    #region Cleaning old results files (csv / zip / html)
        $savedLogs += ("Cleaning old results files" | Set-Log -hold)
        ([ScriptData]::Paths).Results | Get-ChildItem -File |
        Where-object {$_.Name -like "*$($scriptName)*" -and $_.CreationTime -lt $((get-date).adddays(-2))} |
            Remove-Item -Force -ErrorAction SilentlyContinue
    #endregion
    #region updating UserVariables
        $Title = "[" + ($domainInfo.NetBIOSName).ToUpper() + "]["+ [ScriptData]::Env + "][" + [ScriptData]::Category + "][" + [ScriptData]::Class + "] " + [ScriptData]::ShortTitle
        $mailSubject = $Title
        $savedLogs += ("Settings for $($Title) Done" | Set-Log -severity Success -hold)
    #endregion
    #region Let's START
        # Clean the window
        Clear-Host
        # Start Logging
        Start-Transcript -Path ([ScriptData]::Files.Log) -Force | Set-Log
        # Show what happened before
        $savedLogs |
            ForEach {
                Write-Host $_
            }
        $Settings = [ScriptData]::GetSettings()
        $maxNumberOfChar = (@($Settings.PSObject.Properties.Name) | Measure-Object -Maximum -Property Length).Maximum + 1
        ForEach($key in ($Settings.PSObject.Properties.Name)) {
            $numberOfChar = $key.Length
            ("{0,-$($maxNumberOfChar)}: {1}" -f "$($key)","$($Settings.$key)") | Set-Log
        }
    #endregion
#endregion
# This array will hold the output, if any.
$report = @()
#region MAIN
    $ObjectToRemediate = @(
    'user',
    'computer',
    'group',
    'organizationalUnit'
    )
    $objects = (Get-AdObject -filter * -properties Name,isCriticalSystemObject,ObjectCategory,ObjectClass,whenCreated -server $pdc) |
    Where-object {$_.ObjectClass -in $ObjectToRemediate }
    $howMany = ($objects | Measure-Object).Count
    $results=@()
    [ADObjAcl]::DoRemediation = $false
    If ($DoRemediation.isPresent)
    {
    [ADObjAcl]::DoRemediation = $true
    }
    ForEach($object in $objects)
    {
    $progressParam = @{
    id = 101
    Activity = "Getting Owners of $howMany objects found"
    Status = "Currently on $($objects.IndexOf($object)+1)"
    PercentComplete = ((($objects.IndexOf($object) + 1) / $howMany)*100)
    }
    Write-Progress @progressParam
    $_obj = [ADObjAcl]::new($object)
    $results += $_obj
    }
    $report = $results |
    Where-Object {$_.ToChange -eq $true}
    #endregion MAIN
#region End
    if (($report | Measure-Object).Count -gt 0)
    {
        $report | Out-File -FilePath "$([ScriptData]::Paths.Results)\$($domainInfo.DNSRoot).$($scriptName).csv" -Encoding utf8
        "Data preparation is done and will be saved." | Set-Log
        # Will store all files to attach to Email
        $toAttachInMail = @()
        #region CSV
            #export the hash to a csv file
            $Param4CSV.Path = [ScriptData]::Files.Csv
            $report | Export-Csv @Param4CSV -Append -Force
            # display where the file is
            "Results exported to csv : $($Param4CSV.Path)" | Set-Log
            if ($mail.IsPresent) { $toAttachInMail += $Param4CSV.Path }
        #endregion
        #region HTML
            if ($html.IsPresent)
            {
                $htmlParts = @()
                # Create an HTML table for $report
                $htmlReportTable = Set-HtmlTable -TblObjects $report -Divtitle $([ScriptData]::ShortTitle)
                $htmlParts += $htmlReportTable
                # Addings
                $htmlReportFooter = "<div>"
                $htmlReportFooter += "</br>Execution time: $($Timer.Elapsed)"
                $htmlReportFooter += "</br>On: $(([System.Net.Dns]::GetHostByName(($env:computerName)).HostName))</div>"
                $htmlParts += $htmlReportFooter
                # Create an HTML page with this table
                $htmlReport = Set-HtmlPage -Tabtitle $([ScriptData]::ShortTitle) -Fragments $htmlParts
                #export the hash to an html file
                $htmlReport | out-file -FilePath ([ScriptData]::Files.Html)
                # display where the files are
                "`HTML is here : $([ScriptData]::Files.Html)" | Set-Log
            }
        #endregion
        #region ZIP
            # Compress result files in one zip file
            Compress-Archive -Path ([ScriptData]::Files.csv) -DestinationPath ([ScriptData]::Files.Zip) -Force
            if ($mail.IsPresent -and $html.IsPresent)
            {
                Compress-Archive -Path ([ScriptData]::Files.Html) -Update -DestinationPath ([ScriptData]::Files.Zip)
            }
            if ($mail.IsPresent)
            {
                # Stops log
                "Stopping logs to include it in Zip" | Set-log
                Stop-Transcript | Set-Log
                # $Log = $(Get-content ([ScriptData]::Files.Html) | Out-String)
                # $Log | Compress-Archive  -Update -DestinationPath ([ScriptData]::Files.Zip)
                Compress-Archive -Path ([ScriptData]::Files.Log) -Update -DestinationPath ([ScriptData]::Files.Zip)
                # Start Logging
                Start-Transcript -Path ([ScriptData]::Files.Log) -Append | Set-Log
                $toAttachInMail = ((get-item ([ScriptData]::Files.Zip)).FullName)
            }
        #endregion
        #region SENDMAIL
            # Send a mail containing CSV files attached
            if ($mail.IsPresent) {
                $sendMailOptions = @{
                    MailSubject = $mailSubject
                    HTMLBody    = ""
                    Attachment  = $toAttachInMail
                    To          = [ScriptData]::ConfigData.EmailTo
                }
                if ($html)
                {
                    $sendMailOptions.HTMLBody = $(Get-content ([ScriptData]::Files.Html) | Out-String)
                }
                else 
                {
                    $sendMailOptions.HTMLBody = $(Get-content $($Param4CSV.Path) | Out-String)
                }
                Send-Mail @sendMailOptions
            }
        #endregion
        # Show results within HTML file
        if ($html.isPresent) { invoke-expression ([ScriptData]::Files.Html) }
    }
    else
    {
        "Report is NULL. Stopping now." | Set-Log -severity Error
    }
    #region Final
        $Timer.Stop()
        $runTime = "$($Timer.Elapsed)"
        "Script Runtime full: $runtime" | Set-Log
        "The End." | Set-Log
        # Stops log
        Stop-Transcript | Set-Log
        # Restore location
        Pop-Location
    #endregion
#endregion
