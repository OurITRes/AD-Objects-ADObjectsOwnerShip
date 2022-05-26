<#
.NOTES
===========================================================================
Created with:     Windows Powershell ISE
Created on:       05/25/2022 01:00 PM
Organization:     None, Inc.
Contact:          Philippe Grivel, phgrivel@gmail.com
Filename:         Secure-ADObjectOwnerShip.ps1

===========================================================================

.SYNOPSIS
Remediate AD Objects security ownership

.DESCRIPTION
Scan AD Objects selected and remediate security ownership if needed
The option ADDomain is here to receive an ADDomain Object from any reachable domain

.EXAMPLE
  .\Secure-ADObjectOwnerShip.ps1
  Will not do remediation

.EXAMPLE
  .\Secure-ADObjectOwnerShip.ps1 -WithRemediation
  Will do the remediation

.OUTPUTS
Log File: YYYYMMDD.HHmmSS.adobjectownershipremediation.log
Result File: YYYYMMDD.adobjectownershipremediation.csv

#>

[CmdLetBinding()]
Param(
    [Parameter(Mandatory = $false, valueFromPipeLine = $true)]
    [Object]$ADDomain = (Get-ADDomain),
 
    [Parameter(Mandatory = $false)]
    [Switch]$WithRemediation
)

#region Preparation
    $Error.Clear()
    Import-Module activedirectory
    Clear-Host
    #region Functions
        function Get-ScriptName
        {
            return ( Get-ChildItem $MyInvocation.PSCommandPath | Select -Expand Name )
            #return $MyInvocation.ScriptName | Split-Path -Leaf
        }
        $_ScriptName = Get-ScriptName
        function Get-ScriptDirectory {
          return Split-Path -Parent $MyInvocation.PSCommandPath
        }
        $_ScriptLocation = Get-ScriptDirectory
        . "$_ScriptLocation\..\Snippets\FuncGoToEnd.ps1"
    #endregion#>
    #region Enums and Classes
        . "$_ScriptLocation\..\Snippets\SetLogClass.ps1"
        [Log]::Store = @()
        . "$_ScriptLocation\..\Snippets\SetFilesCLass.ps1"
        . "$_ScriptLocation\SetmyADObjectClass.ps1"
        [Log]::Hold("Functions, Enums and Classes Loaded",[States]::Success)
    #endregion#>
    #region Dates and Times
        . "$_ScriptLocation\..\Snippets\SetDateAndTimeInformation.ps1"
        [Log]::Hold("Dates and Times Set",[States]::Success)
    #endregion#>
    #region Files and Paths
        $_ScriptLocation | . "$_ScriptLocation\..\Snippets\SetFiles.ps1"
        $_ScriptName     | . ..\Snippets\SetPaths.ps1
        [Log]::Hold("Files and Paths Set",[States]::Success)
    #endregion#>
#endregion#>
#region Main
    . ..\Snippets\SetStart.ps1
    
        #region Check if Admin
            #Returns a WindowsIdentity object that represents the current Windows user.
            $CurrentWindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            #creating a new object of type WindowsPrincipal, and passing the Windows Identity to the constructor.
            $CurrentWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($CurrentWindowsIdentity)
            #Return True if specific user is Admin else return False
            if ($CurrentWindowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) 
            {
                [Log]::Hold("Admin permission is available and Code is running as administrator",[States]::Success)
            }
            else 
            {
                [Log]::Hold("Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again.",[States]::Error)
                . ..\Snippets\SetEnd.ps1
            }
        #endregion#>

    # Check if we need to do the remediation
    [myADObject]::DoRemediation = $false
    if ($WithRemediation.IsPresent -eq $true)
    {
        [myADObject]::DoRemediation = $true
    }
 
    # List of all AD object types to remediate
    $_ObjectsToRemediate = @(
        'user',
        'computer',
        'group',
        'organizationalUnit'
    )
 
    # List of all properties required for each object
    $_WantedObjProps = @(
        "Name",
        "isCriticalSystemObject",
        "ObjectCategory",
        "ObjectClass",
        "whenCreated"
    )
 
    # Emptying important list variables
    $ADObjects = @()
    $Treated   = @()
   
    # Setting the path for the exported results
    $_paramEC.Path = ($Paths.Results.FullPath)
 
    # Script work is here
    [myADObject]::DomainNetBios = $ADDomain.NetBiosName
    [myADObject]::Pdc = $ADDomain.PDCEmulator
   
    [Log]::new("Query for All Objects to prevent invalid enumeration context,", [States]::Info).Out()
    [Log]::new("must be less than 30 min and should be around 14 min.", [States]::Info).Out()
    $ADObjects = Get-AdObject -filter * -properties $_WantedObjProps -server ([myADObject]::Pdc) -ResultSetSize $null
    [Log]::new("Done.", [States]::Success).Out()
    [Log]::new("Counting objects to deal with.", [States]::Info).Out()
    $howManyGet = ($ADObjects | Measure-Object).Count
    [Log]::new("Piping $($howManyGet.ToString('N0')) Objects to remediation process.", [States]::Info).Out()
    $CounterGet=0
    $counterSel=0
    $ADObjects |
        ForEach-Object {
            Write-Host ([Log]::new("$($_.ObjectClass) Received: $(($counterGet ++)) / $howManyGet $($_.Name)", [States]::Info).Out())
            $_
        } |
        Where-object  {
            $_.ObjectClass -in $_ObjectsToRemediate
        } |
        ForEach-Object {
            Write-Host ([Log]::new("$($_.ObjectClass) $($_.Name) Selected:  #$(($counterSel ++))", [States]::Success).Out())
            $_
        } |
        ForEach-Object {
            $CurrentObject = [myADObject]::new($_)
            $Treated += $CurrentObject
            $CurrentObject
        } |
        Export-Csv @_paramEC
#endregion#>
#region End
 
    #region Final
        $number = $Error.count
        if ( $number -gt 0 )
        {
            [Log]::new("$number Errors happened. Here is the list:",[States]::Title).Out()
            $Error | ForEach-object { $_ }
        }
        [Log]::new("Objects in AD: $howManyGet.", [States]::Info).Out()
        [Log]::new("Objects To Secure: $counterSel.", [States]::Info).Out()
        [Log]::new("Objects Secured: $((($Treated | Where-Object {$_.isToChange -eq $true -and $_.Done -eq $true}) | Measure-Object).Count).", [States]::Info).Out()
        [Log]::new("Errors: $number.", [States]::Info).Out()
        . ..\Snippets\SetEnd.ps1
       
    #endregion#>
   
#endregion#>