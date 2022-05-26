<#
    Exemple d'utilisation

[Log]::new().Out()
[Log]::new("test").Out()
[Log]::new("test2",[States]::Success).Out()
[Log]::Hold()
[Log]::Hold("test3")
[Log]::Hold("test4",[States]::Error)
[Log]::Release()
[Log]::Store

#>
enum States
{
    Success
    Info
    Warning
    Stats
    Error
    No
    Title
}

class Log
{
    static [Log[]]$Store = @()

    [States]$Severity = [States]::Info
    [String]$Message
    [String]$Output
    [Boolean]$isReadyToPrint
    [Boolean]$hasBeenPrinted

    #region Constructors
        Log()
        {
            $this.SetMessage($null).SetSeverity([States]::Info).SetOuput().SetReady()
        }

        Log([String]$txt)
        {
            $this.SetMessage($txt).SetSeverity([States]::Info).SetOuput().SetReady()
        }

        Log([String]$txt, [States]$sev)
        {
            $this.SetMessage($txt).SetSeverity($sev).SetOuput().SetReady()
        }
    #endregion
    #region Set Methods
        [Log] SetMessage([String]$txt)
        {
            if ($txt -ne $null -and $txt -ne "")
            {
                $this.Message = $txt
            }
            return $this
        }

        [Log] SetSeverity([States]$sev)
        {
            $this.Severity = $sev       
            return $this
        }

        [String] SetSpacing([States]$sev)
        {
            $_spaces = @{
                Success = 1
                Info    = 4
                Warning = 1
                Stats   = 3
                Error   = 3
                No      = 11
                Title   = 0
            }
            $_spacing = " "*($_spaces."$($sev)")
            return $_spacing 
        }

        [Log] SetOuput()
        {
            $this.Output = ([datetime]::Now).toString("HH:mm:ss") + " " + $this.SetSpacing($this.Severity)
       
            if ($this.Severity -ne "Title" -and $this.Severity -ne "No")
            {
                $this.Output =  $this.Output + "$($this.Severity)".Toupper() + ":`t" + $this.Message
            }
            else
            {
                $this.Output =  $this.Output + $this.Message
            }
            return $this
        }

        [Log] SetReady()
        {
            $this.isReadyToPrint = $true
            return $this
        }

        [Log] SetPrinted()
        {
            $this.hasBeenPrinted = $true
            return $this
        }
    #endregion
    #region Get Methods
        [String] Out()
        {
            if ($this.isReadyToPrint -eq $true -and $this.hasBeenPrinted -eq $false)
            {
                $this.SetPrinted()
                [Log]::Store += $this
                return $this.Output
            }
            else
            {
                return $null
            }
        }
    #endregion
    #region Static Methods
        static [void] Hold()
        {
            $log = [Log]::new()
            if ($log.isReadyToPrint)
            {
                [Log]::Store += $log
            }
        }

        static [void] Hold([String]$txt)
        {
            $log = [Log]::new([String]$txt)
            if ($log.isReadyToPrint)
            {
                [Log]::Store += $log
            }
        }

        static [void] Hold([String]$txt, [States]$sev)
        {
            $log = [Log]::new([String]$txt, [States]$sev)
            if ($log.isReadyToPrint)
            {
                [Log]::Store += $log
            }
        }

        static [String[]] Release()
        {
            $_output = @()
            [Log]::Store |
                ForEach-Object {
                    if ($_.hasBeenPrinted -eq $false)
                    {
                        $_.hasBeenPrinted = $true
                        $_output += $_.Output
                    }
                }
            return $_output
        }
    #endregion
}