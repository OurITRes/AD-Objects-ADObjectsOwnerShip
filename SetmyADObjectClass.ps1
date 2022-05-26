class myADObject
{
    static [Boolean]$DoRemediation
    static [String]$DomainNetBios
    static [String]$Pdc
    static hidden [String[]]$SecuredOwners

    [String]$ObjectCategory
    [String]$ObjectClass
    [Boolean]$isCriticalSystemObject
    [String]$Name
    [String]$whenCreated
    [String]$CurrentOwner
    [Boolean]$isToChange
    [String]$ChangedTo
    [Boolean]$Done
    [Boolean]$ManualChange
    [String]$DN

    hidden [String]$ObjectGUID
    hidden [object]$CurrentAcl
    hidden [String]$CompoundPath

    #region Constructors
        myADObject(){}

        myADObject([String]$NetBiosName)
        {
            $this.SetDomainNetBios([String]$NetBiosName)
        }
        myADObject([Object]$obj)
        {
            $this.SetDomainNetBios().SetmyADObject($obj).ExamineAcl().RemediateIfNeeded()
        }
    #endregion#>
    #region hidden Methods
        [myADObject] hidden GetDomainPDC([String]$NetBiosName)
        {
            [myADObject]::Pdc = (Get-ADDomain $NetBiosName).PDCEmulator
            return ( $this )   
        }

        [myADObject] hidden SetDomainNetBios()
        {
            if ( [myADObject]::DomainNetBios -ne "" -and
                 [myADObject]::DomainNetBios -ne $null  )
            {
                if ( [myADObject]::Pdc -ne "" -and
                 [myADObject]::Pdc -ne $null  )
                {
                    return $this
                }
                else
                {
                    return ( $this.GetDomainPDC([myADObject]::DomainNetBios) )
                }
            }
            else
            {
                return ( $this.SetDomainNetBios( (Get-ADDomain).NetBIOSName ) )
            }
        }

        [myADObject] hidden UpdateOwner()
        {
            ($this.CurrentAcl).SetOwner([Security.Principal.NTaccount](([myADObject]::SecuredOwners)[0]))
            set-acl -path ($this.CompoundPath) -AclObject ($this.CurrentAcl)
            return $this
        }
        
        [myADObject] hidden GetAcl()
        {
            Try{
                $this.CurrentAcl = ( Get-Acl -Filter {ObjectGUID -eq ($this.ObjectGUID)} -ErrorAction SilentlyContinue )
            }
            catch
            {
                $this.ManualChange = $true
            }
            return $this
        }

        [myADObject] hidden CheckOwner()
        {
            if ($this.CurrentOwner -notin ([myADObject]::SecuredOwners))
            {
                $this.isToChange = $true
            }
            else
            {
                $this.isToChange = $false
            }
            return $this
        }

        [myADObject] hidden CheckUpdate()
        {
            if ( $this.isToChange -eq $true -and
                 $this.CurrentOwner -eq $this.ChangedTo)
            {
                $this.Done = $true
            }
            return $this
        }

        [myADObject] hidden Remediate()
        {
            if ($this.ManualChange -eq $false)
            {
                $this.UpdateOwner().GetAcl()
                if ($this.ManualChange -eq $false)
                {
                    $this.ChangedTo = ($this.CurrentAcl).Owner
                }
            }
            return $this
        }
    #endregion#>
    #region Methods
        [myADObject] SetDomainNetBios([String]$NetBiosName)
        {
            if ( [myADObject]::DomainNetBios -ne "" -and
                 [myADObject]::DomainNetBios -ne $null  )
            {
                if ( [myADObject]::Pdc -eq "" -or
                 [myADObject]::Pdc -eq $null  )
                {
                    $this.GetDomainPDC([myADObject]::DomainNetBios)
                }
            }
            else
            {
                [myADObject]::DomainNetBios = $NetBiosName
                $this.GetDomainPDC([myADObject]::DomainNetBios)
            }
            [myADObject]::SecuredOwners = @(
                "$(([myADObject]::DomainNetBios) + '\Domain Admins')",
                "$(([myADObject]::DomainNetBios) + '\Enterprise Admins')",
                'BUILTIN\Administrators',
                'BUILTIN\Administrateurs',
                'NT AUTHORITY\SYSTEM'
            )
            return $this
        }

        [myADObject] SetmyADObject([Object]$obj)
        {
            $this.ObjectCategory = $obj.ObjectCategory
            $this.ObjectClass    = $obj.ObjectClass
            $this.isCriticalSystemObject = $obj.isCriticalSystemObject
            $this.Name          = $obj.Name
            $this.whenCreated   = $obj.whenCreated
            $this.DN            = $obj.distinguishedname
            $this.ObjectGUID    = $obj.ObjectGUID
            $this.CompoundPath  = ("AD:" + ($this.DN))

            return $this
        }

        [myADObject] ExamineAcl()
        {
            $this = $this.GetAcl()
            if ($this.ManualChange -eq $false)
            {
                $this.CurrentOwner  = ($this.CurrentAcl).Owner
            }
            $this = $this.CheckOwner()
            return $this
        }

        [myADObject] RemediateIfNeeded()
        {
            if ($this.isToChange -eq $true -and
                $this.ManualChange -eq $false -and
                [myADObject]::DoRemediation -eq $true -and
                [myADObject]::DomainNetBios -ne "" -and
                [myADObject]::DomainNetBios -ne $null)
            {
                return ($this.Remediate().CheckUpdate())
            }
            else
            {
                return $this
            }
        }
    #endregion#>
}