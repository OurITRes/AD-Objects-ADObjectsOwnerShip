class Files
{
    static hidden [Int]$FilesToKeep = 7
    static hidden [String]$Base

    [String]$Folder
    [String]$Name
    [String]$FullPath
    #region Constructors
        Files()
        {
            $this.Folder = [Files]::Base
        }

        Files([String]$folderName)
        {
            $this.Folder = $folderName
        }
    #endregion#>
    #region Methods
        [Files] SetName([String]$Name)
        {
            $this.Name = $Name
            return $this
        }

        [Files] Finalize()
        {
            if ([Files]::Base -ne $this.Folder)
            {
                $this.FullPath = [Files]::Base + "\" + $this.Folder + "\" + $this.Name
                $this.BuildIfNeeded()
                $this.CleanIfNeeded()
            }
            else
            {
                $this.FullPath = [Files]::Base + "\" + $this.Name
            }
            return $this
        }

        [void] CleanIfNeeded()
        {
            $_subFolder = [Files]::Base + "\" + $this.Folder
            $_subFolder |
                Get-ChildItem  -File |
                Sort-Object CreationTime -desc|
                Select-Object -Skip ([files]::FilesToKeep) |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }

        [void] BuildIfNeeded()
        {
            $_subFolder = [Files]::Base + "\" + $this.Folder
            if (-not(Test-Path $_subFolder))
            {
                New-Item -Path ([Files]::Base) -Name ($this.Folder) -ItemType "directory" | out-Null
            }
        }
    #endregion#>
}