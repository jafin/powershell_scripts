param([Parameter(Mandatory=$true)][string[]] $ComputerName,
      [switch] $Clobber)


# BSD 3-clause license
# Copyright (C) 2011-2015, Joakim Svendsen
# Author: Joakim Borger Svendsen

# I originally wrote this in an early stage of learning PowerShell, with old habits in play.
# The script needs a complete rewrite, and could do with runspaces for concurrency,
# but I don't feel like it now.

# Originally written around 2011-2012 some time.
# 2015-10-31: Tacked on support for .NET frameworks 4.5.x and 4.6.


##### START OF FUNCTIONS #####

function ql { $args }

function Quote-And-Comma-Join {
    
    param([Parameter(Mandatory=$true)][string[]] $Strings)
    
    # Replace all double quotes in the text with single quotes so the CSV isn't messed up,
    # and remove the trailing newline (all newlines and carriage returns).
    $Strings = $Strings | ForEach-Object { $_ -replace '[\r\n]', '' }
    ($Strings | ForEach-Object { '"' + ($_ -replace '"', "'") + '"' }) -join ','
    
}

##### END OF FUNCTIONS #####

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$StartTime = Get-Date
"Script start time: $StartTime"

$Date = (Get-Date).ToString('yyyy-MM-dd')
$OutputOnlineFile  = ".\DotNetOnline-${date}.txt"
$OutputOfflineFile = ".\DotNetOffline-${date}.txt"
$CsvOutputFile = ".\DotNet-Versions-${date}.csv"

if (-not $Clobber) {
    
    $FoundExistingLog = $false
    
    foreach ($File in $OutputOnlineFile, $OutputOfflineFile, $CsvOutputFile) {
        
        if (Test-Path -PathType Leaf -Path $File) {
            
            $FoundExistingLog = $true
            "$File already exists"
            
        }
    
    }
    
    if ($FoundExistingLog -eq $true) {
        $Answer = Read-Host "The above mentioned log file(s) exist. Overwrite? [yes]"
        if ($Answer -imatch '^n') { 'Aborted'; exit 1 }
    }
}

# Deleting existing log files if they exist (assume they can be deleted...)
Remove-Item $OutputOnlineFile -ErrorAction SilentlyContinue
Remove-Item $OutputOfflineFile -ErrorAction SilentlyContinue
Remove-Item $CsvOutputFile -ErrorAction SilentlyContinue

$Counter    = 0
$DotNetData = @{}
$DotNetVersionStrings = ql v4\Client v4\Full v3.5 v3.0 v2.0.50727 v1.1.4322
$DotNetRegistryBase   = 'SOFTWARE\Microsoft\NET Framework Setup\NDP'

foreach ($Computer in $ComputerName) {
    
    $Counter++
    $DotNetData.$Computer = New-Object PSObject
    
    # Skip malformed lines (well, some of them)
    if ($Computer -notmatch '^\S') {
        Write-Host -Fore Red "Skipping malformed item/line ${Counter}: '$Computer'"
        Add-Member -Name Error -Value "Malformed argument ${Counter}: '$Computer'" -MemberType NoteProperty -InputObject $DotNetData.$Computer
        continue
    }
    
    if (Test-Connection -Quiet -Count 1 $Computer) {
        Write-Host -Fore Green "$Computer is online. Trying to read registry."
        $Computer | Add-Content $OutputOnlineFile
        
        # Suppress errors when trying to open the remote key
        $ErrorActionPreference = 'SilentlyContinue'
        $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        $RegSuccess = $?
        $ErrorActionPreference = 'Stop'
                
        if ($RegSuccess) {
            Write-Host -Fore Green "Successfully connected to registry of ${Computer}. Trying to open keys."
            foreach ($VerString in $DotNetVersionStrings) {
                if ($RegKey = $Registry.OpenSubKey("$DotNetRegistryBase\$VerString")) {
                    #"Successfully opened .NET registry key (SOFTWARE\Microsoft\NET Framework Setup\NDP\$verString)."
                    if ($RegKey.GetValue('Install') -eq '1') {
                        #"$computer has .NET $verString"
                        Add-Member -Name $VerString -Value 'Installed' -MemberType NoteProperty -InputObject $DotNetData.$Computer
                    }
                    else {
                        Add-Member -Name $VerString -Value 'Not installed' -MemberType NoteProperty -InputObject $DotNetData.$Computer
                    }
                }
                else {
                    Add-Member -Name $VerString -Value 'Not installed (no key)' -MemberType NoteProperty -InputObject $DotNetData.$Computer
                }
            }
            # Tacking on 4.5.x and 4.6 detection, as someone requested... this script really needs a rewrite to be
            # more standards-conforming, but I'm mentally exhausted.
            $RegKey = $Null
            if ($RegKey = $Registry.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            {
                if ($DotNet4xRelease = [int] $RegKey.GetValue('Release'))
                {
					#394254 4.6.1 on Windows 10 or 394271 for all other OS
					if ($DotNet4xRelease -ge 394254)  
					{
						$DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value '4.6.1 or later'
					}
                    elseif ($DotNet4xRelease -ge 393295)
                    {
                        $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value '4.6 or later'
                    }
                    elseif ($DotNet4xRelease -ge 379893)
                    {
                        $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value '4.5.2 or later'
                    }
                    elseif ($DotNet4xRelease -ge 378675)
                    {
                        $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value '4.5.1 or later'
                    }
                    elseif ($DotNet4xRelease -ge 378389)
                    {
                        $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value '4.5 or later'
                    }
                    else
                    {
                        $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value 'Universe imploded'
                    }
                }
                else
                {
                    $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value "Error (no key?)"
                }
            }
            else
            {
                $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value 'Not installed (no key)'
            }
        }
        
        # Error opening remote registry
        else {
            Write-Host -Fore Yellow "${Computer}: Unable to open remote registry key."
            Add-Member -Name Error -Value "Unable to open remote registry: $($Error[0].ToString())" -MemberType NoteProperty -InputObject $DotNetData.$Computer
            $DotNetData.$Computer | Add-Member -MemberType NoteProperty -Name '>=4.x' -Value 'Unknown'
        }
    }
    
    # Failed ping test
    else {
        Write-Host -Fore Yellow "${Computer} is offline."
        Add-Member -Name Error -Value "No ping reply" -MemberType NoteProperty -InputObject $DotNetData.$Computer
        $Computer | Add-Content $OutputOfflineFile
    }    
}

$CsvHeaders = @('Computer', '>=4.x') + @($DotNetVersionStrings) + @('Error')
$HeaderLine = Quote-And-Comma-Join $CsvHeaders
Add-Content -Path $CsvOutputFile -Value $HeaderLine

# Process the data and output to manually crafted CSV.
foreach ($Computer in $DotNetData.Keys | Sort-Object) { #| ForEach-Object {
    
    #$Computer = $_.Name
    
    # I'm building a temporary hashtable with all $CsvHeaders
    $TempData = @{}
    $TempData.'Computer' = $Computer
    #Write-Verbose 'Before'
    #Write-Verbose 'After'
    # This means there's an "Error" note property.
    if (Get-Member -InputObject $DotNetData.$Computer -MemberType NoteProperty -Name Error) {
        
        # Add the error to the temp hash.
        $TempData.'Error' = $DotNetData.$Computer.Error
        
        # Populate the .NET version strings with "Unknown".
        foreach ($VerString in $DotNetVersionStrings) {
            $TempData.$VerString = 'Unknown'
        }
        $TempData.'>=4.x' = 'Unknown'
    }
    
    # No errors. Assume all .NET version fields are populated.
    else {
        # Set the error key in the temp hash to "-"
        $TempData.'Error' = '-'
        
        foreach ($VerString in $DotNetVersionStrings) {
            $TempData.$VerString = $DotNetData.$Computer.$VerString
        }
        $TempData.'>=4.x' = $DotNetData.$Computer.'>=4.x'
    
    }
    # Now we should have "complete" $TempData hashes.
    # Manually craft CSV data. Headers were added before the loop.
    
    # The array is for ordering the output predictably.
    $TempArray = @()
    
    foreach ($Header in $CsvHeaders) {
        $TempArray += $TempData.$Header
    }
    
    $CsvLine = Quote-And-Comma-Join $TempArray
    Add-Content -Path $CsvOutputFile -Value $CsvLine

	#dump output to screen
	Import-Csv $CsvOutputFile | ft -auto
}

@"
Script start time: $StartTime
Script end time:   $(Get-Date)
Output files: $CsvOutputFile, $OutputOnlineFile, $OutputOfflineFile
"@
