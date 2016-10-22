#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($hostinvocation -ne $null)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

function Connect-O365 {
    [CmdletBinding()]
    param ()
    process {
        $Params = @{
            ConfigurationName = 'Microsoft.Exchange'
            ConnectionUri = 'https://ps.outlook.com/powershell/'
            Credential = $Global:Credential
            Authentication = 'Basic'
            AllowRedirection = $true
            WarningAction = 'SilentlyContinue'
            Name = 'O365'
            ErrorAction = 'Stop'
        }
        try {
            $Global:o365Session = New-PSSession @Params
            Import-PSSession $Global:o365Session -Prefix 'O365' -AllowClobber -DisableNameChecking -ErrorAction Stop | Out-Null
        }
        catch {
            $Global:o365Session = $Null
        }
        
    }
    
}