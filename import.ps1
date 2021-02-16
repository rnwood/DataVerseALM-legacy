[CmdletBinding()]
param([Switch] $dataonly, [Switch] $nodata, [string] $url)

$ErrorActionPreference="Stop"
$VerbosePreference="Continue"

try {
    $root = (split-path $MyInvocation.MyCommand.Source)
    write-verbose "Root: $root"
    . $root\common.ps1

	if (-not $url) {
		$url = getConfigValue DEVURL
	}

	$connection = getCrmConnection $url
	
    if (-not $dataonly) {

		foreach($dependency in $dependencies) {
			$dependency = (resolve-path $dependency).Path
			importCrmSolution -solutionfile $dependency -connection $connection -managed
		}

		try {
			$tempfile = $null
			$managed=$false
			if (test-path $root\$solutionname.zip) {
				$zipfile = "$root\$solutionname.zip"
				$managed=$true
			} else {
				$zipfile = $tempfile = [IO.Path]::GetTempFileName()
				packCrmSolution -folder $root\solution -zipfile $zipfile
			}
			
			importCrmSolution -solutionfile $zipfile -connection $connection -managed:$managed
		} finally {
			if ($tempfile) {
				remove-item -Force $zipfile
			}
		}

		publishCrmCustomizations -connection $connection
	}

	if (-not $nodata)  {
		importCrmData -connection $connection -datadir $root\data


	}
    
} catch {
    $_.ScriptStackTrace | Write-Host
    throw
}