[CmdletBinding()]
param([Switch] $forcedependencies, [Switch] $dataonly, [Switch] $nodata, [string] $url)

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
		publishCrmCustomizations -connection $connection

		$zipfile = [IO.Path]::GetTempFileName()
		try {
			exportCrmSolution -uniquename $solutionname -solutionfile $zipfile -connection $connection
			unpackCrmSolution -folder $root\solution -zipfile $zipfile
		} finally {
			remove-item -Force $zipfile
		}
	
	}

	if (-not $nodata) {
		exportCrmData -connection $connection -datadir $root\data
	}
	

} catch {
    $_.ScriptStackTrace | Write-Host
    throw
}