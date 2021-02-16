[CmdletBinding()]
param([string] $url, [string] $deployurl, [string] $commitcomment)

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
	
	if ($commitcomment) {
		$solutioninfo = getcrmsolutioninfo -connection $connection $solutionname
		$solutioninfo["version"] = [string] (Read-Host -prompt "Last version is $($solutioninfo["version"]). Enter release version")

		$connection.Update($solutioninfo)
		
		& "$root/export.ps1"

		$process = Start-Process -FilePath "git" -ArgumentList ("-C", "$root", "add", "-A") -NoNewWindow -Wait -PassThru
		if ($process.ExitCode -ne 0) {
			throw "Git add failed. Exit code: $($process.ExitCode)"
		}

		$process = Start-Process -FilePath "git" -ArgumentList ("-C", "$root", "commit", "-m", "`"$commitcomment`"") -NoNewWindow -Wait -PassThru
		if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 1) {
			throw "Git commit failed. Exit code: $($process.ExitCode)"
		}

		$process = Start-Process -FilePath "git" -ArgumentList ("-C", "$root", "push", "origin", "master") -NoNewWindow -Wait -PassThru
		if ($process.ExitCode -ne 0) {
			throw "Git push failed. Exit code: $($process.ExitCode)"
		}
	}


	$version = (getcrmsolutioninfo -connection $connection $solutionname)["version"]
	$version = "$version"
	write-host "Creating release with version: $version"

	$releasepath = "$root\.release\$version"

	if (test-path $releasepath){
		get-childitem $releasepath -Recurse | remove-item -force -Recurse
		remove-item -Recurse -force $releasepath
	}

	new-item -ItemType Directory $releasepath
	copy-item $root\import.ps1 $releasepath
	copy-item $root\common.ps1 $releasepath
	copy-item $root\deploy.ps1 $releasepath
	copy-item -recurse $root\.modules.scripts $releasepath
	copy-item -recurse $root\.packages.scripts $releasepath
	copy-item $root\modules.scripts.config $releasepath
	copy-item $root\packages.scripts.config $releasepath
	copy-item -recurse $root\data $releasepath
	copy-item -recurse $root\dependencies $releasepath
	copy-item -recurse $root\upgradescripts $releasepath

	exportCrmSolution -connection $connection -managed -uniquename $solutionname -solutionfile $releasepath\$solutionname.zip

	$process = Start-Process -FilePath "git" -ArgumentList ("-C", "$root", "tag", "$version") -NoNewWindow -Wait -PassThru
	if ($process.ExitCode -ne 0) {
		throw "Git tag failed. Exit code: $($process.ExitCode)"
	}
	$process = Start-Process -FilePath "git" -ArgumentList ("-C", "$root", "push", "origin", "$version") -NoNewWindow -Wait -PassThru
	if ($process.ExitCode -ne 0) {
		throw "Git push failed. Exit code: $($process.ExitCode)"
	}

} catch {
    $_.ScriptStackTrace | Write-Host
    throw
}

if ($deployurl) {
	& $releasepath\deploy.ps1 -url $deployurl
}