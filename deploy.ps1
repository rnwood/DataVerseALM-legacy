[CmdletBinding()]
param([Parameter(Mandatory)][string] $url)

$ErrorActionPreference="Stop"
$VerbosePreference="Continue"

try {
    $root = (split-path $MyInvocation.MyCommand.Source)
    write-verbose "Root: $root"
    . $root\common.ps1
    & $root\import.ps1 -url $url

} catch {
    $_.ScriptStackTrace | Write-Host
    throw
}