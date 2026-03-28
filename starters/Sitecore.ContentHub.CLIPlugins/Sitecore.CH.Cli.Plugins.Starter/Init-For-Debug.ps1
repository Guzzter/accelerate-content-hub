$pluginName = Split-Path -Path $PSScriptRoot -Leaf
$pluginsDirectoryName = "plugins";

$chCliPath = "${home}\.ch-cli"
$chCliPluginsPath = "${chCliPath}\${pluginsDirectoryName}"
$chCliPluginsLinkPath = "${chCliPluginsPath}\${pluginName}"

$outputPath = "$PSScriptRoot\Output\Debug"
$outputPluginsPath = "${outputPath}\${pluginsDirectoryName}"
$outputPluginsLinkPath = "${outputPluginsPath}\${pluginName}"

If(!(test-path -PathType container $outputPluginsLinkPath))
{
      New-Item -ItemType Directory -Path $outputPluginsLinkPath
}

New-Item -ItemType SymbolicLink -Value $outputPluginsLinkPath -Path $chCliPluginsLinkPath