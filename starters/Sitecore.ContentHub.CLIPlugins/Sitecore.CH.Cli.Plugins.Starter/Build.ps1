param([string]$ProjectDir, [string]$OutputPath, [string]$Configuration, [string]$ProjectName)

$sourceDir = "${ProjectDir}${OutputPath}"
$targetDir = "${ProjectDir}Output\${Configuration}\plugins\${ProjectName}"

# Create the output directory
[void](New-Item -Path $targetDir -ItemType Directory -Force)

function CopyPackage([string]$name) {
	If(($name -ne "") -and ($name -ne "Sitecore.CH.Cli.Core"))
	{
		Write-Output "trying to copy ${name}"
		Copy-Item -Path "${sourceDir}\${name}.*" -Destination $targetDir
	}
}

function ProcessProject([string]$projectFile) {
	Write-Output "Process Project for ${projectFile}"
	[xml]$projectFileData = Get-Content -Path $projectFile

	## Copy nuget dependencies to output folder
	$projectFileData.project.itemgroup.packagereference | foreach-object { CopyPackage $_.include }

	## Copy project references to output folder
	$projectFileData.project.itemgroup.projectreference | where-Object { $_ } | foreach-object { ProcessProject $_.include }

	## Copy project dll to output folder
	$projectName = [io.path]::GetFileNameWithoutExtension($projectFile)
	CopyPackage $projectName
}

ProcessProject "${ProjectDir}\${ProjectName}.csproj"