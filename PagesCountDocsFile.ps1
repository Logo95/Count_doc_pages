$shell = new-object -com shell.application
$pages = Get-ChildItem -Path ".\" -Recurse -Include *.doc, *.docx |
ForEach-Object {
$file = $_
Write-Host $file
$shellFolder = $shell.namespace($_.DirectoryName)
$shellFile = $shellFolder.Items() | Where-Object { $_.path -eq $file.FullName }
$pages = $shellFolder.GetDetailsOf($shellFile, 157) #id-code
Write-Host 'Страниц:' $pages
$pages
}
$pages | Measure-Object -Sum
Read-Host -Prompt "Press Enter to exit"
