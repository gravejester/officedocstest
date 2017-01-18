if (-not (Test-Path '.\Documents')) {
    $documentsFolder = New-Item -Path '.\Documents' -ItemType Directory -Force
}
else {
    $documentsFolder = Get-Item -Path '.\Documents'
}

Get-ChildItem -Path '.\DocumentsUnpacked' -Directory | ForEach-Object {
    
    $pathName = $_.BaseName
    ".\Documents\$($_.BaseName).zip"

    Compress-Archive -Path $_ -DestinationPath ".\Documents\$($_.BaseName).zip" -Force
    Rename-Item -Path "$($documentsFolder.FullName)\$($_.BaseName).zip" -NewName "$($documentsFolder.FullName)\$($_.BaseName).docx" -Force

}