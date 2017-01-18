if (-not (Test-Path '.\Documents')) {
    $documentsFolder = New-Item -Path '.\Documents' -ItemType Directory -Force
}
else {
    $documentsFolder = Get-Item -Path '.\Documents'
}

Get-ChildItem -Path '.\DocumentsUnpacked' -Directory | ForEach-Object {
    
    Compress-Archive -Path "$($_.FullName)\*" -DestinationPath "$($documentsFolder.FullName)\$($_.BaseName).zip" -Force
    Rename-Item -Path "$($documentsFolder.FullName)\$($_.BaseName).zip" -NewName "$($documentsFolder.FullName)\$($_.BaseName).docx" -Force

}