# unpack the document files into the DocumentsUnpacked folder


Get-ChildItem -Path '.\Documents' -File | ForEach-Object {
    $documentName = $_.BaseName
    if (-not (Test-Path -Path ".\DocumentsUnpacked\$($documentName)")) {
        $destination = New-Item -Path ".\DocumentsUnpacked\$($documentName)" -ItemType Directory -Force
    }
    else {
        $destination = Get-Item -Path ".\DocumentsUnpacked\$($documentName)"
        Remove-Item -Path "$($destination.FullName)\*" -Recurse -Force
    }
    $zipFile = Copy-Item -Path $_.FullName -Destination "$($destination.FullName)\$($documentName).zip" -Force -PassThru
    Expand-Archive -Path $zipFile.FullName -DestinationPath $destination -Force
    Remove-Item -Path $zipFile.FullName
}