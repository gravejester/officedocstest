Get-ChildItem -Path '.\DocumentsUnpacked' -Directory | ForEach-Object {
    if (Test-Path -Path "$($_.FullName)\docProps\core.xml") {
        $coreXml = Get-Content -Path "$($_.FullName)\docProps\core.xml" | ConvertTo-Xml
        $documentCreator = (([xml]$coreXml.Objects.InnerText).coreProperties).creator
        $documentLastModifiedBy = (([xml]$coreXml.Objects.InnerText).coreProperties).lastModifiedBy
        $documentRevisions = (([xml]$coreXml.Objects.InnerText).coreProperties).revision
        $documentDescription = (([xml]$coreXml.Objects.InnerText).coreProperties).description
    }
    if (Test-Path -Path "$($_.FullName)\word") {
        $documentType = 'Word'
    }
    elseif (Test-Path -Path "$($_.FullName)\xl") {
        $documentType = 'Excel'
    }

    Write-Output ([PSCustomObject] [Ordered] @{
        Name = $_.BaseName
        DocumentType = $documentType
        Creator = $documentCreator
        LastModifiedBy = $documentLastModifiedBy
        Revisions = $documentRevisions
        Description = $documentDescription
    })
}