param([string]$Docx, [string]$Pdf)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($Docx, $false, $true)
    # 17 = wdFormatPDF
    $doc.SaveAs([ref]$Pdf, [ref]17)
    $doc.Close($false)
    Write-Output "PDF saved: $Pdf"
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
