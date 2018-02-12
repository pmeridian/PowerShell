# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
# Thanks to MFT, takabanana, ComFreek
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Get invocation path

[System.Reflection.Assembly]::LoadFrom('D:\Profiles\obertino\Downloads\itextsharp.dll') | Out-Null

$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
# Create a PowerPoint object
$ppt_app = New-Object -ComObject PowerPoint.Application
# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    $pdf_filename = "$($curr_path)\$($_.BaseName)_tmp.pdf"
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    $opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    $document.SaveAs($pdf_filename, $opt)
    # Close PowerPoint file
    $document.Close()
    $reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $pdf_filename
    $pdf_out_filename = "$($curr_path)\$($_.BaseName).pdf"
    $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader,[System.IO.File]::Create($pdf_out_filename))
    $stamper.addJavaScript("clickAdvance","app.fs.clickAdvances=false;")
    $stamper.Close()
    $reader.Close()
}
# Exit and release the PowerPoint object
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
