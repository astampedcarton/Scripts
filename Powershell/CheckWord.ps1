# Define paths
$inputFolder = "C:\Temp\Temp\temp\"  # Folder containing .doc, .docx, .rtf files
$outputFile = "C:\Temp\Temp\output.txt"    # Output file for strikethrough text

# Create Word COM object
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Array to store strikethrough text
$strikethroughText = @()

# Get all supported files
$files = Get-ChildItem -Path $inputFolder -Include *.doc, *.docx, *.rtf -Recurse
#$files = "C:\Temp\Temp\ae.rtf"


foreach ($file in $files) {
    Write-Host "Processing: $($file.FullName)"
    try {
        # Open the document
        $doc = $word.Documents.Open($file.FullName)

        # Iterate through words in the document
        foreach ($wordObj in $doc.Words) {
            # Check if the word has strikethrough formatting
            $text = $wordObj.Text.Trim()
            if (($wordObj.Font.StrikeThrough -eq $true) -or ($wordObj.HighlightColorIndex -ne 0)  ) {
                # Get the text of the word and trim any trailing spaces/punctuation
                # Get the page number for this word
                $pageNum = $wordObj.Information(3)  # 3 = wdActiveEndPageNumber

                $text = $wordObj.Text.Trim()
                if ($text) {
                    $strikethroughText += "Page $pageNum : $text"
                }
            }
        }

        # Close the document
        $doc.Close()
    }
    catch {
        Write-Host "Error processing $($file.FullName): $_"
    }
}

# Clean up Word COM object
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable word

# Save strikethrough text to file (if any found)
if ($strikethroughText.Count -gt 0) {
    $strikethroughText | Out-File $outputFile -Encoding UTF8
    Write-Host "Strikethrough text saved to $outputFile"
} else {
    "No strikethrough text found." | Out-File $outputFile -Encoding UTF8
    Write-Host "No strikethrough text found."
}