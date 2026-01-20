param(
    [Parameter(Mandatory=$true)]
    [Alias("i")]
    [string]$InputFile,
    
    [Parameter(Mandatory=$true)]
    [Alias("o")]
    [string]$OutputFile
)

# Intermediate file names
$MermaidOutputMd = "temp-mermaid-output.md"
$MermaidOutputPattern = "temp-mermaid-output-*.png"
$MermaidScale = 2
$MermaidFormat = "png"

# Resolve output path before file exists
$OutputFullPath = Join-Path $PWD $OutputFile

# Step 1: Convert Mermaid diagrams to images
Write-Host "Converting Mermaid diagrams..."
mmdc -i $InputFile -o $MermaidOutputMd -e $MermaidFormat -s $MermaidScale

# Step 2: Convert Markdown to Word
Write-Host "Converting Markdown to Word..."
pandoc $MermaidOutputMd -o $OutputFullPath

# Step 3: Add table borders
Write-Host "Adding table borders..."

$word = New-Object -ComObject Word.Application
$word. Visible = $false

try {
    $doc = $word. Documents.Open($OutputFullPath)
    
    foreach ($table in $doc.Tables) {
        $table.Borders.Enable = $true
    }
    
    $doc.Save()
    $doc.Close()
    Write-Host "Borders added successfully"
}
catch {
    Write-Error "Error processing file: $_"
}
finally {
    $word.Quit()
}

# Step 4: Cleanup intermediate files
Write-Host "Cleaning up..."
Remove-Item $MermaidOutputMd -ErrorAction SilentlyContinue
Remove-Item $MermaidOutputPattern -ErrorAction SilentlyContinue


Write-Host "Done!  Output: $OutputFullPath"