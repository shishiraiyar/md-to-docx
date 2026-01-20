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

$MaxImageHeight = 7 * 72  # Convert to points
$MaxImageWidth = 6 * 72

$wdAlignParagraphCenter = 1

# Resolve output path before file exists
$OutputFullPath = Join-Path $PWD $OutputFile

# Step 1: Convert Mermaid diagrams to images
Write-Host "Converting Mermaid diagrams..."
mmdc -i $InputFile -o $MermaidOutputMd -e $MermaidFormat -s $MermaidScale

# Step 2: Convert Markdown to Word
Write-Host "Converting Markdown to Word..."
pandoc $MermaidOutputMd -o $OutputFullPath

$word = New-Object -ComObject Word.Application
$word. Visible = $false

try {
    $doc = $word. Documents.Open($OutputFullPath)

    # Step 3: Add table borders
    Write-Host "Adding table borders..."
    
    foreach ($table in $doc.Tables) {
        $table.Borders.Enable = $true
    }

    Write-Host "Borders added successfully"

    Write-Host "Adjusting image heights..."

    foreach ($shape in $doc.InlineShapes) {
        if ($shape.Height -gt $MaxImageHeight) {
            $ratio = $MaxImageHeight / $shape. Height
            $shape.Height = $MaxImageHeight
            $shape.Width = $shape.Width * $ratio
        }
        
        if ($shape.Width -gt $MaxImageWidth) {
            $ratio = $MaxImageWidth / $shape.Width
            $shape.Width = $MaxImageWidth
            $shape.Height = $shape.Height * $ratio
        }
    }

    Write-Host "Centering images..."
    foreach ($shape in $doc.InlineShapes) {
        $shape.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    }
    $doc.Save()
    $doc.Close()
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