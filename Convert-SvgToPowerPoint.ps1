param(
    [Parameter(Mandatory)][string]$SvgDirectory,
    [Parameter(Mandatory)][string]$Output,
    [string]$NotesJson,
    [string]$PreviewDirectory
)

$ErrorActionPreference = 'Stop'
$culture = [Globalization.CultureInfo]::InvariantCulture
$scale = 0.75
$xlink = 'http://www.w3.org/1999/xlink'
$numberStyle = [Globalization.NumberStyles]::Float

function Number([string]$value) {
    $result = [double]::Parse($value, $numberStyle, $culture)
    if ([double]::IsInfinity($result) -or [double]::IsNaN($result)) { throw 'Expected a finite number.' }
    $result
}

function Color([string]$value) {
    if ($value -notmatch '^#[0-9a-fA-F]{6}$') { throw "Unsupported color: $value" }
    $r = [Convert]::ToInt32($value.Substring(1, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(3, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(5, 2), 16)
    $r + 256 * $g + 65536 * $b
}

$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
if ([IO.Path]::GetExtension($outputPath) -ne '.pptx' -or (Test-Path -LiteralPath $outputPath)) {
    throw 'Specify a new .pptx output file.'
}
if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($outputPath)) -PathType Container)) {
    throw 'Create the output parent directory first.'
}
$files = @(Get-ChildItem -LiteralPath $SvgDirectory -File -Filter '*.svg' | Sort-Object Name)
if (-not $files.Count) { throw 'No SVG files found.' }
$notes = @()
if ($NotesJson) {
    $notes = @(Get-Content -LiteralPath $NotesJson -Raw -Encoding UTF8 | ConvertFrom-Json)
    if ($notes.Count -ne $files.Count -or @($notes | Where-Object { $_ -isnot [string] }).Count) {
        throw 'NotesJson must contain one string per SVG file.'
    }
}
if ($PreviewDirectory) {
    $previewPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PreviewDirectory)
    if (-not (Test-Path -LiteralPath $previewPath -PathType Container)) { throw 'Create the preview directory first.' }
    for ($i = 1; $i -le $files.Count; $i++) {
        if (Test-Path -LiteralPath (Join-Path $previewPath ('{0:D2}.png' -f $i))) {
            throw 'Use a preview directory without existing numbered PNGs.'
        }
    }
}

$settings = [Xml.XmlReaderSettings]::new()
$settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
$settings.XmlResolver = $null
$expectedSize = $null
$slides = foreach ($file in $files) {
    $reader = [Xml.XmlReader]::Create($file.FullName, $settings)
    try {
        $xml = [Xml.XmlDocument]::new()
        $xml.XmlResolver = $null
        $xml.PreserveWhitespace = $true
        $xml.Load($reader)
    } finally { $reader.Close() }
    $root = $xml.DocumentElement
    $box = $root.GetAttribute('viewBox').Trim() -split '\s+'
    if ($root.LocalName -ne 'svg' -or $root.NamespaceURI -ne 'http://www.w3.org/2000/svg' -or
        $box.Count -ne 4 -or (Number $box[0]) -ne 0 -or (Number $box[1]) -ne 0) {
        throw "Expected an SVG with a viewBox starting at 0 0 in $($file.Name)."
    }
    $w = Number $box[2]
    $h = Number $box[3]
    if ($w -le 0 -or $h -le 0) { throw 'The viewBox dimensions must be positive.' }
    if ($expectedSize -and ($expectedSize.W -ne $w -or $expectedSize.H -ne $h)) {
        throw 'All slides must use the same viewBox.'
    }
    $expectedSize = @{ W = $w; H = $h }
    $elements = @($root.ChildNodes | Where-Object {
        $_ -is [Xml.XmlElement] -and $_.LocalName -notin @('title', 'desc')
    })
    if (-not $elements.Count) { throw "No content in $($file.Name)." }
    foreach ($element in $elements) {
        if ($element.LocalName -notin @('rect', 'text', 'image') -or $element.HasAttribute('transform')) {
            throw "Unsupported SVG element in $($file.Name): $($element.LocalName)"
        }
    }
    [pscustomobject]@{ Name = $file.BaseName; Elements = $elements }
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('bob-ppt-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$presentation = $null
$review = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add(0)
    $presentation.PageSetup.SlideWidth = $expectedSize.W * $scale
    $presentation.PageSetup.SlideHeight = $expectedSize.H * $scale
    $slideWidth = $presentation.PageSetup.SlideWidth
    for ($i = 0; $i -lt $slides.Count; $i++) {
        $slide = $presentation.Slides.Add($i + 1, 12)
        $slide.Name = $slides[$i].Name
        foreach ($element in $slides[$i].Elements) {
            $x = (Number ($element.GetAttribute('x'))) * $scale
            $y = (Number ($element.GetAttribute('y'))) * $scale
            if ($element.LocalName -ne 'text') {
                $width = (Number ($element.GetAttribute('width'))) * $scale
                $height = (Number ($element.GetAttribute('height'))) * $scale
            }
            switch ($element.LocalName) {
                'rect' {
                    $shape = $slide.Shapes.AddShape(1, $x, $y, $width, $height)
                    $shape.Line.Visible = 0
                    $shape.Fill.Solid()
                    $shape.Fill.ForeColor.RGB = Color ($element.GetAttribute('fill'))
                }
                'image' {
                    $href = $element.GetAttribute('href', $xlink)
                    if (-not $href) { $href = $element.GetAttribute('href') }
                    if ($href -notmatch '^data:image/(png|jpeg);base64,(.+)$') {
                        throw 'Images must be embedded PNG or JPEG.'
                    }
                    $format = $Matches[1]
                    $bytes = [Convert]::FromBase64String($Matches[2])
                    $name = Join-Path $temp ([guid]::NewGuid().ToString('N') + '.' + $format)
                    [IO.File]::WriteAllBytes($name, $bytes)
                    $shape = $slide.Shapes.AddPicture($name, 0, -1, $x, $y, $width, $height)
                    $shape.LockAspectRatio = -1
                    $title = $element.ChildNodes | Where-Object LocalName -eq 'title' | Select-Object -First 1
                    if ($title) { $shape.AlternativeText = $title.InnerText }
                }
                'text' {
                    if (@($element.ChildNodes | Where-Object { $_ -is [Xml.XmlElement] }).Count) {
                        throw 'Only flat SVG text is supported.'
                    }
                    $size = (Number ($element.GetAttribute('font-size'))) * $scale
                    $shape = $slide.Shapes.AddTextbox(1, $x, ($y - $size * 0.97), ($slideWidth - $x), ($size * 1.45))
                    $frame = $shape.TextFrame2
                    $frame.AutoSize = 0
                    $frame.WordWrap = 0
                    $frame.MarginLeft = $frame.MarginRight = $frame.MarginTop = $frame.MarginBottom = 0
                    $range = $frame.TextRange
                    $range.Text = $element.InnerText
                    $range.ParagraphFormat.Alignment = 1
                    $range.ParagraphFormat.SpaceBefore = $range.ParagraphFormat.SpaceAfter = 0
                    $font = $range.Font
                    $family = $element.GetAttribute('font-family')
                    if (-not $family) { $family = 'Arial' }
                    $font.Name = $font.NameFarEast = $family
                    $font.Size = $size
                    $font.Bold = if ($element.GetAttribute('font-weight') -match '^(bold|[6-9]00)$') { -1 } else { 0 }
                    $font.Italic = if ($element.GetAttribute('font-style') -match '^(italic|oblique)$') { -1 } else { 0 }
                    $font.Fill.ForeColor.RGB = Color ($element.GetAttribute('fill'))
                    if ($element.HasAttribute('letter-spacing')) {
                        $font.Spacing = (Number ($element.GetAttribute('letter-spacing'))) * $scale
                    }
                    $shape.Width = [Math]::Min($slideWidth - $x, [Math]::Max(9, $range.BoundWidth + 3))
                }
            }
        }
        if ($notes.Count) { $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes[$i] }
    }
    $presentation.SaveAs($outputPath, 24, 0)
    $presentation.Close()
    $presentation = $null
    if ($PreviewDirectory) {
        $review = $powerPoint.Presentations.Open($outputPath, -1, 0, 0)
        for ($i = 1; $i -le $review.Slides.Count; $i++) {
            $review.Slides.Item($i).Export(
                (Join-Path $previewPath ('{0:D2}.png' -f $i)), 'PNG',
                [int]$expectedSize.W, [int]$expectedSize.H
            )
        }
    }
    Write-Output $outputPath
} finally {
    if ($review) { $review.Close() }
    if ($presentation) { $presentation.Saved = -1; $presentation.Close() }
    Get-ChildItem -LiteralPath $temp -File | ForEach-Object { Remove-Item -LiteralPath $_.FullName }
    [IO.Directory]::Delete($temp)
}
