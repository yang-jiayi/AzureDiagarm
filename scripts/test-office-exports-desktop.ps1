# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

param([string]$ArtifactDirectory = $env:OFFICE_EXPORT_ARTIFACT_DIR)
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    throw 'Desktop Office checks require Windows and an installed copy of Microsoft PowerPoint.'
}
if (!$ArtifactDirectory) { throw 'Specify -ArtifactDirectory or OFFICE_EXPORT_ARTIFACT_DIR.' }
$directory = (Resolve-Path -LiteralPath $ArtifactDirectory).Path
$files = @(Get-ChildItem -LiteralPath $directory -Filter '*.pptx')
if (!$files.Count) { throw 'No PPTX files found. Run test:exports:browser with OFFICE_EXPORT_ARTIFACT_DIR first.' }
$existingIds = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$application = New-Object -ComObject PowerPoint.Application
$results = [System.Collections.Generic.List[object]]::new()
$overflows = [System.Collections.Generic.List[object]]::new()

function Inspect-TextShape($Shape, [string]$Context) {
    if ($Shape.Type -eq 6) {
        for ($child = 1; $child -le $Shape.GroupItems.Count; $child++) {
            Inspect-TextShape $Shape.GroupItems.Item($child) "$Context/$($Shape.Name)"
        }
    }
    if ($Shape.Type -eq 19) {
        for ($row = 1; $row -le $Shape.Table.Rows.Count; $row++) {
            for ($column = 1; $column -le $Shape.Table.Columns.Count; $column++) {
                Inspect-TextShape $Shape.Table.Cell($row, $column).Shape "$Context/table/$row/$column"
            }
        }
    }
    if ($Shape.HasTextFrame -ne -1 -or $Shape.TextFrame.HasText -ne -1) { return }
    $frame = $Shape.TextFrame2
    $range = $frame.TextRange
    $allowedWidth = $Shape.Width - $frame.MarginLeft - $frame.MarginRight
    $allowedHeight = $Shape.Height - $frame.MarginTop - $frame.MarginBottom
    $text = $range.Text
    $entry = [pscustomobject]@{
        context = $Context
        shape = $Shape.Name
        text = $text
        fontSize = $range.Font.Size
        width = [Math]::Round($allowedWidth, 2)
        height = [Math]::Round($allowedHeight, 2)
        drawnWidth = [Math]::Round($range.BoundWidth, 2)
        drawnHeight = [Math]::Round($range.BoundHeight, 2)
    }
    $results.Add($entry)
    if ($range.BoundWidth -gt $allowedWidth + 2 -or $range.BoundHeight -gt $allowedHeight + 2) {
        $overflows.Add($entry)
    }
}

try {
    foreach ($file in $files) {
        $presentation = $null
        $temporaryCopy = Join-Path $directory (".office-check-" + [Guid]::NewGuid().ToString('N') + '.pptx')
        try {
            Write-Output "Opening $($file.Name)"
            # Never close or modify an original that a user may already have open.
            [IO.File]::Copy($file.FullName, $temporaryCopy)
            $presentation = $application.Presentations.Open($temporaryCopy, $true, $false, $false)
            $renderDirectory = Join-Path $directory ($file.BaseName + '-rendered')
            [IO.Directory]::CreateDirectory($renderDirectory) | Out-Null
            for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
                $slide = $presentation.Slides.Item($index)
                $renderHeight = [Math]::Max(1, [Math]::Round(1600 * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth))
                $slide.Export((Join-Path $renderDirectory ("slide-{0:D2}.png" -f $index)), 'PNG', 1600, $renderHeight)
                for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                    $shape = $slide.Shapes.Item($shapeIndex)
                    Inspect-TextShape $shape "$($file.BaseName)/$index"
                    $angle = $shape.Rotation * [Math]::PI / 180
                    $halfWidth = ([Math]::Abs([Math]::Cos($angle)) * $shape.Width + [Math]::Abs([Math]::Sin($angle)) * $shape.Height) / 2
                    $halfHeight = ([Math]::Abs([Math]::Sin($angle)) * $shape.Width + [Math]::Abs([Math]::Cos($angle)) * $shape.Height) / 2
                    $centerX = $shape.Left + $shape.Width / 2
                    $centerY = $shape.Top + $shape.Height / 2
                    if ($centerX - $halfWidth -lt -2 -or $centerY - $halfHeight -lt -2 -or
                        $centerX + $halfWidth -gt $presentation.PageSetup.SlideWidth + 2 -or
                        $centerY + $halfHeight -gt $presentation.PageSetup.SlideHeight + 2) {
                        throw "Off-slide shape: $($file.Name), slide $index, $($shape.Name)"
                    }
                }
            }
            if ($file.Name -eq 'simple-light.pptx') {
                $slide = $presentation.Slides.Item(1)
                foreach ($id in @('api', 'meta-api', 'label-api')) {
                    if ($slide.Shapes.Item("node-$id").Type -ne 6) {
                        throw "The role-like ID $id lost its independent native service group."
                    }
                }
                $connector = $slide.Shapes.Item('connector-e1')
                if ($connector.ConnectorFormat.BeginConnected -ne -1 -or $connector.ConnectorFormat.EndConnected -ne -1) {
                    throw 'The native connector lost its shape attachments.'
                }
                $rightBefore = $connector.Left + $connector.Width
                $card = $slide.Shapes.Item('node-meta-api')
                $card.Left += 36
                $card.Top += 18
                Start-Sleep -Milliseconds 150
                if ([Math]::Abs(($connector.Left + $connector.Width) - $rightBefore - 36) -gt 2) {
                    throw 'The native connector did not follow its moved service group.'
                }
                Write-Output 'Native connector follows a moved service group.'
            }
            Write-Output "$($file.Name): $($presentation.Slides.Count) slides opened and rendered."
        } finally {
            if ($presentation) {
                $presentation.Saved = $true
                $presentation.Close()
                [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
            }
            if (Test-Path -LiteralPath $temporaryCopy) { Remove-Item -LiteralPath $temporaryCopy -Force }
        }
    }
    $results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'powerpoint-text-bounds.json') -Encoding utf8
    if ($overflows.Count -gt 0) {
        $overflows | Select-Object context, shape, width, height, drawnWidth, drawnHeight | Format-Table -AutoSize
        throw "$($overflows.Count) text blocks exceed their native PowerPoint bounds."
    }
    Write-Output "Inspected $($results.Count) native text blocks."
} finally {
    if ($existingIds.Count -eq 0 -and $application.Presentations.Count -eq 0) { $application.Quit() }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
}
