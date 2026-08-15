# Ground truth for the PowerPoint export: opens the built decks with real
# PowerPoint through COM and fails if any of them will not open.
#
# Why this exists as its own gate. Every offline rule in scripts/exportQualityAudit.ts
# reported 164 fixtures at zero issues while the decks the user downloaded could
# not be opened at all -- PowerPoint answered "Sorry, PowerPoint can't read
# <file>.pptx". The packages were well formed XML, correctly ordered, with unique
# shape ids and valid relationships; they were simply illegal, because a
# <p:cxnSp> may not carry <a:custGeom>. Structural validity is not legality, and
# only PowerPoint itself knows the difference. So this asks PowerPoint.
#
# Usage:
#   npx tsx scripts/exportQualityAudit.ts          # writes tmp-export-audit/*.pptx
#   powershell -File scripts/verifyPptxOpens.ps1   # then prove they open
#
#   powershell -File scripts/verifyPptxOpens.ps1 -Path some.pptx,other.pptx
#
# Requires a local install of PowerPoint, so it is a developer and release gate
# rather than a CI step: the Linux runners have no Office. Exits with the number
# of files that failed to open.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$Path = @('tmp-export-audit/*.pptx')
)

$ErrorActionPreference = 'Stop'

$files = @()
foreach ($pattern in $Path) {
    $resolved = Resolve-Path -Path $pattern -ErrorAction SilentlyContinue
    if ($null -eq $resolved) {
        Write-Output "no match: $pattern"
        continue
    }
    $files += $resolved | ForEach-Object { $_.Path } | Where-Object { $_ -like '*.pptx' }
}

if ($files.Count -eq 0) {
    Write-Output 'Nothing to check. Run the export audit first so there are decks to open.'
    exit 1
}

try {
    $ppt = New-Object -ComObject PowerPoint.Application
}
catch {
    Write-Output "PowerPoint is not available on this machine, so the decks cannot be proven to open: $($_.Exception.Message)"
    exit 1
}

$ppt.DisplayAlerts = 1  # ppAlertsNone: a repair prompt must fail the run, not wait for a click.
$failed = 0

foreach ($full in $files) {
    $leaf = Split-Path $full -Leaf
    try {
        # ReadOnly, not in a window, and never "open as untitled" -- an untitled
        # copy would hide the very repair this is looking for.
        $pres = $ppt.Presentations.Open($full, $true, $false, $false)
        $shapes = 0
        foreach ($slide in $pres.Slides) { $shapes += $slide.Shapes.Count }
        Write-Output ("OPEN-OK   {0}  slides={1} shapes={2}" -f $leaf, $pres.Slides.Count, $shapes)
        $pres.Close()
    }
    catch {
        $failed++
        Write-Output ("OPEN-FAIL {0}  :: {1}" -f $leaf, $_.Exception.Message)
    }
}

$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()

Write-Output ("checked {0} deck(s), {1} failed to open" -f $files.Count, $failed)
exit $failed
