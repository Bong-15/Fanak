# Draws the cartoon fennec fox and writes src\Fanak.ico and src\Fanak.png.
# Only needed when the artwork changes; build.ps1 uses the files it produces.
param([string]$PreviewPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function Color($hex) { [Drawing.ColorTranslator]::FromHtml($hex) }
function Pts($list) { $list | ForEach-Object { New-Object Drawing.PointF ([single]$_[0]), ([single]$_[1]) } }

# All coordinates are on a 1000 x 1000 grid and scaled to the requested size.
function Draw-Fox([int]$size) {
    $bmp = New-Object Drawing.Bitmap $size, $size, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.ScaleTransform($size / 1000.0, $size / 1000.0)

    # Small icons drop the fine detail and get bolder outlines and features.
    $tiny = $size -le 32
    $outlineWidth = if ($tiny) { 46 } else { 22 }
    $outline = New-Object Drawing.Pen (Color '#7A4A22'), $outlineWidth
    $outline.LineJoin = 'Round'
    $fur = New-Object Drawing.SolidBrush (Color '#F2C184')
    $furDark = New-Object Drawing.SolidBrush (Color '#E5A763')
    $cream = New-Object Drawing.SolidBrush (Color '#FFF4DF')
    $earInner = New-Object Drawing.SolidBrush (Color '#F9B7A6')
    $ink = New-Object Drawing.SolidBrush (Color '#2B1D14')
    $white = New-Object Drawing.SolidBrush ([Drawing.Color]::White)
    $blush = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(150, 0xF5, 0x8F, 0x86))

    # Ears: the big ones are what make it a fennec.
    foreach ($side in -1, 1) {
        $ear = New-Object Drawing.Drawing2D.GraphicsPath
        $ear.AddBezier((500 + $side * 90), 430, (500 + $side * 190), 240, (500 + $side * 330), 60, (500 + $side * 440), 30)
        $ear.AddBezier((500 + $side * 440), 30, (500 + $side * 490), 200, (500 + $side * 470), 400, (500 + $side * 370), 560)
        $ear.CloseFigure()
        $g.FillPath($fur, $ear)
        $g.DrawPath($outline, $ear)

        $inner = New-Object Drawing.Drawing2D.GraphicsPath
        $inner.AddBezier((500 + $side * 170), 410, (500 + $side * 240), 280, (500 + $side * 330), 170, (500 + $side * 405), 110)
        $inner.AddBezier((500 + $side * 405), 110, (500 + $side * 430), 240, (500 + $side * 415), 370, (500 + $side * 350), 480)
        $inner.CloseFigure()
        $g.FillPath($earInner, $inner)
        $ear.Dispose(); $inner.Dispose()
    }

    # Head: wide cheeks narrowing to a small chin.
    $head = New-Object Drawing.Drawing2D.GraphicsPath
    $head.AddBezier(130, 620, 130, 400, 300, 320, 500, 320)
    $head.AddBezier(500, 320, 700, 320, 870, 400, 870, 620)
    $head.AddBezier(870, 620, 870, 790, 640, 965, 500, 965)
    $head.AddBezier(500, 965, 360, 965, 130, 790, 130, 620)
    $head.CloseFigure()
    $g.FillPath($fur, $head)

    # Forehead stripe and cream face mask.
    if (-not $tiny) {
        $stripe = New-Object Drawing.Drawing2D.GraphicsPath
        $stripe.AddBezier(440, 325, 455, 400, 475, 450, 500, 500)
        $stripe.AddBezier(500, 500, 525, 450, 545, 400, 560, 325)
        $stripe.CloseFigure()
        $g.FillPath($furDark, $stripe)
        $stripe.Dispose()
    }
    $mask = New-Object Drawing.Drawing2D.GraphicsPath
    $mask.AddBezier(175, 680, 290, 590, 420, 590, 500, 650)
    $mask.AddBezier(500, 650, 580, 590, 710, 590, 825, 680)
    $mask.AddBezier(825, 680, 790, 800, 620, 955, 500, 955)
    $mask.AddBezier(500, 955, 380, 955, 210, 800, 175, 680)
    $mask.CloseFigure()
    $g.FillPath($cream, $mask)
    $mask.Dispose()
    $g.DrawPath($outline, $head)
    $head.Dispose()

    # Eyes: large and round, with the catch-lights that make it look alive.
    $eyeRadius = if ($tiny) { 88 } else { 78 }
    foreach ($cx in 335, 665) {
        $g.FillEllipse($ink, $cx - $eyeRadius, 600 - $eyeRadius, $eyeRadius * 2, $eyeRadius * 2)
        if ($size -ge 32) {
            $g.FillEllipse($white, $cx - 40, 545, 48, 48)
            if (-not $tiny) { $g.FillEllipse($white, $cx + 18, 625, 24, 24) }
        }
    }

    if (-not $tiny) {
        $g.FillEllipse($blush, 195, 690, 110, 66)
        $g.FillEllipse($blush, 695, 690, 110, 66)
    }

    # Nose and mouth.
    $noseScale = if ($tiny) { 1.5 } else { 1.0 }
    $nose = New-Object Drawing.Drawing2D.GraphicsPath
    $nw = 48 * $noseScale; $ny = 735; $nh = 44 * $noseScale
    $nose.AddBezier((500 - $nw), $ny, (500 - $nw / 2), ($ny - 14), (500 + $nw / 2), ($ny - 14), (500 + $nw), $ny)
    $nose.AddBezier((500 + $nw), $ny, (500 + $nw * 0.8), ($ny + $nh * 0.8), (500 + $nw * 0.2), ($ny + $nh), 500, ($ny + $nh))
    $nose.AddBezier(500, ($ny + $nh), (500 - $nw * 0.2), ($ny + $nh), (500 - $nw * 0.8), ($ny + $nh * 0.8), (500 - $nw), $ny)
    $nose.CloseFigure()
    $g.FillPath($ink, $nose)
    $nose.Dispose()

    if (-not $tiny) {
        $mouth = New-Object Drawing.Pen (Color '#2B1D14'), 15
        $mouth.StartCap = 'Round'; $mouth.EndCap = 'Round'
        $g.DrawLine($mouth, 500, 775, 500, 815)
        $g.DrawBezier($mouth, 500, 815, 485, 860, 425, 860, 405, 815)
        $g.DrawBezier($mouth, 500, 815, 515, 860, 575, 860, 595, 815)
        $mouth.Dispose()
    }

    foreach ($d in $g, $outline, $fur, $furDark, $cream, $earInner, $ink, $white, $blush) { $d.Dispose() }
    $bmp
}

function Dib-Bytes($bmp) {
    $s = $bmp.Width
    $ms = New-Object IO.MemoryStream
    $w = New-Object IO.BinaryWriter $ms
    $w.Write([int]40); $w.Write([int]$s); $w.Write([int]($s * 2)); $w.Write([int16]1); $w.Write([int16]32)
    foreach ($i in 1..6) { $w.Write([int]0) }
    for ($y = $s - 1; $y -ge 0; $y--) {
        for ($x = 0; $x -lt $s; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $w.Write([byte]$c.B); $w.Write([byte]$c.G); $w.Write([byte]$c.R); $w.Write([byte]$c.A)
        }
    }
    $w.Write((New-Object byte[] ([int]([math]::Floor(($s + 31) / 32)) * 4 * $s)))
    $w.Flush()
    $ms.ToArray()
}

$src = Join-Path (Split-Path $PSScriptRoot) 'src'
$sizes = 16, 24, 32, 48, 64, 256
$frames = foreach ($s in $sizes) {
    $bmp = Draw-Fox $s
    if ($s -eq 256) {
        $ms = New-Object IO.MemoryStream
        $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png)
        $data = $ms.ToArray()
    } else {
        $data = Dib-Bytes $bmp
    }
    $bmp.Dispose()
    , @($s, $data)
}

$out = New-Object IO.MemoryStream
$w = New-Object IO.BinaryWriter $out
$w.Write([int16]0); $w.Write([int16]1); $w.Write([int16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
foreach ($f in $frames) {
    $e = if ($f[0] -eq 256) { 0 } else { $f[0] }
    $w.Write([byte]$e); $w.Write([byte]$e); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([int16]1); $w.Write([int16]32); $w.Write([int]$f[1].Length); $w.Write([int]$offset)
    $offset += $f[1].Length
}
foreach ($f in $frames) { $w.Write([byte[]]$f[1]) }
$w.Flush()
[IO.File]::WriteAllBytes((Join-Path $src 'Fanak.ico'), $out.ToArray())

$big = Draw-Fox 256
$big.Save((Join-Path $src 'Fanak.png'), [Drawing.Imaging.ImageFormat]::Png)
$big.Dispose()

if ($PreviewPath) {
    $sheet = New-Object Drawing.Bitmap 760, 400
    $g = [Drawing.Graphics]::FromImage($sheet)
    $g.Clear((Color '#E9E9EE'))
    $g.InterpolationMode = 'NearestNeighbor'
    $g.PixelOffsetMode = 'Half'
    $hero = Draw-Fox 360
    $g.DrawImageUnscaled($hero, 20, 20)
    $hero.Dispose()
    $x = 400
    foreach ($s in 16, 24, 32, 48) {
        $b = Draw-Fox $s
        $g.DrawImage($b, $x, 30, 80, 80)
        $g.DrawImageUnscaled($b, $x, 130)
        $b.Dispose()
        $x += 90
    }
    $g.Dispose()
    $sheet.Save($PreviewPath, [Drawing.Imaging.ImageFormat]::Png)
    $sheet.Dispose()
}
Write-Output "Wrote icon: $($out.Length) bytes"
