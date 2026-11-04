param(
    [Parameter(Mandatory = $true)]
    [string]$SourceImage,

    [Parameter(Mandatory = $true)]
    [string]$OutputImage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

function New-Color {
    param(
        [int]$A,
        [int]$R,
        [int]$G,
        [int]$B
    )

    return [System.Drawing.Color]::FromArgb($A, $R, $G, $B)
}

function New-HeartPath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$TopFlatten = 0.0
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $points = New-Object 'System.Drawing.PointF[]' 220

    for ($i = 0; $i -lt 220; $i++) {
        $t = 2.0 * [Math]::PI * $i / 219.0
        $sx = 16 * [Math]::Pow([Math]::Sin($t), 3)
        $sy = 13 * [Math]::Cos($t) - 5 * [Math]::Cos(2 * $t) - 2 * [Math]::Cos(3 * $t) - [Math]::Cos(4 * $t)

        $px = $X + (($sx + 16.0) / 32.0) * $Width
        $py = $Y + ((17.0 - $sy) / 34.0) * $Height
        if ($TopFlatten -gt 0.0) {
            $nx = ($px - $X) / $Width
            $ny = ($py - $Y) / $Height

            if ($ny -lt 0.42) {
                $edgeBias = [Math]::Pow([Math]::Abs($nx - 0.5) * 2.0, 0.85)
                $liftStrength = (0.42 - $ny) / 0.42
                $lift = $TopFlatten * $Height * 0.11 * $liftStrength * (0.35 + 0.65 * $edgeBias)
                $py -= $lift
            }
        }

        $points[$i] = [System.Drawing.PointF]::new([float]$px, [float]$py)
    }

    $path.AddPolygon($points)
    return $path
}

function Draw-HeartTexture {
    param(
        [System.Drawing.Graphics]$Graphics,
        [int]$CanvasWidth,
        [int]$CanvasHeight
    )

    $random = [System.Random]::new(42)

    for ($i = 0; $i -lt 85; $i++) {
        $size = $random.Next(24, 88)
        $x = $random.Next(-20, $CanvasWidth - $size + 20)
        $y = $random.Next(-20, $CanvasHeight - $size + 20)
        $alpha = $random.Next(35, 95)
        $outlineAlpha = [Math]::Min(160, $alpha + 30)

        $fillColor = if ($i % 3 -eq 0) {
            New-Color -A $alpha -R 255 -G 255 -B 255
        } elseif ($i % 3 -eq 1) {
            New-Color -A $alpha -R 255 -G 198 -B 224
        } else {
            New-Color -A $alpha -R 255 -G 164 -B 205
        }

        $strokeColor = New-Color -A $outlineAlpha -R 255 -G 255 -B 255
        $path = New-HeartPath -X $x -Y $y -Width $size -Height $size

        $brush = New-Object System.Drawing.SolidBrush($fillColor)
        $pen = New-Object System.Drawing.Pen($strokeColor, [Math]::Max(1, $size / 24.0))
        $Graphics.FillPath($brush, $path)
        $Graphics.DrawPath($pen, $path)

        $brush.Dispose()
        $pen.Dispose()
        $path.Dispose()
    }
}

function Draw-SoftGlow {
    param(
        [System.Drawing.Graphics]$Graphics,
        [float]$CenterX,
        [float]$CenterY,
        [float]$RadiusX,
        [float]$RadiusY,
        [System.Drawing.Color]$Color
    )

    $steps = 9
    for ($i = $steps; $i -ge 1; $i--) {
        $scale = 1.0 + ($i / 12.0)
        $alpha = [Math]::Max(8, [int]($Color.A / ($i + 1)))
        $brushColor = [System.Drawing.Color]::FromArgb($alpha, $Color.R, $Color.G, $Color.B)
        $brush = New-Object System.Drawing.SolidBrush($brushColor)
        $rect = [System.Drawing.RectangleF]::new(
            [float]($CenterX - ($RadiusX * $scale)),
            [float]($CenterY - ($RadiusY * $scale)),
            [float](2 * $RadiusX * $scale),
            [float](2 * $RadiusY * $scale)
        )
        $Graphics.FillEllipse($brush, $rect)
        $brush.Dispose()
    }
}

function Get-NormalizedBitmap {
    param(
        [string]$Path
    )

    $image = [System.Drawing.Image]::FromFile($Path)
    try {
        $orientationId = 0x0112
        $rotateType = [System.Drawing.RotateFlipType]::RotateNoneFlipNone

        if ($image.PropertyIdList -contains $orientationId) {
            $orientation = $image.GetPropertyItem($orientationId).Value[0]
            switch ($orientation) {
                2 { $rotateType = [System.Drawing.RotateFlipType]::RotateNoneFlipX }
                3 { $rotateType = [System.Drawing.RotateFlipType]::Rotate180FlipNone }
                4 { $rotateType = [System.Drawing.RotateFlipType]::Rotate180FlipX }
                5 { $rotateType = [System.Drawing.RotateFlipType]::Rotate90FlipX }
                6 { $rotateType = [System.Drawing.RotateFlipType]::Rotate90FlipNone }
                7 { $rotateType = [System.Drawing.RotateFlipType]::Rotate270FlipX }
                8 { $rotateType = [System.Drawing.RotateFlipType]::Rotate270FlipNone }
            }
        }

        if ($rotateType -ne [System.Drawing.RotateFlipType]::RotateNoneFlipNone) {
            $image.RotateFlip($rotateType)
        }

        return [System.Drawing.Bitmap]::new($image)
    }
    finally {
        $image.Dispose()
    }
}

$canvasWidth = 1600
$canvasHeight = 900

$bitmap = New-Object System.Drawing.Bitmap($canvasWidth, $canvasHeight)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

try {
    $backgroundRect = [System.Drawing.Rectangle]::new(0, 0, $canvasWidth, $canvasHeight)
    $topColor = New-Color -A 255 -R 255 -G 241 -B 247
    $bottomColor = New-Color -A 255 -R 242 -G 78 -B 156
    $backgroundBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $backgroundRect,
        $topColor,
        $bottomColor,
        90.0
    )
    $graphics.FillRectangle($backgroundBrush, $backgroundRect)
    $backgroundBrush.Dispose()

    Draw-SoftGlow -Graphics $graphics -CenterX 260 -CenterY 170 -RadiusX 260 -RadiusY 170 -Color (New-Color -A 72 -R 255 -G 255 -B 255)
    Draw-SoftGlow -Graphics $graphics -CenterX 1300 -CenterY 700 -RadiusX 300 -RadiusY 210 -Color (New-Color -A 46 -R 255 -G 228 -B 242)
    Draw-HeartTexture -Graphics $graphics -CanvasWidth $canvasWidth -CanvasHeight $canvasHeight

    $source = Get-NormalizedBitmap -Path $SourceImage
    try {
        $heartWidth = 1040.0
        $heartHeight = 980.0
        $heartX = ($canvasWidth - $heartWidth) / 2.0
        $heartY = -55.0

        $shadowPath = New-HeartPath -X ($heartX + 12) -Y ($heartY + 16) -Width $heartWidth -Height $heartHeight -TopFlatten 0.48
        $shadowBrush = New-Object System.Drawing.SolidBrush((New-Color -A 55 -R 118 -G 8 -B 52))
        $graphics.FillPath($shadowBrush, $shadowPath)
        $shadowBrush.Dispose()
        $shadowPath.Dispose()

        $framePath = New-HeartPath -X $heartX -Y $heartY -Width $heartWidth -Height $heartHeight -TopFlatten 0.48
        $frameBrush = New-Object System.Drawing.SolidBrush((New-Color -A 215 -R 255 -G 255 -B 255))
        $graphics.FillPath($frameBrush, $framePath)
        $frameBrush.Dispose()
        $framePath.Dispose()

        $photoInset = 16.0
        $photoX = $heartX + $photoInset
        $photoY = $heartY + $photoInset
        $photoWidth = $heartWidth - (2 * $photoInset)
        $photoHeight = $heartHeight - (2 * $photoInset)
        $photoPath = New-HeartPath -X $photoX -Y $photoY -Width $photoWidth -Height $photoHeight -TopFlatten 0.48

        $clipState = $graphics.Save()
        $graphics.SetClip($photoPath)

        $sourceWidth = [float]$source.Width
        $sourceHeight = [float]$source.Height
        $cropWidth = 3150.0
        $cropHeight = 3000.0
        $cropX = 180.0
        $cropY = 860.0

        if ($cropX + $cropWidth -gt $sourceWidth) { $cropX = $sourceWidth - $cropWidth }
        if ($cropY + $cropHeight -gt $sourceHeight) { $cropY = $sourceHeight - $cropHeight }
        if ($cropX -lt 0) { $cropX = 0 }
        if ($cropY -lt 0) { $cropY = 0 }

        $destRect = [System.Drawing.RectangleF]::new($photoX, $photoY, $photoWidth, $photoHeight)
        $srcRect = [System.Drawing.RectangleF]::new($cropX, $cropY, $cropWidth, $cropHeight)
        $graphics.DrawImage($source, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
        $graphics.Restore($clipState)

        $photoPath.Dispose()

        $outlinePath = New-HeartPath -X $photoX -Y $photoY -Width $photoWidth -Height $photoHeight -TopFlatten 0.48
        $outlinePen = New-Object System.Drawing.Pen((New-Color -A 255 -R 255 -G 255 -B 255), 7.0)
        $graphics.DrawPath($outlinePen, $outlinePath)
        $outlinePen.Dispose()
        $outlinePath.Dispose()

        $accentPath = New-HeartPath -X 85.0 -Y 720.0 -Width 118.0 -Height 118.0
        $accentBrush = New-Object System.Drawing.SolidBrush((New-Color -A 155 -R 255 -G 255 -B 255))
        $graphics.FillPath($accentBrush, $accentPath)
        $accentBrush.Dispose()
        $accentPath.Dispose()

        $accentPath2 = New-HeartPath -X 1410.0 -Y 735.0 -Width 92.0 -Height 92.0
        $accentPen2 = New-Object System.Drawing.Pen((New-Color -A 180 -R 255 -G 255 -B 255), 4.0)
        $graphics.DrawPath($accentPen2, $accentPath2)
        $accentPen2.Dispose()
        $accentPath2.Dispose()
    }
    finally {
        $source.Dispose()
    }

    $borderRect = [System.Drawing.RectangleF]::new(28.0, 28.0, 1544.0, 844.0)
    $borderPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $radius = 42.0
    $diameter = $radius * 2.0
    $borderPath.AddArc($borderRect.X, $borderRect.Y, $diameter, $diameter, 180, 90)
    $borderPath.AddArc($borderRect.Right - $diameter, $borderRect.Y, $diameter, $diameter, 270, 90)
    $borderPath.AddArc($borderRect.Right - $diameter, $borderRect.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $borderPath.AddArc($borderRect.X, $borderRect.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $borderPath.CloseFigure()

    $borderPen = New-Object System.Drawing.Pen((New-Color -A 180 -R 255 -G 247 -B 252), 3.0)
    $graphics.DrawPath($borderPen, $borderPath)
    $borderPen.Dispose()
    $borderPath.Dispose()

    $bitmap.Save($OutputImage, [System.Drawing.Imaging.ImageFormat]::Png)
}
finally {
    $graphics.Dispose()
    $bitmap.Dispose()
}
