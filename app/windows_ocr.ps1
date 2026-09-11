param(
    [Parameter(Mandatory = $true)]
    [string] $ImagePath,
    [string] $LanguageTag = 'zh-Hans-CN'
)

$ErrorActionPreference = 'Stop'
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]

function Await-WinRtOperation {
    param(
        [Parameter(Mandatory = $true)] $Operation,
        [Parameter(Mandatory = $true)] [Type] $ResultType
    )

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.IsGenericMethodDefinition -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
        } |
        Select-Object -First 1

    if (-not $method) {
        throw 'Windows Runtime does not expose the OCR async interface.'
    }

    $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $task.Wait()
    return $task.Result
}

if (-not [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages.LanguageTag.Contains($LanguageTag)) {
    throw "Windows OCR language pack is unavailable: $LanguageTag. Install the OCR feature for Simplified Chinese in Windows Settings."
}

$file = Await-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) ([Windows.Storage.StorageFile])
$stream = Await-WinRtOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language($LanguageTag)))

if (-not $engine) {
    throw "Unable to initialize Windows OCR for $LanguageTag."
}

$recognized = Await-WinRtOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
$lines = @(
    foreach ($line in $recognized.Lines) {
        $words = @()
        $left = [double]::PositiveInfinity
        $top = [double]::PositiveInfinity
        $right = [double]::NegativeInfinity
        $bottom = [double]::NegativeInfinity
        foreach ($word in $line.Words) {
            $x = [Math]::Round($word.BoundingRect.X, 2)
            $y = [Math]::Round($word.BoundingRect.Y, 2)
            $width = [Math]::Round($word.BoundingRect.Width, 2)
            $height = [Math]::Round($word.BoundingRect.Height, 2)
            $words += [PSCustomObject]@{
                text = $word.Text
                x = $x
                y = $y
                width = $width
                height = $height
            }
            $left = [Math]::Min($left, $x)
            $top = [Math]::Min($top, $y)
            $right = [Math]::Max($right, $x + $width)
            $bottom = [Math]::Max($bottom, $y + $height)
        }
        if ($words.Count -eq 0) {
            $left = $line.BoundingRect.X
            $top = $line.BoundingRect.Y
            $right = $left + $line.BoundingRect.Width
            $bottom = $top + $line.BoundingRect.Height
        }
        [PSCustomObject]@{
            text = $line.Text
            x = [Math]::Round($left, 2)
            y = [Math]::Round($top, 2)
            width = [Math]::Round($right - $left, 2)
            height = [Math]::Round($bottom - $top, 2)
            words = $words
        }
    }
)

[PSCustomObject]@{ lines = $lines } | ConvertTo-Json -Depth 6 -Compress
