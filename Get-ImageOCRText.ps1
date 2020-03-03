using namespace Windows.Storage
using namespace Windows.Graphics.Imaging


<#
.Synopsis
   Runs Windows 10 OCR on an image.
.DESCRIPTION
   Takes a path to an image file, with some text on it.
   Runs Windows 10 OCR against the image.
   Returns an [OcrResult], hopefully with a .Text property containing the text
.EXAMPLE
   $result = .\Get-Win10OcrTextFromImage.ps1 -Path 'c:\test.bmp'
   $result.Text
#>
[CmdletBinding()]
Param
(
    # System.drawing.bitmap object
    [Parameter(Mandatory=$false, #Originally $false 
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true, 
                Position=0,
                HelpMessage='Path to an image file, to run OCR on')]
    [ValidateNotNullOrEmpty()]
    $Content
)

Begin {
    # Add assemblies for Toast Notification
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications,         ContentType = WindowsRuntime]
    $null = [Windows.UI.Notifications.ToastNotification,        Windows.UI.Notifications,         ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument,                  Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    
    # Add the WinRT assembly, and load the appropriate WinRT types
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    
    
    $null = [Windows.Storage.StorageFile,                       Windows.Storage,                  ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine,                       Windows.Foundation,               ContentType = WindowsRuntime]
    $null = [Windows.Foundation.IAsyncOperation`1,              Windows.Foundation,               ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap,           Windows.Foundation,               ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.RandomAccessStream,        Windows.Storage.Streams,          ContentType = WindowsRuntime]
    

    # [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    

    # PowerShell doesn't have built-in support for Async operations, 
    # but all the WinRT methods are Async.
    # This function wraps a way to call those methods, and wait for their results.
    $getAwaiterBaseMethod = [WindowsRuntimeSystemExtensions].GetMember('GetAwaiter').
                                Where({
                                        $PSItem.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
                                    }, 'First')[0]

    Function Await {
        param($AsyncTask, $ResultType)

        $getAwaiterBaseMethod.
            MakeGenericMethod($ResultType).
            Invoke($null, @($AsyncTask)).
            GetResult()
    }
    # Pad image with a 30 pixel border of White
    function Get-PaddedImage($p)
    {
        if($p.getType().Name -ne "Bitmap"){
            return $null
        }
        [Int32] $pad = 30
        $clipBoardImage = $p
        $width = $pad * 2 + $clipBoardImage.Width
        $height = $pad * 2 + $clipBoardImage.Height
        $bmp = New-Object System.Drawing.Bitmap($width,$height)
        [System.Drawing.Graphics] $g = [System.Drawing.Graphics]::FromImage($bmp)
        $null = $g.FillRectangle([System.Drawing.Brushes]::Yellow, 0, 0, $width, $height)
        $null = $g.DrawImage($clipBoardImage, $pad, $pad)
        $bmp
    }
    # Send Toast message "notification" of Captured text
    Function Toast-Message($str){


    New-BurntToastNotification -Text $str -silent
    <#
        $APP_ID = "ec03997f-df79-47c6-b6da-62faf59bc54b"

        # Toast format template
        $template = @"
        <toast>
            <visual>
                <binding template="ToastText02">
                    <text id="1">OCR Result:</text>
                    <text id="2">$($str)</text>
                </binding>
            </visual>
            <audio silent = "true" />
        </toast>
"@
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($APP_ID).Show($toast)
        Write-Output "toast, bro"
        #>
    }
}

Process
{
    # If nothing was passed into $Content, assign image in clipboard to $Content
    if($null -eq $Content){
        # Call Snipping tool: Win10 < 1809
        C:\Windows\System32\snippingtool.exe /clip | Wait-Process
        # Call Snip and Sketch: Win10 >= 1809
        # Start-Process ms-screenclip: | Wait-Process
        try{
            $clipboard = Get-Clipboard -Format Image
            if($clipboard.GetType().Name -eq "Bitmap"){
                $Content = $clipboard
            }else{
                $Content = $null
            }
        }catch{
            Write-Output "Clipboard is empty"
            Break
        }
    }

    # Scale to multiple input objects if need be in the future
    foreach ($p in $Content)
    {
        $p = Get-PaddedImage($p)
        if($p -eq $null){
            Write-Output "No image data found"
            break
        }

        # The following converts from System.Drawing.Bitmap to and OCREngine result
        $stream = New-Object System.IO.MemoryStream
        $p.save($stream,"BMP")
        [Streams.IRandomAccessStream] $randomaccessstream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($stream)
        $fileStream = $randomaccessstream

        $params = @{
            AsyncTask  = [BitmapDecoder]::CreateAsync($fileStream)
            ResultType = [BitmapDecoder]
        }
        $bitmapDecoder = Await @params


        $params = @{ 
            AsyncTask = $bitmapDecoder.GetSoftwareBitmapAsync()
            ResultType = [SoftwareBitmap]
        }
        $softwareBitmap = Await @params

        # Get OCR'd text from Engine
        $result = Await $ocrEngine.RecognizeAsync($softwareBitmap) ([Windows.Media.Ocr.OcrResult])
        
        if($result.text -NE ""){
            Write-Output "$($result.Text)"
            $result = $result.Text
        } else {
            Write-Output "No text found"
            $result = "_NTF"
        }
        Set-Clipboard -Value $result
        Toast-Message $result
    }
}