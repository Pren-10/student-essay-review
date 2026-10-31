param([string]$InputPdf,[string]$OutputPrefix)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne 5){throw 'Run this script in Windows PowerShell 5.1'}
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime] | Out-Null
$script:asTask=([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1)
function AwaitOp($operation,$type){$task=$script:asTask.MakeGenericMethod($type).Invoke($null,@($operation)); $task.Wait(); $task.Result}
function AwaitAction($operation){$method=[System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and !$_.IsGenericMethod -and $_.GetParameters().Count -eq 1 }; $task=$method.Invoke($null,@($operation));$task.Wait()}
$file=AwaitOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($InputPdf.Replace('/','\'))) ([Windows.Storage.StorageFile])
$pdf=AwaitOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
for($n=1;$n -le $pdf.PageCount;$n++){if(Test-Path -LiteralPath "$OutputPrefix-$n.png"){throw 'Page image already exists'}}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($OutputPrefix)))
Write-Output "Pages: $($pdf.PageCount)"
for($i=0;$i -lt $pdf.PageCount;$i++){
$page=$pdf.GetPage($i)
$stream=[Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
$options=[Windows.Data.Pdf.PdfPageRenderOptions]::new();$options.DestinationWidth=1600
AwaitAction ($page.RenderToStreamAsync($stream,$options))
$stream.Seek(0)
$reader=[Windows.Storage.Streams.DataReader]::new($stream.GetInputStreamAt(0))
[void](AwaitOp ($reader.LoadAsync([uint32]$stream.Size)) ([uint32]))
$bytes=New-Object byte[] ([int]$stream.Size);$reader.ReadBytes($bytes)
$path="$OutputPrefix-$($i+1).png";[IO.File]::WriteAllBytes($path,$bytes);Write-Output $path
$reader.Dispose();$stream.Dispose();$page.Dispose()
}


