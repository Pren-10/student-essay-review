param(
 [Parameter(Mandatory=$true)][string[]]$InputPaths,
 [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference='Stop'
$target=[IO.Path]::GetFullPath($OutputDir)
[void][IO.Directory]::CreateDirectory($target)
$jobs=@();$seen=@{}
foreach($p in $InputPaths){
 $inputFile=(Get-Item -LiteralPath $p).FullName
 if([IO.Path]::GetExtension($inputFile) -ne '.docx'){throw 'Input must be docx'}
 $pdf=Join-Path $target ([IO.Path]::GetFileNameWithoutExtension($inputFile)+'.pdf')
 if((Test-Path -LiteralPath $pdf) -or $seen.ContainsKey($pdf)){throw "Preview path already used: $pdf"}
 $seen[$pdf]=$true;$jobs+=@{Input=$inputFile;Pdf=$pdf}
}
$word=New-Object -ComObject Word.Application
$word.Visible=$false;$word.DisplayAlerts=0
try {
 foreach($job in $jobs){
  $doc=$null
  try{$doc=$word.Documents.Open($job.Input,$false,$true);$doc.Repaginate();$pages=$doc.ComputeStatistics(2);$doc.ExportAsFixedFormat($job.Pdf,17);[pscustomobject]@{File=$job.Input;Pages=$pages;Preview=$job.Pdf}}
  finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)}}
 }
} finally {$word.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
