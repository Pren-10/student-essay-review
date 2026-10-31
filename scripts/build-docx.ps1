param(
 [Parameter(Mandatory=$true)][string]$ContentPath,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [int]$MinChars=650,
 [int]$MaxChars=750,
 [ValidateSet('Both','WithPunctuation','WithoutPunctuation')][string]$CountMode='Both'
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$template=Join-Path (Split-Path $PSScriptRoot -Parent) 'assets/template.docx'
$data=Get-Content -LiteralPath $ContentPath -Raw -Encoding UTF8 | ConvertFrom-Json
foreach($key in @('title','review')) {
 if($data.$key -isnot [string] -or [string]::IsNullOrWhiteSpace($data.$key) -or $data.$key -match '[\r\n]'){throw "Invalid field: $key"}
}
if($data.student_name -isnot [string] -or $data.student_name -match '[\r\n]'){throw 'student_name must be a single-line string (may be empty)'}
if($data.paragraphs -isnot [Array] -or $data.paragraphs.Count -eq 0){throw 'paragraphs must be a non-empty array'}
foreach($p in $data.paragraphs){if($p -isnot [string] -or [string]::IsNullOrWhiteSpace($p) -or $p -match '[\r\n]'){throw 'Invalid paragraph'}}
if($MinChars -lt 1 -or $MaxChars -lt $MinChars){throw 'Invalid length limits'}
$text=$data.paragraphs -join ''
$with=([regex]::Replace($text,'\s','')).Length
$without=([regex]::Replace($text,'[\s\p{P}]','')).Length
if($CountMode -ne 'WithoutPunctuation' -and ($with -lt $MinChars -or $with -gt $MaxChars)){throw "Body with punctuation out of range: $with"}
if($CountMode -ne 'WithPunctuation' -and ($without -lt $MinChars -or $without -gt $MaxChars)){throw "Body without punctuation out of range: $without"}
$output=[IO.Path]::GetFullPath($OutputPath)
if([IO.Path]::GetExtension($output) -ne '.docx'){throw 'Output must end in .docx'}
if(Test-Path -LiteralPath $output){throw 'Output exists; use a new version name'}
function ReadEntry($zip,$name){$e=$zip.GetEntry($name);if(!$e){throw "Missing package part: $name"};$r=[IO.StreamReader]::new($e.Open());try{$r.ReadToEnd()}finally{$r.Dispose()}}
$source=[IO.Compression.ZipFile]::OpenRead($template)
try {
 $xml=[xml](ReadEntry $source 'word/document.xml')
 $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w',$W)
 $body=$xml.SelectSingleNode('/w:document/w:body',$ns)
 $ps=$body.SelectNodes('w:p',$ns)
 if($ps.Count -ne 6){throw 'Template structure has changed'}
 $section=$body.SelectSingleNode('w:sectPr',$ns).CloneNode($true)
 $nodes=@();foreach($p in $ps){$nodes+=,$p.CloneNode($true)}
 function SetText($p,$value){$texts=$p.SelectNodes('.//w:t',$ns);if($texts.Count -ne 1){throw 'Template paragraph must have one text node'};$texts[0].InnerText=$value}
 SetText $nodes[1] $data.review
 SetText $nodes[4] $data.title
 foreach($c in @($body.ChildNodes)){[void]$body.RemoveChild($c)}
 foreach($i in 0..4){[void]$body.AppendChild($nodes[$i])}
 foreach($p in $data.paragraphs){$node=$nodes[5].CloneNode($true);SetText $node $p;[void]$body.AppendChild($node)}
 [void]$body.AppendChild($section)
 [void][IO.Directory]::CreateDirectory((Split-Path $output -Parent))
 $target=[IO.Compression.ZipFile]::Open($output,[IO.Compression.ZipArchiveMode]::Create)
 try {
  foreach($entry in $source.Entries){
   $dest=$target.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::Optimal);$stream=$dest.Open()
   try {
    if($entry.FullName -eq 'word/document.xml'){$bytes=[Text.UTF8Encoding]::new($false).GetBytes($xml.OuterXml);$stream.Write($bytes,0,$bytes.Length)}
    elseif($entry.FullName -eq 'word/header1.xml'){
     $h=[xml](ReadEntry $source $entry.FullName)
     $hn=[Xml.XmlNamespaceManager]::new($h.NameTable);$hn.AddNamespace('w',$W)
     $h.SelectSingleNode('//w:t',$hn).InnerText=$data.student_name
     $bytes=[Text.UTF8Encoding]::new($false).GetBytes($h.OuterXml);$stream.Write($bytes,0,$bytes.Length)
    } else {$s=$entry.Open();try{$s.CopyTo($stream)}finally{$s.Dispose()}}
   } finally {$stream.Dispose()}
  }
 } finally {$target.Dispose()}
} finally {$source.Dispose()}
$check=[IO.Compression.ZipFile]::OpenRead($output)
try {
 foreach($entry in $check.Entries){if($entry.FullName -match '\.xml$'){$s=ReadEntry $check $entry.FullName;[void][xml]$s;if($s -match '\{\{(STUDENT_NAME|TITLE|REVIEW|BODY)\}\}'){throw 'Template placeholder remains'}}}
} finally {$check.Dispose()}
[pscustomobject]@{Output=$output;BodyWithPunctuation=$with;BodyWithoutPunctuation=$without;ReviewWithPunctuation=([regex]::Replace($data.review,'\s','')).Length;CountMode=$CountMode}
