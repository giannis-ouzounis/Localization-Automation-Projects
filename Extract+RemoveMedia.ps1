[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

for ($LoopCounter = 1; $LoopCounter -le 3; $LoopCounter++) {
  if ( $LoopCounter -eq 1) {
    $fileExtension = "*.docx" 
    $fileNames = Get-ChildItem -Path $PSScriptRoot -Recurse -Include $fileExtension
    $folder = 'word\media'
    $dst = $PSScriptRoot

    foreach ($f in $fileNames) {
      $NewFolderName = $f.Basename
      $dst2 = $dst + "\" + $NewFolderName
      New-Item -ItemType Directory -Path $dst2
      [IO.Compression.ZipFile]::OpenRead($f).Entries | ? {
        $_.FullName -like "$($folder -replace '\\','/')/*.*"
      } | % {
        $file = Join-Path $dst2 $_.FullName
        $parent = Split-Path -Parent $file
        if (-not (Test-Path -LiteralPath $parent)) {
          New-Item -Path $parent -Type Directory | Out-Null
        }
        [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $file, $true)
      }
      $PathToCheck = $dst2 + "\*"
      if ((test-path ($PathToCheck)) -ne $true) {
        Remove-Item -Path $dst2 -Recurse -Force
      }
    }
  }
  Elseif ( $LoopCounter -eq 2) {
    $fileExtension = "*.pptx" 
    $fileNames = Get-ChildItem -Path $PSScriptRoot -Recurse -Include $fileExtension
    $folder = 'ppt\media'
    $dst = $PSScriptRoot

    foreach ($f in $fileNames) {
      $NewFolderName = $f.Basename
      $dst2 = $dst + "\" + $NewFolderName
      New-Item -ItemType Directory -Path $dst2

      # Use explicit zip object to allow cleanup later
      $zip = [IO.Compression.ZipFile]::OpenRead($f.FullName)
      $entries = $zip.Entries | Where-Object {
        $_.FullName -like "$($folder -replace '\\','/')/*.*"
      }
      foreach ($entry in $entries) {
        $file = Join-Path $dst2 $entry.FullName
        $parent = Split-Path -Parent $file
        if (-not (Test-Path -LiteralPath $parent)) {
          New-Item -Path $parent -Type Directory | Out-Null
        }
        [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $file, $true)
      }
      $zip.Dispose()

      $PathToCheck = $dst2 + "\*"
      if ((test-path ($PathToCheck)) -ne $true) {
        Remove-Item -Path $dst2 -Recurse -Force
      }
    }
  }
  Else {
    $fileExtension = "*.xlsx" 
    $fileNames = Get-ChildItem -Path $PSScriptRoot -Recurse -Include $fileExtension
    $folder = 'xl\media'
    $dst = $PSScriptRoot

    foreach ($f in $fileNames) {
      $NewFolderName = $f.Basename
      $dst2 = $dst + "\" + $NewFolderName
      New-Item -ItemType Directory -Path $dst2
      [IO.Compression.ZipFile]::OpenRead($f).Entries | ? {
        $_.FullName -like "$($folder -replace '\\','/')/*.*"
      } | % {
        $file = Join-Path $dst2 $_.FullName
        $parent = Split-Path -Parent $file
        if (-not (Test-Path -LiteralPath $parent)) {
          New-Item -Path $parent -Type Directory | Out-Null
        }
        [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $file, $true)
      }
      $PathToCheck = $dst2 + "\*"
      if ((test-path ($PathToCheck)) -ne $true) {
        Remove-Item -Path $dst2 -Recurse -Force
      }
    }
  }
}

# Delete ppt/media/* from each .pptx after extraction
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

$pptxFiles = Get-ChildItem -Path "$PSScriptRoot" -Recurse -Filter *.pptx

foreach ($z in $pptxFiles) {
  try {
    Write-Host "Updating $($z.Name)"
    $zip = [IO.Compression.ZipFile]::Open($z.FullName, 'Update')

    $entries = $zip.Entries | Where-Object {
      $_.FullName -match '^ppt/media/.+'
    }

    foreach ($entry in $entries) {
      Write-Host "Deleting: $($entry.FullName)"
      $entry.Delete()
    }

    $zip.Dispose()
  }
  catch {
    Write-Host "Could not update $($z.Name): $($_.Exception.Message)"
  }
}

Read-Host -Prompt "Press Enter to exit"
