$src = 'C:\src'

Get-ChildItem -Path $src -Directory | ForEach-Object {
  $normalized = $_.Name -replace '-', '_'
  $parts      = $normalized -split '_'

  if ($parts.Count -lt 2) {
    Write-Warning "Skipping $_ - name format unexpected"
    return
  }

  $imageFile  = Join-Path $src ($normalized + '.wim')
  $name       = $parts[1]

  if (Test-Path $imageFile) {
    Write-Host "Image file $imageFile already exists. Skipping."
    return
  } 

  $args = @(
    '/capture-image'
    "/ImageFile:$imageFile"
    "/CaptureDir:$($_.FullName)"
    "/Name:$name"
    '/CheckIntegrity'
    '/Verify'
    '/Compress:Max'
  )

  Write-Host "Capturing $($_.FullName) to $imageFile with name $name"
  dism @args

  if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to capture image for $($_.FullName). Exit code: $LASTEXITCODE"
    break
  }
}
