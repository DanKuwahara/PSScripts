Get-ChildItem -path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -Recurse | ForEach-Object {
    $CurrentKey = (Get-ItemProperty -Path $_.PsPath)
        If ($CurrentKey -match 'EquatIO'){
        $CurrentKey
        }
    }
