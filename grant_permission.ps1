# Run this script AS ADMINISTRATOR to grant a user write access to the capture input registry key
# After this, the user can switch inputs without any elevation

param(
    [string]$Username = $env:USERNAME
)

$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e96c-e325-11ce-bfc1-08002be10318}\0021"

Write-Host "Granting '$Username' write access to capture input registry key..."
Write-Host "Path: $regPath"
Write-Host ""

try {
    $acl = Get-Acl $regPath
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule(
        $Username,
        "SetValue",
        "Allow"
    )
    $acl.AddAccessRule($rule)
    Set-Acl $regPath $acl

    Write-Host "SUCCESS! User '$Username' can now switch inputs without admin rights."
    Write-Host ""
    Write-Host "Test with: CaptureInputSwitcher.exe switch 0 CY3014"
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Make sure you're running this as Administrator."
    exit 1
}
