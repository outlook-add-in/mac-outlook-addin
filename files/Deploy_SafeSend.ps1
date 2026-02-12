# PayTM Safe Send - VSTO Deployment Scripts
# Run these scripts as Administrator

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 1: Verify Installation
# ═══════════════════════════════════════════════════════════════════════════

function Verify-SafeSendInstallation {
    Write-Host "Checking PayTM Safe Send installation..." -ForegroundColor Cyan
    
    # Check registry
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\PayTMSafeSend"
    
    if (Test-Path $regPath) {
        Write-Host "✓ Registry entry found" -ForegroundColor Green
        Get-Item $regPath | Get-ItemProperty
    } else {
        Write-Host "✗ Registry entry NOT found" -ForegroundColor Red
        Write-Host "  Run: Deploy-SafeSend-Registry" -ForegroundColor Yellow
    }
    
    # Check DLL
    $dllPath = "C:\Program Files\PayTMSafeSend\PayTMSafeSend.dll"
    if (Test-Path $dllPath) {
        Write-Host "✓ DLL file found at: $dllPath" -ForegroundColor Green
    } else {
        Write-Host "✗ DLL file NOT found" -ForegroundColor Red
    }
    
    # Check Outlook
    $outlook = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($outlook) {
        Write-Host "✓ Outlook is running (restart required for add-in to load)" -ForegroundColor Green
    } else {
        Write-Host "ℹ Outlook is not running" -ForegroundColor Yellow
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 2: Registry Deployment (No MSI)
# ═══════════════════════════════════════════════════════════════════════════

function Deploy-SafeSend-Registry {
    param(
        [string]$DLLPath = "C:\Program Files\PayTMSafeSend\PayTMSafeSend.dll"
    )
    
    Write-Host "Deploying PayTM Safe Send via registry..." -ForegroundColor Cyan
    
    # Check DLL exists
    if (-not (Test-Path $DLLPath)) {
        Write-Host "✗ DLL not found at: $DLLPath" -ForegroundColor Red
        Write-Host "Please copy DLL first or specify correct path" -ForegroundColor Yellow
        return
    }
    
    # Create registry path
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\PayTMSafeSend"
    
    if (-not (Test-Path $regPath)) {
        Write-Host "Creating registry key..." -ForegroundColor Gray
        New-Item -Path $regPath -Force | Out-Null
    }
    
    # Set registry values
    Write-Host "Setting registry values..." -ForegroundColor Gray
    
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -PropertyType DWord
    Set-ItemProperty -Path $regPath -Name "Description" -Value "PayTM Safe Send" -PropertyType String
    Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "PayTM Safe Send" -PropertyType String
    Set-ItemProperty -Path $regPath -Name "Manifest" -Value $DLLPath -PropertyType String
    
    Write-Host "✓ Registry deployment complete" -ForegroundColor Green
    Write-Host "✓ Restart Outlook for changes to take effect" -ForegroundColor Green
    
    Verify-SafeSendInstallation
}

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 3: Remove Installation
# ═══════════════════════════════════════════════════════════════════════════

function Remove-SafeSend {
    Write-Host "Removing PayTM Safe Send..." -ForegroundColor Cyan
    
    # Stop Outlook
    Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Remove registry
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\PayTMSafeSend"
    if (Test-Path $regPath) {
        Write-Host "Removing registry entry..." -ForegroundColor Gray
        Remove-Item -Path $regPath -Force
        Write-Host "✓ Registry entry removed" -ForegroundColor Green
    }
    
    # Remove DLL
    $dllPath = "C:\Program Files\PayTMSafeSend\PayTMSafeSend.dll"
    if (Test-Path $dllPath) {
        Write-Host "Removing DLL..." -ForegroundColor Gray
        Remove-Item -Path $dllPath -Force -ErrorAction SilentlyContinue
        Write-Host "✓ DLL removed" -ForegroundColor Green
    }
    
    Write-Host "✓ Removal complete" -ForegroundColor Green
}

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 4: Group Policy Deployment
# ═══════════════════════════════════════════════════════════════════════════

function Deploy-SafeSend-GroupPolicy {
    param(
        [string]$MSIPath = "\\domain\share\addins\PayTMSafeSend.msi",
        [string]$OU = "OU=Computers,DC=domain,DC=com"
    )
    
    Write-Host "Creating Group Policy for PayTM Safe Send..." -ForegroundColor Cyan
    Write-Host "Manual steps required in Group Policy Editor:" -ForegroundColor Yellow
    
    Write-Host @"
1. Open Group Policy Editor (gpedit.msc)
2. Navigate to:
   Computer Configuration > Software Settings > Software Installation
3. Right-click → New → Package
4. Select: $MSIPath
5. Choose: "Assigned"
6. Click OK
7. Run gpupdate /force on target machines
8. Machines will auto-install on next login

Or use PowerShell on domain controller:
  Import-Module GroupPolicy
  New-GPO -Name "Deploy PayTM Safe Send" | New-GPLink -Target "$OU"
"@
}

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 5: Check Installation on Multiple Machines
# ═══════════════════════════════════════════════════════════════════════════

function Check-SafeSend-Remote {
    param(
        [string[]]$ComputerNames = @("computer1", "computer2", "computer3")
    )
    
    Write-Host "Checking SafeSend installation on remote machines..." -ForegroundColor Cyan
    
    foreach ($computer in $ComputerNames) {
        Write-Host "`n[$computer]" -ForegroundColor Cyan
        
        $regPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\PayTMSafeSend"
        
        try {
            $session = New-PSSession -ComputerName $computer
            
            $result = Invoke-Command -Session $session -ScriptBlock {
                Test-Path "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\PayTMSafeSend"
            }
            
            if ($result) {
                Write-Host "  ✓ Installed" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Not installed" -ForegroundColor Red
            }
            
            Remove-PSSession $session
        }
        catch {
            Write-Host "  ✗ Error: $_" -ForegroundColor Red
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# SCRIPT 6: View SafeSend Logs
# ═══════════════════════════════════════════════════════════════════════════

function Get-SafeSend-Logs {
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [int]$LineCount = 50
    )
    
    $logPath = "\\$ComputerName\c$\Users\$env:USERNAME\AppData\Roaming\PayTMSafeSend\log.txt"
    
    if (Test-Path $logPath) {
        Write-Host "Latest $LineCount log entries:" -ForegroundColor Cyan
        Get-Content $logPath -Tail $LineCount
    } else {
        Write-Host "Log file not found at: $logPath" -ForegroundColor Red
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# USAGE EXAMPLES
# ═══════════════════════════════════════════════════════════════════════════

<#

# Verify installation on current machine
Verify-SafeSendInstallation

# Deploy via registry
Deploy-SafeSend-Registry -DLLPath "C:\Program Files\PayTMSafeSend\PayTMSafeSend.dll"

# Remove installation
Remove-SafeSend

# Check multiple machines
Check-SafeSend-Remote -ComputerNames @("DESKTOP-001", "DESKTOP-002", "LAPTOP-001")

# View logs
Get-SafeSend-Logs -ComputerName $env:COMPUTERNAME

# Deploy via Group Policy (manual guide)
Deploy-SafeSend-GroupPolicy -MSIPath "\\domain\share\addins\PayTMSafeSend.msi"

#>
