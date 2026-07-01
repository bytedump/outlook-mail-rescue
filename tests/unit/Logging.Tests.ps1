#Requires -Version 5.1
# Unit tests for src/Logging.ps1 (Initialize-Logging, Set-LogQueue, Write-Log).
# All pure / filesystem / in-memory queue work - no COM, GUI, or admin needed.

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
}

Describe 'Initialize-Logging' {
    It 'returns the explicit path verbatim when -Path is given' {
        $p = Join-Path $TestDrive 'explicit.log'
        Initialize-Logging -Path $p | Should -Be $p
    }
    It 'builds a default path under LOCALAPPDATA with the tag and a .log extension' {
        $p = Initialize-Logging -Tag 'unittest'
        $p | Should -Match ([regex]::Escape($env:LOCALAPPDATA))
        $p | Should -Match 'mailrescue_unittest_\d{8}-\d{6}\.log$'
    }
    It 'defaults the tag to session' {
        Initialize-Logging | Should -Match 'mailrescue_session_\d{8}-\d{6}\.log$'
    }
}

Describe 'Write-Log' {
    BeforeEach {
        $logPath = Join-Path $TestDrive ('w_' + [guid]::NewGuid().ToString('N') + '.log')
        Initialize-Logging -Path $logPath | Out-Null
        # Reset the SUT's shared script-scope accumulators between cases.
        $script:MrErrors = [System.Collections.Generic.List[string]]::new()
        Set-LogQueue $null
    }

    It 'writes a timestamped, leveled line to the log file' {
        Write-Log -Level INFO -Message 'hello world'
        (Get-Content -Path $logPath -Raw) |
            Should -Match '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \[INFO\] hello world'
    }

    It 'defaults the level to INFO' {
        Write-Log -Message 'no level'
        (Get-Content -Path $logPath -Raw) | Should -Match '\[INFO\] no level'
    }

    It 'rejects an out-of-set level (ValidateSet guard)' {
        { Write-Log -Level NOPE -Message 'x' } | Should -Throw
    }

    It 'accumulates <level> lines into the error list' -ForEach @(
        @{ level = 'ERROR' }
        @{ level = 'FATAL' }
    ) {
        Write-Log -Level $level -Message "boom $level"
        $script:MrErrors.Count | Should -Be 1
        $script:MrErrors[0]     | Should -Match "\[$level\] boom $level"
    }

    It 'does not accumulate INFO or WARN into the error list' {
        Write-Log -Level INFO -Message 'fine'
        Write-Log -Level WARN -Message 'careful'
        $script:MrErrors.Count | Should -Be 0
    }

    It 'enqueues a {Level,Line} object to the attached queue' {
        $q = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        Set-LogQueue $q
        Write-Log -Level WARN -Message 'queued'
        $q.Count | Should -Be 1
        $item = $null
        $q.TryDequeue([ref]$item) | Should -BeTrue
        $item.Level | Should -Be 'WARN'
        $item.Line  | Should -Match '\[WARN\] queued'
    }

    It 'is a no-op-safe call with no log file and no queue (headless helper run)' {
        $script:MrLogFile = $null
        Set-LogQueue $null
        { Write-Log -Message 'headless' } | Should -Not -Throw
    }
}

Describe 'Set-LogQueue' {
    It 'detaches the queue when passed $null' {
        $q = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        Set-LogQueue $q
        Set-LogQueue $null
        Write-Log -Message 'after detach'
        $q.Count | Should -Be 0
    }
}
