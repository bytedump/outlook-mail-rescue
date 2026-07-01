@{
    # PSScriptAnalyzer policy for Outlook Mail Rescue.
    # Rules below are excluded deliberately - each is an intentional, defensible choice
    # for this tool, NOT a hidden defect. Everything else stays on.
    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        # Best-effort cleanup/probe blocks: releasing COM objects, reading optional
        # registry values, GetAttributes on a dir we then skip, log file appends. A
        # failure there is genuinely ignorable; rethrowing would be wrong.
        'PSAvoidUsingEmptyCatchBlock',

        # This is an interactive GUI tool; confirmation happens in the UI (MessageBox),
        # not via -WhatIf/-Confirm. Adding ShouldProcess to every New-/Set-/Start-
        # helper would add ceremony with no user-facing benefit.
        'PSUseShouldProcessForStateChangingFunctions',

        # Write-Log intentionally echoes to the console (Write-Host) in addition to the
        # log file and the GUI pane; the console stream is a deliberate output here.
        'PSAvoidUsingWriteHost',

        # False positive: 'Write-Log' is our own function, not a shipped cmdlet.
        'PSAvoidOverwritingBuiltInCmdlets',

        # Collective nouns (Get-OutlookStore*s*, Remove-InvalidFileNameChar*s*) read
        # naturally and follow the data they return.
        'PSUseSingularNouns'
    )
}
