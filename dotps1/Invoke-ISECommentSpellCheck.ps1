<#
    .SYNOPSIS
        Checks the spelling of Comments.
    .DESCRIPTION
        Checks the spelling of Comments using a Microsoft Office Word COM object.
    .INPUTS
        String.
    .OUTPUTS
    .EXAMPLE
        PS C:\> Invoke-ISECommentSpellCheck

        Spelling Error: 'mispeled'.  Suggestions: 'misspelled'
    .NOTES
        Here is a mispeled word to check.
    .LINK
        http://dotps1.github.io
    .LINK
        http://grposh.github.io
#>

Function Invoke-ISECommentSpellCheck {
    [CmdletBinding()]
    [OutputType()]

    Param (
        [Parameter(
            ValueFromPipeline = $true
        )]
        [String[]]
        $Code = $psISE.CurrentFile.Editor.Text
    )

    Begin {
        $winword = New-Object -ComObject Word.Application
        $winword.Documents.Add() |
            Out-Null
    }

    Process {
        foreach ($item in $Code) {
            $comments = [System.Management.Automation.PSParser]::Tokenize($item, [ref]$null) | Where-Object { 
                $_.Type -eq 'Comment' 
            }

            foreach ($comment in $comments) {
                foreach ($word in $comment.Content -split ' ') {
                    if (-not $winword.CheckSpelling($word)) {
                        Write-Output -InputObject "Spelling Error: '$word'.  Suggestions: '$($winword.GetSpellingSuggestions($word) | %{ $_.Name -join '' })'"
                    }
                }
            }
        }
    }

    End {
        Get-Process -Name winword -ErrorAction SilentlyContinue |
            Stop-Process -Force
    }
}
