Remove-Module -Name PSCodeStats -ErrorAction SilentlyContinue
Import-Module -Name ./PSCodeStats
Describe "PSCodeStats Tests" {
  InModuleScope -ModuleName PSCodeStats {
    BeforeAll {
      $mockFunction = {
        [CmdletBinding()]
        Param(
          [bool]$Test = $true
        )
        Process
        {
          Try # Try
          {
            if ($Test -and $false) # if (+1) and operator (+1)
            {
              if ($null -like $false) # if (+1)
              {
                Get-Command -Name Get-Command
              }
              else
              {
                # No elseif
              }
            }
            elseif ($Test -or $false) # elseif (+1) or operator (+1)
            {
              # elseif 
            }
            elseif ($false) # elseif (+1)
            {
              Test-Path -Path Env:
              switch ($test) # Switch
              {
                $true {break} # Clause (+1)
                $false {break} # Clause (+1)
                default {break}
              }
            }
            else
            {
              while ($false) # While (+1)
              {
                Try
                {
                  #Try
                }
                Catch # Catch (+1)
                {
                  # CatchAll No Finally
                }
              }
              do
              {
                # Do / While
              } while ($false -xor 1) # xor operator (+1)
            }
          }
          Catch [Runtime.InteropServices.ExternalException] # Typed Catch (+1)
          {
            # Typed Catch
          }
          Catch [Exception] # Typed Catch (+1)
          {
            # Typed Catch
          }
          Catch # Catch (+1)
          {
            # Catch all 
          }
          Finally
          {
            # Finally
          }
        }
      }

    }

    Context "Find-IfStatement" {
      It "Finds If clause" {
        $clause = Find-IfStatement -ScriptBlock $mockFunction -IncludeNestedClause
        $clause | Should -Not -Be $null
      }
      It "Finds no If clause" {
        $clause = Find-IfStatement -ScriptBlock {'hi'}
        $clause | Should -Be $null
      }
    }

    Context "Find-TryStatement" {
      It "Finds Try statement" {
        $clause = Find-TryStatement -ScriptBlock $mockFunction -IncludeNestedClause
        $clause | Should -Not -Be $null
      }
      It "Finds no Try statement" {
        $clause = Find-TryStatement -ScriptBlock {'hi'} 
        $clause | Should -Be $null
      }
    }

    Context "Find-WhileStatement" {
      It "Finds While statement" {
        $clause = Find-WhileStatement -ScriptBlock $mockFunction -IncludeNestedClause
        $clause | Should -Not -Be $null
      }
      It "Finds no While statement" {
        $clause = Find-WhileStatement -ScriptBlock {'hi'}
        $clause | Should -Be $null
      }
    }

    Context "Find-SwitchStatement" {
      It "Finds Switch statement" {
        $clause = Find-SwitchStatement -ScriptBlock $mockFunction -IncludeNestedClause
        $clause | Should -Not -Be $null
      }
      It "Finds no Switch statement" {
        $clause = Find-SwitchStatement -ScriptBlock {'hi'}
        $clause | Should -Be $null
      }
    }

    Context "ClauseStatistics" {
      It "Has ToString method" {
        $classInstance = New-ClauseStatisticsClassInstance
        $classInstance.ToString() | Should -Be '{CodePaths = 0...}'
      }
      It "Has GetLineCount method" {
        $classInstance = New-ClauseStatisticsClassInstance
        $classInstance.GetLineCount() | Should -Be '1'
      }
    }

    Context "Get-IfClauseStatistics" {
      It "Gets correct If clause statistics" {
        $ifClause = Find-IfStatement -IncludeNestedClause -ScriptBlock $mockFunction
        $stats = Get-IfClauseStatistics -Clause $ifClause[0]
        $stats.IfStatements | Should -Be 1
        $stats.ElseStatements | Should -Be 1
        $stats.ElseIfStatements | Should -Be 2
        $stats.CodePaths | Should -Be 3
      }
      It "Gets no Else clause" {
        $noElseSb = [scriptblock]::Create({if ($true) {'hi'}})
        $ifClause = Find-IfStatement -ScriptBlock $noElseSb
        $stats = Get-IfClauseStatistics -Clause $ifClause[0]
        $stats.ElseStatements | Should -Be 0
      }
    }

    Context "Get-TryCatchClauseStatistics" {
      It "Gets correct Try clause statistics" {
        $tryClause = Find-TryStatement -IncludeNestedClause -ScriptBlock $mockFunction
        $stats = Get-TryCatchClauseStatistics -Clause $tryClause[0]
        $stats.CatchStatements | Should -Be 3
        $stats.TypedCatchStatements | Should -Be 2
        $stats.CatchAllStatements | Should -Be 1
        $stats.HasFinally | Should -Be $true
        $stats.CodePaths | Should -Be 3
      }
      It "Gets no Finally clause" {
        $tryClause = Find-TryStatement -IncludeNestedClause -ScriptBlock $mockFunction
        $stats = Get-TryCatchClauseStatistics -Clause $tryClause[1]
        $stats.HasFinally | Should -Be $false
      }
    }

    Context "Get-ScriptBlockToken" {
      It "Gets correct token kind" {
        $tokenKind = [Management.Automation.Language.TokenKind]::LCurly
        $sb = [scriptblock]::Create({if ($true){'hi'} else {'bye'}})
        $sbTokens = Get-ScriptBlockToken -TokenKind $tokenKind -ScriptBlock $sb
        $sbTokens.Count | Should -Be 2
        foreach ($token in $sbTokens)
        {
          $token.Kind | Should -Be 'LCurly'
        }
      }
      
    }

    Context "Get-BoolOperatorStatistics" {
      It "Gets correct Operator statistics" {
        $stats = Get-BoolOperatorStatistics -ScriptBlock $mockFunction
        $stats.CodePaths | Should -Be 3
        $stats.AndOperators | Should -Be 1
        $stats.OrOperators | Should -Be 1
        $stats.XOrOperators | Should -Be 1
      }
      It "Gets no Operators correctly" {
        $sb = [scriptblock]::Create({'hi'})
        $stats = Get-BoolOperatorStatistics -ScriptBlock $sb
        $stats.CodePaths | Should -Be 0
      }
    }

    Context "Get-MaxNestedDepth" {
      It "Gets correct MaxNestedDepth" {
        $depth = Get-MaxNestedDepth -ScriptBlock $mockFunction
        $depth | Should -Be 5
      }
      It "Gets correct depth with no nested blocks" {
        $sb = [scriptblock]::Create({})
        $depth = Get-MaxNestedDepth -ScriptBlock $sb
        $depth | Should -Be 0
      }
      It "Throws if Get-ScriptBlockToken throws" {
        mock -CommandName Get-ScriptBlockToken -MockWith {throw 'Error'}
        $sb = [scriptblock]::Create({})
        {Get-MaxNestedDepth -ScriptBlock $sb -ErrorAction Stop} | Should -Throw
      }
    }
  }
}
