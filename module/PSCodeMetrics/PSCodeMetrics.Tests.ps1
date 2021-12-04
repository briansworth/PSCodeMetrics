$moduleName = 'PSCodeMetrics'
Remove-Module -Name $moduleName -ErrorAction SilentlyContinue
Import-Module -Name $moduleName


InModuleScope -ModuleName PSCodeMetrics {
  BeforeAll {
    $emptySb = [scriptblock]::Create({})
    $mockExtent = New-Object -TypeName PSObject -Property @{
      'StartLineNumber' = 1;
      'EndLineNumber' = 2;
    }
  }

  Describe 'Find-CommandStatement' {
    BeforeAll {
      $cmdSb = [scriptblock]::Create({Test-Path -Path 'test'})

      $nestedCmdSb = [scriptblock]::Create({
        Test-Path -Path ({Join-Path -Path 'test' -ChildPath 'path'})
      })
    }

    Context 'Without IncludeNestedClause' {
      It 'Finds 1 command' {
        $command = Find-CommandStatement -ScriptBlock $cmdSb

        $command | Should -Not -BeNullOrEmpty $null
        $command | Should -HaveCount 1
      }
      It 'Finds 1 command (nested)' {
        $command = Find-CommandStatement -ScriptBlock $nestedCmdSb

        $command | Should -HaveCount 1
      }
      It 'Finds no commands (empty)' {
        $command = Find-CommandStatement -ScriptBlock $emptySb

        $command | Should -BeNullOrEmpty
      }
    }

    Context 'With IncludeNestedClause' {
      It 'Finds 2 commands (nested)' {
        $command = Find-CommandStatement -ScriptBlock $nestedCmdSb -IncludeNestedClause

        $command | Should -HaveCount 2
      }
    }
  }

  Describe 'Find-IfStatement' {
    BeforeAll {
      $ifSb = [scriptblock]::Create({
        $exist = Test-Path -Path 'test'
        if ($exist){
          return $true
        }
      })

      $nestedIfSb = [scriptblock]::Create({
        $exist = Test-Path -Path 'test'
        {if ($exist){return $true}}
      })
    }
    Context 'Without IncludeNestedClause' {
      It 'Gets if statement' {
        $if = Find-IfStatement -ScriptBlock $ifSb

        $if | Should -HaveCount 1
      }
      It 'Gets no if statement (nested)' {
        $if = Find-IfStatement -ScriptBlock $nestedIfSb

        $if | Should -BeNullOrEmpty
      }
      It 'Gets no if statement (empty)' {
        $if = Find-IfStatement -ScriptBlock $emptySb

        $if | Should -BeNullOrEmpty
      }
    }
    Context 'With IncludeNestedClause' {
      It 'Gets if statement (nested)' {
        $if = Find-IfStatement -ScriptBlock $nestedIfSb -IncludeNestedClause

        $if | Should -HaveCount 1
      }
    }
  }

  Describe 'Find-TryStatement' {
    BeforeAll {
      $trySb = [scriptblock]::Create({
        Try{}Catch{}
      })

      $nestedTrySb = [scriptblock]::Create({
        {Try{}Catch{}}
      })
    }

    Context 'Without IncludeNestedClause' {
      It 'Gets Try statement' {
        $try = Find-TryStatement -ScriptBlock $trySb

        $try | Should -Not -BeNullOrEmpty
        $try | Should -HaveCount 1
      }
      It 'Gets no Try statement (nested)' {
        $try = Find-TryStatement -ScriptBlock $nestedTrySb

        $try | Should -BeNullOrEmpty
      }
      It 'Gets no Try statement (empty)' {
        $try = Find-TryStatement -ScriptBlock $emptySb

        $try | Should -BeNullOrEmpty
      }
    }

    Context 'With IncludeNestedClause' {
      It 'Gets Try statement (nested)' {
        $try = Find-TryStatement -ScriptBlock $nestedTrySb -IncludeNestedClause

        $try | Should -HaveCount 1
      }
    }
  }

  Describe 'Find-WhileStatement' {
    BeforeAll {
      $whileSb = [scriptblock]::Create({while($false){}})

      $nestedWhileSb = [scriptblock]::Create({
        {while($false){}}
      })
    }

    Context 'With IncludeNestedClause' {
      It 'Gets while statement' {
        $while = Find-WhileStatement -ScriptBlock $whileSb

        $while | Should -HaveCount 1
      }
      It 'Gets no while statement (nested)' {
        $while = Find-WhileStatement -ScriptBlock $nestedWhileSb

        $while | Should -BeNullOrEmpty
      }
      It 'Gets no while statement (empty)' {
        $while = Find-WhileStatement -ScriptBlock $emptySb

        $while | Should -BeNullOrEmpty
      }
    }
    Context 'Without IncludeNestedClause' {
      It 'Gets while statement (nested)' {
        $while = Find-WhileStatement -ScriptBlock $nestedWhileSb -IncludeNestedClause

        $while | Should -Not -BeNullOrEmpty
        $while | Should -HaveCount 1
      }
    }
  }

  Describe 'Find-SwitchStatement' {
    BeforeAll {
      $switchSb = [scriptblock]::Create({
        $a = 'test'
        switch($a){default{}}
      })

      $nestedSwitchSb = [scriptblock]::Create({
        $a = 'test'
        {switch($a){default{}}}
      })
    }

    Context 'Without IncludeNestedClause' {
      It 'Gets Switch statement' {
        $swtch = Find-SwitchStatement -ScriptBlock $switchSb

        $swtch | Should -Not -BeNullOrEmpty
        $swtch | Should -HaveCount 1
      }
      It 'Gets no Switch statement (nested)' {
        $swtch = Find-SwitchStatement -ScriptBlock $nestedSwitchSb

        $swtch | Should -BeNullOrEmpty
      }
      It 'Gets no Switch statement (empty)' {
        $swtch = Find-SwitchStatement -ScriptBlock $emptySb

        $swtch | Should -BeNullOrEmpty
      }
    }

    Context 'With IncludeNestedClause' {
      It 'Gets Switch statement (nested)' {
        $swtch = Find-SwitchStatement -ScriptBlock $nestedSwitchSb -IncludeNestedClause

        $swtch | Should -Not -BeNullOrEmpty
        $swtch | Should -HaveCount 1
      }
    }
  }

  Describe 'ClauseMetrics' {
    BeforeAll {
      $type = 'ClauseMetrics'
      $clauseStats = New-ClauseMetricsClassInstance -TypeName $type
      $clauseStats.Cc = 1
      $clauseStats.StartLineNumber = 1
      $clauseStats.EndLineNumber = 1
    }

    Context 'GetLineCount' {
      It 'Produces accurate line count' {
        $lc = $clauseStats.GetLineCount()

        $lc | Should -BeOfType [int]
        $lc | Should -Be 1
      }
    }

    Context 'ToString' {
      It 'Shows code paths in string' {
        $str = $clauseStats.ToString()

        $str | Should -BeOfType [string]
        $str | Should -BeLike '*Cc*'
      }
    }
  }

  Describe 'Get-IfClauseMetrics' {
    BeforeAll {
      $ifType = 'Management.Automation.Language.IfStatementAst'

      $mockElse = New-MockObject -Type $ifType
      $mockElse | Add-Member -Name Clauses `
        -MemberType NoteProperty `
        -Value @(1) `
        -Force
      $mockElse | Add-Member -Name Extent `
        -MemberType NoteProperty `
        -Value $mockExtent `
        -Force
      $mockElse | Add-Member -Name ElseClause `
        -MemberType NoteProperty `
        -Value (New-Object PSObject -Property @{'Extent' = $mockExtent}) `
        -Force
      $mockElse.ElseClause.Extent.EndLineNumber = 3

      $mockElseIf = New-MockObject -Type $ifType
      $mockElseIf | Add-Member -Name Clauses `
        -MemberType NoteProperty `
        -Value @(1, 2) `
        -Force
      $mockElseIf | Add-Member -Name Extent `
        -MemberType NoteProperty `
        -Value $mockExtent `
        -Force
    }

    Context 'Has Else (no ElseIf)' {
      It 'Has correct Else count' {
        $result = Get-IfClauseMetrics -Clause $mockElse

        $result.IfStatements | Should -BeExactly 1
        $result.ElseStatements | Should -BeExactly 1
        $result.ElseIfStatements | Should -BeExactly 0
        $result.EndLineNumber |
          Should -BeExactly $mockElse.ElseClause.Extent.EndLineNumber
      }
    }

    Context 'Has ElseIf (no Else)' {
      It 'Has correct ElseIf count' {
        $result = Get-IfClauseMetrics -Clause $mockElseIf

        $result.IfStatements | Should -BeExactly 1
        $result.ElseIfStatements | Should -BeExactly 1
        $result.ElseStatements | Should -BeExactly 0
      }
    }
  }

  Describe 'Get-TryCatchClauseMetrics' {
    BeforeEach {
      $tryType = 'Management.Automation.Language.TryStatementAst'
      $catchClauseType = 'Management.Automation.Language.CatchClauseAst'
      $finallyType = 'Management.Automation.Language.StatementBlockAst'
      $mockTry = New-MockObject -Type $tryType
      $mockCatch = New-MockObject -Type $catchClauseType
      $mockFinally = New-MockObject -Type $finallyType
      $mockTry | Add-Member -Name Extent `
        -MemberType NoteProperty `
        -Value $mockExtent `
        -Force
    }

    It 'Has Finally set correctly (true)' {
      $mockTry | Add-Member -Name CatchClauses `
        -MemberType NoteProperty `
        -Value @($mockCatch) `
        -Force
      $mockTry | Add-Member -Name Finally `
        -MemberType NoteProperty `
        -Value $mockFinally `
        -Force

      $try = Get-TryCatchClauseMetrics -Clause $mockTry

      $try.HasFinally | Should -Be $true
    }
    It 'Has Finally set correctly (false)' {
      $mockTry | Add-Member -Name CatchClauses `
        -MemberType NoteProperty `
        -Value @($mockCatch) `
        -Force

      $try = Get-TryCatchClauseMetrics -Clause $mockTry

      $try.HasFinally | Should -Be $false
    }
    It 'Has Catch All set correctly' {
      $mockCatch | Add-Member -Name IsCatchAll `
        -MemberType NoteProperty `
        -Value $true `
        -Force
      $mockTry | Add-Member -Name CatchClauses `
        -MemberType NoteProperty `
        -Value @($mockCatch) `
        -Force
      $mockTry | Add-Member -Name Finally `
        -MemberType NoteProperty `
        -Value $mockFinally `
        -Force

        $try = Get-TryCatchClauseMetrics -Clause $mockTry

        $try.CatchAllStatements | Should -BeExactly 1
    }
    It 'Has Catch All / Typed Catch set correctly' {
      $typedCatch = New-MockObject -Type $catchClauseType
      $mockCatch | Add-Member -Name IsCatchAll `
        -MemberType NoteProperty `
        -Value $true `
        -Force
      $mockTry | Add-Member -Name CatchClauses `
        -MemberType NoteProperty `
        -Value @($mockCatch, $typedCatch) `
        -Force
      $mockTry | Add-Member -Name Finally `
        -MemberType NoteProperty `
        -Value $mockFinally `
        -Force

        $try = Get-TryCatchClauseMetrics -Clause $mockTry

        $try.CatchAllStatements | Should -BeExactly 1
        $try.TypedCatchStatements | Should -BeExactly 1
        $try.Cc | Should -BeExactly 2
    }
  }

  Describe 'Get-SwitchClauseMetrics' {
    BeforeEach {
      $switchType = 'Management.Automation.Language.SwitchStatementAst'
      $mockSwitch = New-MockObject -Type $switchType
      $mockSwitch | Add-Member -Name Extent `
        -MemberType NoteProperty `
        -Value $mockExtent `
        -Force
      $mockSwitch | Add-Member -Name Clauses `
        -MemberType NoteProperty `
        -Value @(1) `
        -Force
    }

    It 'Calculates correct number of Switch Clauses' {
      $mockSwitch | Add-Member -Name Clauses `
        -MemberType NoteProperty `
        -Value @(1, 2, 3) `
        -Force

      $result = Get-SwitchClauseMetrics -Clause $mockSwitch

      $result.SwitchStatements | Should -BeExactly 1
      $result.SwitchClauses | Should -BeExactly 3
      $result.Cc | Should -BeExactly 3
    }

    It 'Correctly identifies default clause (true)' {
      $mockSwitch | Add-Member -Name Default `
        -MemberType NoteProperty `
        -Value 1 `
        -Force

      $result = Get-SwitchClauseMetrics -Clause $mockSwitch

      $result.HasDefault | Should -Be $true
    }
    It 'Correctly identifies default clause (false)' {
      $result = Get-SwitchClauseMetrics -Clause $mockSwitch

      $result.HasDefault | Should -Be $false
    }
  }

  Describe 'Get-WhileClauseMetrics' {
    BeforeEach {
      $whileClauseType = 'Management.Automation.Language.WhileStatementAst'
      $mockWhile = New-MockObject -Type $whileClauseType
      $mockWhile | Add-Member -Name Extent `
        -MemberType NoteProperty `
        -Value $mockExtent `
        -Force
    }
    It 'Calculates correct Cc' {
      $result = Get-WhileClauseMetrics -Clause $mockWhile

      $result.WhileStatements | Should -BeExactly 1
      $result.Cc | Should -BeExactly 1
    }
  }

  Describe 'Get-ScriptBlockToken' {
    BeforeAll {
      $mockSb = [scriptblock]::Create({
        # Comment
        Write-Output 'Hello world!'
      })
    }
    It 'Returns all tokens' {
      $result = Get-ScriptBlockToken -ScriptBlock $mockSb

      $result | Should -Not -BeNullOrEmpty
      $result.Count | Should -BeGreaterThan 4
    }
    It 'Returns valid token' {
      $result = Get-ScriptBlockToken -ScriptBlock $mockSb -TokenKind Comment

      $result | Should -Not -BeNullOrEmpty
      $result.Kind | Should -Be 'Comment'
    }
    It 'Returns null if token is not found' {

    }
    It 'Returns multiple tokens' {
      $kinds = @('StringLiteral', 'Comment')
      $result = Get-ScriptBlockToken -ScriptBlock $mockSb -TokenKind $kinds

      $result | Should -HaveCount 2
      foreach ($kind in $kinds)
      {
        $result | Select-Object -ExpandProperty Kind | Should -Contain $kind
      }
    }
  }

  Describe 'Get-BoolOperatorMetrics' {
    BeforeEach {
      $mockToken = New-MockObject -Type 'Management.Automation.Language.Token'
      Mock -CommandName Get-ScriptBlockToken -MockWith {return}
    }

    It 'Properly handles no returned tokens' {
      $result = Get-BoolOperatorMetrics -ScriptBlock $emptySb

      $result.Cc | Should -BeExactly 0
    }

    It 'Reports proper code path count' {
      $mockToken1 = New-MockObject -Type 'Management.Automation.Language.Token'
      $mockToken2 = New-MockObject -Type 'Management.Automation.Language.Token'
      $mockTokens = @($mockToken, $mockToken1, $mockToken2)
      $kinds = @('And', 'Or', 'XOr')
      for ($i=0; $i -lt 3; $i++)
      {
        $mockTokens[$i] | Add-Member -Name Kind `
          -MemberType NoteProperty `
          -Value $kinds[$i] `
          -Force
      }
      Mock -CommandName Get-ScriptBlockToken -MockWith {return $mockTokens}

      $result = Get-BoolOperatorMetrics -ScriptBlock $emptySb

      $result.Cc | Should -BeExactly 3
    }

    Context 'BoolOperatorMetrics ToString' {
      It 'Should return a string with ToString' {
        $result = Get-BoolOperatorMetrics -ScriptBlock $emptySb

        $result.ToString() | Should -BeOfType [string]
      }
    }
  }

  Describe 'Get-MaxNestedDepth' {
    BeforeAll {
      $lCurlyKind = [Management.Automation.Language.TokenKind]::LCurly
      $rCurlyKind = [Management.Automation.Language.TokenKind]::RCurly

      $lCurly = New-MockObject -Type Management.Automation.Language.Token
      $rCurly = New-MockObject -Type Management.Automation.Language.Token

      $lCurly | Add-Member -Name Kind `
        -MemberType NoteProperty `
        -Value $lCurlyKind `
        -Force
      $rCurly | Add-Member -Name Kind `
        -MemberType NoteProperty `
        -Value $rCurlyKind `
        -Force
    }
    BeforeEach {
      Mock -CommandName Get-ScriptBlockToken -MockWith {
        return @($lCurly, $rCurly, $lCurly, $rCurly)
      }
    }

    It 'Returns positive integer' {
      $result = Get-MaxNestedDepth -ScriptBlock $emptySb

      $result | Should -BeOfType [int]
      $result | Should -BeGreaterThan 0
    }

    It 'Throws if Get-ScriptBlockToken throws' {
      $emsg = 'Bad'
      Mock -CommandName Get-ScriptBlockToken -MockWith {
        Write-Error -Message $emsg -ErrorAction Stop
      }

      {Get-MaxNestedDepth -ScriptBlock $emptySb -ErrorAction Stop} | Should -Throw $emsg
    }
  }

  Describe 'Measure-IfStatementMetrics' {
    BeforeAll {
      $measureObjectType = 'Microsoft.PowerShell.Commands.GenericMeasureInfo'
    }
    BeforeEach {
      $measure = New-MockObject -Type $measureObjectType
      $ifStats = New-MockObject -Type ([IfClauseMetrics])

      $measure | Add-Member -Name Sum `
        -MemberType NoteProperty `
        -Value 1 `
        -Force
      $ifStats | Add-Member -Name GetLineCount `
        -MemberType ScriptMethod `
        -Value {return 1} `
        -Force

      Mock -CommandName Measure-Object {
        return $measure
      }
    }

    It 'Successfully measures IfClauseMetrics' {
      $result = Measure-IfStatementMetrics -IfClauseMetrics $ifStats

      $result | Should -BeOfType ([TotalIfClauseMetrics])
      $result.LargestStatementLineCount | Should -BeExactly 1
    }
  }

  Describe 'Measure-TryCatchClauseMetrics' {
    BeforeAll {
      $measureObjectType = 'Microsoft.PowerShell.Commands.GenericMeasureInfo'
    }
    BeforeEach {
      $measure = New-MockObject -Type $measureObjectType
      $stats = New-MockObject -Type ([TryClauseMetrics])

      $measure | Add-Member -Name Sum `
        -MemberType NoteProperty `
        -Value 1 `
        -Force
      $stats | Add-Member -Name GetLineCount `
        -MemberType ScriptMethod `
        -Value {return 1} `
        -Force
      $stats | Add-Member -Name HasFinally `
        -MemberType NoteProperty `
        -Value $true `
        -Force

      Mock -CommandName Measure-Object {
        return $measure
      }
    }

    It 'Successfully measure TryCatchMetrics' {
      $result = Measure-TryCatchClauseMetrics -TryClauseMetrics $stats

      $result | Should -BeOfType ([TotalTryClauseMetrics])
      $result.LargestStatementLineCount | Should -BeExactly 1
      $result.FinallyStatementTotal | Should -BeExactly 1
    }
  }

  Describe 'Measure-SwitchClauseMetrics' {
    BeforeAll {
      $measureObjectType = 'Microsoft.PowerShell.Commands.GenericMeasureInfo'
    }
    BeforeEach {
      $measure = New-MockObject -Type $measureObjectType
      $stats = New-MockObject -Type ([SwitchClauseMetrics])

      $measure | Add-Member -Name Sum `
        -MemberType NoteProperty `
        -Value 1 `
        -Force
      $stats | Add-Member -Name GetLineCount `
        -MemberType ScriptMethod `
        -Value {return 1} `
        -Force
      $stats | Add-Member -Name HasDefault `
        -MemberType NoteProperty `
        -Value $true `
        -Force

      Mock -CommandName Measure-Object {
        return $measure
      }
    }

    It 'Successfully measures SwitchClauseMetrics' {
      $result = Measure-SwitchClauseMetrics -SwitchClauseMetrics $stats

      $result | Should -BeOfType ([TotalSwitchClauseMetrics])
      $result.DefaultClauseTotal | Should -BeExactly 1
    }
  }

  Describe 'Measure-WhileClauseMetrics' {
    BeforeAll {
      $measureObjectType = 'Microsoft.PowerShell.Commands.GenericMeasureInfo'
    }
    BeforeEach {
      $measure = New-MockObject -Type $measureObjectType
      $stats = New-MockObject -Type ([WhileClauseMetrics])

      $measure | Add-Member -Name Sum `
        -MemberType NoteProperty `
        -Value 1 `
        -Force
      $stats | Add-Member -Name GetLineCount `
        -MemberType ScriptMethod `
        -Value {return 1} `
        -Force

      Mock -CommandName Measure-Object {
        return $measure
      }
    }

    It 'Successfully measures WhileClauseMetrics' {
      $result = Measure-WhileClauseMetrics -WhileClauseMetrics $stats

      $result | Should -BeOfType ([TotalWhileClauseMetrics])
    }
  }

  Describe 'Get-ScriptBlockCommandMetrics' {
    BeforeAll {
      $commandAstType = 'Management.Automation.Language.CommandAst'
      $cmdType = 'Management.Automation.Language.StringConstantExpressionAst'
      $varType = 'Management.Automation.Language.VariableExpressionAst'
    }
    BeforeEach {
      $commandAst = New-MockObject -Type $commandAstType

      $cmdElement = New-MockObject -Type $cmdType
      $elementList = New-Object -TypeName Collections.Generic.List[$cmdType]
      $elementList.Add($cmdElement)

      $commandAst | Add-Member -Name CommandElements `
        -MemberType NoteProperty `
        -Value $elementList `
        -Force
    }

    It 'Returns TotalCommandMetrics if no commands are found' {
      Mock -CommandName Find-CommandStatement -MockWith {return $null}

      $result = Get-ScriptBlockCommandMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalCommandMetrics])
      $result.CommandCount | Should -BeExactly 0
    }

    It 'Successfully measures script block command statistics' {
      Mock -CommandName Find-CommandStatement -MockWith {
        return $commandAst
      }

      $result = Get-ScriptBlockCommandMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalCommandMetrics])
      $result.CommandCount | Should -BeExactly $elementList.Count
    }

    It 'Successfully measures script block with Variable expression' {
      $element = New-MockObject -Type $varType
      $elements = New-Object -TypeName Collections.Generic.List[$varType]

      $element | Add-Member -Name VariablePath `
        -MemberType NoteProperty `
        -Value (New-Object -TypeName PSObject -Property @{'UserPath' = 'a';}) `
        -Force
      $elements.Add($element)

      $commandAst | Add-Member -Name CommandElements `
        -MemberType NoteProperty `
        -Value $elements `
        -Force

      Mock -CommandName Find-CommandStatement -MockWith {return $commandAst}

      $result = Get-ScriptBlockCommandMetrics -ScriptBlock $emptySb

      $result.CommandCount | Should -BeExactly $elements.Count
      $result.ToString() | Should -BeOfType [string]
    }
  }

  Describe 'Get-ScriptBlockIfMetrics' {
    BeforeAll {
      $clauseType = 'Management.Automation.Language.IfStatementAst'
      $clauseList = New-Object -TypeName Collections.Generic.List[$clauseType]
      $clause = New-MockObject -Type $clauseType
      $clauseList.Add($clause)

      $stat = New-MockObject -Type ([IfClauseMetrics])
      $total = New-MockObject -Type ([TotalIfClauseMetrics])
    }
    BeforeEach {
      Mock -CommandName Find-IfStatement -MockWith {return $clauseList}
      Mock -CommandName Get-IfClauseMetrics -MockWith {return $stat}
      Mock -CommandName Measure-IfStatementMetrics -MockWith {return $total}
    }

    It 'Succesfully returns when if statements are found' {
      $result = Get-ScriptBlockIfMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalIfClauseMetrics])
      $result.ToString() | Should -BeOfType [string]
    }

    It 'Succesfully returns when no if statements are found' {
      Mock -CommandName Find-IfStatement -MockWith {return $null}

      $result = Get-ScriptBlockIfMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalIfClauseMetrics])
    }
  }

  Describe 'Get-ScriptBlockTryCatchMetrics' {
    BeforeAll {
      $clauseType = 'Management.Automation.Language.TryStatementAst'
      $clauseList = New-Object -TypeName Collections.Generic.List[$clauseType]
      $clause = New-MockObject -Type $clauseType
      $clauseList.Add($clause)

      $stat = New-MockObject -Type ([TryClauseMetrics])
      $total = New-MockObject -Type ([TotalTryClauseMetrics])
    }
    BeforeEach {
      Mock -CommandName Find-TryStatement -MockWith {return $clauseList}
      Mock -CommandName Get-TryCatchClauseMetrics -MockWith {return $stat}
      Mock -CommandName Measure-TryCatchClauseMetrics `
        -MockWith {return $total}
    }

    It 'Succesfully returns when Try/Catch statements are found' {
      $result = Get-ScriptBlockTryCatchMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalTryClauseMetrics])
    }
    It 'Succesfully returns when no Try/Catch statements are found' {
      Mock -CommandName Find-TryStatement -MockWith {return $null}

      $result = Get-ScriptBlockTryCatchMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalTryClauseMetrics])
    }
  }

  Describe 'Get-ScriptBlockSwitchMetrics' {
    BeforeAll {
      $clauseType = 'Management.Automation.Language.SwitchStatementAst'
      $clauseList = New-Object -TypeName Collections.Generic.List[$clauseType]
      $clause = New-MockObject -Type $clauseType
      $clauseList.Add($clause)

      $stat = New-MockObject -Type ([SwitchClauseMetrics])
      $total = New-MockObject -Type ([TotalSwitchClauseMetrics])
    }
    BeforeEach {
      Mock -CommandName Find-SwitchStatement -MockWith {return $clauseList}
      Mock -CommandName Get-SwitchClauseMetrics -MockWith {return $stat}
      Mock -CommandName Measure-SwitchClauseMetrics `
        -MockWith {return $total}
    }

    It 'Succesfully returns when switch statements are found' {
      $result = Get-ScriptBlockSwitchMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalSwitchClauseMetrics])
    }
    It 'Succesfully returns when no switch statements are found' {
      Mock -CommandName Find-SwitchStatement -MockWith {return $null}

      $result = Get-ScriptBlockSwitchMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalSwitchClauseMetrics])
    }
  }

  Describe 'Get-ScriptBlockWhileMetrics' {
    BeforeAll {
      $clauseType = 'Management.Automation.Language.WhileStatementAst'
      $clauseList = New-Object -TypeName Collections.Generic.List[$clauseType]
      $clause = New-MockObject -Type $clauseType
      $clauseList.Add($clause)

      $stat = New-MockObject -Type ([WhileClauseMetrics])
      $stat | Add-Member -Name WhileStatements `
        -MemberType NoteProperty `
        -Value 1 `
        -Force
      $total = New-MockObject -Type ([TotalWhileClauseMetrics])
    }
    BeforeEach {
      Mock -CommandName Find-WhileStatement -MockWith {return $clauseList}
      Mock -CommandName Get-WhileClauseMetrics -MockWith {return $stat}
      Mock -CommandName Measure-WhileClauseMetrics `
        -MockWith {return $total}
    }

    It 'Succesfully returns when while statements are found' {
      $result = Get-ScriptBlockWhileMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalWhileClauseMetrics])

      Assert-MockCalled -CommandName Get-WhileClauseMetrics -Times 1 -Scope It
    }
    It 'Succesfully returns when no while statements are found' {
      Mock -CommandName Find-WhileStatement -MockWith {return $null}

      $result = Get-ScriptBlockWhileMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([TotalWhileClauseMetrics])
    }
  }

  Describe 'ConvertTo-ScriptBlock' {
    BeforeEach {
      $mockCommand = New-MockObject -Type Management.Automation.ApplicationInfo
      $mockCommand | Add-Member -Name CommandType `
        -MemberType NoteProperty `
        -Value 'Function' `
        -Force

      Mock -CommandName Get-Command -MockWith {return $mockCommand}
      Mock -CommandName Join-Path -MockWith {return 'path'}
      Mock -CommandName Get-Content -MockWith {return $emptySb}
    }

    It 'Throws when command is not found' {
      $emsg = 'Command Not Found'
      $exceptType = 'Management.Automation.CommandNotFoundException'
      Mock -CommandName Get-Command -MockWith {
        $except = New-Object -TypeName $exceptType -ArgumentList @($emsg)
        Write-Error -Exception $except
      }

      {ConvertTo-ScriptBlock -FunctionName 'Test' -ErrorAction Stop} |
        Should -Throw $emsg -ExceptionType $exceptType
    }

    It 'Throws if command is not a function' {
      $exceptType = 'Microsoft.PowerShell.Commands.WriteErrorException'
      $mockCommand | Add-Member -Name CommandType `
        -MemberType NoteProperty `
        -Value 'Cmdlet' `
        -Force
      Mock -CommandName Get-Command -MockWith {return $mockCommand}

      {ConvertTo-ScriptBlock -FunctionName 'Test' -ErrorAction Stop} |
        Should -Throw -ExceptionType $exceptType
    }

    It 'Returns a script block if successful' {
      $result = ConvertTo-ScriptBlock -FunctionName 'Test'

      $result | Should -BeOfType [scriptblock]
    }
  }

  Describe 'Get-FunctionMetrics' {
    BeforeAll {
      $ifs = New-MockObject -Type ([TotalIfClauseMetrics])
      $trys = New-MockObject -Type ([TotalTryClauseMetrics])
      $switches = New-MockObject -Type ([TotalSwitchClauseMetrics])
      $whiles = New-MockObject -Type ([TotalWhileClauseMetrics])
      $operators = New-MockObject -Type ([BoolOperatorMetrics])
      $commands = New-MockObject -Type ([TotalCommandMetrics])
    }
    BeforeEach {
      Mock -CommandName Get-ScriptBlockIfMetrics -MockWith {return $ifs}
      Mock -CommandName Get-ScriptBlockTryCatchMetrics `
        -MockWith {return $trys}
      Mock -CommandName Get-ScriptBlockSwitchMetrics `
        -MockWith {return $switches}
      Mock -CommandName Get-ScriptBlockWhileMetrics `
        -MockWith {return $whiles}
      Mock -CommandName Get-BoolOperatorMetrics -MockWith {return $operators}
      Mock -CommandName Get-ScriptBlockCommandMetrics `
        -MockWith {return $commands}

      Mock -CommandName ConvertTo-ScriptBlock -MockWith {return $emptySb}
    }

    It 'Successfully measures all components of scriptblock' {
      $result = Get-FunctionMetrics -ScriptBlock $emptySb

      $result | Should -BeOfType ([FunctionMetrics])
      Assert-MockCalled -CommandName ConvertTo-ScriptBlock -Times 0 -Scope It
    }

    It 'Successfully measures function statistics' {
      $result = Get-FunctionMetrics -FunctionName 'Test'

      $result | Should -BeOfType ([FunctionMetrics])
      Assert-MockCalled -CommandName ConvertTo-ScriptBlock -Times 1 -Scope It
    }

    It 'Throws if ConvertTo-ScriptBlock throws' {
      $exceptType = 'Microsoft.PowerShell.Commands.WriteErrorException'
      $emsg = 'Error'
      Mock -CommandName ConvertTo-ScriptBlock -MockWith {
        Write-Error -Message $emsg
      }

      {Get-FunctionMetrics -FunctionName 'Test' -ErrorAction Stop} |
        Should -Throw $emsg -ExceptionType $exceptType
    }
  }
}

$mockFunction = {
  [CmdletBinding()]
  Param(
    [bool]$Test = $true
  )
  <#
  Comment block
  #>
  Process
  {
    Try # Try
    {
      if ($Test -and $false) # if (+1) and operator (+1)
      {
        if ($null -like $false) # if (+1)
        {
          Get-Command -Name Get-Command `
            -ErrorAction Stop
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
        foreach ($test in @(0, 1)) # Foreach (+1)
        {
          continue
        }
        for ($i = 0; $i -lt 3; $i++) # For (+1)
        {
          continue
        }
        for ($i = 0;;$i++) # For no condition (+0)
        {
          break
        }
        for (;;) # For no initializer condition or iterator (+0)
        {
          break
        }
        for (($i = 0),($j = 0); $i -lt 3 -and $j -lt 3; $i++,$j++) # For (+1) -and (+1)
        {
          "`$i:$i"
          "`$j:$j"
        }
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
