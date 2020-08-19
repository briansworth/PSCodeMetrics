Set-StrictMode -Version latest

Function Find-CommandStatement
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $commands = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.CommandAst]},
    $IncludeNestedClause.ToBool()
  )
  return $commands
}

Function Find-IfStatement
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ifClause = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.IfStatementAst]},
    $IncludeNestedClause.ToBool()
  )

  return $ifClause
}

Function Find-TryStatement
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $catch = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.TryStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $catch
}

Function Find-WhileStatement
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $while = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.WhileStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $while
}

Function Find-SwitchStatement
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $switchStatement = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.SwitchStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $switchStatement
}


class ClauseStatistics
{
  [int]$CodePaths
  [int]$StartLineNumber
  [int]$EndLineNumber

  [int] GetLineCount()
  {
    return $this.EndLineNumber - $this.StartLineNumber + 1
  }

  [string] ToString()
  {
    return "{CodePaths = $($this.CodePaths)...}"
  }
}


class IfClauseStatistics: ClauseStatistics
{
  [int]$IfStatements
  [int]$ElseIfStatements
  [int]$ElseStatements
}


class TryClauseStatistics: ClauseStatistics
{
  [int]$CatchStatements
  [int]$CatchAllStatements
  [int]$TypedCatchStatements
  [bool]$HasFinally
}


class SwitchClauseStatistics: ClauseStatistics
{
  [int]$SwitchStatements
  [int]$SwitchClauses
  [bool]$HasDefault
}


class WhileClauseStatistics: ClauseStatistics
{
  [int]$WhileStatements
}


Function New-ClauseStatisticsClassInstance
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0)]
    [ValidateSet(
      'ClauseStatistics',
      'IfClauseStatistics',
      'TryCatchClauseStatistics',
      'SwitchClauseStatistics',
      'WhileClauseStatistics'
    )]
    [string]$TypeName = 'ClauseStatistics'
  )
  $classInstance = New-Object -TypeName $TypeName
  return $classInstance
}

Function Get-IfClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.IfStatementAst]$Clause
  )
  $ifStats = New-Object -TypeName IfClauseStatistics
  $ifStats.IfStatements = 1

  $ifStats.ElseIfStatements = $Clause.Clauses.Count - $ifStats.IfStatements
  $ifStats.CodePaths = $Clause.Clauses.Count
  $ifStats.StartLineNumber = $Clause[0].Extent.StartLineNumber

  if ($null -eq $Clause.ElseClause)
  {
    $ifStats.EndLineNumber = $Clause[-1].Extent.EndLineNumber
  }
  else
  {
    $ifStats.ElseStatements = 1
    $ifStats.EndLineNumber = $Clause.ElseClause.Extent.EndLineNumber
  }
  return $ifStats
}

Function Get-TryCatchClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.TryStatementAst]$Clause
  )
  $catchStats = New-Object -TypeName TryClauseStatistics

  $catchStats.CatchStatements = $Clause.CatchClauses.Count
  $catchStats.CodePaths = $catchStats.CatchStatements
  $catchStats.StartLineNumber = $Clause.Extent.StartLineNumber
  $catchStats.EndLineNumber = $Clause.Extent.EndLineNumber
  $catchStats.HasFinally = $false

  $catchAlls = $Clause.CatchClauses | Where-Object {$_.IsCatchAll} | Measure-Object
  $catchStats.CatchAllStatements = $catchAlls.Count
  $catchStats.TypedCatchStatements = $catchStats.CatchStatements - $catchAlls.Count
  if ($null -ne $Clause.Finally)
  {
    $catchStats.HasFinally = $true
  }

  return $catchStats
}

Function Get-SwitchClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.SwitchStatementAst]$Clause
  )
  $switchStats = New-Object -TypeName SwitchClauseStatistics

  $switchStats.SwitchStatements = 1
  $switchStats.StartLineNumber = $Clause.Extent.StartLineNumber
  $switchStats.EndLineNumber = $Clause.Extent.EndLineNumber
  $switchStats.HasDefault = $false

  $switchClauses = $Clause.Clauses | Measure-Object
  $switchStats.SwitchClauses = $switchClauses.Count
  $switchStats.CodePaths = $switchStats.SwitchClauses

  if ($null -ne $Clause.Default)
  {
    $switchStats.HasDefault = $true
  }

  return $switchStats
}

Function Get-WhileClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.WhileStatementAst]$Clause
  )
  $whileStats = New-Object -TypeName WhileClauseStatistics

  $whileStats.WhileStatements = 1
  $whileStats.CodePaths = 1
  $whileStats.StartLineNumber = $Clause.Extent.StartLineNumber
  $whileStats.EndLineNumber = $Clause.Extent.EndLineNumber

  return $whileStats
}

Function Get-ScriptBlockToken
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [Parameter(Position=1)]
    [Management.Automation.Language.TokenKind[]]$TokenKind
  )
  $tokens = $null
  $errors = $null

  [void][Management.Automation.Language.Parser]::ParseInput(
    $ScriptBlock,
    [ref]$tokens,
    [ref]$errors
  )

  if ($PSBoundParameters.ContainsKey('TokenKind'))
  {
    $tokens = $tokens | Where-Object {$_.Kind -in $TokenKind}
  }
  return $tokens
}


class BoolOperatorStatistics
{
  [int]$CodePaths
  [int]$AndOperators
  [int]$OrOperators
  [int]$XOrOperators

  [string] ToString()
  {
    return "{CodePaths = $($this.CodePaths)...}"
  }
}


Function Get-BoolOperatorStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $operators = Get-ScriptBlockToken -ScriptBlock $ScriptBlock `
    -TokenKind @('and', 'or', 'xor')

  $orCount = 0
  $andCount = 0
  $xorCount = 0
  $operatorCount = 0

  foreach ($operator in $operators)
  {
    $operatorCount++
    if ($operator.Kind -eq 'And')
    {
      $andCount++
    }
    if ($operator.Kind -eq 'Or')
    {
      $orCount++
    }
    if ($operator.Kind -eq 'XOr')
    {
      $xorCount++
    }
  }

  $stats = New-Object -TypeName BoolOperatorStatistics
  $stats.CodePaths = $operatorCount
  $stats.AndOperators = $andCount
  $stats.OrOperators = $orCount
  $stats.XOrOperators = $xorCount

  return $stats
}

Function Get-MaxNestedDepth
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  Try
  {
    $max = 0
    $stack = New-Object -TypeName Collections.Stack

    $curlyTokens = Get-ScriptBlockToken -ScriptBlock $ScriptBlock `
      -TokenKind AtCurly, LCurly, RCurly `
      -ErrorAction Stop

    foreach ($token in $curlyTokens){
      if ($token.Kind -in @('AtCurly', 'LCurly'))
      {
        $stack.Push($token)
      }
      elseif ($token.Kind -eq 'RCurly')
      {
        [void]$stack.Pop()
      }
      if ($stack.Count -gt $max)
      {
        $max++
      }
    }
    return $max
  }
  Catch
  {
    Write-Error -ErrorRecord $_
  }
}


class TotalCommandStatistics
{
  [int]$CommandCount
  [int]$UniqueCommandCount
  [Collections.Generic.List[string]]$CommandNames

  [string] ToString()
  {
    return "{CommandCount = $($this.CommandCount)...}"
  }
}


class TotalClauseStatistics
{
  [int]$CodePathTotal
  [int]$LargestStatementLineCount

  [string] ToString()
  {
    return "{CodePathTotal = $($this.CodePathTotal)...}"
  }
}


class TotalIfClauseStatistics: TotalClauseStatistics
{
  [int]$IfStatementTotal
  [int]$ElseIfStatementTotal
  [int]$ElseStatementTotal
}


class TotalTryClauseStatistics: TotalClauseStatistics
{
  [int]$TryStatementTotal
  [int]$CatchStatementTotal
  [int]$CatchAllStatementTotal
  [int]$TypedCatchStatementTotal
  [int]$FinallyStatementTotal
}


class TotalSwitchClauseStatistics: TotalClauseStatistics
{
  [int]$SwitchStatementTotal
  [int]$SwitchClauseTotal
  [int]$DefaultClauseTotal
}


class TotalWhileClauseStatistics: TotalClauseStatistics
{
  [int]$WhileStatementTotal
}


Function Measure-IfStatementStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [IfClauseStatistics[]]$IfClauseStatistics
  )
  $longestStatement = 0

  $codePaths = $IfClauseStatistics | Measure-Object -Property CodePaths -Sum
  $ifs = $IfClauseStatistics | Measure-Object -Property IfStatements -Sum
  $elifs = $IfClauseStatistics | Measure-Object -Property ElseIfStatements -Sum
  $elses = $IfClauseStatistics | Measure-Object -Property ElseStatements -Sum

  foreach ($clause in $IfClauseStatistics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }
  $totalStats = New-Object -TypeName TotalIfClauseStatistics
  $totalStats.CodePathTotal = $codePaths.Sum
  $totalStats.IfStatementTotal = $ifs.Sum
  $totalStats.ElseIfStatementTotal = $elifs.Sum
  $totalStats.ElseStatementTotal = $elses.Sum
  $totalStats.LargestStatementLineCount = $longestStatement
  return $totalStats
}

Function Measure-TryCatchClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [TryClauseStatistics[]]$TryClauseStatistics
  )
  $longestStatement = 0
  $totalFinallyCount = 0

  $codePaths = $TryClauseStatistics | Measure-Object -Property CodePaths -Sum
  $trys = $TryClauseStatistics | Measure-Object
  $catches = $TryClauseStatistics | Measure-Object -Property CatchStatements -Sum
  $catchAlls = $TryClauseStatistics | Measure-Object -Property CatchAllStatements -Sum
  $typedCatches = $TryClauseStatistics | Measure-Object -Property TypedCatchStatements -Sum

  foreach ($clause in $TryClauseStatistics)
  {
    if ($clause.HasFinally)
    {
      $totalFinallyCount++
    }
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }

  $totalStats = New-Object -TypeName TotalTryClauseStatistics
  $totalStats.CodePathTotal = $codePaths.Sum
  $totalStats.TryStatementTotal = $trys.Count
  $totalStats.CatchStatementTotal = $catches.Sum
  $totalStats.CatchAllStatementTotal = $catchAlls.Sum
  $totalStats.TypedCatchStatementTotal = $typedCatches.Sum
  $totalStats.FinallyStatementTotal = $totalFinallyCount
  $totalStats.LargestStatementLineCount = $longestStatement

  return $totalStats
}

Function Measure-SwitchClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [SwitchClauseStatistics[]]$SwitchClauseStatistics
  )
  $longestStatement = 0
  $totalDefaultClauseCount = 0

  $switches = $SwitchClauseStatistics | Measure-Object
  $codePaths = $SwitchClauseStatistics | Measure-Object -Property CodePaths -Sum
  $switchClauses = $SwitchClauseStatistics | Measure-Object -Property SwitchClauses -Sum

  foreach ($clause in $SwitchClauseStatistics)
  {
    if ($clause.HasDefault)
    {
      $totalDefaultClauseCount++
    }
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }

  $totalStats = New-Object -TypeName TotalSwitchClauseStatistics
  $totalStats.CodePathTotal = $codePaths.Sum
  $totalStats.SwitchStatementTotal = $switches.Count
  $totalStats.SwitchClauseTotal = $switchClauses.Sum
  $totalStats.DefaultClauseTotal = $totalDefaultClauseCount
  $totalStats.LargestStatementLineCount = $longestStatement

  return $totalStats
}

Function Measure-WhileClauseStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [WhileClauseStatistics[]]$WhileClauseStatistics
  )
  $longestStatement = 0

  $codePaths = $WhileClauseStatistics | Measure-Object -Property CodePaths -Sum
  $whiles = $WhileClauseStatistics | Measure-Object

  foreach ($clause in $WhileClauseStatistics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }

  $totalStats = New-Object -TypeName TotalWhileClauseStatistics
  $totalStats.CodePathTotal = $codePaths.Sum
  $totalStats.WhileStatementTotal = $whiles.Count
  $totalStats.LargestStatementLineCount = $longestStatement

  return $totalStats
}

Function Get-ScriptBlockCommandStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $commands = Find-CommandStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  $stats = New-Object -TypeName TotalCommandStatistics
  if ($null -eq $commands)
  {
    return $stats
  }

  $commandCount = $commands | Measure-Object

  $commandNames = New-Object -TypeName 'Collections.Generic.List[string]'
  foreach ($command in $commands)
  {
    $cmdElement = $command.CommandElements[0]
    if ($cmdElement -is [Management.Automation.Language.StringConstantExpressionAst])
    {
      [void]$commandNames.Add($cmdElement[0].Value)
    }
    elseif ($cmdElement -is [Management.Automation.Language.VariableExpressionAst])
    {
      [void]$commandNames.Add($cmdElement[0].VariablePath.UserPath)
    }
  }
  $uniqueCount = $commandNames | Select-Object -Unique | Measure-Object

  $stats.CommandCount = $commandCount.Count
  $stats.UniqueCommandCount = $uniqueCount.Count
  $stats.CommandNames = $commandNames

  return $stats
}

Function Get-ScriptBlockIfStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'IfClauseStatistics'
  $ifClause = Find-IfStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $ifClause)
  {
    $stats = New-Object -TypeName $statsTypeName
  }
  else
  {
    $stats = New-Object -TypeName "Collections.Generic.List[$statsTypeName]"
    foreach ($clause in $ifClause)
    {
      $statistics = Get-IfClauseStatistics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalStatistics = Measure-IfStatementStatistics -IfClauseStatistics $stats
  return $totalStatistics
}

Function Get-ScriptBlockTryCatchStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'TryClauseStatistics'
  $tryClause = Find-TryStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $tryClause)
  {
    $stats = New-Object -TypeName $statsTypeName
  }
  else
  {
    $stats = New-Object -TypeName "Collections.Generic.List[$statsTypeName]"
    foreach ($clause in $tryClause)
    {
      $statistics = Get-TryCatchClauseStatistics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalStatistics = Measure-TryCatchClauseStatistics -TryClauseStatistics $stats
  return $totalStatistics
}

Function Get-ScriptBlockSwitchStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'SwitchClauseStatistics'
  $switchClause = Find-SwitchStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $switchClause)
  {
    $stats = New-Object -TypeName $statsTypeName
  }
  else
  {
    $stats = New-Object -TypeName "Collections.Generic.List[$statsTypeName]"
    foreach ($clause in $switchClause)
    {
      $statistics = Get-SwitchClauseStatistics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalStatistics = Measure-SwitchClauseStatistics -SwitchClauseStatistics $stats
  return $totalStatistics
}

Function Get-ScriptBlockWhileStatistics
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'WhileClauseStatistics'
  $whileClause = Find-WhileStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $whileClause)
  {
    $stats = New-Object -TypeName $statsTypeName
  }
  else
  {
    $stats = New-Object -TypeName "Collections.Generic.List[$statsTypeName]"
    foreach ($clause in $whileClause)
    {
      $statistics = Get-WhileClauseStatistics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalStatistics = Measure-WhileClauseStatistics -WhileClauseStatistics $stats
  return $totalStatistics
}

Function ConvertTo-ScriptBlock
{
  [CmdletBinding()]
  Param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$FunctionName
  )
  Try
  {
    $command = Get-Command -Name $FunctionName -ErrorAction Stop
    if ($command.CommandType -ne 'Function')
    {
      $emsg = [string]::Format(
        'Function: [{0}] is a [{1}]. Only PowerShell functions are supported',
        $FunctionName,
        $command.CommandType
      )
      Write-Error -Message $emsg -Category InvalidType -ErrorAction Stop
    }
    $functionPath = Join-Path -Path Function: -ChildPath $FunctionName
    $scriptBlock = Get-Content -Path $functionPath -ErrorAction Stop
    return $scriptBlock
  }
  Catch [Management.Automation.RuntimeException]
  {
    $emsg = [string]::Format(
      'Function: [{0}] does not exist or is not imported into the session',
      $FunctionName
    )
    Write-Error -Message $emsg -Exception $_.Exception -Category InvalidArgument
  }
  Catch
  {
    Write-Error -ErrorRecord $_
  }
}


class FunctionStatistics
{
  [string]$Name
  [int]$CodePathTotal
  [int]$MaxNestedDepth
  [int]$LineCount
  [int]$CommandCount
  [TotalIfClauseStatistics]$IfElseStatistics
  [TotalTryClauseStatistics]$TryCatchStatistics
  [TotalSwitchClauseStatistics]$SwitchStatistics
  [TotalWhileClauseStatistics]$WhileStatistics
  [BoolOperatorStatistics]$BoolOperatorStatistics
  [TotalCommandStatistics]$CommandStatistics
}


Function Get-FunctionStatistics
{
  <#
  .SYNOPSIS
  Get statistics about PowerShell functions.

  .DESCRIPTION
  Get statistics about PowerShell functions, including CodePaths, LineCount, etc.

  .PARAMETER FunctionName
  Name of the function to retrieve statistics of.
  The named function cannot be a Cmdlet, and must be imported into the current session

  .PARAMETER ScriptBlock
  A PowerShell script block to retrieve statistics for.

  .EXAMPLE
  $sb = [scriptblock]::Create({if ($true){return $true}else {'Hmmm...'}})
  Get-PSCSFunctionStatistics -ScriptBlock $sb

  Get statistics about the newly created script block.

  .EXAMPLE
  Get-PSCSFunctionStatistics -FunctionName Get-PSCSFunctionStatistics

  Get the function statistics about itself.
  #>
  [CmdletBinding(DefaultParameterSetName='Function')]
  Param(
    [Parameter(Position=0, Mandatory=$true, ParameterSetName='Function')]
    [string]$FunctionName,

    [Parameter(Position=0, Mandatory=$true, ParameterSetName='ScriptBlock')]
    [scriptblock]$ScriptBlock
  )
  Try
  {
    if ($PSCmdlet.ParameterSetName -eq 'Function')
    {
      $ScriptBlock = ConvertTo-ScriptBlock -FunctionName $FunctionName `
        -ErrorAction Stop
    }
    else
    {
      $FunctionName = ''
    }

    $lineCount = $ScriptBlock.Ast.Extent.EndLineNumber -`
      $ScriptBlock.Ast.Extent.StartLineNumber + 1

    $ifStats = Get-ScriptBlockIfStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $tryStats = Get-ScriptBlockTryCatchStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $switchStats = Get-ScriptBlockSwitchStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $whileStats = Get-ScriptBlockWhileStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $operatorStats = Get-BoolOperatorStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $commandStats = Get-ScriptBlockCommandStatistics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $totalCodePaths = $ifStats.CodePathTotal + `
      $tryStats.CodePathTotal + `
      $switchStats.CodePathTotal + `
      $whileStats.CodePathTotal + `
      $operatorStats.CodePaths

    $nestedDepth = Get-MaxNestedDepth -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $functionStats = New-Object -TypeName FunctionStatistics
    $functionStats.Name = $FunctionName
    $functionStats.LineCount = $lineCount
    $functionStats.CodePathTotal = $totalCodePaths
    $functionStats.CommandCount = $commandStats.CommandCount
    $functionStats.MaxNestedDepth = $nestedDepth
    $functionStats.IfElseStatistics = $ifStats
    $functionStats.TryCatchStatistics = $tryStats
    $functionStats.SwitchStatistics = $switchStats
    $functionStats.WhileStatistics = $whileStats
    $functionStats.BoolOperatorStatistics = $operatorStats
    $functionStats.CommandStatistics = $commandStats

    return $functionStats
  }
  Catch
  {
    Write-Error -ErrorRecord $_
  }
}
