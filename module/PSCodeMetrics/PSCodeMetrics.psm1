Set-StrictMode -Version latest

function Get-FunctionMetrics
{
  <#
  .SYNOPSIS
  Get metrics about PowerShell functions.

  .DESCRIPTION
  Get metrics about PowerShell functions, including Cc, LineCount, etc.

  .PARAMETER FunctionName
  Name of the function to retrieve metrics of.
  The named function cannot be a Cmdlet, and must be imported into the current session

  .PARAMETER ScriptBlock
  A PowerShell script block to retrieve metrics for.

  .EXAMPLE
  $sb = [scriptblock]::Create({if ($true){return $true}else {'Hmmm...'}})
  Get-PSCSFunctionMetrics -ScriptBlock $sb

  Get metrics about the newly created script block.

  .EXAMPLE
  Get-PSCSFunctionMetrics -FunctionName Get-PSCSFunctionMetrics

  Get the function metrics about itself.
  #>
  [CmdletBinding(DefaultParameterSetName='Function')]
  param(
    [Parameter(Position=0, Mandatory=$true, ParameterSetName='Function')]
    [string]$FunctionName,

    [Parameter(Position=0, Mandatory=$true, ParameterSetName='ScriptBlock')]
    [scriptblock]$ScriptBlock
  )
  try
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

    $nestedDepth = Get-MaxNestedDepth -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $commandStats = Get-ScriptBlockCommandMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $cCMetrics = Get-ScriptBlockCyclomaticComplexity -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $functionStats = [FunctionMetrics]@{
      'Name' = $FunctionName;
      'LineCount' = $lineCount;
      'Cc' = $cCMetrics.Cc;
      'CcMetrics' = $cCMetrics;
      'CommandCount' = $commandStats.CommandCount;
      'MaxNestedDepth' = $nestedDepth;
      'CommandMetrics' = $commandStats;
    }
    return $functionStats
  }
  catch
  {
    Write-Error -ErrorRecord $_
  }
}

function ConvertTo-ScriptBlock
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$FunctionName
  )
  try
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
  catch [Management.Automation.RuntimeException]
  {
    $emsg = [string]::Format(
      'Function: [{0}] does not exist or is not imported into the session',
      $FunctionName
    )
    Write-Error -Message $emsg -Exception $_.Exception -Category InvalidArgument
  }
  catch
  {
    Write-Error -ErrorRecord $_
  }
}


class FunctionMetrics
{
  [string]$Name
  [int]$Cc
  [CyclomaticComplexityMetrics]$CcMetrics
  [int]$MaxNestedDepth
  [int]$LineCount
  [int]$CommandCount
  [TotalCommandMetrics]$CommandMetrics
}

function Get-MaxNestedDepth
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  try
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
  catch
  {
    Write-Error -ErrorRecord $_
  }
}

function Get-ScriptBlockCommandMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $commands = Find-CommandStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  $stats = New-Object -TypeName TotalCommandMetrics
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

function Get-ScriptBlockCyclomaticComplexity
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
    $ifStats = Get-ScriptBlockIfMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $foreachStats = Get-ScriptBlockForeachMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $tryStats = Get-ScriptBlockTryCatchMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $switchStats = Get-ScriptBlockSwitchMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $whileStats = Get-ScriptBlockWhileMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $operatorStats = Get-BoolOperatorMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $totalCc = $ifStats.Cc + `
      $tryStats.Cc + `
      $switchStats.Cc + `
      $whileStats.Cc + `
      $foreachStats.Cc + `
      $operatorStats.Cc + 1 # Always at least 1 code path

    $cCMetrics = [CyclomaticComplexityMetrics]@{
      'Cc' = $totalCc;
      'IfElse' = $ifStats;
      'TryCatch' = $tryStats;
      'Switch' = $switchStats;
      'While' = $whileStats;
      'Foreach' = $foreachStats;
      'BoolOperator' = $operatorStats;
    }
    return $cCMetrics
}


class CyclomaticComplexityMetrics
{
  [int]$Cc
  [TotalIfClauseMetrics]$IfElse
  [TotalTryClauseMetrics]$TryCatch
  [TotalSwitchClauseMetrics]$Switch
  [TotalWhileClauseMetrics]$While
  [TotalForeachClauseMetrics]$Foreach
  [BoolOperatorMetrics]$BoolOperator
}

function Get-ScriptBlockIfMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'IfClauseMetrics'
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
      $statistics = Get-IfClauseMetrics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalMetrics = Measure-IfStatementMetrics -IfClauseMetrics $stats
  return $totalMetrics
}

function Get-ScriptBlockForeachMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'ForeachClauseMetrics'
  $foreachClause = Find-ForeachStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $foreachClause)
  {
    $stats = New-Object -TypeName $statsTypeName
  }
  else
  {
    $stats = New-Object -TypeName "Collections.Generic.List[$statsTypeName]"
    foreach ($clause in $foreachClause)
    {
      $statistics = Get-ForeachClauseMetrics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalMetrics = Measure-ForeachClauseMetrics -ForeachClauseMetrics $stats
  return $totalMetrics
}

function Get-ScriptBlockTryCatchMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'TryClauseMetrics'
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
      $statistics = Get-TryCatchClauseMetrics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalMetrics = Measure-TryCatchClauseMetrics -TryClauseMetrics $stats
  return $totalMetrics
}

function Get-ScriptBlockSwitchMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'SwitchClauseMetrics'
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
      $statistics = Get-SwitchClauseMetrics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalMetrics = Measure-SwitchClauseMetrics -SwitchClauseMetrics $stats
  return $totalMetrics
}

function Get-ScriptBlockWhileMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $statsTypeName = 'WhileClauseMetrics'
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
      $statistics = Get-WhileClauseMetrics -Clause $clause
      [void]$stats.Add($statistics)
    }
  }

  $totalMetrics = Measure-WhileClauseMetrics -WhileClauseMetrics $stats
  return $totalMetrics
}

function Get-BoolOperatorMetrics
{
  [CmdletBinding()]
  param(
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

  $stats = [BoolOperatorMetrics]@{
    'Cc' = $operatorCount;
    'AndOperators' = $andCount;
    'OrOperators' = $orCount;
    'XOrOperators' = $xorCount;
  }

  return $stats
}


class BoolOperatorMetrics
{
  [int]$Cc
  [int]$AndOperators
  [int]$OrOperators
  [int]$XOrOperators

  [string] ToString()
  {
    return "{Cc = $($this.Cc)...}"
  }
}

function Measure-IfStatementMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [IfClauseMetrics[]]$IfClauseMetrics
  )
  $longestStatement = 0

  $codePaths = $IfClauseMetrics | Measure-Object -Property Cc -Sum
  $ifs = $IfClauseMetrics | Measure-Object -Property IfStatements -Sum
  $elifs = $IfClauseMetrics | Measure-Object -Property ElseIfStatements -Sum
  $elses = $IfClauseMetrics | Measure-Object -Property ElseStatements -Sum

  foreach ($clause in $IfClauseMetrics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }
  $totalStats = [TotalIfClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'IfStatementTotal' = $ifs.Sum;
    'ElseIfStatementTotal' = $elifs.Sum;
    'ElseStatementTotal' = $elses.Sum;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalStats
}

function Measure-ForeachClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [ForeachClauseMetrics[]]$ForeachClauseMetrics
  )
  $longestStatement = 0

  $codePaths = $ForeachClauseMetrics | Measure-Object -Property Cc -Sum
  $foreachs = $ForeachClauseMetrics | Measure-Object
  $variableNames = New-Object -TypeName Collections.Generic.List[string]
  $conditions = New-Object -TypeName Collections.Generic.List[string]

  foreach ($clause in $ForeachClauseMetrics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
    $variableNames.Add($clause.VariableName)
    $conditions.Add($clause.Condition)
  }

  $totalStats = [TotalForeachClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'LargestStatementLineCount' = $longestStatement;
    'ForeachStatementTotal' = $foreachs.Count;
    'VariableNames' = $variableNames;
    'Conditions' = $conditions;
  }
  return $totalStats
}

function Measure-TryCatchClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [TryClauseMetrics[]]$TryClauseMetrics
  )
  $longestStatement = 0
  $totalFinallyCount = 0

  $codePaths = $TryClauseMetrics | Measure-Object -Property Cc -Sum
  $trys = $TryClauseMetrics | Measure-Object
  $catches = $TryClauseMetrics | Measure-Object -Property CatchStatements -Sum
  $catchAlls = $TryClauseMetrics | Measure-Object -Property CatchAllStatements -Sum
  $typedCatches = $TryClauseMetrics | Measure-Object -Property TypedCatchStatements -Sum

  foreach ($clause in $TryClauseMetrics)
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

  $totalStats = [TotalTryClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'TryStatementTotal' = $trys.Count;
    'CatchStatementTotal' = $catches.Sum;
    'CatchAllStatementTotal' = $catchAlls.Sum;
    'TypedCatchStatementTotal' = $typedCatches.Sum;
    'FinallyStatementTotal' = $totalFinallyCount;
    'LargestStatementLineCount' = $longestStatement;
  }

  return $totalStats
}

function Measure-SwitchClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [SwitchClauseMetrics[]]$SwitchClauseMetrics
  )
  $longestStatement = 0
  $totalDefaultClauseCount = 0

  $switches = $SwitchClauseMetrics | Measure-Object
  $codePaths = $SwitchClauseMetrics | Measure-Object -Property Cc -Sum
  $switchClauses = $SwitchClauseMetrics | Measure-Object -Property SwitchClauses -Sum

  foreach ($clause in $SwitchClauseMetrics)
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

  $totalStats = [TotalSwitchClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'SwitchStatementTotal' = $switches.Count;
    'SwitchClauseTotal' = $switchClauses.Sum;
    'DefaultClauseTotal' = $totalDefaultClauseCount;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalStats
}

function Measure-WhileClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [WhileClauseMetrics[]]$WhileClauseMetrics
  )
  $longestStatement = 0

  $codePaths = $WhileClauseMetrics | Measure-Object -Property Cc -Sum
  $whiles = $WhileClauseMetrics | Measure-Object

  foreach ($clause in $WhileClauseMetrics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
  }

  $totalStats = [TotalWhileClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'WhileStatementTotal' = $whiles.Count;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalStats
}


class TotalCommandMetrics
{
  [int]$CommandCount
  [int]$UniqueCommandCount
  [Collections.Generic.List[string]]$CommandNames

  [string] ToString()
  {
    return "{CommandCount = $($this.CommandCount)...}"
  }
}


class TotalClauseMetrics
{
  [int]$Cc
  [int]$LargestStatementLineCount

  [string] ToString()
  {
    return "{Cc = $($this.Cc)...}"
  }
}


class TotalIfClauseMetrics: TotalClauseMetrics
{
  [int]$IfStatementTotal
  [int]$ElseIfStatementTotal
  [int]$ElseStatementTotal
}


class TotalTryClauseMetrics: TotalClauseMetrics
{
  [int]$TryStatementTotal
  [int]$CatchStatementTotal
  [int]$CatchAllStatementTotal
  [int]$TypedCatchStatementTotal
  [int]$FinallyStatementTotal
}


class TotalSwitchClauseMetrics: TotalClauseMetrics
{
  [int]$SwitchStatementTotal
  [int]$SwitchClauseTotal
  [int]$DefaultClauseTotal
}


class TotalWhileClauseMetrics: TotalClauseMetrics
{
  [int]$WhileStatementTotal
}


class TotalForeachClauseMetrics: TotalClauseMetrics
{
  [int]$ForeachStatementTotal
  [string[]]$VariableNames
  [string[]]$Conditions
}

function Get-IfClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.IfStatementAst]$Clause
  )
  $clauseCount = $Clause.Clauses.Count

  $ifStats = [IfClauseMetrics]@{
    'IfStatements' = 1;
    'ElseIfStatements' = $clauseCount - 1;
    'Cc' = $clauseCount;
    'StartLineNumber' = $Clause[0].Extent.StartLineNumber;
  }

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

function Get-ForeachClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.ForeachStatementAst]$Clause
  )
  $foreachStats = [ForeachClauseMetrics]@{
    'ForeachStatements' = 1;
    'Cc' = 1;
    'VariableName' = $Clause.Variable.Extent.Text;
    'Condition' = $Clause.Condition.Extent.Text;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $foreachStats
}

function Get-TryCatchClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.TryStatementAst]$Clause
  )
  $catchAlls = $Clause.CatchClauses | Where-Object {$_.IsCatchAll} | Measure-Object
  $catchCount = $Clause.CatchClauses.Count

  $catchStats = [TryClauseMetrics]@{
    'CatchStatements' = $catchCount;
    'Cc' = $catchCount;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
    'HasFinally' = $false;
    'CatchAllStatements' = $catchAlls.Count;
    'TypedCatchStatements' = $catchCount - $catchAlls.Count;
  }

  if ($null -ne $Clause.Finally)
  {
    $catchStats.HasFinally = $true
  }
  return $catchStats
}

function Get-SwitchClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.SwitchStatementAst]$Clause
  )
  $switchClauses = $Clause.Clauses | Measure-Object

  $switchStats = [SwitchClauseMetrics]@{
    'SwitchStatements' = 1;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
    'HasDefault' = $false;
    'SwitchClauses' = $switchClauses.Count;
    'Cc' = $switchClauses.Count;
  }

  if ($null -ne $Clause.Default)
  {
    $switchStats.HasDefault = $true
  }
  return $switchStats
}

function Get-WhileClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.WhileStatementAst]$Clause
  )
  $whileStats = [WhileClauseMetrics]@{
    'WhileStatements' = 1;
    'Cc' = 1;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $whileStats
}


class ClauseMetrics
{
  [int]$Cc
  [int]$StartLineNumber
  [int]$EndLineNumber

  [int] GetLineCount()
  {
    return $this.EndLineNumber - $this.StartLineNumber + 1
  }

  [string] ToString()
  {
    return "{Cc = $($this.Cc)...}"
  }
}


class IfClauseMetrics: ClauseMetrics
{
  [int]$IfStatements
  [int]$ElseIfStatements
  [int]$ElseStatements
}


class ForeachClauseMetrics: ClauseMetrics
{
  [int]$ForeachStatements
  [string]$VariableName
  [string]$Condition
}


class TryClauseMetrics: ClauseMetrics
{
  [int]$CatchStatements
  [int]$CatchAllStatements
  [int]$TypedCatchStatements
  [bool]$HasFinally
}


class SwitchClauseMetrics: ClauseMetrics
{
  [int]$SwitchStatements
  [int]$SwitchClauses
  [bool]$HasDefault
}


class WhileClauseMetrics: ClauseMetrics
{
  [int]$WhileStatements
}

function Find-CommandStatement
{
  [CmdletBinding()]
  param(
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

function Find-IfStatement
{
  [CmdletBinding()]
  param(
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

function Find-ForeachStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $switchStatement = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.ForeachStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $switchStatement
}

function Find-TryStatement
{
  [CmdletBinding()]
  param(
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

function Find-SwitchStatement
{
  [CmdletBinding()]
  param(
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

function Find-WhileStatement
{
  [CmdletBinding()]
  param(
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

function Get-ScriptBlockToken
{
  [CmdletBinding()]
  param(
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

function New-ClauseMetricsClassInstance
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0)]
    [ValidateSet(
      'ClauseMetrics',
      'IfClauseMetrics',
      'TryCatchClauseMetrics',
      'SwitchClauseMetrics',
      'ForeachClauseMetrics',
      'WhileClauseMetrics'
    )]
    [string]$TypeName = 'ClauseMetrics'
  )
  $classInstance = New-Object -TypeName $TypeName
  return $classInstance
}

