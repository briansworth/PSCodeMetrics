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

    $commandMetrics = Get-ScriptBlockCommandMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $cCMetrics = Get-ScriptBlockCyclomaticComplexity -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $functionMetric = [FunctionMetrics]@{
      'Name' = $FunctionName;
      'LineCount' = $lineCount;
      'Cc' = $cCMetrics.Cc;
      'CcMetrics' = $cCMetrics;
      'CommandCount' = $commandMetrics.CommandCount;
      'MaxNestedDepth' = $nestedDepth;
      'CommandMetrics' = $commandMetrics;
    }
    return $functionMetric
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

  $metrics = New-Object -TypeName TotalCommandMetrics
  if ($null -eq $commands)
  {
    return $metrics
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

  $metrics.CommandCount = $commandCount.Count
  $metrics.UniqueCommandCount = $uniqueCount.Count
  $metrics.CommandNames = $commandNames

  return $metrics
}

function Get-ScriptBlockCyclomaticComplexity
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
    $ifMetric = Get-ScriptBlockIfMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $foreachMetric = Get-ScriptBlockForeachMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $forMetric = Get-ScriptBlockForMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $tryMetric = Get-ScriptBlockTryCatchMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $switchMetric = Get-ScriptBlockSwitchMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $whileMetric = Get-ScriptBlockWhileMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop
    $operatorMetric = Get-BoolOperatorMetrics -ScriptBlock $ScriptBlock `
      -ErrorAction Stop

    $totalCc = $ifMetric.Cc + `
      $foreachMetric.Cc + `
      $forMetric.Cc + `
      $tryMetric.Cc + `
      $switchMetric.Cc + `
      $whileMetric.Cc + `
      $operatorMetric.Cc + 1 # Always at least 1 code path

    $cCMetrics = [CyclomaticComplexityMetrics]@{
      'Cc' = $totalCc;
      'IfElse' = $ifMetric;
      'Foreach' = $foreachMetric;
      'For' = $forMetric;
      'TryCatch' = $tryMetric;
      'Switch' = $switchMetric;
      'While' = $whileMetric;
      'BoolOperator' = $operatorMetric;
    }
    return $cCMetrics
}


class CyclomaticComplexityMetrics
{
  [int]$Cc
  [TotalIfClauseMetrics]$IfElse
  [TotalForeachClauseMetrics]$Foreach
  [TotalForClauseMetrics]$For
  [TotalTryClauseMetrics]$TryCatch
  [TotalSwitchClauseMetrics]$Switch
  [TotalWhileClauseMetrics]$While
  [BoolOperatorMetrics]$BoolOperator
}

function Get-ScriptBlockIfMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'IfClauseMetrics'
  $ifClause = Find-IfStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $ifClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $ifClause)
    {
      $metric = Get-IfClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-IfStatementMetrics -IfClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockForeachMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'ForeachClauseMetrics'
  $foreachClause = Find-ForeachStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $foreachClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $foreachClause)
    {
      $metric = Get-ForeachClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-ForeachClauseMetrics -ForeachClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockForMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'ForClauseMetrics'
  $clauses = Find-ForStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $clauses)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $clauses)
    {
      $metric = Get-ForClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-ForClauseMetrics -ForClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockTryCatchMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'TryClauseMetrics'
  $tryClause = Find-TryStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $tryClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $tryClause)
    {
      $metric = Get-TryCatchClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-TryCatchClauseMetrics -TryClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockSwitchMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'SwitchClauseMetrics'
  $switchClause = Find-SwitchStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $switchClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $switchClause)
    {
      $metric = Get-SwitchClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-SwitchClauseMetrics -SwitchClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockWhileMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'WhileClauseMetrics'
  $whileClause = Find-WhileStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $whileClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $whileClause)
    {
      $metric = Get-WhileClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-WhileClauseMetrics -WhileClauseMetrics $metrics
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

  $metrics = [BoolOperatorMetrics]@{
    'Cc' = $operatorCount;
    'AndOperators' = $andCount;
    'OrOperators' = $orCount;
    'XOrOperators' = $xorCount;
  }

  return $metrics
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
  $totalMetrics = [TotalIfClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'IfStatementTotal' = $ifs.Sum;
    'ElseIfStatementTotal' = $elifs.Sum;
    'ElseStatementTotal' = $elses.Sum;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalMetrics
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

  $totalMetrics = [TotalForeachClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'LargestStatementLineCount' = $longestStatement;
    'ForeachStatementTotal' = $foreachs.Count;
    'VariableNames' = $variableNames;
    'Conditions' = $conditions;
  }
  return $totalMetrics
}

function Measure-ForClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [ForClauseMetrics[]]$ForClauseMetrics
  )
  $longestStatement = 0

  $codePaths = $ForClauseMetrics | Measure-Object -Property Cc -Sum
  $fors = $ForClauseMetrics | Measure-Object
  $initializers = New-Object -TypeName Collections.Generic.List[string]
  $iterators = New-Object -TypeName Collections.Generic.List[string]
  $conditions = New-Object -TypeName Collections.Generic.List[string]

  foreach ($clause in $ForClauseMetrics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
    $initializers.Add($clause.Initializer)
    $iterators.Add($clause.Iterator)
    $conditions.Add($clause.Condition)
  }

  $totalMetrics = [TotalForClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'LargestStatementLineCount' = $longestStatement;
    'ForStatementTotal' = $fors.Count;
    'Initializers' = $initializers;
    'Iterators' = $iterators;
    'Conditions' = $conditions;
  }
  return $totalMetrics
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

  $totalMetrics = [TotalTryClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'TryStatementTotal' = $trys.Count;
    'CatchStatementTotal' = $catches.Sum;
    'CatchAllStatementTotal' = $catchAlls.Sum;
    'TypedCatchStatementTotal' = $typedCatches.Sum;
    'FinallyStatementTotal' = $totalFinallyCount;
    'LargestStatementLineCount' = $longestStatement;
  }

  return $totalMetrics
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

  $totalMetrics = [TotalSwitchClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'SwitchStatementTotal' = $switches.Count;
    'SwitchClauseTotal' = $switchClauses.Sum;
    'DefaultClauseTotal' = $totalDefaultClauseCount;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalMetrics
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

  $totalMetrics = [TotalWhileClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'WhileStatementTotal' = $whiles.Count;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalMetrics
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


class TotalForeachClauseMetrics: TotalClauseMetrics
{
  [int]$ForeachStatementTotal
  [string[]]$VariableNames
  [string[]]$Conditions
}


class TotalForClauseMetrics: TotalClauseMetrics
{
  [int]$ForStatementTotal
  [string[]]$Initializers
  [string[]]$Iterators
  [string[]]$Conditions
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

function Get-IfClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.IfStatementAst]$Clause
  )
  $clauseCount = $Clause.Clauses.Count

  $ifMetrics = [IfClauseMetrics]@{
    'IfStatements' = 1;
    'ElseIfStatements' = $clauseCount - 1;
    'Cc' = $clauseCount;
    'StartLineNumber' = $Clause[0].Extent.StartLineNumber;
  }

  if ($null -eq $Clause.ElseClause)
  {
    $ifMetrics.EndLineNumber = $Clause[-1].Extent.EndLineNumber
  }
  else
  {
    $ifMetrics.ElseStatements = 1
    $ifMetrics.EndLineNumber = $Clause.ElseClause.Extent.EndLineNumber
  }
  return $ifMetrics
}

function Get-ForeachClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.ForeachStatementAst]$Clause
  )
  $metrics = [ForeachClauseMetrics]@{
    'ForeachStatements' = 1;
    'Cc' = 1;
    'VariableName' = $Clause.Variable.Extent.Text;
    'Condition' = $Clause.Condition.Extent.Text;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $metrics
}

function Get-ForClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.ForStatementAst]$Clause
  )
  $cc = 0
  $condition = [string]::Empty
  $initializer = [string]::Empty
  $iterator = [string]::Empty

  if ($null -ne $Clause.Condition)
  {
    $cc = 1
    $condition = $Clause.Condition.Extent.Text
  }
  if ($null -ne $Clause.Iterator)
  {
    $iterator = $Clause.Iterator.Extent.Text
  }
  if ($null -ne $Clause.Initializer)
  {
    $initializer = $Clause.Initializer.Extent.Text
  }

  $metrics = [ForClauseMetrics]@{
    'ForStatements' = 1;
    'Cc' = $cc;
    'Initializer' = $initializer;
    'Iterator' = $iterator;
    'Condition' = $condition;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $metrics
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

  $metrics = [TryClauseMetrics]@{
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
    $metrics.HasFinally = $true
  }
  return $metrics
}

function Get-SwitchClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.SwitchStatementAst]$Clause
  )
  $switchClauses = $Clause.Clauses | Measure-Object

  $metrics = [SwitchClauseMetrics]@{
    'SwitchStatements' = 1;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
    'HasDefault' = $false;
    'SwitchClauses' = $switchClauses.Count;
    'Cc' = $switchClauses.Count;
  }

  if ($null -ne $Clause.Default)
  {
    $metrics.HasDefault = $true
  }
  return $metrics
}

function Get-WhileClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.WhileStatementAst]$Clause
  )
  $metrics = [WhileClauseMetrics]@{
    'WhileStatements' = 1;
    'Cc' = 1;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $metrics
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


class ForClauseMetrics: ClauseMetrics
{
  [int]$ForStatements
  [string]$Iterator
  [string]$Initializer
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
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.CommandAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-IfStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.IfStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-ForeachStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.ForeachStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-ForStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.ForStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-TryStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.TryStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-SwitchStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.SwitchStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
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

