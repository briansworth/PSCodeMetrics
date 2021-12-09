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

    $nestedDepth = Get-MaxNestedDepth -ScriptBlock $ScriptBlock
    $locMetrics = Get-ScriptBlockLocMetrics -ScriptBlock $ScriptBlock
    $commandMetrics = Get-ScriptBlockCommandMetrics -ScriptBlock $ScriptBlock
    $cCMetrics = Get-ScriptBlockCyclomaticComplexity -ScriptBlock $ScriptBlock

    $functionMetric = [FunctionMetrics]@{
      'Name' = $FunctionName;
      'CcMetrics' = $cCMetrics;
      'LocMetrics' = $locMetrics;
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
        '{0} is a [{1}]. Only PowerShell functions are supported',
        $FunctionName,
        $command.CommandType
      )
      $exception = [ArgumentException]::new($emsg)
      Write-Error -Exception $exception -Category InvalidArgument -ErrorAction Stop
    }
    $functionPath = Join-Path -Path Function: -ChildPath $FunctionName
    $scriptBlock = Get-Content -Path $functionPath -ErrorAction Stop
    return $scriptBlock
  }
  catch [ArgumentException]
  {
    Write-Error -ErrorRecord $_
  }
  catch
  {
    $emsg = [string]::Format(
      'Function: [{0}] does not exist or is not imported into the session',
      $FunctionName
    )
    $exception = [ArgumentException]::new($emsg, $_.Exception)
    Write-Error -Exception $exception -Category InvalidArgument
  }
}


class FunctionMetrics
{
  [string]$Name
  [int]$MaxNestedDepth
  [CyclomaticComplexityMetrics]$CcMetrics
  [LocMetrics]$LocMetrics
  [TotalCommandMetrics]$CommandMetrics
}

function Get-MaxNestedDepth
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
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

function Get-ScriptBlockLocMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )

  $sloc = $ScriptBlock.Ast.Extent.EndLineNumber - `
    $ScriptBlock.Ast.Extent.StartLineNumber + 1
  $comments = Get-ScriptBlockCommentMetrics -ScriptBlock $ScriptBlock
  $emptyLines = Get-ScriptBlockEmptyLineMetrics -ScriptBlock $ScriptBlock
  $overlap = Measure-CommentEmptyLines -Comments $comments

  $lloc = $sloc - $comments.CommentLineCount - $emptyLines.TotalEmptyLines + $overlap.Count

  $metrics = [LocMetrics]@{
    'Lloc' = $lloc;
    'Sloc' = $sloc;
    'EmptyLines' = $emptyLines;
    'Comments' = $comments;
  }
  return $metrics
}


class LocMetrics
{
  [int]$Lloc
  [int]$Sloc
  [TotalEmptyLineMetrics]$EmptyLines
  [TotalCommentMetrics]$Comments
  
  [string] ToString()
  {
    return "{Sloc: $($this.Sloc)...}"
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
  $ifMetric = Get-ScriptBlockIfMetrics -ScriptBlock $ScriptBlock
  $foreachMetric = Get-ScriptBlockForeachMetrics -ScriptBlock $ScriptBlock
  $forMetric = Get-ScriptBlockForMetrics -ScriptBlock $ScriptBlock
  $tryMetric = Get-ScriptBlockTryCatchMetrics -ScriptBlock $ScriptBlock
  $switchMetric = Get-ScriptBlockSwitchMetrics -ScriptBlock $ScriptBlock
  $whileMetric = Get-ScriptBlockWhileMetrics -ScriptBlock $ScriptBlock
  $trapMetric = Get-ScriptBlockTrapMetrics -ScriptBlock $ScriptBlock
  $operatorMetric = Get-BoolOperatorMetrics -ScriptBlock $ScriptBlock

  $totalCc = $ifMetric.Cc + `
    $foreachMetric.Cc + `
    $forMetric.Cc + `
    $tryMetric.Cc + `
    $switchMetric.Cc + `
    $whileMetric.Cc + `
    $trapMetric.Cc + `
    $operatorMetric.Cc + 1 # Always at least 1 code path
  $grade = Get-CyclomaticComplexityGrade -Cc $totalCc

  $cCMetrics = [CyclomaticComplexityMetrics]@{
    'Grade' = $grade;
    'Cc' = $totalCc
    'IfElse' = $ifMetric;
    'Foreach' = $foreachMetric;
    'For' = $forMetric;
    'TryCatch' = $tryMetric;
    'Switch' = $switchMetric;
    'While' = $whileMetric;
    'Trap' = $trapMetric;
    'BoolOperator' = $operatorMetric;
  }
  return $cCMetrics
}

function Get-CyclomaticComplexityGrade
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [int]$Cc
  )
  if ($Cc -le 5)
  {
    $grade = [CcGrade]::A
  }
  elseif ($Cc -le 10)
  {
    $grade = [CcGrade]::B
  }
  elseif ($Cc -le 20)
  {
    $grade = [CcGrade]::C
  }
  elseif ($Cc -le 30)
  {
    $grade = [CcGrade]::D
  }
  elseif ($Cc -le 40)
  {
    $grade = [CcGrade]::E
  }
  else
  {
    $grade = [CcGrade]::F
  }
  return $grade
}


enum CcGrade
{
  A
  B
  C
  D
  E
  F
}


class CyclomaticComplexityMetrics
{
  [int]$Cc
  [CcGrade]$Grade
  [TotalIfClauseMetrics]$IfElse
  [TotalForeachClauseMetrics]$Foreach
  [TotalForClauseMetrics]$For
  [TotalTryClauseMetrics]$TryCatch
  [TotalSwitchClauseMetrics]$Switch
  [TotalWhileClauseMetrics]$While
  [TotalTrapClauseMetrics]$Trap
  [BoolOperatorMetrics]$BoolOperator

  [string] ToString()
  {
    return "{Cc: $($this.Cc)...}"
  }
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

function Get-ScriptBlockTrapMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $typeName = 'TrapClauseMetrics'
  $trapClause = Find-TrapStatement -ScriptBlock $ScriptBlock `
    -IncludeNestedClause

  if ($null -eq $trapClause)
  {
    $metrics = New-Object -TypeName $typeName
  }
  else
  {
    $metrics = New-Object -TypeName "Collections.Generic.List[$typeName]"
    foreach ($clause in $trapClause)
    {
      $metric = Get-TrapClauseMetrics -Clause $clause
      [void]$metrics.Add($metric)
    }
  }

  $totalMetrics = Measure-TrapClauseMetrics -TrapClauseMetrics $metrics
  return $totalMetrics
}

function Get-ScriptBlockCommentMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $details = Get-CommentDetailMetrics -ScriptBlock $ScriptBlock
  $totalMetrics = Measure-CommentMetrics -CommentDetails $details
  return $totalMetrics
}

function Get-ScriptBlockEmptyLineMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $emptyLines = Get-EmptyLineMetric -ScriptBlock $ScriptBlock
  $totalMetrics = Measure-EmptyLineMetrics -EmptyLineMetrics $emptyLines
  return $totalMetrics
}

function Measure-CommentEmptyLines
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [TotalCommentMetrics]$Comments
  )
  $overlapCounter = 0
  $overlapLines = New-Object -TypeName Collections.Generic.List[string]
  $blockComments = $Comments.Details | Where-Object {$_.Type -eq 'Block'}
  foreach ($comment in $blockComments)
  {
    $commentSb = [scriptblock]::Create($comment.Token.Extent.Text)
    $emptyLines = Get-ScriptBlockEmptyLineMetrics -ScriptBlock $commentSb
    $overlapCounter += $emptyLines.TotalEmptyLines
  }
  $output = [PSObject]@{
    'Count' = $overlapCounter;
  }
  return $output
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
    return "{Cc: $($this.Cc)...}"
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

function Measure-TrapClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [TrapClauseMetrics[]]$TrapClauseMetrics
  )
  $longestStatement = 0

  $codePaths = $TrapClauseMetrics | Measure-Object -Property Cc -Sum
  $traps = $TrapClauseMetrics | Measure-Object
  $trapTypes = New-Object -TypeName Collections.Generic.List[string]

  foreach ($clause in $TrapClauseMetrics)
  {
    $lineCount = $clause.GetLineCount()
    if ($lineCount -gt $longestStatement)
    {
      $longestStatement = $lineCount
    }
    $trapTypes.Add($clause.Type)
  }

  $totalMetrics = [TotalTrapClauseMetrics]@{
    'Cc' = $codePaths.Sum;
    'TrapStatementTotal' = $traps.Count;
    'TrapTypes' = $trapTypes;
    'LargestStatementLineCount' = $longestStatement;
  }
  return $totalMetrics
}

function Measure-CommentMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [AllowEmptyCollection()]
    [CommentDetails[]]$CommentDetails
  )
  $metrics = [TotalCommentMetrics]@{
    'TotalComments' = ($CommentDetails | Measure-Object).Count;
    'Details' = $CommentDetails;
  }
  $lineCount = 0

  foreach ($comment in $CommentDetails)
  {
    if ($comment.Type -ne [CommentType]::Appended)
    {
      $lineCount += $comment.Context.Length
    }
  }
  $metrics.CommentLineCount = $lineCount
  return $metrics
}

function Measure-EmptyLineMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [AllowEmptyCollection()]
    [EmptyLineMetrics[]]$EmptyLineMetrics
  )
  $count = ($EmptyLineMetrics | Measure-Object).Count
  $metrics = [TotalEmptyLineMetrics]@{
    'TotalEmptyLines' = $count;
    'EmptyLines' = $EmptyLineMetrics;
  }
  return $metrics
}


class TotalEmptyLineMetrics
{
  [int]$TotalEmptyLines
  [EmptyLineMetrics[]]$EmptyLines

  [string] ToString()
  {
    return "{EmptyLines: $($this.TotalEmptyLines)...}"
  }
}


class TotalCommentMetrics
{
  [int]$TotalComments
  [int]$CommentLineCount
  [CommentDetails[]]$Details

  [string] ToString()
  {
    return "{Comments: $($this.TotalComments)...}"
  }
}


class TotalCommandMetrics
{
  [int]$CommandCount
  [int]$UniqueCommandCount
  [Collections.Generic.List[string]]$CommandNames

  [string] ToString()
  {
    return "{CommandCount: $($this.CommandCount)...}"
  }
}


class TotalClauseMetrics
{
  [int]$Cc
  [int]$LargestStatementLineCount

  [string] ToString()
  {
    return "{Cc: $($this.Cc)...}"
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

class TotalTrapClauseMetrics: TotalClauseMetrics
{
  [int]$TrapStatementTotal
  [string[]]$TrapTypes
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

function Get-TrapClauseMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [Management.Automation.Language.TrapStatementAst]$Clause
  )
  $trapType = [string]::Empty
  if ($null -ne $Clause.TrapType)
  {
    $trapType = $Clause.TrapType.ToString()
  }

  $metrics = [TrapClauseMetrics]@{
    'TrapStatements' = 1;
    'Cc' = 1;
    'Type' = $trapType;
    'StartLineNumber' = $Clause.Extent.StartLineNumber;
    'EndLineNumber' = $Clause.Extent.EndLineNumber;
  }
  return $metrics
}

function Get-CommentDetailMetrics
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $tokenComments = Get-ScriptBlockToken -ScriptBlock $ScriptBlock -TokenKind Comment
  $comments = New-Object -TypeName Collections.Generic.List[CommentDetails]
  foreach ($comment in $tokenComments)
  {
    $detail = Get-CommentDetails -ScriptBlock $ScriptBlock -CommentToken $comment
    $comments.Add($detail)
  }
  return ,$comments
}

function Get-EmptyLineMetric
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock
  )
  $endLine = $ScriptBlock.Ast.Extent.EndLineNumber
  $startLine = $ScriptBlock.Ast.Extent.StartLineNumber
  $range = [Range]::new($startLine, $endLine)

  Write-Debug -Message "Get EmptyLine context. Range: [$range]"
  $context = Get-ScriptBlockContext -ScriptBlock $ScriptBlock -LineRange $range
  $metrics = New-Object -TypeName Collections.Generic.List[EmptyLineMetrics]
  for ($i = 0; $i -lt $context.Context.Length; $i++)
  {
    $line = $context.Context[$i]
    $isEmpty = Test-IsEmptyLine -Text $context.Context[$i]
    if ($isEmpty.Result)
    {
      $metric = [EmptyLineMetrics]@{
        'LineNumber' = $i + $context.Offset;
        'HasWhitespace' = $isEmpty.HasWhitespace;
      }
      $metrics.Add($metric)
    }
  }
  return ,$metrics
}

function Test-IsEmptyLine
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [AllowEmptyString()]
    [string]$Text
  )
  $hasWhitespace = $false
  if ($Text -eq [string]::Empty)
  {
    $isEmpty = $true
  }
  elseif ($Text -match '^\s+$')
  {
    $isEmpty = $true
    $hasWhitespace = $true
  }
  else
  {
    $isEmpty = $false
  }
  $output = New-Object -TypeName PSObject -Property @{
    'Result' = $isEmpty;
    'HasWhitespace' = $hasWhitespace;
  }
  return $output
}


class CommentDetails
{
  [CommentType]$Type
  [Management.Automation.Language.Token]$Token
  [string[]]$Context
}


enum CommentType
{
  Default
  Appended
  Block
}


class EmptyLineMetrics
{
  [int]$LineNumber
  [bool]$HasWhitespace
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


class TrapClauseMetrics: ClauseMetrics
{
  [int]$TrapStatements
  [string]$Type
}

function Get-CommentDetails
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [Parameter(Position=1, Mandatory=$true)]
    [Management.Automation.Language.Token]$CommentToken
  )
  $sbStartLine = $ScriptBlock.Ast.Extent.StartLineNumber

  $startLine = $CommentToken.Extent.StartLineNumber + $sbStartLine
  $endLine = $CommentToken.Extent.EndLineNumber + $sbStartLine
  $range = [Range]::new($startLine, $endLine)
  
  Write-Debug -Message "Get comment context. Range: [$range]"
  $context = Get-ScriptBlockContext -ScriptBlock $ScriptBlock -LineRange $range
  $commentType = Resolve-CommentType -Context $context.Context
  $details = [CommentDetails]@{
    'Type' = $commentType;
    'Token' = $CommentToken;
    'Context' = $context.Context;
  }
  return $details
}

function Get-ScriptBlockContext
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [Parameter(Position=1, Mandatory=$true)]
    [range]$LineRange
  )
  $offset = $ScriptBlock.Ast.Extent.StartScriptPosition.LineNumber
  Write-Debug -Message "ScriptBlock Offset: [$offset]"
  Write-Debug -Message "Target LineRange: [$LineRange]"
  $startIndex = $LineRange.Start.Value - $offset
  $endIndex = $LineRange.End.Value - $offset

  $scriptLines = $ScriptBlock.Ast.Extent.Text.Split("`n")
  Write-Debug -Message "Index range: [$startIndex..$endIndex]"
  Write-Debug -Message "ScriptBlock length: [$($scriptLines.Length)]"
  $context = $scriptLines["$startIndex".."$endIndex"]
  $output = New-Object -TypeName PSObject -Property @{
    'Context' = $context;
    'Offset' = $offset;
  }
  return $output
}

function Resolve-CommentType
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [AllowEmptyString()]
    [string[]]$Context
  )
  if ($Context.Count -gt 1)
  {
    $type = [CommentType]::Block
  }
  elseif ($Context -notmatch '^\s*#.*')
  {
    $type = [CommentType]::Appended
  }
  else
  {
    $type = [CommentType]::Default
  }
  return $type
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
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.WhileStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
}

function Find-TrapStatement
{
  [CmdletBinding()]
  param(
    [Parameter(Position=0, Mandatory=$true)]
    [scriptblock]$ScriptBlock,

    [switch]$IncludeNestedClause
  )
  $ast = $ScriptBlock.Ast.FindAll(
    {$args[0] -is [Management.Automation.Language.TrapStatementAst]},
    $IncludeNestedClause.ToBool()
  )
  return $ast
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

