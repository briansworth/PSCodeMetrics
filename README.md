## PS Code Statistics

PowerShell module for getting statistics about PowerShell functions
(or script blocks).

## Using the module


### For functions

```powershell
Get-PSCSFunctionStatistics -FunctionName Get-PSCSFunctionStatistics
```

```
Name                       CodePathTotal MaxNestedDepth LineCount CommandCount
----                       ------------- -------------- --------- ------------
Get-PSCSFunctionStatistics             2              2        92           10
```

### For script files

```powershell
$rawScript = Get-Content -Path ./Path_to_script.ps1 -Raw
$scriptBlock = [scriptblock]::Create($rawScript)
Get-PSCSFunctionStatistics -ScriptBlock $scriptBlock
```

