[![codecov](https://codecov.io/gh/briansworth/PSCodeMetrics/branch/master/graph/badge.svg)](https://codecov.io/gh/briansworth/PSCodeMetrics)


## PS Code Metrics

PowerShell module for getting metrics about PowerShell functions
(or script blocks).

## Using the module


### For functions

```powershell
Get-PSCMFunctionMetrics -FunctionName Get-PSCSFunctionMetrics
```

```
Name                    Cc MaxNestedDepth Lloc CommandCount CcMetrics
----                    -- -------------- ---- ------------ ---------
Get-PSCMFunctionMetrics  3              2   59            6 {IfElse: 1..}

```

### For script files

```powershell
$rawScript = Get-Content -Path ./Path_to_script.ps1 -Raw
$scriptBlock = [scriptblock]::Create($rawScript)
Get-PSCMFunctionMetrics -ScriptBlock $scriptBlock
```

