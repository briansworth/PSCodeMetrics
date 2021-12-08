$moduleName = 'PSCodeMetrics'
Remove-Module -Name $moduleName -ErrorAction SilentlyContinue
Import-Module -Name $moduleName


BeforeAll {
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
      trap {oh no!} # trap no type (+1)
      trap [Exception] {again?!} # trap (+1)
      $varCmd = 'arbitrary_cmd'
      & $varCmd

      try # Try
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
            try
            {
              #Try
            }
            catch # Catch (+1)
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
            break
          }
        }
      }
      catch [Runtime.InteropServices.ExternalException] # Typed Catch (+1)
      {
        # Typed Catch
      }
      catch [Exception] # Typed Catch (+1)
      {
        # Typed Catch
      }
      catch # Catch (+1)
      {
        # Catch all
      }
      finally
      {
        # Finally (+0)
      }
    }
  }
}


Describe Get-PSCMFunctionMetrics {
  It 'Correctly calculates Cc' {
    $expectedCc = 21

    $result = Get-PSCMFunctionMetrics -ScriptBlock $mockFunction

    $result.CcMetrics.Cc | Should -BeExactly $expectedCc
    $result.CcMetrics.Grade | Should -Be 'D'
  }

  It 'Correctly analyzes its own function code' {
    {Get-PSCMFunctionMetrics -FunctionName Get-PSCMFunctionMetrics -ErrorAction Stop} |
    Should -Not -Throw
  }

  It 'Throws exception when function does not exist' {
    $funcName = 'NotExists'
    $expectedMsg = "Function: *$funcName* does not exist*"

    {Get-PSCMFunctionMetrics -FunctionName $funcName -ErrorAction Stop} |
      Should -Throw -ExceptionType ([ArgumentException]) -ExpectedMessage $expectedMsg
  }

  It 'Throws exception when function name is a cmdlet' {
    $funcName = 'Get-ChildItem'
    $expectedMsg = "$funcName is a *Cmdlet*"

    {Get-PSCMFunctionMetrics -FunctionName $funcName -ErrorAction Stop} |
      Should -Throw -ExceptionType ([ArgumentException]) -ExpectedMessage $expectedMsg
  }

  It 'Successfully handles an empty scriptblock' {
    {Get-PSCMFunctionMetrics -ScriptBlock {} -ErrorAction Stop} |
      Should -Not -Throw
  }
}

