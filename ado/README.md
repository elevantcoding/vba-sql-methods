# VBA-T-SQL Integration Framework

## 🚀 Key Features

* **ADO Stored Procedure Wrapper:** A robust `SQLCmdSP` function that utilizes `ParamArray` injections to eliminate boilerplate code. Handles complex `adDecimal` types and output parameters with ease.
* **T-SQL Logic Helpers:** Custom `IsIn()` and `IsNotIn()` functions that bring T-SQL syntax convenience to VBA business logic, featuring strict type-safety checks.

## 📋 Usage Example
```vba
' Example of a clean, type-safe Stored Procedure call
SP_Exec = SQLCmdSP("dbo.UpdateLineItem", _
          SPParam("@ID", adInteger, adParamInput, 0, Me.ID), _
          SPParam("@Amt", adDecimal, adParamInput, 0, Me.Amount, 18, 2))
