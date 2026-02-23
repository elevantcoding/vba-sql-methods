# VBA-T-SQL Integration Framework
### A High-Performance Bridge for MS Access & SQL Server

This repository contains a suite of professional VBA modules designed to transform MS Access into a high-performance "Thin Client" for SQL Server. Instead of relying on standard linked tables and the ACE/Jet engine, these tools utilize direct ADO communication and T-SQL-inspired logic.

## 🚀 Key Features

* **ADO Stored Procedure Wrapper:** A robust `SQLCmdSP` function that utilizes `ParamArray` injections to eliminate boilerplate code. Handles complex `adDecimal` types and output parameters with ease.
* **T-SQL Logic Helpers:** Custom `IsIn()` and `IsNotIn()` functions that bring T-SQL syntax convenience to VBA business logic, featuring strict type-safety checks.
* **Transactional Integrity:** Direct SQL DML execution logic that programmatically audits `RecordsAffected` to ensure atomic updates and prevent unintended bulk changes.
* **Deep-Object Auditing:** The `UtilizationV1` utility for scanning hidden dependencies within Form RecordSources and Control RowSources—areas often missed by standard Access tools.

## 🛠 Why this exists
Access SQL is limited. By moving complex logic into SQL Server Stored Procedures and using this framework to bridge the gap, developers can leverage:
* Window Functions (OVER/PARTITION)
* Clean, non-nested JOIN syntax
* Superior execution plans and server-side performance

## 📋 Usage Example
```vba
' Example of a clean, type-safe Stored Procedure call
SP_Exec = SQLCmdSP("dbo.UpdateLineItem", _
          SPParam("@ID", adInteger, adParamInput, 0, Me.ID), _
          SPParam("@Amt", adDecimal, adParamInput, 0, Me.Amount, 18, 2))
