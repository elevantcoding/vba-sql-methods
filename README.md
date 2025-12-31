# vba-sql-methods
VBA that talks directly to SQL Server

Enterprise-grade VBA modules for direct SQL Server communication via ADO.

This repository contains reusable, production-ready VBA components designed for Microsoft Access
front-ends that communicate directly with SQL Server back-ends

## Features
Shared ADO connection and command management  
Clean separation between Access UI and SQL Server logic  
High-performance design for enterprise-scale systems  
Simple, consistent calling patterns  

## Architecture Overview
The library is designed to be embedded into an Access application and
expects the host application to provide the SQL Server connection string.

**Required configuration:**

The host application must define:

```vb
Public Const ADOConnect As String = "<your SQL Server connection string>"
```

## Modules

| Module | Purpose |
|--------|---------|
| `SQLConnection.bas` | Manages shared ADO connection and command lifecycle |
| `SQLMethods.bas` | Core SQL execution and helper routines |
| `Common.bas` | General utility functions |
| `Interaction.bas` | Higher-level interaction helpers |

## Note
Portions of this library were developed with the assistance of AI-based tooling as part of the iterative design and review process.
Core architecture, abstractions, and problem-solving logic — including routines such as SQLInfo, IsValidSQLObject, 
and SearchPredicate — were designed and implemented as part of the original library development.
