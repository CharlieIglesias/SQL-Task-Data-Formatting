# SQL Server – Task_Data Setup for Team Task Data

## Overview
This SQL script sets up the **Task_Data** SQL Server database and a secure **TeamTaskData** table for storing departmental operational data. It also creates a dedicated SQL login and database user for Excel to connect and read/write data safely. The solution provides a centralised, secure repository that supports downstream Excel reporting for productivity, utilisation, and capacity metrics.

## Key Features & Benefits
- **Centralised database:** Creates a dedicated `Task_Data` database to store all team task data.  
- **Secure data storage:** Implements a SQL login (`ExcelUser`) with read/write permissions, ensuring controlled access.  
- **Structured data table:** `TeamTaskData` table includes columns for `Agent`, `Date`, `Task`, `TaskDescription`, `Pass`, and `Fail`.  
- **Ready for automation:** Supports Excel macros for scheduled reads and writes, facilitating reliable data transfer and reporting.  

## Usage
1. Open **SQL Server Management Studio (SSMS)**.  
2. Add the SQL script below into a new query window.  
3. Execute the script to create the database, table, and Excel user with proper permissions.  
4. Update your Excel VBA connection string to use the `ExcelUser` login and the defined password.  

## Impact
- Eliminates manual setup for a centralised departmental database.  
- Ensures controlled and secure access for Excel-based macros.  
- Provides a robust foundation for automated reporting of productivity, utilisation, and capacity.  
- Supports reliable integration with Excel for downstream calculations and dashboards.
