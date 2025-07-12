---


# 📊 SQL Server Agent Job: Automate Excel/CSV to SQL Import using SSIS

This job uses **SQL Server Agent** to automatically import Excel or CSV files into a SQL Server database using an **SSIS Script Task** or `.dtsx` package.

[![SSIS](https://img.shields.io/badge/SSIS-Automation-green.svg)](#)
[![SQL Server](https://img.shields.io/badge/SQL%20Server-Integration-blue)](#)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE.txt)
[![Maintainer](https://img.shields.io/badge/Maintained%20by-Abhijeet%20Raj-blue)](https://github.com/abhijeetraj22)

---

## 🚀 Objective

Automate file-to-database loading using SSIS and schedule it via SQL Server Agent.

---
## 📦 Folder Structure

```plaintext
├── Integration Services Project1/
│   ├── LoadExcelDataDynaSpreedSheetData.dtsx
│   ├── LoadExcelDataDynaSpreedSheetDupDataInst.dtsx
│   ├── LoadExcelDataInSqlOneStaticTableInsertDupData.dtsx
│   ├── LoadExcelDataInSqlOneTableInsert.dtsx
│   ├── LoadExcelFilesDynamicExcelFileDataInsert.dtsx
│   ├── LoadExcelFilesDynamicExcelFileDataInsertDupData.dtsx
│   ├── LoadFlatFilesDynami.dtsx
│   ├── Project.params
│   └── Integration Services Project1.dtproj
├── .gitattributes
├── .gitignore
├── Integration Services Project1.sln
├── LICENSE.txt
├── README.md
└── SSIS_ConnectionDocument.txt
```
---

## 📁 Components

| Component         | Description                                                  |
|------------------|--------------------------------------------------------------|
| SSIS Script Task | Reads Excel/CSV → dynamically creates table → inserts data   |
| SQL Server Agent | Automates the execution of the SSIS package                  |
| SQL Job Schedule | Runs the job on a recurring basis (daily, hourly, etc.)      |

---

## ⚙️ Required SSIS Variables

| Variable Name          | Description                              |
|------------------------|------------------------------------------|
| `User::SourceFolder`   | Folder containing input files            |
| `User::FileExtension`  | File extension (e.g. `.csv` or `.xlsx`)  |
| `User::ArchiveFolder`  | Path to move processed files             |
| `User::FileDelimiter`  | (For CSV only) delimiter like `,`        |
| `User::ColumnsDataType`| SQL datatype for all columns (`NVARCHAR(MAX)`) |
| `User::SchemaName`     | Target SQL schema                        |
| `User::LogFolder`      | Log path for error handling              |

---
## 📦 Package Descriptions

| SSIS Package Name                                        | Description                                                  |
| -------------------------------------------------------- | ------------------------------------------------------------ |
| **LoadExcelDataDynaSpreedSheetData.dtsx**                | Dynamically creates tables and loads Excel data              |
| **LoadExcelDataDynaSpreedSheetDupDataInst.dtsx**         | Same as above but avoids inserting duplicate records         |
| **LoadExcelDataInSqlOneStaticTableInsert.dtsx**          | Inserts Excel data into a static predefined table            |
| **LoadExcelDataInSqlOneStaticTableInsertDupData.dtsx**   | Same as above, allows duplicates                             |
| **LoadExcelFilesDynamicExcelFileDataInsert.dtsx**        | Loads data from multiple Excel files to corresponding tables |
| **LoadExcelFilesDynamicExcelFileDataInsertDupData.dtsx** | Same as above but skips duplicates                           |
| **LoadFlatFilesDynami.dtsx**                             | Loads data from CSV files with dynamic column creation       |


---

## 🧰 Prerequisites

- SQL Server Agent running
- SSIS package deployed (.dtsx file)
- Connection manager configured (e.g. `DB_Conn_TechBrothersIT`)
- Required drivers installed:
  - For Excel: `Microsoft.ACE.OLEDB.16.0`
  - For CSV: .NET StreamReader logic
---

## 🛠️ Step-by-Step: Create SSIS Script Task in Visual Studio

### ✅ Prerequisites

* Visual Studio 2019 or later with **SSDT** (SQL Server Data Tools) installed
* SQL Server Instance + Database
* Access to folders for source and archive
* Microsoft.ACE.OLEDB.16.0 installed (for Excel)
* Necessary SQL Server login rights

---

### 🔧 Step 1: Create a New SSIS Project

1. Open **Visual Studio**
2. Click on `Create a new project`
3. Search for **Integration Services Project**
4. Select it and click **Next**
5. Name the project (e.g., `ExcelToSQLAutomation`)
6. Click **Create**

---

### 📦 Step 2: Add a New SSIS Package

1. In **Solution Explorer**, right-click on `SSIS Packages` → `Add New SSIS Package`
2. Rename it (e.g., `LoadExcelToSQL.dtsx`)

---

### 📄 Step 3: Add and Configure Variables

1. Open the new package
2. Press `Ctrl + K, V` to open **Variables Pane**
3. Add the following **User:: variables**:

   * `SourceFolder` → `String` → e.g., `C:\SSIS\Input\`
   * `ArchiveFolder` → `String` → e.g., `C:\SSIS\Archive\`
   * `FileExtension` → `String` → `.xlsx` or `.csv`
   * `SchemaName` → `String` → `dbo`
   * `LogFolder` → `String` → e.g., `C:\SSIS\Logs\`
   * `FileDelimiter` → `String` → `,` (for CSVs)
   * `ColumnsDataType` → `String` → `NVARCHAR(MAX)`

---

### 🧱 Step 4: Add Connection Manager

1. Right-click **Connection Managers** → `New OLE DB Connection`
2. Choose your SQL Server, DB, and test the connection
3. Name it (e.g., `DB_Conn_TechBrothersIT`)

---

### 🔁 Step 5: Add a Script Task

1. From **SSIS Toolbox**, drag a **Script Task** onto the Control Flow
2. Double-click the task to configure

---

### ✍️ Step 6: Configure Script Task

#### A. Script Task Editor

* **ReadOnlyVariables**:
  `User::SourceFolder,User::ArchiveFolder,User::FileExtension,User::SchemaName,User::LogFolder,User::FileDelimiter,User::ColumnsDataType`

* **Connections**:
  Select the OLE DB connection (e.g., `DB_Conn_TechBrothersIT`)

* Click **Edit Script**

---

### 💻 Step 7: Write Script Logic

1. Inside `ScriptMain.cs`, paste one of the previously provided scripts:

   * Excel Loader (OLEDB)
   * CSV Loader (FileReader)
2. Save and Close the script editor
3. Click OK to close the Script Task Editor

---

### 📤 Step 8: Execute the Package

1. Right-click on the `LoadExcelToSQL.dtsx` → `Execute Package`
2. Monitor progress via the **Output** and **Execution Results** window
3. Check SQL Server tables and archive folder after execution

---

### 🕓 Step 9: Schedule via SQL Server Agent

After testing, deploy the `.dtsx` or project to file system or SSISDB and follow the SQL Server Agent steps mentioned in the README:

```bash
SSMS → SQL Server Agent → Jobs → New Job → Add Step → Type: SSIS → Package Path
```

---

## 🐛 Troubleshooting Tips

| Purpose               | Tips                                               |
| --------------------- | -------------------------------------------------- |
| **Excel Processing**  | Install ACE OLEDB 16.0 (64-bit)                    |
| **CSV Parsing**       | Use `FileDelimiter` and `StreamReader`             |
| **Error Handling**    | Log errors to `LogFolder/ErrorLog_<timestamp>.log` |
| **Duplication Check** | Add `IF NOT EXISTS` in insert logic                |
| **SQL Performance**   | Use `SqlBulkCopy` (advanced) for large files       |
| **Permissions**       | SQL Agent must access folders used in the script   |

---

Would you like this guide saved as a PDF or Markdown (`.md`) file for your repo?


---

## 🧪 Job Setup (Step-by-Step)

### ✅ Step 1: Enable SQL Server Agent
- Open SSMS
- Right-click `SQL Server Agent` → Start

### ✅ Step 2: Create Job
- In SSMS → Expand `SQL Server Agent` → Right-click `Jobs` → **New Job**
  - **Name:** `Load_Excel_To_SQL_Job`
  - **Owner:** Your login (with Agent rights)
  - **Category:** Optional

### ✅ Step 3: Add Step
- Click **Steps** tab → **New…**
  - **Type:** SQL Server Integration Services Package
  - **Package source:** File System / SSISDB
  - **Package path:** Browse and select your `.dtsx`
  - **Run As:** SQL Agent Service Account

### ✅ Step 4: Add Schedule
- Click **Schedules** tab → **New…**
  - **Name:** `Daily_Load`
  - **Schedule Type:** Recurring
  - **Frequency:** Daily / Hourly / Weekly
  - **Start time:** `02:00 AM`

### ✅ Step 5: Notifications (Optional)
- **Notify** via email, event log, etc. (requires DB mail setup)

---

## 🔄 What the Job Does

1. Picks up `.xlsx` or `.csv` files from input folder
2. Dynamically reads column names from file
3. Creates SQL table with `NVARCHAR(MAX)` fields
4. Inserts records row-by-row
5. Moves processed files to archive folder with timestamp
6. Logs any errors into the log folder

---

## 🐛 Troubleshooting Tips

| Problem | Solution |
|--------|----------|
| Agent not starting | Run SSMS as Admin, check SQL Agent service |
| Cannot access file | Ensure SQL Agent has folder/file permissions |
| Login error (15404) | Use SQL authentication or fix AD permission |
| Excel errors | Ensure ACE.OLEDB driver installed |
| SSIS package not found | Deploy `.dtsx` properly or use `dtexec` |

---

## 🧪 Optional: Run Package from T-SQL

```sql
EXEC xp_cmdshell 'dtexec /f "C:\SSIS\LoadExcelToSQL.dtsx"';
````

---

## 📌 T-SQL Script to Auto-Create Job

```sql
USE msdb;
GO

EXEC sp_add_job @job_name = N'LoadExcelToSQLJob';

EXEC sp_add_jobstep @job_name = N'LoadExcelToSQLJob',
  @step_name = N'Run SSIS',
  @subsystem = N'SSIS',
  @command = N'/File "C:\SSIS\LoadExcelToSQL.dtsx"',
  @database_name = N'master';

EXEC sp_add_schedule @schedule_name = N'DailySchedule',
  @freq_type = 4,  -- Daily
  @active_start_time = 20000;  -- 8 PM

EXEC sp_attach_schedule @job_name = N'LoadExcelToSQLJob',
  @schedule_name = N'DailySchedule';

EXEC sp_add_jobserver @job_name = N'LoadExcelToSQLJob';
```
---
## Connect with me! 🌐

[<img target="_blank" src="https://img.icons8.com/bubbles/100/000000/linkedin.png" title="LinkedIn">](https://www.linkedin.com/in/rajabhijeet22/)       [<img target="_blank" src="https://img.icons8.com/bubbles/100/000000/github.png" title="Github">](https://github.com/abhijeetraj22)     [<img target="_blank" src="https://img.icons8.com/bubbles/100/000000/instagram-new.png" title="Instagram">](https://www.instagram.com/abhijeet_raj_/?hl=en) [<img target="_blank" src="https://img.icons8.com/bubbles/100/000000/twitter-circled.png" title="Twitter">](https://twitter.com/abhijeet_raj_/)
