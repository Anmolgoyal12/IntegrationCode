# PostgreSQL connection details
$hostname = "localhost"
$port = "5432"
$dbname = "procedure_db"
$username = "anmol_ta"

# Prompt for password securely
$password = Read-Host -Prompt "Enter PostgreSQL password" -AsSecureString

# Convert the SecureString password to plain text for psql (necessary for external commands)
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Define the psql query to create table
$tableCreationQuery = @"
CREATE TABLE IF NOT EXISTS Integration (
    id Serial PRIMARY KEY,
    SheetName TEXT,
    FieldName TEXT,
    CSVName TEXT,
    Comment TEXT,
    created_by INTEGER,
    updated_by INTEGER,
    Deleted Boolean DEFAULT FALSE,
    same_env_change Boolean DEFAULT False,
    diff_env_change Boolean DEFAULT False,
    TabeleName TEXT,
    ColumnName TEXT,
    created_dt TIMESTAMP DEFAULT NOW(),
    updated_dt TIMESTAMP
);
"@

# Define the psql command to run query
$psqlCommand = @"
psql -h $hostname -d $dbname -p $port -U $username -w -c `"$tableCreationQuery`"
"@

# Set PGPASSWORD environment variable temporarily (only for the current process)
$env:PGPASSWORD = $plainPassword

# Execute the command to create the table 
Invoke-Expression $psqlCommand

# Clean up the environment variable after use
Remove-Item Env:PGPASSWORD

# Define the Excel file path
$excelFilePath = "D:\Integration\Automation Parameters.xlsx"

# Read Excel data into PowerShell using Import-Excel (make sure to install the ImportExcel module)
$excelData = Import-Excel -Path $excelFilePath -WorksheetName "Sheet1"

# SQL insert statements generation
$sqlStatements = @()

foreach ($row in $excelData) {
    # Prepare SQL INSERT statement with handling for blank values
  $insertStatement = @"
INSERT INTO integration (sheetname, fieldname, csvname, comment, created_by, updated_by, deleted, same_env_change, diff_env_change, tabelename, columnname, created_dt, updated_dt)
VALUES (
    '$($row.'Sheet Name')',
    '$($row.'Field Name')',
    '$($row.'CSV Name')',
    '$($row.Comment)',
    $(if ($row.'created_by') {"'$($row.'created_by')'"} else {'NULL'}),
    $(if ($row.'updated_by') {"'$($row.'updated_by')'"} else {'NULL'}),
    $(if ($row.'Deleted' -eq $true) {'TRUE'} else {'FALSE'}),
    $(if ($row.'same_env_change' -eq $true) {'TRUE'} else {'FALSE'}),
    $(if ($row.'diff_env_change' -eq $true) {'TRUE'} else {'FALSE'}),
    $(if ($row.'TabeleName') {"'$($row.'TabeleName')'"} else {'NULL'}),
    $(if ($row.'ColumnName') {"'$($row.'ColumnName')'"} else {'NULL'}),
    $(if ($row.'created_dt') {"'$($row.'created_dt')'"} else {'NOW()'}),
    $(if ($row.'updated_dt') {"'$($row.'updated_dt')'"} else {'NULL'})
);
"@

    $sqlStatements += $insertStatement
}

# Combine all insert statements into one SQL script
$sqlScript = $sqlStatements -join "`n"

# Write SQL script to a file (optional for debugging or re-use)
$sqlFilePath = "D:\Integration\insert_records.sql"
Set-Content -Path $sqlFilePath -Value $sqlScript

# Set PGPASSWORD environment variable again for running the insert script
$env:PGPASSWORD = $plainPassword

# Run SQL script through psql to insert data into the table
$psqlCommand = "psql -h $hostname -p $port -d $dbname -U $username -f `"$sqlFilePath`""
Invoke-Expression $psqlCommand

# Clean up the environment variable after use
Remove-Item Env:PGPASSWORD
