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
CREATE TABLE IF NOT EXISTS scripts_parameter
(
scripts_parameter_id SERIAL PRIMARY KEY,
param_group_name VARCHAR(255) not null,
sheet_name VARCHAR(255) not null,
name VARCHAR(255) not null,
display_name VARCHAR(255) not null,
datatype VARCHAR(255) not null,
csv_name VARCHAR(255),
description TEXT,
updated_dt timestamp with time zone ,
created_dt timestamp with time zone default now(),
created_by VARCHAR(255),
updated_by VARCHAR(255),
same_env_change	BOOLEAN,
diff_env_change	BOOLEAN,
table_name VARCHAR(255),
column_name VARCHAR(255),
deleted BOOLEAN default false
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
INSERT INTO scripts_parameter (param_group_name,sheet_name, name,display_name,datatype,
description, updated_dt,created_dt,created_by, updated_by, 
same_env_change, diff_env_change, table_name, column_name,deleted)
VALUES (
    '$($row.'param_group_name')',
    '$($row.'sheet_name')',
    '$($row.'field_name')',
    '$($row.'field_name')',
    '$($row.'datatype')',
    '$($row.comment)',
    $(if ($row.'updated_dt') {"'$($row.'updated_dt')'"} else {'NULL'}),
     $(if ($row.'created_dt') {"'$($row.'created_dt')'"} else {'NOW()'}),
    $(if ($row.'created_by') {"'$($row.'created_by')'"} else {'NULL'}),
$(if ($row.'updated_by') {"'$($row.'updated_by')'"} else {'NULL'}),
    $(if ($row.'same_env_change' -eq $true) {'TRUE'} else {'FALSE'}),
    $(if ($row.'diff_env_change' -eq $true) {'TRUE'} else {'FALSE'}),
    $(if ($row.'table_name') {"'$($row.'table_name')'"} else {'NULL'}),
    $(if ($row.'column_name') {"'$($row.'column_name')'"} else {'NULL'}),
     $(if ($row.'deleted' -eq $true) {'TRUE'} else {'FALSE'})
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
