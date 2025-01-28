# PostgreSQL connection details
$hostname = "localhost"
$port = "5432"
$dbname = "procedure_db"
$username = "anmol_ta"

# Secure prompt for Password
$password = Read-Host -Prompt "Postgresql Password" -AsSecureString

# Convert Secure String to plain text
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Define the psql query to create table
$tableCreationQuery = @"
CREATE TABLE IF NOT EXISTS Integration (
    SheetName TEXT,
    FieldName TEXT,
    CSVName TEXT,
    Comment TEXT
);
"@

# Set environment variable for password
$env:PGPASSWORD = $plainPassword

# Define the psql command to run query
$psqlCommand = @"
psql -h $hostname -d $dbname -p $port -U $username -c `"$tableCreationQuery`"
"@

# Execute the command to create table 
Invoke-Expression $psqlCommand

# Define the Excel file path
$excelFilePath = "D:\Integration\Automation Parameters.xlsx"

# Read Excel data into PowerShell using Import-Excel (make sure to install the ImportExcel module)
$excelData = Import-Excel -Path $excelFilePath -WorksheetName "Sheet1"

# SQL insert statements ko generate karna
$sqlStatements = @()

foreach ($row in $excelData) {
    # Check if the required fields have valid values (non-empty)
    if (![string]::IsNullOrWhiteSpace($row.'Sheet Name') -and 
        (![string]::IsNullOrWhiteSpace($row.'Field Name')) -and
        (![string]::IsNullOrWhiteSpace($row.'CSV Name'))) {
        
        # Prepare SQL INSERT statement only for valid rows
        $insertStatement = @"
        INSERT INTO integration (sheetname, fieldname, csvname, comment)
        VALUES ('$($row.'Sheet Name')', '$($row.'Field Name')', '$($row.'CSV Name')', '$($row.Comment)');
"@
        $sqlStatements += $insertStatement
    }
}

# Combine all insert statements into one SQL script
$sqlScript = $sqlStatements -join "`n"

# Write SQL script to a file (optional for debugging or re-use)
$sqlFilePath = "D:\Integration\insert_records.sql"
Set-Content -Path $sqlFilePath -Value $sqlScript

# Run SQL script through psql to insert data into the table
$psqlCommand = "psql -h $hostname -p $port -d $dbname -U $username -f `"$sqlFilePath`""
Invoke-Expression $psqlCommand

# Clean up the environment variable
Remove-Item Env:PGPASSWORD
