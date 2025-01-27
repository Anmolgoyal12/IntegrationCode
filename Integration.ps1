# PostgreSQL connection details
$hostname = "localhost"
$port = "5432"
$dbname = "procedure_db"
$username = "anmol_ta"
$password = "Anmol@1819"

# Define the psql query to create table
$tableCreationQuery = @"
CREATE TABLE IF NOT EXISTS Integration (
    SheetName varchar(100),
    FieldName varchar(100),
    CSVName varchar(100),
    Comment varchar(100)
);
"@

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
    # Prepare SQL INSERT statement
    $insertStatement = @"
    INSERT INTO integration (sheetname, fieldname, csvname, comment)
    VALUES ('$($row.'Sheet Name')', '$($row.'Field Name')', '$($row.'CSV Name')', '$($row.Comment)');
"@
    $sqlStatements += $insertStatement
}

# Combine all insert statements into one SQL script
$sqlScript = $sqlStatements -join "`n"

# Write SQL script to a file (optional for debugging or re-use)
$sqlFilePath = "D:\Integration\insert_records.sql"
Set-Content -Path $sqlFilePath -Value $sqlScript

# Run SQL script through psql to insert data into the table
$psqlCommand = "psql -h $hostname -p $port -d $dbname -U $username -f `"$sqlFilePath`""
Invoke-Expression $psqlCommand
