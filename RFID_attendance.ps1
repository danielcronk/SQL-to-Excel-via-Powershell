# set up SQL
$SQLServer = "###.###.###.###"
$DBName = "databaseNameHere"
$usr ="usernameHere"
$pwd = "password123"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = "Server = $SQLServer; Database = $DBName; uid = $usr; pwd = $pwd;"
$connection.open()
Write-Host "`nYou need to be in this window before scanning any badges. Make sure the reader beeps when a badge is scanned."
Write-Host "Enter 0 at any time to exit the program - don't just close the window.`n"

# set up Excel
$excel = new-object -ComObject excel.application
$excel.visible = $false
$workbook = $excel.Workbooks.Add()
$ws1 = $workbook.Worksheets.Item("Sheet1")
$rng = $ws1.Range("B2","E2")
$ws1.Cells.Item(2,2) = 'Last Name'
$ws1.Cells.Item(2,3) = 'First Name'
$ws1.Cells.Item(2,4) = 'Bagde ID'
$ws1.Cells.Item(2,5) = 'Associate ID'
$rng.Font.Bold = $True
$rng.Interior.ColorIndex = 33
$usedRange = $ws1.UsedRange	
$usedRange.EntireColumn.AutoFit() | Out-Null

#loop to read badges and write to excel

[Int] $row = 3;
[Int] $col = 2;

DO {

    $badgeRead = Read-Host -Prompt 'Scan Badge or enter 0 to finish logging'

    $query = "select LAST_NAME, FIRST_NAME, BADGEID, EMPLID from tableNameHere where columnNameHere like '$badgeRead'"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $result = $command.ExecuteReader()

    #not sure why these 2 lines are needed but the script doesn't work without them
    $table = new-object System.Data.DataTable
    $table.Load($result)

    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
    $SqlAdapter.SelectCommand = $command;
    $DataSet = New-Object System.Data.DataSet;
    $SqlAdapter.Fill($DataSet);
    $DataSetTable = $DataSet.Tables["Table"];

    [Array] $getColumnNames = $DataSetTable.Columns | Select ColumnName;

    foreach ($rec in $DataSetTable.Rows)
    {
        foreach ($Coln in $getColumnNames)
        {
        # Populating columns:
        $ws1.Cells.Item($row, $col) = `
        $rec.$($Coln.ColumnName).ToString();
        $Col++;
        };
    $row++; $Col = 2;
    };

    # Adjusting columns in the Excel sheet:
    $usedRange = $ws1.UsedRange	
    $usedRange.EntireColumn.AutoFit() | Out-Null

} While($badgeRead -ne 0)


$connection.Close()
$excel.visible = $true
Read-Host -Prompt "`nSave your excel file now before closing. Press Enter to exit or just close the window."
