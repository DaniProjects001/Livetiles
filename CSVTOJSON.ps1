$path = "C:\Users\Manoj LiveTiles\Documents\CoM\COM_IA.csv"
$csv = Import-csv -path $path


#region headers
# Get the row of headers
$Headers = @{}
$NumberOfColumns = 0
$FoundHeaderValue = $true
while ($FoundHeaderValue -eq $true) {
    $cellValue = $theSheet.Cells.Item(1, $NumberOfColumns+1).Text
    if ($cellValue.Trim().Length -eq 0) {
        $FoundHeaderValue = $false
    } else {
        $NumberOfColumns++
        $Headers.$NumberOfColumns = $cellValue
    }
}
#endregion headers

# Count the number of rows in use, ignore the header row
$rowsToIterate = $theSheet.UsedRange.Rows.Count

#region rows
$results = @()
foreach ($rowNumber in 2..$rowsToIterate+1) {
    if ($rowNumber -gt 1) {
        $result = @{}
        foreach ($columnNumber in $Headers.GetEnumerator()) {
            $ColumnName = $columnNumber.Value
            $CellValue = $theSheet.Cells.Item($rowNumber, $columnNumber.Name).Value2
            $result.Add($ColumnName,$cellValue)
        }
        $results += $result
    }
}
#endregion rows


$results | ConvertTo-Json | Out-File -Encoding ASCII -FilePath $OutputFileName