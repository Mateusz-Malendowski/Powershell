#Powershell script that lists all .xlsx files in the directory and attempts to find data that can be cast to datetime format and overwrites original xlsx files with potentialy updated data.
#Useful when Excel does not recognize date as a date, as after casting using Powershell it will work as intended within Excel.
#ImportExcel seems to save January 1st 1900 as January 2nd 1900 within Excel, hence the hardcoded date within if.

Get-ChildItem -Filter *.xlsx | ForEach-Object  {
    Write-Host "Processing" $_.Name -ForegroundColor Cyan
    $file = Import-Excel -Path $_.Name
    $list = $file[0].PSObject.Properties | ForEach-Object {$_.Name}
    $list = [System.Collections.Generic.List[String]]$list
    for($i=0; $i -lt $file.Count; $i++)
    {
        if ($i -eq 0)
        {
            $tempList = [System.Collections.Generic.List[String]]::new()
            $list | ForEach-Object {$tempList.Add($_)}
        }
        foreach($column in $list)
        {
            try
            {
                if(([datetime]$file[$i].($column)) -eq ([datetime]"01.01.1900")){$file[$i].($column)=([datetime]$file[$i].($column)).AddDays(-1)}
                else{$file[$i].($column)=([datetime]$file[$i].($column))}
            }
            catch { if ($i -eq 0){$tempList.Remove($column) | Out-Null } }
        }
        if ($i -eq 0) {$list = $tempList}
    }
    
    $worksheet = Get-ExcelSheetInfo $_.Name
    $file | Export-Excel -Path $path -AutoFilter -NoNumberConversion * -WorksheetName $worksheet[0].Name
    Write-Host $_.Name "processed successfully" -ForegroundColor Green
}
