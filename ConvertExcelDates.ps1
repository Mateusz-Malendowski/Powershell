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