
. ".\officefile.ps1"

$excelfiles = @(
    @{'filename' = "a.xlsx"; 'pass' = '123'},
    # "b.xlsx",
    @{'filename' = "b.pptx"; 'readonly' = $false},
    "notfound.xlsx"
)

foreach($file in $excelfiles){
    OpenOfficeFile $file
}

OpenExcelFile -FileName "b.xlsx"

. ".\set-windows-taskbar-small-icon.ps1"

#[PowerShellによるレジストリの操作例 - Qiita](https://qiita.com/mima_ita/items/1e6c74c7fb641852edff)
