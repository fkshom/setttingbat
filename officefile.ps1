function OpenOfficeFile {
    Param( 
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [object] $file
    )
    $fileobj = @{}
    if($file.GetType().Name -eq 'HashTable'){
        Write-Host $file['filename']
        $fileobj = $file
    }
    elseif($file.GetType().Name -eq 'String'){
        Write-Host $file
        $fileobj['filename'] = $file
    }
    else{
        Write-Error -ErrorAction Continue -Message "filename type is invalid"
        return
    }
    if( -not (Test-Path $fileobj['filename'])) {
        Write-Host $fileobj['filename'] is not found
        return
    }

    $ext = [System.IO.Path]::GetExtension($fileobj['filename'])
    if($ext -eq '.xlsx'){
        OpenExcelFile -FileName $fileobj['filename'] -Password $fileobj['pass']
    }
    elseif($ext -eq '.pptx'){
        OpenPowerPointFile -FileName $fileobj['filename'] -ReadOnly $fileobj['readonly']
    }
}

function OpenExcelFile {
    Param( 
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string] $FileName,
        [string] $Password = $null,
        [ValidateSet($null, $true, $false)]
        [object] $ReadOnly = $null
    )

    if( -not (Test-Path $FileName)) {
        Write-Host $FileName is not found
        return
    }

    [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")
    $xlWindowState = [Microsoft.Office.Interop.Excel.XlWindowState]
    $xlMaximized = $xlWindowState::xlMaximized  #-4137
    $xlMinimized = $xlWindowState::xlMinimized  #-4140
    $xlNormal    = $xlWindowState::xlNormal  #-4143

    $FileName = Convert-Path $FileName
    $UpdateLinks = [System.Reflection.Missing]::Value
    $_ReadOnly = [System.Reflection.Missing]::Value
    $Format = [System.Reflection.Missing]::Value
    $Password = if($Password){ $Password } else { [System.Reflection.Missing]::Value }
    $WriteResPassword = [System.Reflection.Missing]::Value
    $Ignorereadonlyrecommended = [System.Reflection.Missing]::Value
    $Origin = [System.Reflection.Missing]::Value
    $Delimiter = [System.Reflection.Missing]::Value
    $Editable = [System.Reflection.Missing]::Value
    $Notify = [System.Reflection.Missing]::Value
    $Converter = [System.Reflection.Missing]::Value
    $AddToMru = [System.Reflection.Missing]::Value
    $Local = [System.Reflection.Missing]::Value
    $CorruptLoad = [System.Reflection.Missing]::Value
    $app = New-Object -ComObject Excel.Application
    $book = $app.Workbooks.Open($FileName,
       $UpdateLinks, $_ReadOnly, $Format, $Password, $WriteResPassword, $Ignorereadonlyrecommended,
       $Origin, $Delimiter, $Editable, $Notify, $Converter, $AddToMru, $Local, $CorruptLoad
      )
    $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
}

function OpenPowerPointFile {
    Param( 
        [Parameter(Mandatory, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string] $FileName,
        [ValidateSet($null, $true, $false)]
        [object] $ReadOnly = $null
    )

    if( -not (Test-Path $FileName)) {
        Write-Host $FileName is not found
        return
    }
    
    [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.PowerPoint")
    $ppWindowState = [Microsoft.Office.Interop.PowerPoint.ppWindowState]
    $ppWindowMaximized = $ppWindowState::ppWindowMaximized  #3
    $ppWindowMinimized = $ppWindowState::ppWindowMinimized  #2
    $ppWindowNormal    = $ppWindowState::ppWindowNormal  #1

    $FileName = Convert-Path $FileName
    $_ReadOnly = if($null -ne $ReadOnly){ $ReadOnly } else { [System.Reflection.Missing]::Value }
    $Untitled = [System.Reflection.Missing]::Value
    $WithWindow = [System.Reflection.Missing]::Value

    $app = New-Object -ComObject PowerPoint.Application
    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    # $app.EnableResize = $true
    # $app.WindowState = $ppWindowMinimized
    $presentation = $app.Presentations.Open(
        $FileName,
        $_ReadOnly,
        $Untitled,
        $WithWindow)
    # $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
}