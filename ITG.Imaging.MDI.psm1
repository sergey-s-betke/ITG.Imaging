function Copy-Mdi2Tiff { 
	<#
		.Synopsis
		    Преобразует MDI файл в TIFF формат
		.Description
		    Преобразует MDI файл в TIFF формат
		.Parameter sourceFile
		    Исходный .mdi файл (объект).
		.Parameter destination
			Полный путь (к папке), в который будет сохранён полученный .tiff файл.
            Если параметр не указан, будет создан файл с тем же именем, но с расширением .tiff
            в каталоге исходного файла
		.Parameter newName
            Имя создаваемого файла.
			Если параметр не указан, будет создан файл с тем же именем, но с расширением .tiff
		.Example
			Обработка файлов в каталоге с сохранением tiff файлов в тот же каталог:
			dir 'c:\temp\*.mdi' | Convert-MDI2TIFF
	#>
    
    param (
		[Parameter(
			Mandatory=$true,
			Position=0,
			ValueFromPipeline=$true,
			HelpMessage='Исходный .mdi файл (объект).'
		)]
        [System.IO.FileInfo] $sourceFile
		, [Parameter(
			Mandatory=$false,
			Position=1,
			ValueFromPipeline=$false,
			HelpMessage='Полный путь (к папке), в который будет сохранён полученный .tiff файл.'
		)]
        [System.IO.DirectoryInfo] $destination = ($sourceFile.Directory)
		, [Parameter(
			Mandatory=$false,
			Position=2,
			ValueFromPipeline=$false,
			HelpMessage='Имя создаваемого файла.'
		)]
  		[string] $newName = ([System.IO.Path]::GetFileNameWithoutExtension($sourceFile.name) + '.tif')
        , [switch] $PassThru
	)
    begin {
        $mdiDoc = new-object -comObject 'MODI.Document';
    }
    process {
        $mdiDoc.Create( $sourceFile.FullName );
        [System.IO.FileInfo] $tifFile = ( join-path -path $destination.Fullname -childPath $newName );
        $singlePage = ( $mdiDoc.Images.count -eq 1 );
        $mdiDoc.SaveAs( $tifFile.FullName, 2, 0 ); ## обязательно - без потерь и без сжатия, чтобы не jpeg компрессия была
        $mdiDoc.Close( $false );
        if ( $PassThru ) { $tifFile };
    }
    end {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mdiDoc) | out-null;
    }
};  

Export-ModuleMember `
    Copy-Mdi2Tiff `
;
