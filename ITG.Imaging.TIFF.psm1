add-type -assembly 'System.Drawing'; 
add-type -assembly 'PresentationCore';

function Copy-Tiff2Tiff {
    <#
        .Synopsis
            "Исправляет" одностраничные TIFF файлы после mdi
        .Description
            "Исправляет" одностраничные TIFF файлы после mdi
        .Parameter sourceFile
            Исходный .tiff файл (объект).
        .Parameter destination
            Полный путь (к папке), в который будет сохранён полученный .tiff файл.
            Если параметр не указан, будет создан файл с тем же именем, но с расширением .tiff
            в каталоге исходного файла
        .Parameter newName
            Имя создаваемого файла.
            Если параметр не указан, будет создан файл с тем же именем, но с расширением .tiff
        .Example
            Обработка файлов в каталоге с сохранением tiff файлов в тот же каталог:
            dir 'c:\temp\*.tif' | Convert-TIFF2TIFF
    #>
    
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$true,
            HelpMessage='Исходный .tiff файл (объект).'
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

    process {
        $sourceStream = new-object -typeName System.IO.FileStream -argumentList `
            ( $sourceFile.FullName ) `
            , ( [System.IO.FileMode]::Open ) `
            , ( [System.IO.FileAccess]::Read  ) `
            , ( [System.IO.FileShare]::Read  ) `
            , 16384 `
            , ( [System.IO.FileOptions]::SequentialScan ) `
        ;
        [System.IO.FileInfo] $destFile = ( join-path -path $destination.Fullname -childPath $newName );

        $decoder = new-object -typeName System.Windows.Media.Imaging.TiffBitmapDecoder -argumentList `
            ( $sourceStream ) `
            , ( [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat `
                -bor [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache `
            ) `
            , ( [System.Windows.Media.Imaging.BitmapCacheOption]::None ) `
        ;
        $framesCount = $decoder.Frames.count;

        if ( $framesCount -eq 1 ) {
            $bmp = new-object -typeName System.Drawing.Bitmap -argumentList $sourceStream;
            $tempStream = new-object -TypeName System.IO.MemoryStream;
            $bmp.Save( $tempStream, ([System.Drawing.Imaging.ImageFormat]::Tiff) );
            $sourceStream.Close();
            $bmp.Dispose();

            $destStream = new-object -typeName System.IO.FileStream -argumentList `
                ( $destFile.FullName ) `
                , ( [System.IO.FileMode]::Create ) `
                , ( [System.IO.FileAccess]::Write  ) `
                , ( [System.IO.FileShare]::None  ) `
                , 16384 `
                , ( [System.IO.FileOptions]::SequentialScan ) `
            ;
            $tempStream.WriteTo( $destStream );
            $destStream.Close();
            $tempStream.Close();
        } else {
            $sourceStream.Close();
            if ( $destination.fullName -ne $sourceFile.Directory.fullName ) {
                $sourceFile.CopyTo( $destFile.fullName, $true );
            };
        };

        if ( $PassThru ) { $destFile };
    }
};  

function Copy-Tiff2TiffBlackWhite {
    <#
        .Synopsis
            Преобразует TIFF файл в монохромный TIFF
        .Description
            Преобразует TIFF файл в монохромный TIFF
        .Parameter sourceFile
            Исходный файл (объект).
        .Parameter destination
            Полный путь (к папке), в который будет сохранён полученный .tiff файл.
            Если параметр не указан, будет создан файл с тем же именем
            в каталоге исходного файла
        .Parameter newName
            Имя создаваемого файла.
            Если параметр не указан, будет создан файл с тем же именем
        .Example
            Обработка файлов каталога с их перезаписью "на месте":
            dir 'c:\temp\*.tif' | Convert-Tiff2TiffBlackWhite
    #>
    
    param (
        [Parameter(
            Mandatory=$true,
            Position=0,
            ValueFromPipeline=$true,
            HelpMessage='Исходный файл (объект).'
        )]
        [System.IO.FileInfo] $sourceFile
        , [Parameter(
            Mandatory=$false,
            Position=1,
            ValueFromPipeline=$false,
            HelpMessage='Полный путь (к папке), в который будет сохранён полученный файл.'
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
    
    process {
        $sourceStream = new-object -typeName System.IO.FileStream -argumentList `
            ( $sourceFile.FullName ) `
            , ( [System.IO.FileMode]::Open ) `
            , ( [System.IO.FileAccess]::Read  ) `
            , ( [System.IO.FileShare]::Read  ) `
            , 16384 `
            , ( [System.IO.FileOptions]::SequentialScan ) `
        ;
        [System.IO.FileInfo] $destFile = ( join-path -path $destination.Fullname -childPath $newName );

        $decoder = new-object -typeName System.Windows.Media.Imaging.TiffBitmapDecoder -argumentList `
            ( $sourceStream ) `
            , ( [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat `
                -bor [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache `
            ) `
            , ( [System.Windows.Media.Imaging.BitmapCacheOption]::None ) `
        ;
        $encoder = new-object -typeName System.Windows.Media.Imaging.TiffBitmapEncoder;

        foreach ($frame in $decoder.Frames) {
            $encoder.Frames.Add( [System.Windows.Media.Imaging.BitmapFrame]::Create(
                ( new-object -typeName System.Windows.Media.Imaging.FormatConvertedBitmap -argumentList `
                    ($frame) `
                    , ([System.Windows.Media.PixelFormats]::BlackWhite) `
                    , ([System.Windows.Media.Imaging.BitmapPalettes]::BlackAndWhite) `
                    , 1.0 `
                )
            ) );
        };

        $tempStream = new-object -TypeName System.IO.MemoryStream;
        $encoder.Compression = [System.Windows.Media.Imaging.TiffCompressOption]::Ccitt4;
        $encoder.Save( $tempStream );
        $sourceStream.Close();

        $destStream = new-object -typeName System.IO.FileStream -argumentList `
            ( $destFile ) `
            , ( [System.IO.FileMode]::Create ) `
            , ( [System.IO.FileAccess]::Write ) `
            , ( [System.IO.FileShare]::None ) `
            , 16384 `
            , ( [System.IO.FileOptions]::SequentialScan ) `
        ;
        $tempStream.WriteTo( $destStream );
        $destStream.Close();
        $tempStream.Close();

        if ( $PassThru ) { $destFile };
    }
};  

Export-ModuleMember `
    Copy-Tiff2TiffBlackWhite `
    , Copy-Tiff2Tiff `
;
