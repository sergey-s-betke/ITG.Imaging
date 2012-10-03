add-type -assembly 'System.Drawing'; 
add-type -assembly 'PresentationCore';
add-type -path (join-path -path $PSScriptRoot -childPath 'PDFsharp\PdfSharp-WPF.dll');

function Copy-Tiff2PDF {
	<#
		.Synopsis
		    ����������� TIFF ���� � PDF
		.Description
		    ����������� TIFF ���� � PDF
		.Parameter sourceFile
		    �������� ���� (������).
		.Parameter destination
			������ ���� (� �����), � ������� ����� ������� ���������� ����.
            ���� �������� �� ������, ����� ������ ���� � ��� �� ������
            � �������� ��������� �����
		.Parameter newName
            ��� ������������ �����.
			���� �������� �� ������, ����� ������ ���� � ��� �� ������
		.Example
			��������� ������ ��������:
			dir 'c:\temp\*.tif' | Copy-Tiff2PDF
	#>
    
    param (
		[Parameter(
			Mandatory=$true,
			Position=0,
			ValueFromPipeline=$true,
			HelpMessage='�������� ���� (������).'
		)]
        [System.IO.FileInfo] $sourceFile
		, [Parameter(
			Mandatory=$false,
			Position=1,
			ValueFromPipeline=$false,
			HelpMessage='������ ���� (� �����), � ������� ����� ������� ���������� ����.'
		)]
        [System.IO.DirectoryInfo] $destination = ($sourceFile.Directory)
		, [Parameter(
			Mandatory=$false,
			Position=2,
			ValueFromPipeline=$false,
			HelpMessage='��� ������������ �����.'
		)]
  		[string] $newName = ([System.IO.Path]::GetFileNameWithoutExtension($sourceFile.name) + '.pdf')
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
        $decoder = new-object -typeName System.Windows.Media.Imaging.TiffBitmapDecoder -argumentList `
            ( $sourceStream ) `
            , ( [System.Windows.Media.Imaging.BitmapCreateOptions]::None ) `
            , ( [System.Windows.Media.Imaging.BitmapCacheOption]::None ) `
        ;
        [System.IO.FileInfo] $destFile = ( join-path -path $destination.Fullname -childPath $newName );
        $pdf = new-object -typeName PdfSharp.Pdf.PdfDocument; 
        
        foreach ($frame in $decoder.Frames) {
            [System.IO.FileInfo] $tempFile = [System.IO.Path]::GetTempFileName();
            $encoder = new-object -typeName System.Windows.Media.Imaging.TiffBitmapEncoder;
            $encoder.Frames.Add( $frame.Clone() );
            $tempStream = new-object -typeName System.IO.FileStream -argumentList `
                ( $tempFile ) `
                , ( [System.IO.FileMode]::Create ) `
            ;
            $encoder.Save( $tempStream );
            $tempStream.Close();

            [PdfSharp.Pdf.PdfPage] $page = $pdf.AddPage(); 
            $page.Height = [PdfSharp.Drawing.XUnit]::FromPresentation( $frame.Height );
            $page.Width = [PdfSharp.Drawing.XUnit]::FromPresentation( $frame.Width );
            $pdfFrame = [PdfSharp.Drawing.XImage]::FromFile( $tempFile.FullName );
            ( [PdfSharp.Drawing.XGraphics]::FromPdfPage( $page ) ).DrawImage( `
                $pdfFrame `
                , 0, 0 `
            );
            ## $page.Close();
            ## $tempFile.Delete();
        };

        $pdf.Save( $destFile.FullName ); 
        $pdf.Dispose();
        $sourceStream.Close();

        if ( $PassThru ) { $destFile };
    }
};  

Export-ModuleMember `
    Copy-Tiff2PDF `
;
