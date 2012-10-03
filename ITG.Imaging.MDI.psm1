function Copy-Mdi2Tiff { 
	<#
		.Synopsis
		    ����������� MDI ���� � TIFF ������
		.Description
		    ����������� MDI ���� � TIFF ������
		.Parameter sourceFile
		    �������� .mdi ���� (������).
		.Parameter destination
			������ ���� (� �����), � ������� ����� ������� ���������� .tiff ����.
            ���� �������� �� ������, ����� ������ ���� � ��� �� ������, �� � ����������� .tiff
            � �������� ��������� �����
		.Parameter newName
            ��� ������������ �����.
			���� �������� �� ������, ����� ������ ���� � ��� �� ������, �� � ����������� .tiff
		.Example
			��������� ������ � �������� � ����������� tiff ������ � ��� �� �������:
			dir 'c:\temp\*.mdi' | Convert-MDI2TIFF
	#>
    
    param (
		[Parameter(
			Mandatory=$true,
			Position=0,
			ValueFromPipeline=$true,
			HelpMessage='�������� .mdi ���� (������).'
		)]
        [System.IO.FileInfo] $sourceFile
		, [Parameter(
			Mandatory=$false,
			Position=1,
			ValueFromPipeline=$false,
			HelpMessage='������ ���� (� �����), � ������� ����� ������� ���������� .tiff ����.'
		)]
        [System.IO.DirectoryInfo] $destination = ($sourceFile.Directory)
		, [Parameter(
			Mandatory=$false,
			Position=2,
			ValueFromPipeline=$false,
			HelpMessage='��� ������������ �����.'
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
        $mdiDoc.SaveAs( $tifFile.FullName, 2, 0 ); ## ����������� - ��� ������ � ��� ������, ����� �� jpeg ���������� ����
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
