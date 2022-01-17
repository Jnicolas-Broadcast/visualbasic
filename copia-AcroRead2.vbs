'this script takes 2 arguments => "Source" and "C:\" and uses this to copy a the file
"\\10.1.3.171\c$\Program Files (x86)\Lansweeper\PackageShare\Installers\Nuevo Acrobat Reader\AcroRead.msi" = WScript.Arguments.Item(0)
"C:\" = WScript.Arguments.Item(1)

Set fso = CreateObject("Scripting.FileSystemObject")
'Check to see if the file already exists in the "C:\" folder
If fso.FileExists("C:\") Then
	'Check to see if the file is read-only
	If Not fso.GetFile("C:\").Attributes And 1 Then 
			fso.CopyFile "\\10.1.3.171\c$\Program Files (x86)\Lansweeper\PackageShare\Installers\Nuevo Acrobat Reader\AcroRead.msi", "C:\", True
	Else 
		'The file exists and is read-only.
		fso.GetFile("C:\").Attributes = fso.GetFile("C:\").Attributes - 1
			fso.CopyFile "\\10.1.3.171\c$\Program Files (x86)\Lansweeper\PackageShare\Installers\Nuevo Acrobat Reader\AcroRead.msi", "C:\", True
	End If
Else
		fso.CopyFile "\\10.1.3.171\c$\Program Files (x86)\Lansweeper\PackageShare\Installers\Nuevo Acrobat Reader\AcroRead.msi", "C:\", True
End If
Set fso = Nothing
