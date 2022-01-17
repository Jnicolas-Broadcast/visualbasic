Dim fso
Const OverwriteExisting = True
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile "\\10.1.3.171\c$\Program Files (x86)\Lansweeper\PackageShare\Installers\Nuevo Acrobat Reader\AcroRead.msi", "C:\", OverwriteExisting
'el origen del archivo lo toma desde "\\10.1.3.171......" y lo pasa al "C:\"