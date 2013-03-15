Set fso = CreateObject("Scripting.FileSystemObject")
Source = "C:\Users\Andrea\SkyDrive\Documenti\Developer\Sergenti\.NET CMS\"

on error resume next

'===================================
' incrementa ma versione x.nn
'===================================


Set oReadObj  = CreateObject("Scripting.FileSystemObject")
Set oRead = oReadObj.OpenTextFile(source & "readme.txt", 1)
Readme = oRead.ReadAll
oRead.Close()  ' Close input file

part1=left(Readme,11)
vminor=mid(Readme,12,2)
vminor=right("00" & vminor+1,2)
part2=mid(Readme,14)
Readme=part1 & vminor & part2

Set oWriteObj = CreateObject("Scripting.FileSystemObject")
Set oWrite = oWriteObj.CreateTextFile(source & "readme.txt", True)
oWrite.Write(Readme) 
oWrite.Close() ' Close output file


'===================================
' cancella directory da aggiornare
'===================================

fso.DeleteFolder ".\cms\App_WebReferences"
fso.DeleteFolder ".\cms\App_Code"
fso.DeleteFolder ".\cms\Resources"
fso.DeleteFolder ".\cms\ClientBin"
fso.DeleteFolder ".\cms\App_Data\Skins"

'===================================
' aggiorna le directory
'===================================

fso.CopyFolder Source & "App_WebReferences"	,".\cms\App_WebReferences"
fso.CopyFolder Source & "App_Code" 		, ".\cms\App_Code"
fso.CopyFolder Source & "Resources" 		, ".\cms\Resources"
fso.CopyFolder Source & "ClientBin" 		, ".\cms\ClientBin"
fso.CopyFolder Source & "App_Data\Skins" 	, ".\cms\App_Data\Skins"

'===================================
' cancella il dizionario delle lingue
'===================================

fso.DeleteFile ".\cms\App_Data\localization.mdb"
fso.DeleteFile ".\cms\App_Dode\General.PhraseBooks.xml"

'===================================
' copia il dizionario delle lingue
'===================================

fso.CopyFile Source & "App_Data\localization.mdb" 		, ".\cms\App_Data\"
fso.CopyFile Source & "App_Data\General.PhraseBooks.xml"	, ".\cms\App_Data\"

'===================================
' cancella i file nella root
'===================================

fso.DeleteFile ".\cms\*.vb"
fso.DeleteFile ".\cms\*.master"
fso.DeleteFile ".\cms\*.asax"
fso.DeleteFile ".\cms\*.aspx"
fso.DeleteFile ".\cms\*.ascx"
fso.DeleteFile ".\cms\*.txt"
fso.DeleteFile ".\cms\*.sln"
fso.DeleteFile ".\cms\*.zip"
fso.DeleteFile "*.zip"

'===================================
' copia i file sorgente nella root
'===================================

fso.CopyFile Source & "*.vb" 		, ".\cms\"
fso.CopyFile Source & "*.master" 	, ".\cms\"
fso.CopyFile Source & "*.asax" 		, ".\cms\"
fso.CopyFile Source & "*.aspx" 		, ".\cms\"
fso.CopyFile Source & "*.ascx" 		, ".\cms\"
fso.CopyFile Source & "*.txt" 		, ".\cms\"
fso.CopyFile Source & "*.sln" 		, ".\cms\"
fso.CopyFile Source & "*.html" 		, ".\cms\"


'===================================
' cancella file non distribuibili
'===================================

'fso.DeleteFile ".\cms\.NET CMS.sln"

fso.DeleteFile ".\cms\*.exclude"
fso.DeleteFile ".\cms\app_code\*.exclude"

fso.DeleteFile ".\cms\*.suo"

fso.DeleteFile ".\cms\app_code\Plugins\NewsBombing.vb"
fso.DeleteFile ".\cms\app_code\Plugins\NewsSending.vb"

fso.DeleteFolder ".\cms\app_code\Private"

fso.DeleteFile ".\cms\accommodation.aspx" 
fso.DeleteFile ".\cms\accommodation.aspx.vb" 

fso.DeleteFile ".\cms\dating.aspx" 
fso.DeleteFile ".\cms\dating.aspx.vb"
 
fso.DeleteFile ".\cms\download.aspx" 
fso.DeleteFile ".\cms\download.aspx.vb"
 
fso.DeleteFile ".\cms\guarantee.aspx" 
fso.DeleteFile ".\cms\guarantee.aspx.vb"
 
fso.DeleteFile ".\cms\proxy.aspx" 
fso.DeleteFile ".\cms\proxy.aspx.vb"
 
fso.DeleteFile ".\cms\services.aspx" 
fso.DeleteFile ".\cms\services.aspx.vb"
 
fso.DeleteFile ".\cms\showcode.aspx" 
fso.DeleteFile ".\cms\showcode.aspx.vb"
 
fso.DeleteFile ".\cms\test.aspx" 
fso.DeleteFile ".\cms\test.aspx.vb"
 
fso.DeleteFile ".\cms\Visa.aspx" 
fso.DeleteFile ".\cms\Visa.aspx.vb"

fso.DeleteFile ".\cms\VisaModule.ascx" 
fso.DeleteFile ".\cms\VisaModule.ascx.vb"

fso.DeleteFile ".\cms\PostalService.aspx" 
fso.DeleteFile ".\cms\PostalService.aspx.vb"

fso.DeleteFile ".\cms\Fricchettone.*"

fso.DeleteFile ".\cms\AndroidBenchmark.*"




dim objShell
'===================================
' Zipa tutto
'===================================

Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("""\Program Files\7-Zip\7z.exe"" a CMS4." & vminor & ".zip *  -x!""desktop.ini"" -x!""Thumbs.db"" -x!""update cms.vbs""")
'objShell.Run("""\Program Files\7-Zip\7z.exe"" a CMS.zip .\cms\*")


'===================================
' Lancia internet explorer
'===================================


objShell.Run("iexplore.exe https://sourceforge.net/projects/cmsaspnet/upload/")

Set objShell = Nothing


on error goto 0


'msgbox("End!")

