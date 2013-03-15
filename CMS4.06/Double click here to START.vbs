'===============================================
' Double click to this file to RUN .NET CMS
' StartUp Sequence V.1.03
' GNU General Public License
' FREE SUPPORT WWW.CMSASPNET.COM
'===============================================

'RUN THE WEB SERVER
Set Shell = WScript.CreateObject( "WScript.Shell" )
Shell.Run(".\CassiniDev4.exe /a:.\CMS  /pm:Specific /p:7777")


'OPEN THE BROWSER AND GO TO loading.html PAGE
Shell.Run("CMS\loading.html")

'	OR

'OPEN THE BROWSER AND GO TO DEFAULT PAGE
'Shell.Run("http://localhost:7777/loading.html")
