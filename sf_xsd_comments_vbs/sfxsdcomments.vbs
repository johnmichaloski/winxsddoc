
Dim oShell, strExe, strXsd, strOriginalCD ,  oNothing 
Dim  scriptdir , ServiceName , crcldir 
Dim objFolder, objShell, objFSO, xsdFolder, colFiles , sf
Dim html,foldername 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set oShell = CreateObject ("WScript.Shell") 

scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strExe = scriptdir & "\msxsl.exe" 
strXslt = scriptdir & "\xs3p.xsl" 

Main


Function GetParentFolderNamepath(path )
    result = objFSO.GetParentFolderName(path)
    GetParentFolderName = Mid(result, InStrRev(result, "\") + 1)
End Function

Function SubFolder(path)
        Dim bArray, vArray, bsize, vsize, i
        bArray = Split(crcldir, "\")
        vArray = Split(path, "\")
        bsize = UBound(bArray)
        vsize = UBound(vArray)
        SubFolder = ""
        i=bSize
        while i< vsize
          SubFolder=  SubFolder & vArray(i+1) & "/" 
          i = i + 1 
        wend
End Function


Function TraverseFolders( strpath ) 
  set fldr = objFSO.GetFolder( strpath )
  foldername = objFSO.GetBasename( strpath  )
  foldersubpath = SubFolder(strpath)
  'MsgBox("name = " & foldername & " path = " & foldersubpath )

  html = html  & "<ul>" 
  html = html  & "<h1>" & UCASE(foldername) & " XSD Documentation</h1>" & vbCrLf


  curdir = strpath 

  For Each objFile in fldr.Files
    If UCase(objFSO.GetExtensionName(objFile.name)) = "XSD" Then
         strXsd =  curdir & "\" & objFile.Name
         outfile = curdir & "\" & objFSO.GetBaseName(objFile.Name) & ".html"  
	 html = html  & "<li><a href=""./" &  foldersubpath & objFSO.GetBaseName(objFile.Name) _
                      & ".html""> " & objFSO.GetBaseName(objFile.Name) _
                      & "</a></li>" &vbCrLf

	 oShell.run "%comspec% /c " & "cd  " & scriptdir & " & msxsl.exe """ & strXsd & """ """ & strXslt & """ -o """ & outfile & "" , 1, true
    End If
  Next

  For Each sf In fldr.SubFolders
    TraverseFolders sf.path
  Next

  html = html  & "</ul>" 
End Function


Sub Main()
Set oNothing = Nothing
Set objFolder = objShell.BrowseForFolder(0, "Please select the folder.", 1, "")
if oNothing  is objFolder = True Then
	Wscript.Quit(0)
End if
crcldir = objFolder.Self.path

html = ""
html = html  & "<html>" & vbCrLf
html = html  &"<body>" & vbCrLf


TraverseFolders crcldir


html = html  & "</body>" & vbCrLf
html = html  & "</html>" & vbCrLf
set objFileToWrite = objFSO.OpenTextFile(crcldir & "\index.html",2,true)
objFileToWrite.Write(html)
objFileToWrite.Close
    

End Sub


