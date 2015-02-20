@echo off
echo strMdbFile = WScript.Arguments.Item(0) > runconv.vbs
echo set fs = Wscript.CreateObject("Scripting.FileSystemObject") >> runconv.vbs
echo Set dbeng = CreateObject("DAO.DBEngine.36") >> runconv.vbs
echo Set db = dbeng.OpenDatabase(strMdbFile) >> runconv.vbs
echo For  tbl = 0 To db.TableDefs.Count - 1 >> runconv.vbs
echo If db.TableDefs(tbl).Attributes ^<^> 0 Then >> runconv.vbs
echo Else >> runconv.vbs
echo   dbTable = db.TableDefs(tbl).Name  >> runconv.vbs
echo   strTextOut = strMdbFile ^& "_" ^& dbTable ^& ".csv" >> runconv.vbs
echo   strTextOut = Replace(strTextOut, "/", "")   >> runconv.vbs
echo   Set rs = db.OpenRecordset(dbTable) >> runconv.vbs
echo   rs.movefirst >> runconv.vbs
echo   If rs.EOF = true Then >> runconv.vbs
echo     quit >> runconv.vbs
echo   End If >> runconv.vbs
echo   Set ts = fs.OpenTextFile(strTextOut, 2, True)  >> runconv.vbs
echo   strOutText = rs.Fields(0).Name >> runconv.vbs
echo   For  n = 1 To rs.Fields.Count - 1 >> runconv.vbs
echo     fName = rs.Fields(n).Name >> runconv.vbs
echo     strOutText = strOutText ^& "	" ^& fName >> runconv.vbs
echo   Next   >> runconv.vbs
echo     ts.Writeline strOutText >> runconv.vbs
echo   do while rs.EOF = false >> runconv.vbs
echo     strOutText = rs.Fields(0).Value >> runconv.vbs
echo     For  n = 1 To rs.Fields.Count - 1 >> runconv.vbs
echo       strOutText = strOutText ^& "	" ^& rs.Fields(n).Value >> runconv.vbs
echo     Next   >> runconv.vbs
echo     ts.Writeline strOutText >> runconv.vbs
echo     rs.movenext >> runconv.vbs
echo   loop >> runconv.vbs
echo   ts.close >> runconv.vbs
echo   rs.close >> runconv.vbs
echo   End if >> runconv.vbs
echo Next >> runconv.vbs
%SystemRoot%\SysWOW64\wscript.exe runconv.vbs %1
del runconv.vbs