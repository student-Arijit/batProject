@echo off
setlocal
set "input=%~1"
set "ext=%~x1"
set "name=%~n1"
set "fullpath=%cd%\%input%"
set "pdf=%cd%\%name%.pdf"

:: Word Documents
if /I "%ext%"==".doc"  goto ConvertWord
if /I "%ext%"==".docx" goto ConvertWord

:: Excel Files
if /I "%ext%"==".xls"  goto ConvertExcel
if /I "%ext%"==".xlsx" goto ConvertExcel

:: PowerPoint Files
if /I "%ext%"==".ppt"  goto ConvertPPT
if /I "%ext%"==".pptx" goto ConvertPPT

:: Image Files
if /I "%ext%"==".jpg"  goto ConvertImage
if /I "%ext%"==".jpeg" goto ConvertImage
if /I "%ext%"==".png"  goto ConvertImage

echo Unsupported file type: %ext%
goto :eof

:ConvertWord
powershell -Command "$word=New-Object -ComObject Word.Application; $word.Visible=$false; $doc=$word.Documents.Open('%fullpath%'); $doc.SaveAs('%pdf%', 17); $doc.Close(); $word.Quit()"
goto :eof

:ConvertExcel
powershell -Command "$excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $wb=$excel.Workbooks.Open('%fullpath%'); $wb.ExportAsFixedFormat(0, '%pdf%'); $wb.Close($false); $excel.Quit()"
goto :eof

:ConvertPPT
powershell -Command "$ppt=New-Object -ComObject PowerPoint.Application; $ppt.Visible=$false; $pres=$ppt.Presentations.Open('%fullpath%', $false, $false, $false); $pres.SaveAs('%pdf%', 32); $pres.Close(); $ppt.Quit()"
goto :eof

:ConvertImage
powershell -Command ^
  "$word = New-Object -ComObject Word.Application;" ^
  "$word.Visible = $false;" ^
  "$doc = $word.Documents.Add();" ^
  "$selection = $word.Selection;" ^
  "$selection.InlineShapes.AddPicture('%fullpath%');" ^
  "$doc.SaveAs([ref]'%pdf%', 17);" ^
  "$doc.Close([ref]$false);" ^
  "$word.Quit();" ^
  "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null;" ^
  "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null;" ^
  "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null;" ^
  "[GC]::Collect(); [GC]::WaitForPendingFinalizers();"
goto :eof


timeout /t 3 >nul
goto :eof
