cd /d %~dp0
rem Do not use whitespace between the name and value; SET foo = bar will not work

set strExe=msxsl.exe
set strXslt=C:\Users\michalos\Documents\bin\vbs\xsd comments subfolder\xs3p.xsl 
set strXsd=C:\Users\michalos\Documents\XSD\xsd\qif\QIFApplications\QIFDocument.xsd
set outfile=C:\Users\michalos\Documents\XSD\xsd\qif\QIFApplications\QIFDocument.html
echo  %strExe%   %strXsd%  %strXslt%  -o  %outfile% 
%strExe%   "%strXsd%"  "%strXslt%"  -o  "%outfile%"

pause