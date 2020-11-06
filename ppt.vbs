set fso=CreateObject("Scripting.FileSystemObject")
set fs=fso.getfolder(".").files
For Each argv In fs
Call myfor(argv)
next
Sub myfor(input)
filename = input
if right(filename, 3) = "ppt" or right(filename, 3) = "pps" then
pdfname = left(filename, len(filename)-3) + "pdf"
elseif right(filename, 4) = "pptx" then
pdfname = left(filename, len(filename)-4) + "pdf"
else
exit sub
MSGBOX
end if
Set pptApp = CreateObject("PowerPoint.Application")
Set MyPress = pptApp.Presentations.Open(filename)
ppSaveAsPDF = 32
MyPress.SaveAs pdfname, ppSaveAsPDF, false
pptApp.Quit
end sub