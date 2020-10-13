Call Spoof
Sub Spoof()
dim filePath,fileName,exten,FileLen,revExten,dest,NewFileName
set fs = CreateObject("Scripting.FileSystemObject")
filePath=inputbox("Enter The FilePath")
exten=inputbox("Enter the Extension to spoof" & vbCrlf & _
"Example : .jpg, .mp3 , .avi , etc..")
filepath=Replace(filepath,Chr(34),"")
fileName=fs.GetFileName(filepath)
filePath=fs.GetParentFolderName(filepath)
FileLen=Len(fileName)-4
Dim spclChar
spclChar = ChrW(8238)
revExten=StrReverse(exten)
NewFileName=inputbox("Enter the new file name max 5 chars")
dest=NewFileName & spclChar & revExten & mid(fileName,FileLen+2)

fs.copyFile filePath & "\" & Filename,filePath & "\" & dest
msgbox "Extension Spoofed Successfully!!",vbInformation
end sub
