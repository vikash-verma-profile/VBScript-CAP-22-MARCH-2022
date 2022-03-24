'created variables
dim fso,myfile,filename,strtextfilepath

'path where we can create the file
strtextfilepath="C:\Users\om\Desktop\VBScript-CAP-22-MARCH-2022\Day-3\createdfile.txt"

'create object of filesystem
set fso=CreateObject("Scripting.FileSystemObject")

'use method CreateTextFile to create a file at the path we have assigned into a variable
set myfile=fso.CreateTextFile(strtextfilepath,true)

'check if file is created or not with exists function
If fso.FileExists(strtextfilepath) Then
MsgBox "FIle is created"

'write into the file using writline method
myfile.WriteLine("Hi i am writing data into a text file")

'close the file as while writing we have opened the file through code
myfile.Close

MsgBox "content is written into the file"
End IF