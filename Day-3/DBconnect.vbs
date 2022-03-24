'variables to contain the db objects
Dim objConn,oRs
'Dim adstateOpen:adstateOpen=1
'create ado connection object
set objConn=CreateObject("ADODB.Connection")

'using dsn
objConn.Open "VBSCourse"

'If Not objConn.State=adstateOpen Then  
'sql query to get data from database
strSql="Select * from Emp"

'execute the query into database using ado object
set oRs=objConn.Execute(strSql)

'reading the data recived into the object after running the query to database 
Do While Not oRs.EOF
    strRep="Name :" & oRs("Name").Value & vbNewLine &"Age :" & oRs("Age").Value
    MsgBox strRep,0,"Employee Name"
    oRs.MoveNext
Loop
'closing the session we have created with database
oRs.Close
objConn.Close
'End If
