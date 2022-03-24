Set EmployeeData=CreateObject("Scripting.Dictionary")
Id=101'
EmployeeData.Add (Id,"Vikash Verma : 40 : Male")
EmployeeData.Add "102","Sumit Kumar : 21 : Male"
EmployeeData.Add "103","Raj Kumar : 22 : Male"

dim output
For each ele in EmployeeData.Items
    output=output & ele & vbNewLine
Next

MsgBox output,0,"Employee Details"