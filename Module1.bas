Attribute VB_Name = "Module1"
'MODULE THAT USES THE VLRECORD CLASS TO READ AND WRITE RECORDS TO
'DAO AND ADO DATABASES.  YOU MUST HAVE REFERENCES TO DAO AND ADO


'FUNCTION:  ReadDAORecord
'DESCRIPTION:  Reads from DAO type Recordset to VLRecord
Public Function ReadDAORecord(DAORecordset As DAO.Recordset) As VLRecord
    Dim Field As DAO.Field
    Dim Record As New VLRecord
    
    For Each Field In DAORecordset.Fields
        Record(Field.Name) = Field.Value
    Next
    Set ReadDAORecord = Record
End Function

'FUNCTION:  WriteDAORecord
'DESCRIPTION:  Writes record to DAO type recordset
Public Function WriteDAORecord(DAORecordset As DAO.Recordset, Record As VLRecord)
    Dim Field As DAO.Field

    For Each Field In DAORecordset.Fields
        Field.Value = Record(Field.Name)
    Next
End Function

'FUNCTION:  ReadADORecord
'DESCRIPTION:  Read
Public Function ReadADORecord(ADORecordset As ADODB.Recordset) As VLRecord
    Dim Field As ADODB.Field
    Dim Record As New VLRecord
    
    For Each Field In ADORecordset.Fields
        Record(Field.Name) = Field.Value
    Next
    Set ReadADORecord = Record
End Function

'FUNCTION:  WriteADORecord
'DESCRIPTION:  Writes record to ADO type recordset
Public Function WriteADORecord(ADORecordset As ADODB.Recordset, Record As VLRecord)
    Dim Field As ADODB.Field

    For Each Field In ADORecordset.Fields
        Field.Value = Record(Field.Name)
    Next
End Function




Public Sub Main()
'DEMO OF USE OF VLRECORD CLASS
''EXAMPLE USE:
Dim George As New VLRecord
Dim Dennis As New VLRecord
'
    George("First Name") = "George"
   George("Last Name") = "Wilson"
   George("Age") = 45
   George("Demeanor") = "Irate"

   Dennis("First Name") = "Dennis"
   Dennis("Last Name") = "The Meanace"
   Dennis("Age") = "Young"

   Set Dennis("Target") = George

   Debug.Print Dennis("Target")("First Name")  'Prints George

   Dennis.Remove "Target"

   George.RemoveAll

End Sub
