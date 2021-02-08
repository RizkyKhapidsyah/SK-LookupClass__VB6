VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VLRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FILE:  VLRecord.cls
'Author:  Stephen Goguen
'Email :  gsteve@hotmail.com

'Instructions:  Insert this into a class file

'CLASS:  VLRecord

'DESCRIPTION:  Provides the user with a simple Key->Value lookup interface,
'where the user can retrieve and set keys, but cannot enumerate keys or
'values like a hash table.

'EXAMPLE USE:
'Dim George As New VLRecord
'Dim Dennis As New VLRecord
'
'   George("First Name") = "George"
'   George("Last Name") = "Wilson"
'   George("Age") = 45
'   George("Demeanor") = "Irate"

'   Dennis("First Name") = "Dennis"
'   Dennis("Last Name") = "The Meanace"
'   Dennis("Age") = "Young"

'   Set Dennis("Target") = George

'   Debug.Print Dennis("Target")("First Name")  'Prints George

'   Dennis.Remove "Target"

'   George.RemoveAll

'NOTES:  After reading Advanced Microsoft Visual Basic 5 on the MSDN CD's,
'I rethought using first class data types and using variants instead for
'a number of compelling reasons.
'
'IMPORTANT:  Use 'Tools > Procedure Attributes' to define Item as [Default]
'method.


Option Explicit
Dim mKeyLookup As New Collection

'PROPERTY: Get Item
'DESCRIPTION:  Retrieves a Variant VALUE given the KEY.  If value exists
'for the key Null is returned...
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    On Error Resume Next
    Key = CStr(Key)
    If IsObject(mKeyLookup(Key)) = True Then
        Set Item = mKeyLookup(Key)
    Else
        Item = mKeyLookup(Key)
    End If
    If IsEmpty(Item) = True Then
        Item = Null
    End If
End Property

'PROPERTY: Let Item
'DESCRIPTION:  Sets the value of a Key for a simple data type.
Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    On Error Resume Next
    Key = CStr(Key)
    mKeyLookup.Add Value, Key
    If Err.Number <> 0 Then
        mKeyLookup.Remove Key
        mKeyLookup.Add Value, Key
    End If
End Property

'PROPERTY: Set Item
'DESCRIPTION:  Sets the value of a Key for object data types.
Public Property Set Item(ByVal Key As Variant, ByVal Value As Variant)
    On Error Resume Next
    Key = CStr(Key)
    mKeyLookup.Add Value, Key
    If Err.Number <> 0 Then
        mKeyLookup.Remove Key
        mKeyLookup.Add Value, Key
    End If
End Property

'FUNCTION:  Remove
'DESCRIPTION:  Removes a Key->Value pair
Public Function Remove(ByVal Key As Variant)
    On Error Resume Next
    mKeyLookup.Remove Key
End Function

'FUNCTION:  ClearAll
'DESCRIPTION:  Removes all Key->Value pairs
Public Function RemoveAll()
    Set mKeyLookup = New Collection
End Function


