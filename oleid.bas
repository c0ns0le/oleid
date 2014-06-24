Attribute VB_Name = "OLEID"
' Export an item from Outlook based on the EntryId
'
' Copyright (c) 2014 Malte Stretz <stretz@silpion.de> (Silpion IT-Solutions GmbH)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Sub DumpItemByEntryID()
    Dim EntryID As String
    EntryID = InputBox("EntryID", "Get EntryID")
    If IsEmpty(EntryID) Then
        Exit Sub
    End If
    Dim Result As String
    Result = SaveItemByEntryID(EntryID, Environ("TEMP"))
    If Not Result = "" Then
        MsgBox "Item saved as " & Result, vbInformation, EntryID
    Else
        MsgBox "Item not found", vbExclamation, EntryID
    End If
End Sub

Function SaveItemByEntryID(EntryID As String, Dir As String) As String
    Dim Item As Object
    Set Item = GetItemByEntryID(EntryID)
    If Not Item Is Nothing Then
        EntryID = Item.EntryID
        Debug.Print Item & " " & TypeName(Item) & " " & EntryID
        
        Dim Path As String
        Path = Dir & "\" & EntryID & ".msg"
        Item.SaveAs Path, olMSGUnicode
        
        SaveItemByEntryID = Path
    Else
        Debug.Print "- Nothing " & EntryID
        
        SaveItemByEntryID = ""
    End If
End Function

Function GetItemByEntryID(ID As String) As Object
    On Error GoTo Error
    Set GetItemByEntryID = Nothing
    
    Dim Session As Outlook.NameSpace
    Set Session = Application.Session
    
    Dim Stores As Outlook.Stores
    Set Stores = Session.Stores
    
    Dim Store As Outlook.Store
    For Each Store In Stores
        On Error Resume Next
        Set Item = Session.GetItemFromID(ID, Store.StoreID)
        If Not Item Is Nothing Then
            If Not IsEmpty(Item) Then
                Set GetItemByEntryID = Item
                Exit Function
            End If
        End If
        On Error GoTo Error
    Next Store
    
Error:
    Exit Function
End Function

