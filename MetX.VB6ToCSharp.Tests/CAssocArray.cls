VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAssocArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ********************************************************************************
' Class Module      CAssocArray
'
' Filename          CAssocArray.cls
'
' Copyright         1998 by Firm Solutions
'                   All rights reserved.
'
' Author            William M. Rawls
'                   Firm Solutions
'
' Created On        4/30/1998 1:23 PM
'
' Description
'
'    The Reality Matrix, Dimention 2 of 3
'       "Associative array" like abilities
'
'    What's "assosiative array" like abilities mean ?
'       Email = wrawls@firmsolutions.com to find out more.
'    Web page you ask ?
'       URL = http://www.firmsolutions.com/RealityMatrix.html
'    Why does this read like an e-mail ?
'       Because = It does
'
' Revisions
'
' <RevisionDate>, <RevisedBy>
' <Description of Revision>
'
' 4/30/1998, William M. Rawls
' Entered into public domain as freeware. Use at your own risk.
'
' ********************************************************************************

Public mCol              As Collection
Private CurItem          As CAssocItem
Public Section           As String
Public ItemDelimiter     As String
Public KeyValueDelimiter As String
Public FieldDelimiter    As String

Public Property Get All() As String
    Static sAllKeyValues As String
 
    For Each CurItem In mCol
        With CurItem
             sAllKeyValues = sAllKeyValues + .Key + KeyValueDelimiter + .Value + ItemDelimiter
        End With
    Next CurItem
    Set CurItem = Nothing

    All = sAllKeyValues
    sAllKeyValues = ""
End Property

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Public Function Add(Key As String, Optional Value As String) As CAssocItem
       Set CurItem = New CAssocItem
           With CurItem
                If Len(Value) = 0 And InStr(Key, Me.KeyValueDelimiter) > 0 Then
                   .Key = sGetToken(Key, 1, Me.KeyValueDelimiter)
                   .Value = sAfter(Key, 1, Me.KeyValueDelimiter)
                Else
                   .Key = Key
                   .Value = Value
                End If
           End With
           mCol.Add CurItem, Key
           Set Add = CurItem
       Set CurItem = Nothing
End Function

Public Function Column(Optional iCol As Long = 1, Optional ByVal sDelimiter As String) As String
    Static sColumn As String

    If Len(sDelimiter) = 0 Then sDelimiter = ItemDelimiter

    For Each CurItem In mCol
        sColumn = sColumn & sDelimiter & CurItem.F(iCol, FieldDelimiter)
    Next CurItem
    Set CurItem = Nothing

    Column = Mid$(sColumn, 2)
    sColumn = ""
End Function

Public Sub FillList(lstCtrl As Object, Optional sItemToChoose As String = "*NONE*", Optional ByVal iColumn As Integer = 0, Optional ByVal bClearList As Boolean = True)
    If bClearList Then lstCtrl.Clear
    
    If iColumn = 0 Then
       For Each CurItem In mCol
           lstCtrl.AddItem CurItem.Key
       Next CurItem
    Else
       For Each CurItem In mCol
           lstCtrl.AddItem CurItem.F(CLng(iColumn), FieldDelimiter)
       Next CurItem
    End If
    Set CurItem = Nothing
    
    If StrComp(sItemToChoose, "*NONE*") <> 0 Then
       SetListIndex lstCtrl, sItemToChoose
    End If
End Sub

Public Sub FillListView(lvwCtrl As Object, Optional bFullLine As Boolean = True)
    Static lvwX As ListView
    Static NewItem As ListItem
    Static SubItems As Integer
    Static CurSubItem As Integer
    Static sT As String

    Set lvwX = lvwCtrl
        lvwX.ListItems.Clear
        If bFullLine Then ExtendListView lvwX.hWnd

        For Each CurItem In mCol
            With CurItem
                 SubItems = iTokenCount(.Value, FieldDelimiter)
                 If SubItems > 0 Then
                    sT = sGetToken(.Value, 1, FieldDelimiter)
                    Set NewItem = lvwX.ListItems.Add(, "Key " & .Key, .Key) ', sT, sT)
                    For CurSubItem = 2 To SubItems
                        NewItem.SubItems(CurSubItem - 1) = sGetToken(.Value, CurSubItem, FieldDelimiter)
                    Next CurSubItem
                 Else
                    Set NewItem = lvwX.ListItems.Add(, .Key, .Key)
                 End If
            End With
        Next CurItem
        Set CurItem = Nothing
        Set NewItem = Nothing
    Set lvwX = Nothing
    sT = ""
End Sub

Public Sub FillListViewColumns(lvwCtrl As Object)
    Static lvwX As ListView
    Set lvwX = lvwCtrl
        With lvwX.ColumnHeaders
             .Clear
             For Each CurItem In mCol
                 .Add , CurItem.Key, CurItem.Key, Val(CurItem.Value)
             Next CurItem
             Set CurItem = Nothing
        End With
    Set lvwX = Nothing
End Sub

Public Sub FillTreeNode(tvwToFill As Object, nodCur As Object, sImage As String, bExpanded As Boolean)
    Dim nodX As Node
    Dim PrevNode As Node
    Dim tvwX As TreeView
    Dim sKey As String
    Dim sTag As String
    Dim sIcon As String
    Dim nChildDelimiter As Integer
    Dim nEndChildDelimiter As Integer

    Dim NodeStack() As Node
    ReDim NodeStack(0 To 0) As Node

    nChildDelimiter = Len(ChildDelimiter)
    nEndChildDelimiter = Len(EndChildDelimiter)

    Set tvwX = tvwToFill
        Set nodX = nodCur
            For Each CurItem In mCol
                sKey = sGetToken(CurItem.Key, 1, IconDelimiter)
                sIcon = sGetToken(CurItem.Key, 2, IconDelimiter)
                sTag = sGetToken(sKey, 2, TagDelimiter)
                sKey = sGetToken(sKey, 1, TagDelimiter)
                If Len(sIcon) = 0 Then sIcon = sImage

                If nChildDelimiter > 0 Then
                   If Left$(sKey, nChildDelimiter) = ChildDelimiter Then
                      sKey = Mid$(sKey, nChildDelimiter + 1)
                      If Not PrevNode Is Nothing Then
                         If UBound(NodeStack) > 0 Then
                            ReDim Preserve NodeStack(1 To UBound(NodeStack) + 1) As Node
                         Else
                           ReDim NodeStack(1 To 1) As Node
                         End If
                         Set NodeStack(UBound(NodeStack)) = nodX
                         Set nodX = PrevNode
                      End If
                   End If
                End If
On Error Resume Next
                If nodX Is Nothing Then
                   Set PrevNode = tvwX.Nodes.Add(, , sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
                Else
                   Set PrevNode = tvwX.Nodes.Add(nodX.Key, tvwChild, nodX.Key & "_" & sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
                End If
                With PrevNode
                     .ExpandedImage = sGetToken(sIcon, 1, EndChildDelimiter)
                     .Expanded = bExpanded
                     .Tag = sTag
                End With
                
                If nEndChildDelimiter > 0 Then
                   Do While Right$(sIcon, nEndChildDelimiter) = EndChildDelimiter
                      sIcon = Left$(sIcon, Len(sIcon) - nEndChildDelimiter)
                      Set nodX = NodeStack(UBound(NodeStack))
                      If UBound(NodeStack) > 1 Then
                         ReDim Preserve NodeStack(1 To UBound(NodeStack) - 1) As Node
                      Else
                         ReDim NodeStack(0 To 0) As Node
                      End If
                   Loop
                End If
            Next CurItem
        Set CurItem = Nothing
        Set PrevNode = Nothing
        Set nodX = Nothing
    Set tvwX = Nothing
End Sub

Public Property Get Item(sIndexKey As String) As CAssocItem
Attribute Item.VB_UserMemId = 0

On Error GoTo EH_CAssocArray_Item
       Set Item = mCol(sIndexKey)
 
EH_CAssocArray_Item_Continue:
       Exit Property
 
EH_CAssocArray_Item:
       Set Item = Add(sIndexKey)
       Err.Clear
       Resume EH_CAssocArray_Item_Continue
End Property

Public Property Get Count() As Long
       Count = mCol.Count
End Property


Public Property Let Keys(ByVal KeyValuePairs As String)
    Dim TokenCount As Long
    Dim CurToken As Long
    Dim sCurToken As String
    
    Me.Clear
    
    TokenCount = iTokenCount(KeyValuePairs, Me.ItemDelimiter)
    For CurToken = 1 To TokenCount
        Me.Add sGetToken(KeyValuePairs, CurToken, ItemDelimiter)
    Next CurToken

End Property

Public Property Get Keys() As String
    Static sKeys As String
    Dim sDelimiter As String

    sDelimiter = ItemDelimiter

    For Each CurItem In mCol
        sKeys = sKeys & sDelimiter & CurItem.Key
    Next CurItem

    Keys = Mid$(sKeys, 2)
    sKeys = ""
End Property

Public Sub ListViewToAll(lvwCtrl As Object)
    Static CurSubItem As Integer
    Static sAll As String

    Static lvwX As ListView
    Static CurListItem As ListItem

    Set lvwX = lvwCtrl
        Clear
        sAll = ""
        For Each CurListItem In lvwX.ListItems
            With CurListItem
                 sAll = sAll & ItemDelimiter & .Key & KeyValueDelimiter & .Icon
                 For CurSubItem = 1 To lvwX.ColumnHeaders.Count - 1
                     sAll = sAll & FieldDelimiter & .SubItems(CurSubItem)
                 Next CurSubItem
                 sAll = Mid$(sAll, Len(FieldDelimiter))
            End With
        Next CurListItem
        All = Mid$(sAll, Len(ItemDelimiter) + 1)
    Set CurListItem = Nothing
    Set lvwX = Nothing
    sAll = ""
End Sub

Public Sub RecordsetToAll(rst As Recordset)
    Dim CurField As Long
    Me.Clear
    
    Do Until rst.EOF
       With Me.Add(rst.Fields(0))
            For CurField = 1 To rst.Fields.Count - 1
                .Value = .Value & sReplace("" & rst.Fields(CurField), Me.FieldDelimiter, ".") & Me.FieldDelimiter
            Next CurField
            .Value = Left(.Value, Len(.Value) - 1)
       End With
       rst.MoveNext
    Loop
End Sub

Public Sub Remove(sIndexKey As String)
       mCol.Remove sIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

       Set NewEnum = mCol.[_NewEnum]
End Property


Public Property Let All(ByVal sAllKeyValues As String)
    Static sT As String
    Static sKey As String
    Static sValue As String

    Clear
    Do While Len(sAllKeyValues) > 0
       sT = sGetToken(sAllKeyValues, 1, ItemDelimiter)
       If InStr(sT, KeyValueDelimiter) > 0 Then
          sKey = sGetToken(sT, 1, KeyValueDelimiter)
          sValue = sGetToken(sT, 2, KeyValueDelimiter)
          Add sKey, sValue
       Else
          Add sT
       End If
       sAllKeyValues = sAfter(sAllKeyValues, 1, ItemDelimiter)
    Loop

    sT = ""
    sKey = ""
    sValue = ""

End Property

Public Sub TreeToAll(tvwToGet As Object)
    Dim sAll As String

    Dim CurNode As Node
    Dim tvwX As TreeView

    Set tvwX = tvwToGet
        Clear
        For Each CurNode In tvwX.Nodes
            With CurNode
                 If .Parent Is Nothing Then
                    If Len(.Tag) = 0 Then
                       sAll = sAll & ItemDelimiter & .Text & IconDelimiter & .Image
                    Else
                       sAll = sAll & ItemDelimiter & .Text & TagDelimiter & .Tag & IconDelimiter & .Image
                    End If
                    TreeToAll_AddChildren sAll, CurNode
                 End If
            End With
        Next CurNode
        All = Mid$(sAll, 3)
        sAll = ""
    Set CurNode = Nothing
    Set tvwX = Nothing
End Sub

Private Sub TreeToAll_AddChildren(sAll As String, CurNode As Node)
    Dim CurChild As Node
    Set CurChild = CurNode.Child
        If Not CurChild Is Nothing Then
           If Len(CurChild.Tag) = 0 Then
              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
           Else
              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
           End If
           
           If Not CurChild.Child Is Nothing Then
              TreeToAll_AddChildren sAll, CurChild
           End If
           
           Set CurChild = CurChild.Next
        
           Do Until CurChild Is Nothing
              If Len(CurChild.Tag) = 0 Then
                 sAll = sAll & ItemDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
              Else
                 sAll = sAll & ItemDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
              End If
              If Not CurChild.Child Is Nothing Then
                 TreeToAll_AddChildren sAll, CurChild
              End If
              Set CurChild = CurChild.Next
           Loop
           sAll = sAll & EndChildDelimiter
        End If
    Set CurChild = Nothing
End Sub

Private Sub Class_Initialize()
        Set mCol = New Collection
        ItemDelimiter = "~"
        KeyValueDelimiter = "="
        FieldDelimiter = ","
End Sub

Private Sub Class_Terminate()
        Set mCol = Nothing
End Sub

