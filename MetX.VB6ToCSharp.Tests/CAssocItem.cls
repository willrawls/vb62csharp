VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAssocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ********************************************************************************
' Class Module      CAssocItem
'
' Filename          CAssocItem.cls
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
' The Reality Matrix, Dimention 3 of 3
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

Public Key As String
Private m_sValue As String

Public Property Let Value(sNewValue As String)
    m_sValue = sNewValue
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = m_sValue
End Property

'
' Retrieves the Nth delimited token from the value.
' If delimiter is ommited, then a space is assumed as the delimiter.
' NOTE: sGetToken required for proper use
'
Public Property Get F(Index As Long, Optional sDelimiter As String = " ") As String
    F = sGetToken(Value, Index, sDelimiter)
End Property


