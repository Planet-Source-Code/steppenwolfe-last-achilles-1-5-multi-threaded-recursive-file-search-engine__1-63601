VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Results_Process()
'/* scan for an exact match

Dim S As String
Dim x As Long
Dim l As Long
Dim V As Variant

On Error Resume Next
    
    With c_CReturn
        If c_CForward.Count = 0 Then Exit Sub
        For Each V In c_CForward
            '/* look up key
           S = c_Storage.Item(CStr(V))
            If Not LenB(S) > 0 Then GoTo Skip
            x = LenB(S)
            '/* if exists add it to results
            If Not LenB(x) = 0 Then
                .Add S
                '/* test for additional keys
                Do Until x = 0
                    l = l + 1
                    S = vbNullString
                    S = c_Storage.Item(V & Chr$(31) & l)
                    x = LenB(S)
                    If Not LenB(S) > 0 Then Exit Do
                    .Add S
                Loop
            End If
Skip:
        Next V
    End With

On Error GoTo 0

End Sub

Public Sub Results_Pattern()
'/* search for a pattern within items

Dim V As Variant
Dim M As Variant
Dim S As Variant
Dim l As Long

On Error Resume Next

    If c_CForward.Count = 0 Then Exit Sub
        For Each V In c_CForward
            For Each M In c_Storage
                If Not LenB(M) > 0 Then GoTo Skip
                S = Mid$(M, InStrRev(M, Chr$(92)) + 1)
                If InStr(1, S, V) > 0 Then
                    c_CReturn.Add M
                End If
Skip:
            Next M
       Next V

On Error GoTo 0

End Sub

