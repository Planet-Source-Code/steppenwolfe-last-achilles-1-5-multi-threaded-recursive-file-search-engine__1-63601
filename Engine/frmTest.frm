VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Achilles 1.0 - Test Harness"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPattern 
      Caption         =   "Pattern >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7410
      TabIndex        =   10
      Top             =   5130
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6060
      TabIndex        =   9
      Top             =   5130
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4710
      TabIndex        =   8
      Top             =   5130
      Width           =   1215
   End
   Begin VB.ListBox lstSearch 
      Height          =   4545
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   2745
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   4575
      Left            =   3000
      TabIndex        =   1
      Top             =   300
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8760
      TabIndex        =   0
      Top             =   5130
      Width           =   1215
   End
   Begin VB.Label lblFound 
      AutoSize        =   -1  'True
      Caption         =   "Found:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   5310
      Width           =   585
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      Caption         =   "25 Common Windows Files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblEngine 
      AutoSize        =   -1  'True
      Caption         =   "Idle.."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4260
      TabIndex        =   3
      Top             =   90
      Width           =   450
   End
   Begin VB.Label lblEngineCap 
      AutoSize        =   -1  'True
      Caption         =   "Engine Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   90
      Width           =   1245
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cEngine      As clsEngine
Attribute cEngine.VB_VarHelpID = -1

Private cLocal                  As New Collection
Private cResult                 As New Collection
Private tTime                   As Double
Private lItem                   As ListItem
Private lCounter                As Long

Private Sub cEngine_eEComplete()
    Debug.Print "scan complete"
End Sub

Private Sub cEngine_eECount()
    Debug.Print "count"
End Sub

Private Sub cEngine_eECountMax(lMax As Long)
Debug.Print "count max " & lMax
End Sub

Private Sub cEngine_eEDump()
Debug.Print "dump complete"
End Sub

Private Sub cEngine_eEEngaged()
Debug.Print "engine engaged"
End Sub

Private Sub cEngine_eEReload()
Debug.Print "reload complete"
End Sub

Private Sub cEngine_eEReset()
Debug.Print "reset complete"
End Sub

Private Sub cEngine_eERestore()
Debug.Print "index restored"
End Sub

Private Sub cEngine_eEReturn()
Debug.Print "index returned"
End Sub

Private Sub cEngine_eEReturned()
Debug.Print "index return"
End Sub

Private Sub cEngine_eEStoreReset()
Debug.Print "storage reset"
End Sub

Private Sub cmdBuild_Click()
    
    With cEngine
        .p_BuildPath = "c:\"
        .p_EngineTask = Index_Build
        .Start
    End With

End Sub

Private Sub cmdPattern_Click()

    With cEngine
        List
        Set .p_CForward = cLocal
        .p_EngineTask = Search_Pattern
        .Start
    End With

End Sub

Private Sub cEngine_eEPatternComplete()
Debug.Print "pattern search complete"

Dim vItem As Variant

    With cEngine
        Debug.Print .p_CReturn.Count
        For Each vItem In .p_CReturn
            Debug.Print vItem
        Next
    End With

End Sub

Private Sub cmdSearch_Click()
    
    With cEngine
        List
        Set .p_CForward = cLocal
        .p_EngineTask = Search_Exact
        .Start
    End With

End Sub

Private Sub cEngine_eEProcessComplete()
Debug.Print "process scan complete"

Dim vItem As Variant

    With cEngine
        Debug.Print .p_CReturn.Count
        For Each vItem In .p_CReturn
            Debug.Print vItem
        Next
    End With
    
End Sub

Private Sub cmdRestore_Click()
    
    With cEngine
        .p_IndexPath = App.Path & "\index1.dat"
        .p_EngineTask = Index_Restore
        .Start
    End With

End Sub

Private Sub cmdSave_Click()
    
    With cEngine
        .p_IndexId = 1
        .p_EngineTask = Index_Save
        .Start
    End With

End Sub

Private Sub Form_Load()

    '//dimension objects/set options
    Set cEngine = New clsEngine
    Set cLocal = New Collection
    Set cResult = New Collection

End Sub

Private Sub List()
'//common files for test

    With cLocal
        .Add "test"
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'//release resource

    Set cEngine = Nothing

End Sub

Private Sub optSearch_Click(Index As Integer)

    lstResults.ListItems.Clear
    lstResults.ColumnHeaders.Clear
    '//display toggle
    Select Case Index
    Case 0
        lstResults.ColumnHeaders.Add 1, , "Results", (lstResults.Width)
    Case 1
        With lstResults
            .ColumnHeaders.Add 1, , "Path", (.Width / 4) * 2
            .ColumnHeaders.Add 2, , "Created", (.Width / 4)
            .ColumnHeaders.Add 3, , "Modified", (.Width / 8)
            .ColumnHeaders.Add 4, , "Size", (.Width / 8)
        End With
    End Select

End Sub
