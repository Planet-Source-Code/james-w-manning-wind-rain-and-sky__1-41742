VERSION 5.00
Begin VB.Form frmAirport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Location Wizard"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   30
      Width           =   6375
      Begin VB.CheckBox chkDontShow 
         Caption         =   "Do not show this page next time."
         Height          =   240
         Left            =   0
         TabIndex        =   7
         Top             =   4440
         Width           =   2715
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Please click next to continue..."
         Height          =   210
         Left            =   285
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Welcome!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15
         TabIndex        =   5
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to the Select Location Wizard!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1350
         TabIndex        =   4
         Top             =   450
         Width           =   3390
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   1
      Left            =   2265
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ListBox lstCountry 
         Height          =   4050
         Left            =   195
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   405
         Width           =   6060
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select a country:"
         Height          =   210
         Left            =   195
         TabIndex        =   9
         Top             =   120
         Width           =   1740
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Top             =   45
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ListBox lstProvince 
         Height          =   4050
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   6060
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Now, please select a province/state:"
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   270
         Width           =   2655
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   3
      Left            =   2280
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ListBox lstCity 
         Height          =   3630
         Left            =   225
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   645
         Width           =   5625
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Finally, please select a city:"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   315
         Width           =   1980
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   4
      Left            =   2280
      TabIndex        =   18
      Top             =   60
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkDefault 
         Caption         =   "Make this location the default when I run this program next time."
         Height          =   270
         Left            =   150
         TabIndex        =   20
         Top             =   4335
         Width           =   4920
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   $"frmAirport.frx":0000
         Height          =   480
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   5940
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4830
      TabIndex        =   2
      Top             =   5115
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6045
      TabIndex        =   1
      Top             =   5115
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7410
      TabIndex        =   0
      Top             =   5100
      Width           =   1215
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   210
      Left            =   150
      TabIndex        =   17
      Top             =   5145
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   4740
      Left            =   60
      Top             =   60
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   182
      X2              =   8642
      Y1              =   4930
      Y2              =   4930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   180
      X2              =   8640
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmAirport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iStep As Integer
Dim xData As ADODB.Connection
Dim xRs As ADODB.Recordset
Dim SkipCity As Boolean

Private Sub cmdBack_Click()
    If iStep = fraWizard.UBound Then
        If Not SkipCity Then
            iStep = iStep - 1
        Else
            iStep = iStep - 2
        End If
    Else
        iStep = iStep - 1
    End If
    SetStep
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If cmdNext.Caption <> "Finish" Then
        iStep = iStep + 1
        SetStep
    Else
        Location = IIf(SkipCity, Trim(Mid$(lstProvince.Text, 1, 5)), Trim(Mid$(lstCity.Text, 1, 5)))
        frmMain.mRefresh True
        If Not NoData Then
            If chkDefault.Value = 1 Then
                SaveSetting "WindRainSky", "DefaultLocation", "ICAO", Location
            End If
            Unload Me
        End If
        If chkDontShow.Value = 1 Then
            SaveSetting "WindRainSky", "LocationWizard", "Show", "0"
        Else
            SaveSetting "WindRainSky", "LocationWizard", "Show", "1"
        End If
    End If
End Sub

Private Sub Form_Load()
    lblResults.Caption = "Nothing Selected yet."
    If GetSetting("WindRainSky", "LocationWizard", "Show", "1") = "1" Then
        iStep = 0
    Else
        iStep = 1
    End If
    SetStep
End Sub

Private Sub SetStep()
    Dim i As Integer
    Screen.MousePointer = 11
    Select Case iStep
        Case 1
            Set xData = New ADODB.Connection
            Set xRs = New ADODB.Recordset
            With xData
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Locations.mdb;Persist Security Info=False"
                .Open
            End With
            With xRs
                .Open "SELECT * FROM Countries", xData, adOpenKeyset, adLockOptimistic
                lstCountry.Clear
                .MoveFirst
                Do Until .EOF
                    DoEvents
                    lstCountry.AddItem !CountryName
                    lstCountry.ItemData(lstCountry.NewIndex) = !CountryID
                    .MoveNext
                Loop
            End With
            xRs.Close
            xData.Close
            Set xRs = Nothing
            Set xData = Nothing
        Case 2
            Set xData = New ADODB.Connection
            Set xRs = New ADODB.Recordset
            With xData
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Locations.mdb;Persist Security Info=False"
                .Open
            End With
            With xRs
                .Open "SELECT * FROM ProvState WHERE CountryID=" & lstCountry.ItemData(lstCountry.ListIndex) & ";", xData, adOpenKeyset, adLockOptimistic
                lstProvince.Clear
                .MoveFirst
                Do Until .EOF
                    DoEvents
                    lstProvince.AddItem !ProvState
                    lstProvince.ItemData(lstProvince.NewIndex) = !ProvStateID
                    .MoveNext
                Loop
            End With
            xRs.Close
            xData.Close
            Set xRs = Nothing
            Set xData = Nothing
        Case 3
            Set xData = New ADODB.Connection
            Set xRs = New ADODB.Recordset
            With xData
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Locations.mdb;Persist Security Info=False"
                .Open
            End With
            With xRs
                .Open "SELECT * FROM Cities WHERE ProvStateID=" & lstProvince.ItemData(lstProvince.ListIndex) & ";", xData, adOpenKeyset, adLockOptimistic
                If Not .EOF And Not .BOF Then
                    lstCity.Clear
                    .MoveFirst
                    Do Until .EOF
                        DoEvents
                        lstCity.AddItem !CityName
                        lstCity.ItemData(lstCity.NewIndex) = !CityID
                        .MoveNext
                    Loop
                    SkipCity = False
                Else
                    iStep = iStep + 1
                    SkipCity = True
                End If
            End With
            xRs.Close
            xData.Close
            Set xRs = Nothing
            Set xData = Nothing
    End Select
    For i = 0 To fraWizard.UBound
        If i = iStep Then
            fraWizard(i).Visible = True
        Else
            fraWizard(i).Visible = False
        End If
    Next i
    If iStep <> 0 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    If iStep <> fraWizard.UBound Then
        cmdNext.Enabled = True
    Else
        cmdNext.Caption = "Finish"
    End If
    If iStep < fraWizard.UBound Then
        cmdNext.Caption = "Next >>"
    End If
    Screen.MousePointer = 0
    lblResults.Caption = IIf(lstCountry.Text <> vbNullString, lstCountry.Text, "") & IIf(lstProvince.Text <> vbNullString, ", " & lstProvince.Text, "") & IIf(lstCity.Text <> vbNullString, lstCity.Text, "")
End Sub
