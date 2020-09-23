VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wind, Rain and Sky"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   5985
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   585
      Top             =   105
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   -30
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin WindRainSky.LED7Seg Bar7Seg 
      Height          =   1020
      Index           =   0
      Left            =   4035
      TabIndex        =   0
      Top             =   795
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin WindRainSky.LED7Seg Bar7Seg 
      Height          =   1020
      Index           =   1
      Left            =   4755
      TabIndex        =   1
      Top             =   780
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin WindRainSky.LED7Seg Bar7Seg 
      Height          =   1020
      Index           =   2
      Left            =   5475
      TabIndex        =   2
      Top             =   780
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin WindRainSky.LED7Seg Bar7Seg 
      Height          =   1020
      Index           =   3
      Left            =   6210
      TabIndex        =   3
      Top             =   780
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin VB.Label lblConditions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1980
      TabIndex        =   11
      Top             =   4755
      Width           =   1575
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   1980
      TabIndex        =   10
      Top             =   4560
      Width           =   1560
   End
   Begin VB.Label lblWGust 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1710
      TabIndex        =   9
      Top             =   3555
      Width           =   690
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   5355
      Shape           =   3  'Circle
      Top             =   3975
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   3135
      Width           =   390
   End
   Begin VB.Label lblImperial 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   5325
      TabIndex        =   8
      Top             =   5100
      Width           =   135
   End
   Begin VB.Label lblMetric 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5385
      TabIndex        =   7
      Top             =   4755
      Width           =   105
   End
   Begin VB.Label lblWSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1725
      TabIndex        =   6
      Top             =   2910
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Location Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1980
      TabIndex        =   5
      Top             =   4380
      Width           =   1530
   End
   Begin VB.Label lblVis 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Km"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   450
      TabIndex        =   4
      Top             =   810
      Width           =   3090
   End
   Begin VB.Line lnWindDir 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   1035
      X2              =   1035
      Y1              =   4905
      Y2              =   4320
   End
   Begin VB.Line lnTemp 
      BorderWidth     =   2
      Index           =   0
      X1              =   5565
      X2              =   4260
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line lnWindSpd 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   2040
      X2              =   1185
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line lnWindSpd 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   2040
      X2              =   1185
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line lnTemp 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   1
      X1              =   5565
      X2              =   4260
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu itmMetric 
         Caption         =   "&Metric"
         Checked         =   -1  'True
      End
      Begin VB.Menu itmImperial 
         Caption         =   "&Imperial"
      End
      Begin VB.Menu itmSep1 
         Caption         =   "-"
      End
      Begin VB.Menu itmRefNow 
         Caption         =   "&Refresh Now"
      End
      Begin VB.Menu itmRefWhen 
         Caption         =   "Refresh &When..."
      End
      Begin VB.Menu itmHistGraph 
         Caption         =   "&Show History Graph..."
      End
      Begin VB.Menu itmSep2 
         Caption         =   "-"
      End
      Begin VB.Menu itmViewMetar 
         Caption         =   "&View Metar Strip..."
      End
      Begin VB.Menu itmLocDetails 
         Caption         =   "V&iew Location Details..."
      End
      Begin VB.Menu itmLocation 
         Caption         =   "&Choose Location..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intUnits As Units

Private Sub Form_Load()
    Location = GetSetting("WindRainSky", "DefaultLocation", "ICAO", "CYHM")
    mRefresh
End Sub

Private Sub itmExit_Click()
    Unload Me
End Sub

Private Sub itmHistGraph_Click()
    MsgBox "To be implemented in the future.", vbInformation + vbOKOnly, "View History Graph"
End Sub

Private Sub itmImperial_Click()
    If Not itmImperial.Checked Then
        itmMetric.Checked = False
        itmImperial.Checked = True
    End If
    intUnits = Imperial
    mRefresh
End Sub

Private Sub itmLocation_Click()
    frmAirport.Show
End Sub

Private Sub itmLocDetails_Click()
    MsgBox GetMetarData(Conditions, intUnits), vbInformation + vbOKOnly, "Location Details"
End Sub

Private Sub itmMetric_Click()
    If Not itmMetric.Checked Then
        itmMetric.Checked = True
        itmImperial.Checked = False
    End If
    intUnits = Metric
    mRefresh
End Sub

Public Sub mRefresh(Optional fromWeb As Boolean = False)
    Dim intTemp As Integer
    'Refresh from Web or just with existing data
        If fromWeb Then
            GetMetar
        End If
        If NoData Then
            MsgBox "No data available for this location, sorry.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
    'Visibility
        lblVis.Caption = GetMetarData(Visibility, intUnits) & IIf(intUnits = Metric, " Kilometers", " Miles")
    'Barometer
        SetBarDigits
    'Temperature
        intTemp = GetMetarData(Temperature, Metric)
        lblMetric.Caption = intTemp
        lnTemp(0).X2 = lnTemp(0).X1 + Sin(3.14159 * (Round(intTemp / 120 * 360) - 17) / 180) * 1305
        lnTemp(0).Y2 = lnTemp(0).Y1 - Cos(3.14159 * (Round(intTemp / 120 * 360) - 17) / 180) * 1305
        intTemp = GetMetarData(Temperature, Imperial)
        lblImperial.Caption = intTemp
    'Dewpoint - BLUE POINTER ON GAUGE
        intTemp = GetMetarData(DewPoint, Metric)
        lnTemp(1).X2 = lnTemp(1).X1 + Sin(3.14159 * (Round(intTemp / 120 * 360) - 17) / 180) * 1305
        lnTemp(1).Y2 = lnTemp(1).Y1 - Cos(3.14159 * (Round(intTemp / 120 * 360) - 17) / 180) * 1305
    'WindSpeed
        intTemp = GetMetarData(WindSpeed, Imperial)
        lblWSpeed.Caption = GetMetarData(WindSpeed, intUnits) & IIf(intUnits = Metric, " Km/h", " mph")
        lnWindSpd(0).X2 = lnWindSpd(0).X1 + Sin(3.14159 * (Round(intTemp / 195 * 360) - 140) / 180) * 855
        lnWindSpd(0).Y2 = lnWindSpd(0).Y1 - Cos(3.14159 * (Round(intTemp / 195 * 360) - 140) / 180) * 855
    'WindGusts - RED POINTER ON GAUGE
        intTemp = GetMetarData(WindGusts, Imperial)
        lblWGust.Caption = GetMetarData(WindGusts, intUnits) & IIf(intUnits = Metric, " Km/h", " mph")
        lnWindSpd(1).X2 = lnWindSpd(1).X1 + Sin(3.14159 * (Round(intTemp / 195 * 360) - 140) / 180) * 855
        lnWindSpd(1).Y2 = lnWindSpd(1).Y1 - Cos(3.14159 * (Round(intTemp / 195 * 360) - 140) / 180) * 855
    'WindDirection
        intTemp = GetMetarData(WindDirection, Imperial)
        lnWindDir.X2 = lnWindDir.X1 + Sin(3.14159 * intTemp / 180) * 585
        lnWindDir.Y2 = lnWindDir.Y1 - Cos(3.14159 * intTemp / 180) * 585
    'Location
        lblLocation.Caption = Location
    'Conditions
        lblConditions.Caption = GetMetarData(Conditions, intUnits)
End Sub

Private Sub itmRefNow_Click()
    mRefresh True
End Sub

Sub SetBarDigits()
    Dim strTmp As String
    Dim i As Integer, j As Integer
    
    strTmp = Trim(Str(GetMetarData(Barometer, intUnits)))
    If Len(strTmp) < 5 Then
        If InStr(1, strTmp, ".") Then
            For i = 1 To 5 - Len(strTmp)
                strTmp = strTmp & "0"
            Next i
        Else
            strTmp = strTmp & ".0"
        End If
    End If
    j = 1
    For i = 0 To 3
        If j < 4 And Mid$(strTmp, j + 1, 1) = "." Then
            Bar7Seg(i).DrawLED Mid$(strTmp, j, 2)
            j = j + 1
        Else
            Bar7Seg(i).DrawLED Mid$(strTmp, j, 1)
        End If
        j = j + 1
    Next i
    
End Sub

Private Sub itmRefWhen_Click()
    MsgBox "To be implemented in the future.", vbInformation + vbOKOnly, "Refresh When"
End Sub

Private Sub itmViewMetar_Click()
    MsgBox Metar, vbInformation + vbOKOnly, "Metar Strip"
End Sub
