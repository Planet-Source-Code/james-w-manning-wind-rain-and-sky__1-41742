Attribute VB_Name = "mdlMain"
Option Explicit

Public Metar As String
Public MetarTokens() As String
Public Location As String
Public NoData As Boolean

Enum MetType
    Temperature = 0
    Barometer = 1
    Visibility = 2
    WindSpeed = 3
    WindDirection = 4
    WindGusts = 5
    Conditions = 6
    DewPoint = 7
End Enum

Enum Units
    Metric = 0
    Imperial = 1
End Enum

Public Function GetMetarData(Token As MetType, dUnits As Units) As Variant
    Dim sToken As String
    Dim intTmp As Integer
    Dim dblTmp As Double
    Dim strTmp As String
    Dim i As Integer
    
    If Metar = vbNullString Then
        If Location = vbNullString Then
            GetMetar "CYHM"
        Else
            GetMetar
        End If
    End If
    
    For i = 0 To UBound(MetarTokens(), 1)
        Select Case Token
            Case Temperature, DewPoint
                If InStr(1, MetarTokens(i), "/") <> 0 Then
                    sToken = MetarTokens(i)
                    Exit For
                End If
            Case Barometer
                If InStr(1, MetarTokens(i), "A") <> 0 And Len(MetarTokens(i)) = 5 Then
                    sToken = MetarTokens(i)
                    Exit For
                End If
            Case Visibility
                If InStr(1, MetarTokens(i), "SM") <> 0 Then
                    If Len(MetarTokens(i - 1)) = 1 And InStr(1, MetarTokens(i), "/") <> 0 Then
                        sToken = MetarTokens(i - 1) & " " & MetarTokens(i)
                    Else
                        sToken = MetarTokens(i)
                    End If
                    Exit For
                End If
            Case WindSpeed, WindDirection, WindGusts
                If InStr(1, MetarTokens(i), "K") <> 0 Then
                    sToken = MetarTokens(i)
                    Exit For
                End If
        End Select
    Next i
    
    If sToken <> vbNullString Or Token = Conditions Then
        Select Case Token
            Case Temperature
                If Mid$(sToken, 1, 1) = "M" Then
                    intTmp = -Val(Mid$(sToken, 2, 2))
                Else
                    intTmp = Val(Mid$(sToken, 1, 2))
                End If
                If dUnits = Imperial Then
                    intTmp = Round((1.8 * intTmp) + 32, 0)
                End If
                GetMetarData = intTmp
            Case DewPoint
                If Mid$(sToken, InStr(1, sToken, "/") + 1, 1) = "M" Then
                    intTmp = -Val(Mid$(sToken, InStr(1, sToken, "/") + 2, 2))
                Else
                    intTmp = Val(Mid$(sToken, InStr(1, sToken, "/") + 1, 2))
                End If
                If dUnits = Imperial Then
                    intTmp = Round((1.8 * intTmp) + 32, 0)
                End If
                GetMetarData = intTmp
            Case Barometer
                dblTmp = Val(Mid$(sToken, 2, 2) & "." & Mid$(sToken, 4, 2))
                If dUnits = Metric Then
                    dblTmp = Round(dblTmp * 3.38638815789, 1)
                End If
                GetMetarData = dblTmp
            Case Visibility
                If InStr(1, sToken, "/") <> 0 Then
                    If InStr(1, sToken, " ") <> 0 Then
                        dblTmp = Mid$(sToken, 1, 1)
                        dblTmp = dblTmp + Val("0." & (Val(Mid$(sToken, 3, 1)) / Val(Mid$(sToken, 5, 1))))
                    Else
                        dblTmp = Val("0." & (Val(Mid$(sToken, 3, 1)) / Val(Mid$(sToken, 5, 1))))
                    End If
                Else
                    dblTmp = Val(Mid$(sToken, 1, InStr(1, sToken, "SM") - 1))
                End If
                If dUnits = Metric Then
                    dblTmp = Round(dblTmp * 1.609344, 0)
                End If
                GetMetarData = dblTmp
            Case WindSpeed
                If InStr(1, sToken, "G") = 0 Then
                    intTmp = Val(Mid$(sToken, 4, InStr(4, sToken, "K") - 1))
                Else
                    intTmp = Val(Mid$(sToken, InStr(4, sToken, "G") + 1, InStr(InStr(4, sToken, "G") + 1, sToken, "K") - 1))
                End If
                If dUnits = Metric Then
                    intTmp = intTmp * 1.852
                Else
                    intTmp = intTmp * 1.15
                End If
                GetMetarData = intTmp
            Case WindDirection
                intTmp = Val(Mid$(sToken, 1, 3))
                GetMetarData = intTmp
            Case WindGusts
                If InStr(1, sToken, "G") <> 0 Then
                    intTmp = Val(Mid$(sToken, 4, InStr(4, sToken, "G") - 1))
                    If dUnits = Metric Then
                        intTmp = intTmp * 1.852
                    Else
                        intTmp = intTmp * 1.15
                    End If
                Else
                    intTmp = 0
                End If
                GetMetarData = intTmp
            Case Conditions
                strTmp = ""
                For i = 0 To UBound(MetarTokens(), 1)
                    If Mid$(MetarTokens(i), 1, 1) = "+" Then
                        strTmp = strTmp & "Heavy "
                    ElseIf Mid$(MetarTokens(i), 1, 1) = "-" Then
                        strTmp = strTmp & "Light "
                    End If
                    If InStr(1, MetarTokens(i), "OVC") <> 0 Then
                        If dUnits = Imperial Then
                            strTmp = strTmp & "Overcast clouds at " & Val(Mid$(MetarTokens(i), 4, 3)) * 100 & " feet. "
                        Else
                            strTmp = strTmp & "Overcast clouds at " & Round(Val(Mid$(MetarTokens(i), 4, 3)) * 100 * 0.3048, 0) & " meters. "
                        End If
                    End If
                    If InStr(1, MetarTokens(i), "SCT") <> 0 Then
                        If dUnits = Imperial Then
                            strTmp = strTmp & "Scattered clouds at " & Val(Mid$(MetarTokens(i), 4, 3)) * 100 & " feet. "
                        Else
                            strTmp = strTmp & "Scattered clouds at " & Round(Val(Mid$(MetarTokens(i), 4, 3)) * 100 * 0.3048, 0) & " meters. "
                        End If
                    End If
                    If InStr(1, MetarTokens(i), "FEW") <> 0 Then
                        If dUnits = Imperial Then
                            strTmp = strTmp & "A few clouds at " & Val(Mid$(MetarTokens(i), 4, 3)) * 100 & " feet. "
                        Else
                            strTmp = strTmp & "A few clouds at " & Round(Val(Mid$(MetarTokens(i), 4, 3)) * 100 * 0.3048, 0) & " meters. "
                        End If
                    End If
                    If InStr(1, MetarTokens(i), "BKN") <> 0 Then
                        If dUnits = Imperial Then
                            strTmp = strTmp & "Broken clouds at " & Val(Mid$(MetarTokens(i), 4, 3)) * 100 & " feet. "
                        Else
                            strTmp = strTmp & "Broken clouds at " & Round(Val(Mid$(MetarTokens(i), 4, 3)) * 100 * 0.3048, 0) & " meters. "
                        End If
                    End If
                    If InStr(1, MetarTokens(i), "CLR") <> 0 Then
                        strTmp = strTmp & "Clear sky. "
                    End If
                    If InStr(1, MetarTokens(i), "FG") <> 0 Then
                        strTmp = strTmp & "Foggy. "
                    End If
                    If InStr(1, MetarTokens(i), "HZ") <> 0 Then
                        strTmp = strTmp & "Hazy. "
                    End If
                    If InStr(1, MetarTokens(i), "DZ") <> 0 Then
                        strTmp = strTmp & "Drizzle. "
                    End If
                    If Mid$(MetarTokens(i), 1, 1) = "+" Or Mid$(MetarTokens(i), 1, 1) = "-" Then
                        If InStr(2, MetarTokens(i), "RA") <> 0 Then
                            strTmp = strTmp & "Rain. "
                        End If
                        If InStr(2, MetarTokens(i), "SN") <> 0 Then
                            strTmp = strTmp & "Snow. "
                        End If
                        If InStr(2, MetarTokens(i), "SH") <> 0 Then
                            strTmp = strTmp & "Showers. "
                        End If
                    Else
                        If InStr(1, MetarTokens(i), "RA") <> 0 Then
                            strTmp = strTmp & "Moderate Rain. "
                        End If
                        If InStr(1, MetarTokens(i), "SN") <> 0 Then
                            strTmp = strTmp & "Moderate Snow. "
                        End If
                        If InStr(1, MetarTokens(i), "SH") <> 0 Then
                            strTmp = strTmp & "Moderate Showers. "
                        End If
                    End If
                Next i
                GetMetarData = strTmp
        End Select
    End If
End Function

Public Sub GetMetar(Optional mAirport As String = vbNullString)
    Dim i As Integer
    Dim Airport As String
    On Error GoTo GetMetar_err
    If mAirport = vbNullString Then
        If Len(Location) > 4 Then
            Airport = Mid$(Location, 1, 4)
        Else
            Airport = Location
        End If
    Else
        If Len(mAirport) > 4 Then
            Airport = Mid$(mAirport, 1, 4)
        Else
            Airport = mAirport
        End If
    End If
    
    With frmMain.Inet1
        Metar = .OpenURL("http://weather.noaa.gov/cgi-bin/mgetmetar.pl?cccc=" & Airport, icString)
    End With
    i = InStr(1, Metar, "<P>The observation is:</P>")
    i = InStr(i, Metar, Mid$(Airport, 1, 4))
    If InStr(1, Metar, "RMK") <> 0 Then
        Metar = Mid$(Metar, i, InStr(i, Metar, "RMK") - i)
    Else
        Metar = Mid$(Metar, i, InStr(i, Metar, Chr(10)) - i)
    End If
    
    MetarTokens = Split(Mid$(Metar, 6), " ")

    NoData = False

    Exit Sub

GetMetar_err:

    NoData = True
    
End Sub
