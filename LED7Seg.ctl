VERSION 5.00
Begin VB.UserControl LED7Seg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   41
   ToolboxBitmap   =   "LED7Seg.ctx":0000
End
Attribute VB_Name = "LED7Seg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum lColors
    lRed = 0
    lGreen = 1
    lBlue = 2
    lPurple = 3
    lYellow = 4
    lWhite = 5
End Enum

Private bColor As Long
Private dColor As Long
'Default Property Values:
Const m_def_Color = lRed
'Property Variables:
Dim m_Color As lColors



Private Sub GetColor(Clr As lColors)
    Select Case Clr
        Case lRed
            bColor = RGB(255, 0, 0)
            dColor = RGB(95, 0, 0)
        Case lGreen
            bColor = RGB(0, 255, 0)
            dColor = RGB(0, 95, 0)
        Case lBlue
            bColor = RGB(0, 0, 255)
            dColor = RGB(0, 0, 95)
        Case lPurple
            bColor = RGB(255, 0, 255)
            dColor = RGB(95, 0, 95)
        Case lYellow
            bColor = RGB(255, 255, 0)
            dColor = RGB(95, 95, 0)
        Case lWhite
            bColor = RGB(255, 255, 255)
            dColor = RGB(95, 95, 95)
    End Select
End Sub

Private Sub lTop(lColor As Long)
    'Top
    Line (2, 2)-(35, 2), lColor
    Line (3, 3)-(34, 3), lColor
    Line (4, 4)-(33, 4), lColor
End Sub

Private Sub lTopLeft(lColor As Long)
    'Top Left
    Line (2, 5)-(2, 32), lColor
    Line (3, 6)-(3, 31), lColor
    Line (4, 7)-(4, 30), lColor
End Sub

Private Sub lTopRight(lColor As Long)
    'Top Right
    Line (37, 4)-(37, 32), lColor
    Line (36, 5)-(36, 31), lColor
    Line (35, 6)-(35, 30), lColor
End Sub

Private Sub lMiddle(lColor As Long)
    'Middle
    Line (4, 32)-(33, 32), lColor
    Line (3, 33)-(34, 33), lColor
    Line (2, 34)-(35, 34), lColor
    Line (3, 35)-(34, 35), lColor
    Line (4, 36)-(33, 36), lColor
End Sub

Private Sub lBottomLeft(lColor As Long)
    'Bottom Left
    Line (2, 36)-(2, 63), lColor
    Line (3, 37)-(3, 62), lColor
    Line (4, 38)-(4, 61), lColor
End Sub

Private Sub lBottomRight(lColor As Long)
    'Bottom Right
    Line (37, 36)-(37, 63), lColor
    Line (36, 37)-(36, 62), lColor
    Line (35, 38)-(35, 61), lColor
End Sub

Private Sub lBottom(lColor As Long)
    'Bottom
    Line (3, 64)-(33, 64), lColor
    Line (2, 65)-(34, 65), lColor
    Line (1, 66)-(35, 66), lColor
End Sub

Private Sub lDec(lColor As Long)
    'Decimal
    Line (42, 64)-(44, 66), lColor, BF
End Sub

Public Sub DrawLED(Optional Char As Variant)
    Dim i As Integer
    Dim Tmp As String
    GetColor m_Color
    DrawWidth = 2
    'Draw initially
    lTop dColor
    lTopLeft dColor
    lTopRight dColor
    lMiddle dColor
    lBottomLeft dColor
    lBottomRight dColor
    lBottom dColor
    lDec dColor
    'Then draw lit segments
    If Not IsMissing(Char) Then
        If Not IsNumeric(Char) Then
            For i = 1 To Len(Char)
                Tmp = Mid$(Char, i, 1)
                Select Case UCase(Tmp)
                    Case "A"
                        lTop bColor
                        lTopLeft bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                    Case "B"
                        lTopLeft bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "C"
                        lTop bColor
                        lTopLeft bColor
                        lBottomLeft bColor
                        lBottom bColor
                    Case "D"
                        lTopRight bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "E"
                        lTop bColor
                        lTopLeft bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottom bColor
                    Case "F"
                        lTop bColor
                        lTopLeft bColor
                        lMiddle bColor
                        lBottomLeft bColor
                End Select
                If i > 0 And Len(Char) > 1 Then
                    If Mid$(Char, i + 1, 1) = "." Then
                        lDec (bColor)
                        Exit For
                    End If
                End If
            Next i
        Else
            For i = 1 To Len(Char)
                Tmp = Mid$(Char, i, 1)
                Select Case Tmp
                    Case "0"
                        lTop bColor
                        lTopLeft bColor
                        lTopRight bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "1"
                        lTopRight bColor
                        lBottomRight bColor
                    Case "2"
                        lTop bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottom bColor
                    Case "3"
                        lTop bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "4"
                        lTopLeft bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomRight bColor
                    Case "5"
                        lTop bColor
                        lTopLeft bColor
                        lMiddle bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "6"
                        lTop bColor
                        lTopLeft bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "7"
                        lTop bColor
                        lTopRight bColor
                        lBottomRight bColor
                    Case "8"
                        lTop bColor
                        lTopLeft bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomLeft bColor
                        lBottomRight bColor
                        lBottom bColor
                    Case "9"
                        lTop bColor
                        lTopLeft bColor
                        lTopRight bColor
                        lMiddle bColor
                        lBottomRight bColor
                        lBottom bColor
                End Select
                If i > 0 And Len(Char) > 1 Then
                    If Mid$(Char, i + 1, 1) = "." Then
                        lDec (bColor)
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
    UserControl.Refresh
End Sub
Private Sub UserControl_Initialize()
    Width = 715
    Height = 1020
    DrawLED
End Sub

Private Sub UserControl_Resize()
    Width = 715
    Height = 1020
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Color() As lColors
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As lColors)
    m_Color = New_Color
    PropertyChanged "Color"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color = PropBag.ReadProperty("Color", m_def_Color)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
End Sub

