VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmOth 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00973E1E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversi"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4560
   Icon            =   "frmOth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Text            =   "255"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Text            =   "128"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Text            =   "51"
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Text            =   "51"
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Text            =   "255"
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   4440
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picV 
      AutoRedraw      =   -1  'True
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   0
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   0
      Top             =   240
      Width           =   4350
   End
   Begin VB.PictureBox picB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H006C2D16&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3810
      Left            =   90
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   1
      Top             =   360
      Width           =   4110
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1155
   End
   Begin VB.PictureBox picA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   6060
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   270
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   4290
      Width           =   4335
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   4305
      Width           =   4335
   End
   Begin VB.Menu game 
      Caption         =   "&Game"
      Index           =   1
      Begin VB.Menu new 
         Caption         =   "&New"
      End
      Begin VB.Menu colourplayer 
         Caption         =   "&Player Colours"
         Begin VB.Menu p1 
            Caption         =   "Player &1"
         End
         Begin VB.Menu p2 
            Caption         =   "Player &2"
         End
      End
      Begin VB.Menu colour 
         Caption         =   "&Background Colour..."
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Index           =   2
      Begin VB.Menu plaiyingmga 
         Caption         =   "&Playing the Game"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About Reversi..."
      End
   End
End
Attribute VB_Name = "frmOth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Othello
'
' Used names:
' picV   = playing field
' picA   = background
' picB   = buffer
' "mij"  = the computer as a player ("mijn")
' "jij"  = the user (you) as a player ("jouw")
' R,K    = row, column

DefInt A-Z                       ' not declared var are integers
Dim C(9, 9)                      ' playing field : White, Black or Empty
Dim P(9, 9)                      ' priority of the piece-positions
Dim RIx(8), RIy(8)               ' search directions in x and y values
Dim CK(9), CR(9), AantalCRKs     ' storage found possibilities "mij"
Dim R1, K1, R2, K2               ' search-frame border - enlarges during the game

Const Rood = 255                 ' red - used colors
Dim Wit As Long    ' white
Dim Zwart As Long               ' black
Const Grijs = 14671839           ' gray

Const bWit = 0                   ' status piece-poss. on the playing field
Const bZwart = 3
Const bLeeg = 6

Dim YourColor, MyColor           ' used black/white choise

Dim IAmWaitingForYou As Boolean
Dim jR, jK                       ' jouw row/column
Dim mR, mK                       ' mijn row/column

Dim W                            ' size of one piece-poss. (pixels)

Dim DoStop As Boolean            ' always True when not playing
                                 ' serves to stop the game
Dim Beurt                        ' count each players-turn (max. 60)

Dim ZeurInfo As Boolean          ' this info serves to cheat and to test the progr.

Dim ie As String                 ' internet browser-help

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


' sets the priorities according to the current state of the playing field
Private Sub UpdatePriorities(prior)
   Dim I, J
   
   If C(2, 2) = YourColor And C(3, 1) = MyColor Or C(1, 3) = MyColor Then P(3, 1) = 1: P(1, 3) = 1 ': Stop
   If C(7, 7) = YourColor And (C(8, 6) = MyColor Or C(6, 8) = MyColor) Then P(8, 6) = 1: P(6, 8) = 1 ': Stop
   If C(2, 7) = YourColor And (C(1, 6) = MyColor Or C(3, 8) = MyColor) Then P(1, 6) = 1: P(3, 8) = 1 ': Stop
   If C(7, 2) = YourColor And (C(6, 1) = MyColor Or C(8, 3) = MyColor) Then P(6, 1) = 1: P(8, 3) = 1 ': Stop
   
   ' examin if chosen cell is on the border
   If Not (IsBorderCell(jR, jK) Or IsBorderCell(mR, mK)) Then Exit Sub
   For J = 1 To 8 Step 7
      For I = 2 To 7
         If C(I, J) = MyColor Then P(I + 1, J) = 21: P(I - 1, J) = 21
         If C(J, I) = MyColor Then P(J, I + 1) = 21: P(J, I - 1) = 21
      Next I
      For I = 2 To 7
         If C(I, J) = YourColor Then P(I + 1, J) = 2: P(I - 1, J) = 2
         If C(J, I) = YourColor Then P(J, I + 1) = prior: P(J, I - 1) = 2
      Next I
   Next J
   P(1, 2) = 1: P(1, 7) = 1: P(2, 1) = 1: P(7, 1) = 1
   P(2, 8) = 1: P(7, 8) = 1: P(8, 2) = 1: P(8, 7) = 1
   For I = 2 To 7
      If C(1, I - 1) = YourColor And C(1, I + 1) = YourColor Then P(1, I) = 2
      If C(8, I - 1) = YourColor And C(8, I + 1) = YourColor Then P(8, I) = 2
      If C(I - 1, 1) = YourColor And C(I + 1, 1) = YourColor Then P(I, 1) = 2
      If C(I - 1, 8) = YourColor And C(I + 1, 8) = YourColor Then P(I, 8) = 2
   Next I
   Dim Ic
   For J = 1 To 8 Step 7
     For I = 4 To 8
       If C(J, I) = MyColor Then
         Ic = I - 1
         If C(J, Ic) <> bLeeg Then
           While C(J, Ic) = YourColor: Ic = Ic - 1: Wend
           If Ic > 0 Then
              If C(J, Ic) = bLeeg And Ic <> 0 And Not (C(J, I + 1) = YourColor And C(J, Ic - 1) = bLeeg) Then P(J, Ic) = 26
              End If
           End If
         End If
       If C(I, J) = MyColor Then
         Ic = I - 1
         If C(Ic, J) <> bLeeg Then
           While C(Ic, J) = YourColor: Ic = Ic - 1: Wend
           If Ic > 0 Then
              If C(Ic, J) = bLeeg And Ic <> 0 And Not (C(I + 1, J) = YourColor And C(Ic - 1, J) = bLeeg) Then P(Ic, J) = 26
              End If
           End If
         End If
     Next I
     For I = 1 To 5
       If C(J, I) = MyColor Then
         Ic = I + 1
         If C(J, Ic) <> bLeeg Then
           While C(J, Ic) = YourColor: Ic = Ic + 1: Wend
           If Ic < 9 Then
              If C(J, Ic) = bLeeg And Not (C(J, I - 1) = YourColor And C(J, Ic + 1) = bLeeg) Then P(J, Ic) = 26
              End If
           End If
         End If
       If C(I, J) = MyColor Then
         Ic = I + 1
         If C(Ic, J) = bLeeg Then GoTo L2440
         While C(Ic, J) = YourColor: Ic = Ic + 1: Wend
         If Ic < 9 Then
            If C(Ic, J) = bLeeg And Not (C(I - 1, J) = YourColor And C(Ic + 1, J) = bLeeg) Then P(Ic, J) = 26
            End If
       End If
     Next I
   Next J
'-----
L2440:
   If C(1, 1) = MyColor Then For I = 2 To 6: P(1, I) = 20: P(I, 1) = 20: Next I
   If C(1, 8) = MyColor Then For I = 2 To 6: P(I, 8) = 20: P(1, 9 - I) = 20: Next I
   If C(8, 1) = MyColor Then For I = 2 To 6: P(9 - I, 1) = 20: P(8, I) = 20: Next I
   If C(8, 8) = MyColor Then For I = 3 To 6: P(I, 8) = 20: P(8, I) = 20: Next I
   If C(1, 1) <> bLeeg Then P(2, 2) = 5
   If C(1, 8) <> bLeeg Then P(2, 7) = 5
   If C(8, 1) <> bLeeg Then P(7, 2) = 5
   If C(8, 8) <> bLeeg Then P(7, 7) = 5
   P(1, 1) = 30: P(1, 8) = 30: P(8, 1) = 30: P(8, 8) = 30
   For I = 3 To 6
      If C(1, I) = MyColor Then P(2, I) = 4
      If C(8, I) = MyColor Then P(7, 1) = 4
      If C(I, 1) = MyColor Then P(1, 2) = 4
      If C(I, 8) = MyColor Then P(1, 7) = 4
   Next I
   If C(7, 1) = YourColor And C(4, 1) = MyColor And C(6, 1) = bLeeg And C(5, 1) = bLeeg Then P(6, 1) = 26
   If C(1, 7) = YourColor And C(1, 4) = MyColor And C(1, 6) = bLeeg And C(1, 5) = bLeeg Then P(1, 6) = 26
   If C(2, 1) = YourColor And C(5, 1) = MyColor And C(3, 1) = bLeeg And C(4, 1) = bLeeg Then P(3, 1) = 26
   If C(1, 2) = YourColor And C(1, 5) = MyColor And C(1, 3) = bLeeg And C(1, 4) = bLeeg Then P(1, 3) = 26
   If C(8, 2) = YourColor And C(8, 5) = MyColor And C(8, 3) = bLeeg And C(5, 1) = bLeeg Then P(6, 1) = 26
   If C(2, 8) = YourColor And C(5, 8) = MyColor And C(3, 8) = bLeeg And C(1, 5) = bLeeg Then P(1, 6) = 26
   If C(8, 7) = YourColor And C(8, 4) = MyColor And C(8, 5) = bLeeg And C(8, 6) = bLeeg Then P(8, 6) = 26
   If C(7, 8) = YourColor And C(4, 8) = MyColor And C(8, 8) = bLeeg And C(6, 8) = bLeeg Then P(6, 8) = 26
End Sub

' it's the computers turn, looks here for the best choise
Private Sub SearchBestCells(priorMax, VrMax)
   Dim I, R, K
   Dim Vr, VrCel  ' angulars, counter or total of the cell
   Dim KK, RR     ' start R en K, zoek richting lus
   
   ' replace search frame ?
   If Not (R1 * K1 = 1 And R2 * K2 = 64) Then
      For I = 2 To 7
         If C(2, I) <> bLeeg Then R1 = 1
         If C(7, I) <> bLeeg Then R2 = 8
         If C(I, 2) <> bLeeg Then K1 = 1
         If C(I, 7) <> bLeeg Then K2 = 8
      Next I
      If ZeurInfo Then picV.Line (K1 * W, R1 * W)-(K2 * W + W, R2 * W + W), RGB(255, 0, 0), B
      End If

   ' go over all cells, select empty one's and examin them
   VrCel = 0
   priorMax = 0: VrMax = 0
   For R = R1 To R2: For K = K1 To K2
      If C(R, K) = bLeeg Then
         If P(R, K) < priorMax Then GoTo ZBCvolgende:
         VrCel = 0
         For I = 0 To 8                    ' all directions
            Vr = 0
            KK = K: RR = R
            KK = KK + RIx(I)
            RR = RR + RIy(I)
            While C(RR, KK) = YourColor    ' count cells in this
               Vr = Vr + 1                 ' direction
               KK = KK + RIx(I)
               RR = RR + RIy(I)
            Wend
            If C(RR, KK) <> bLeeg And Vr <> 0 Then VrCel = VrCel + Vr
         Next I
         If VrCel <> 0 Then
            If P(R, K) > priorMax Then
               priorMax = P(R, K)
               AantalCRKs = 0              ' higher priority
               VrMax = VrCel               ' so restart
               CK(0) = K: CR(0) = R
               GoTo ZBCvolgende:
               End If
            If VrMax > VrCel Then GoTo ZBCvolgende:
            If VrMax < VrCel Then
               AantalCRKs = 0              ' more gain
               VrMax = VrCel               ' so restart
               CK(0) = K
               CR(0) = R
               Else
               AantalCRKs = AantalCRKs + 1 ' same priority
               CK(AantalCRKs) = K          ' same gain
               CR(AantalCRKs) = R
               End If
            End If
         End If
            
ZBCvolgende:
      DoEvents
   Next K: Next R
   If ZeurInfo Then For I = 0 To AantalCRKs: ShowLittleCircle CR(I), CK(I), Rood: Next I
      
End Sub

' display message at the bottom
Private Sub Message(txt As String, Kleur As Long)
   lblMessage(0).ForeColor = Kleur
   lblMessage(0).Caption = txt
   lblMessage(1).Caption = txt
End Sub

' serves to display via the cmdVorigeStand button
' the previous state (mostly of "mij")
Private Sub BufferPrevState()
   picB.Picture = picV.Image
End Sub

Private Function IsOutOfField(R, K) As Boolean
   IsOutOfField = IIf((R < 1 Or R > 8 Or K < 1 Or K > 8), True, False)
End Function

Private Function IsBorderCell(R, K) As Boolean
   IsBorderCell = IIf((R = 1 Or R = 8 Or K = 1 Or K = 8), True, False)
End Function

' count turns and check?
Private Function PlayTurnsLeft() As Boolean
   Beurt = Beurt + 1
   If Beurt = 60 Then
      CalcWinner "All turns are played."
      ResetGame
      PlayTurnsLeft = False
      Else
      PlayTurnsLeft = True
      End If
End Function

Private Function YouHaveToPass() As Boolean
   Dim Vr, I, R, K, RR, KK
   
   YouHaveToPass = True
   For R = 1 To 8: For K = 1 To 8   ' of all cells
      If C(R, K) = bLeeg Then       ' look for empty one's
         For I = 1 To 8             ' in all directions
            Vr = 0: KK = K: RR = R  ' remember RR,KK startcell
            Do
               RR = RR + RIy(I): KK = KK + RIx(I)
               If IsOutOfField(RR, KK) Then Exit Do
               If C(RR, KK) = MyColor Then Vr = Vr + 1
            Loop Until C(RR, KK) <> MyColor
            If C(RR, KK) = YourColor And Vr > 0 Then
               YouHaveToPass = False
               If ZeurInfo Then
                  ShowLittleCircle R, K, RGB(255, 0, 0)
                  Else
                  Exit Function
                  End If
               End If
         Next I
         End If
   Next K, R
End Function



     

Private Sub PrintAt(pic As PictureBox, txt As String, X, Y, Color As Long)
   With pic
      .ForeColor = Color: .CurrentX = X: .CurrentY = Y
   End With
   pic.Print txt
End Sub

Private Sub CalcWinner(sMsg As String)
   Dim R, K, mTot, jTot
   Dim sTot As String, sWinnaar As String
   
   DoStop = True
   mTot = 0: jTot = 0
   For R = 1 To 8: For K = 1 To 8
      If C(R, K) = MyColor Then mTot = mTot + 1 Else jTot = jTot + 1
   Next K, R
   
   If mTot = jTot Then
      sTot = "There's an equal number of player 1 and player 2 blocks."
      sWinnaar = "Deadlock"
      End If
   If mTot < jTot Then
      sTot = "You have" & jTot & " squares, while the pc has " & mTot & "."
      sWinnaar = "You'we won!"
      End If
   If mTot > jTot Then
      sTot = "You have" & jTot & " squares, while the pc has " & mTot & "."
      sWinnaar = "You've lost."
      End If
   MsgBox sMsg & vbCrLf & vbCrLf & sTot & vbCrLf & vbCrLf & sWinnaar
      
End Sub

Private Sub ResetGame()
   cmdStart.Caption = "&Start"
End Sub

Private Sub SetConstants()
   Dim I, R, K
   Dim txt As String
   Dim RK(1 To 8) As String
   
   ' "empty" outerborder
   For I = 0 To 9
      C(I, 0) = bLeeg
      C(0, I) = bLeeg
      C(9, I) = bLeeg
      C(I, 9) = bLeeg
   Next I
   
   ' initial search-frame
   R1 = 2: K1 = 2: R2 = 7: K2 = 7
   
   ' directions xy
   For I = 1 To 8
      RIx(I) = Choose(I, 1, 1, 0, -1, -1, -1, 0, 1)
      RIy(I) = Choose(I, 0, 1, 1, 1, 0, -1, -1, -1)
   Next I
   
   ' initial priority values
   RK(1) = "30 01 20 10 10 20 01 30"
   RK(2) = "01 01 03 03 03 03 01 01"
   RK(3) = "20 03 05 05 05 05 03 20"
   RK(4) = "10 03 05 00 00 05 03 10"
   RK(5) = "10 03 05 00 00 05 03 10"
   RK(6) = "20 03 05 05 05 05 03 20"
   RK(7) = "01 01 03 03 03 03 01 01"
   RK(8) = "30 01 20 10 10 20 01 30"
   For R = 1 To 8: For K = 1 To 8
      P(R, K) = Val(Mid(RK(R), (K - 1) * 3 + 1, 3))
      C(R, K) = bLeeg
   Next K: Next R
      
   ' put first 4 pieces
   C(4, 4) = 3: C(4, 5) = 0
   C(5, 4) = 0: C(5, 5) = 3
   
   ' ready to start 60 turns
   Beurt = 0
End Sub

'
Private Sub Play()
   Dim priorMax            ' current highest priority
   Dim Winst               ' gain
   Dim Passen As Boolean   ' True = "jij" has to pass
   Dim CRKrnd              ' chosing "mijn" RK via this random-index
   
   SetConstants
   ShowField

   If MyColor = bZwart Then GoTo AanMij:
   
AanJou:
   Message " ", Zwart
   ShowWhoIsOn "jou"
   Do
      WaitingForYou
      If DoStop = True Then Exit Sub
      If C(jR, jK) <> bLeeg Then MsgBox "This place is already occupied!", Rood: Wait 1: Message " ", Zwart
   Loop Until C(jR, jK) = bLeeg
   ShowCross jR, jK
   Winst = GainPieces(jR, jK, YourColor)
      If Winst = 0 Then Screen.MousePointer = 13
   If Winst = 0 Then MsgBox "Not a valid move", Rood: Wait 1: ShowField: GoTo AanJou:
   If Winst = 1 Then Screen.MousePointer = 2
   C(jR, jK) = YourColor
   ShowGain Winst
   BufferPrevState
   ShowField
   If PlayTurnsLeft() = False Then Exit Sub
   '------------------------------------
AanMij:
   Message " ", Zwart
   ShowWhoIsOn "mij"
   UpdatePriorities priorMax
   SearchBestCells priorMax, Winst         ' are both first set to null and then filled in
   ' found a good cell
   If priorMax > 0 Then
      If AantalCRKs > 0 Then               ' more possibilities
         AantalCRKs = AantalCRKs + 1       ' see manner of counting in SearchBestCells
         CRKrnd = Int(Rnd(1) * AantalCRKs) ' chose one
         mK = CK(CRKrnd): mR = CR(CRKrnd)
         Else
         mK = CK(0): mR = CR(0)            ' one possibility
         End If
      ShowCross mR, mK
      Winst = GainPieces(mR, mK, MyColor)
      C(mR, mK) = MyColor
      ShowGain Winst
      BufferPrevState
      ShowField
      If PlayTurnsLeft() = False Then Exit Sub
      Passen = YouHaveToPass()
      If Passen = True Then MsgBox "You have to pass", Rood: Wait 2000: GoTo AanMij:
      GoTo AanJou:
      End If
   ' I didn't find one and didn't earlier
   If Passen = True Then
      MsgBox "You've won... The game is over!"
      ResetGame
      Exit Sub
      End If
   ' I didn't find one but you did earlier
   MsgBox "I pass.", Rood: Wait 2
   ' you an still go on
   Passen = YouHaveToPass()
   If Passen = True Then
      MsgBox "You've lost... The game is over!"
      ResetGame
      Exit Sub
      End If
   ' you can go on
   GoTo AanJou:
   
   
End Sub

Private Sub StartStop()
   Dim ipb As Variant
   Dim txt As String
   
   If cmdStart.Caption = "&Start" Then
      
      frmStart.Show 1, Me
      If frmStart.OK = False Then Exit Sub
      If frmStart.ChosenColor = 1 Then
         YourColor = bWit: MyColor = bZwart
         Else
         YourColor = bZwart: MyColor = bWit
         End If
      cmdStart.Caption = "&Stop"
      DoStop = False
      Play
      
      Else
      If MsgBox("Are you sure you want to stop the game?", vbOKCancel, "Stop") = vbCancel Then Exit Sub
      ResetGame
      Message " ", Rood
      
      End If

End Sub

Private Sub ShowField()
   Dim R, K, X, Y
   Dim txt As String
   
   With picV
      .Cls
      ' background
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Me.hDC, .Left, .Top, SRCCOPY
      
      .FontSize = 18
      ' horizontal - columns
      
      ' vertical - rows
      
      
      ' cells of playing field
      For R = 1 To 8: For K = 1 To 8
         Select Case C(R, K)
           Case bZwart: .FillColor = Zwart
           Case bWit:   .FillColor = Wit
           Case bLeeg:  .FillColor = Grijs
         End Select
         picV.Line (W * K, W * R)-Step(W, W), .FillColor, BF
         picV.Line (W * K, W * R)-Step(W, W), QBColor(7), B
      Next K, R
      If ZeurInfo Then

         For R = 1 To 8: For K = 1 To 8
            txt = Format(P(R, K))
            Select Case C(R, K)
              Case bZwart: .ForeColor = Wit
              Case bWit:   .ForeColor = Zwart
              Case bLeeg:  .ForeColor = QBColor(8)
            End Select
            .CurrentX = K * W + (W - .TextWidth(txt)) / 2
            .CurrentY = R * W + (W - .TextHeight(txt)) / 2 + 1
            picV.Print txt
         Next K, R
         End If
   End With
End Sub

' at the start of this program
Private Sub DrawBackground()

   
   'Random background

   
   'Tile
 
   
   'Title
   
End Sub

Private Sub ShowWhoIsOn(Wie As String)

End Sub

Private Sub ShowLittleCircle(R, K, Kleur As Long)
   picV.Circle (K * W + W / 2, R * W + W / 2), W / 2 - 2, Kleur
End Sub

Private Sub ShowCross(R, K)

End Sub

Private Sub ShowGain(Vrk)

End Sub

Private Function GainPieces(R, K, ZoekKleur)
   Dim Vr               ' cells per direction
   Dim VrTot            ' in total
   Dim KK, RR           '
   Dim I
   Dim DraaiKleur       ' to turn color
   
   DraaiKleur = IIf(ZoekKleur = bZwart, bWit, bZwart)
   
   For I = 1 To 8       ' inspect 8 directions
      Vr = 0            ' reset Vr per direction counter
      KK = K: RR = R    ' always start from this particular cell
      KK = KK + RIx(I)
      RR = RR + RIy(I)
      While C(RR, KK) = DraaiKleur
         Vr = Vr + 1
         KK = KK + RIx(I)
         RR = RR + RIy(I)
      Wend
      If C(RR, KK) <> bLeeg And Vr <> 0 Then
         VrTot = VrTot + Vr
         KK = KK - RIx(I)
         RR = RR - RIy(I)
         While C(RR, KK) <> bLeeg
            C(RR, KK) = ZoekKleur   ' turn over
            KK = KK - RIx(I)
            RR = RR - RIy(I)
         Wend
         End If
   Next I
   GainPieces = VrTot
End Function

Private Sub Wait(tos)
   Dim tijd As Variant
   tijd = Timer
   While Timer - tijd < tos / 1000: DoEvents: Wend
End Sub

' it is yuur turn, I'm waiting for a click a keydown
Private Sub WaitingForYou()
WaitingForYouOpnieuw:
   jR = 0: jK = 0
   IAmWaitingForYou = True
   While IAmWaitingForYou = True And DoStop = False: DoEvents: Wend
   If DoStop = True Then Exit Sub
   If IsOutOfField(jR, jK) Then MsgBox "Buiten veld !": GoTo WaitingForYouOpnieuw:
End Sub

Private Sub about_Click()
Call ShellAbout(0, "Reversi", "Microsoft Windows Reversi - Version 5.2 Release Candidate 1 (build 15) - Codename Sapphire", frmOth.Icon)
End Sub

Private Sub cmdStart_Click()
   StartStop
End Sub

Private Sub colour_Click()
  ' Set Cancel to True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  'Set the Flags property
  CommonDialog1.Flags = cdlCCRGBInit
  ' Display the Color Dialog box
  CommonDialog1.ShowColor
  ' Set the form's background color to selected color
  frmOth.BackColor = CommonDialog1.Color
  picB.BackColor = CommonDialog1.Color
  picV.BackColor = CommonDialog1.Color
     picV.BorderStyle = 0
   picB.BorderStyle = 0
   lblMessage(0).BorderStyle = 0
   lblMessage(1).BorderStyle = 0
   DrawBackground
   W = Int(picV.ScaleWidth / 9.5)
   ShowField
   Show
   cmdStart.SetFocus
   cmdStart.Visible = False
  Exit Sub

ErrHandler:
  ' User pressed the Cancel button
End Sub


Private Sub exit_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then imgHelp_Click
   If KeyCode = vbKeyF12 And Shift <> 0 Then
      ZeurInfo = IIf(ZeurInfo, False, True)
      ShowField
      KeyCode = 0
      End If
   If KeyCode = vbKeyV Then lblMessage_MouseDown 0, 1, 0, 5, 5
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If IAmWaitingForYou = True Then
      If KeyAscii > Asc("0") And KeyAscii < Asc("9") Then
         If jR = 0 Then jR = Val(Chr(KeyAscii)): Exit Sub
         If jK = 0 Then jK = Val(Chr(KeyAscii)): IAmWaitingForYou = False
         KeyAscii = 0
         End If
      End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyV Then lblMessage_MouseUp 0, 1, 0, 5, 5
End Sub

Private Sub Form_Load()
Zwart = RGB(Text1.Text, Text2.Text, Text3.Text)
Wit = RGB(Text4.Text, Text5.Text, Text6.Text)
Text7.Text = Zwart
Text8.Text = Wit
   picV.BorderStyle = 0
   picB.BorderStyle = 0
   lblMessage(0).BorderStyle = 0
   lblMessage(1).BorderStyle = 0
   DrawBackground
   W = Int(picV.ScaleWidth / 9.5)
   SetConstants
   ShowField
   Show
   cmdStart.SetFocus
   cmdStart.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub imgHelp_Click()
   On Error Resume Next
   AppActivate "OTHELLO help", False
   If Err = 0 Then Exit Sub
   Err = 0
   If ie = "" Then
      ie = "c:\program files\internet explorer\iexplore.exe"
      ie = InputBox("Is this the correct Internet Browser you are using?", "Othello.help", ie)
      If ie = "" Then Exit Sub
      End If
   Shell ie & " " & App.Path & "\othhelp.htm", vbNormalFocus
   If Err <> 0 Then MsgBox Err.Description
   On Error GoTo 0
End Sub

Private Sub Index_Click()
Dialog.Show
End Sub

Private Sub lblMessage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picV.Visible = False
End Sub

Private Sub lblMessage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   picV.Visible = True
End Sub

Private Sub new_Click()
   StartStop
   frmStart.optColor(0).ForeColor = Text7.Text
   frmStart.optColor(1).ForeColor = Text8.Text
End Sub

Private Sub p1_Click()
  ' Set Cancel to True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  'Set the Flags property
  CommonDialog1.Flags = cdlCCRGBInit
  ' Display the Color Dialog box
  CommonDialog1.ShowColor
  ' Set the form's background color to selected color
 Zwart = CommonDialog1.Color
 Text1.Text = CommonDialog1.Color
 Text2.Text = CommonDialog1.Color
 Text3.Text = CommonDialog1.Color
 Text7.Text = CommonDialog1.Color
     picV.BorderStyle = 0
   picB.BorderStyle = 0
   lblMessage(0).BorderStyle = 0
   lblMessage(1).BorderStyle = 0
   DrawBackground
   W = Int(picV.ScaleWidth / 9.5)
   ShowField
   Show
   cmdStart.SetFocus
   cmdStart.Visible = False
  Exit Sub

ErrHandler:
  ' User pressed the Cancel button
End Sub

Private Sub p2_Click()
  ' Set Cancel to True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  'Set the Flags property
  CommonDialog1.Flags = cdlCCRGBInit
  ' Display the Color Dialog box
  CommonDialog1.ShowColor
  ' Set the form's background color to selected color
 Wit = CommonDialog1.Color
 Text8.Text = CommonDialog1.Color
     picV.BorderStyle = 0
   picB.BorderStyle = 0
   lblMessage(0).BorderStyle = 0
   lblMessage(1).BorderStyle = 0
   DrawBackground
   W = Int(picV.ScaleWidth / 9.5)
   ShowField
   Show
   cmdStart.SetFocus
   cmdStart.Visible = False
  Exit Sub

ErrHandler:
  ' User pressed the Cancel button
End Sub

Private Sub picV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If IAmWaitingForYou = True Then
      jK = Int(X / W)
      jR = Int(Y / W)
      IAmWaitingForYou = False
      End If
End Sub


Private Sub plaiyingmga_Click()
Dialog.Show
End Sub

Private Sub Text1_Change()
Zwart = Text1.Text
End Sub

Private Sub Timer1_Timer()
   Dim priorMax            ' current highest priority
   Dim Winst               ' gain
   Dim Passen As Boolean   ' True = "jij" has to pass
   Dim CRKrnd              ' chosing "mijn" RK via this random-index
If Winst = 0 Then Screen.MousePointer = 0
If Winst = 1 Then Screen.MousePointer = 2

   frmStart.optColor(0).ForeColor = Text7.Text
   frmStart.optColor(1).ForeColor = Text8.Text
End Sub
