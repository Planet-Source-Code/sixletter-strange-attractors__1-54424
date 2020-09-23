VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Strange Attractors"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   26
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Iterations"
      Height          =   855
      Left            =   7800
      TabIndex        =   11
      Top             =   3240
      Width           =   2295
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   100
         Left            =   240
         Max             =   10000
         Min             =   1000
         TabIndex        =   12
         Top             =   360
         Value           =   1000
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Progress"
      Height          =   2655
      Left            =   7800
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
      Begin VB.PictureBox Picture2 
         Height          =   1815
         Left            =   240
         ScaleHeight     =   1755
         ScaleWidth      =   1755
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   855
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Strange Attractors"
      Height          =   7575
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   7455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   6975
         Left            =   240
         ScaleHeight     =   6945
         ScaleWidth      =   6945
         TabIndex        =   7
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coefficients"
      Height          =   2895
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "Random"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   1815
      End
      Begin VB.HScrollBar CoScroll 
         Height          =   255
         Index           =   5
         LargeChange     =   100
         Left            =   480
         Max             =   5000
         Min             =   -5000
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.HScrollBar CoScroll 
         Height          =   255
         Index           =   4
         LargeChange     =   100
         Left            =   480
         Max             =   5000
         Min             =   -5000
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.HScrollBar CoScroll 
         Height          =   255
         Index           =   3
         LargeChange     =   100
         Left            =   480
         Max             =   5000
         Min             =   -5000
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar CoScroll 
         Height          =   255
         Index           =   2
         LargeChange     =   100
         Left            =   480
         Max             =   5000
         Min             =   -5000
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.HScrollBar CoScroll 
         Height          =   255
         Index           =   1
         LargeChange     =   100
         Left            =   480
         Max             =   5000
         Min             =   -5000
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label CoLbl 
         Caption         =   "Label4"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   24
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label CoLbl 
         Caption         =   "Label4"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   23
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label CoLbl 
         Caption         =   "Label4"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label CoLbl 
         Caption         =   "Label4"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.Label CoLbl 
         Caption         =   "Label4"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "B:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "C:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "D:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "E:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim I As Integer
    Me.Picture1.BackColor = vbWhite
    Me.HScroll1.Value = 2000
    Me.Label2.Caption = Me.HScroll1.Value

    Randomize
    Me.CoScroll(1).Value = 10000 * Rnd - 5000
    Me.CoScroll(2).Value = 10000 * Rnd - 5000
    Me.CoScroll(3).Value = 10000 * Rnd - 5000
    Me.CoScroll(4).Value = 10000 * Rnd - 5000
    Me.CoScroll(5).Value = 10000 * Rnd - 5000
    
    For I = 1 To 5
        Me.CoLbl(I).Caption = Format(Me.CoScroll(I).Value / 1000, "0.000")
    Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Command1_Click()

    Dim Its As Integer '# of iterations
    
    Dim A As Single 'coefficient
    Dim B As Single 'coefficient
    Dim C As Single  'coefficient
    Dim D As Single  'coefficient
    Dim E As Single  'coefficient
    
    Dim I As Integer 'iteration
    Dim J As Integer 'iteration
    Dim K As Single
    
    Dim X As Single  'Point
    Dim Y As Single  'Point
    Dim Z As Single  'Point
    
    Dim Xp As Single 'Plot Point
    Dim Yp As Single 'Plot Point
    Dim Zp As Single 'Plot Point
    
    Dim MaxX As Single 'Scaling factor
    Dim MaxY As Single  'Scaling factor
    
    Dim PCol As Long 'Point Color


    On Error Resume Next
    Me.Picture1.SetFocus
    Me.Command1.Enabled = False
    
    A = Me.CoScroll(1).Value / 1000
    B = Me.CoScroll(2).Value / 1000
    C = Me.CoScroll(3).Value / 1000
    D = Me.CoScroll(4).Value / 1000
    E = Me.CoScroll(5).Value / 1000
    
    Its = Me.HScroll1.Value
    
    MaxX = 0
    MaxY = 0
    
    X = 0
    Y = 0
    Z = 0
    
    'go through iterations to scale picture
    For J = 1 To Its
        For I = 1 To 100
    
            Xp = Sin(A * Y) - Z * Cos(B * X)
            Yp = Z * Sin(C * X) - Cos(D * Y)
            Zp = E * Sin(X)
    
            X = Xp
            Y = Yp
            Z = Zp
    
            If Xp > MaxX Then MaxX = Xp
            If Yp > MaxY Then MaxY = Yp
    
        Next I
        DoEvents
    Next J
   
    
    If MaxX > 0 And MaxY > 0 Then
        MaxX = RoundUp(MaxX)
        MaxY = RoundUp(MaxY)
        
        
        Me.Picture1.Cls
        'scale picture
        Me.Picture1.Scale (-MaxX, MaxY)-(MaxX, -MaxY)
        
        'scale progress bar
        Me.Picture2.Scale (0, Its)-(100, 0)
        Me.Picture3.Move 0, 0, 100, 0
        Me.Picture3.BackColor = RGB(255, 0, 0)
        Me.Label3.Caption = "0%"
        Me.Picture2.Visible = True
        Me.Label3.Visible = True
        
        'go through iterations and plot points
        For J = 1 To Its
            For I = 1 To 100
                Xp = Sin(A * Y) - Z * Cos(B * X)
                Yp = Z * Sin(C * X) - Cos(D * Y)
                Zp = E * Sin(X)
                X = Xp
                Y = Yp
                Z = Zp
                
                'bring color down on individual points until they are completely black
                PCol = Picture1.Point(Xp, Yp) - RGB(16, 16, 16)
                If PCol < vbBlack Then
                    PCol = vbBlack
                End If
                Picture1.PSet (Xp, Yp), PCol
                
            Next I
            'move progress bar
            Me.Picture3.BackColor = RGB(Int(255 - ((J / Its) * 255)), Int(255 * (J / Its)), 0)
            Me.Picture3.Move 0, J, 100, J
            Me.Label3.Caption = Format(J / Its, "0%")
            
            DoEvents
        Next J
    Else
        'sometimes this happens
        MsgBox "The Coefficients You Have Chosen Are Beyond My Capability.", vbOKOnly, "Coefficients"
    End If
  
    
    Me.Picture2.Visible = False
    Me.Command1.Enabled = True
    Me.Label3.Visible = False
    On Error GoTo 0
End Sub

Private Sub Command2_Click()
    Dim I As Integer
    
    Randomize
    Me.CoScroll(1).Value = 10000 * Rnd - 5000
    Me.CoScroll(2).Value = 10000 * Rnd - 5000
    Me.CoScroll(3).Value = 10000 * Rnd - 5000
    Me.CoScroll(4).Value = 10000 * Rnd - 5000
    Me.CoScroll(5).Value = 10000 * Rnd - 5000
    
    For I = 1 To 5
        Me.CoLbl(I).Caption = Format(Me.CoScroll(I).Value / 1000, "0.000")
    Next I
End Sub

Private Sub CoScroll_Change(Index As Integer)
    Me.CoLbl(Index).Caption = Format(Me.CoScroll(Index).Value / 1000, "0.000")
End Sub

Private Sub CoScroll_Scroll(Index As Integer)
    CoScroll_Change Index
End Sub

Private Sub HScroll1_Change()
    Me.Label2.Caption = Me.HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Function RoundUp(ByVal Vals As Double)
    RoundUp = -Int(-Vals)
End Function
