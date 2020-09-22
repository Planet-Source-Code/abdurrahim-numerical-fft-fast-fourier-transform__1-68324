VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FFT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "FFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   998
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THE FAST FOURIER TRANSFORM
'The frequency domain signals, held in REX[ ] and IMX[ ], are calculated from
'the time domain signal, held in XX[ ].
Option Explicit
Dim XX() As Double 'XX[ ] holds the time domain signal
Dim REX() As Double 'REX[ ] holds the real part of the frequency domain
Dim IMX() As Double 'IMX[ ] holds the imaginary part of the frequency domain
Dim K As Double
Dim i As Integer
Dim C As Integer
Dim Frekans() As Double
Dim b() As Double
Dim Value() As Currency
Private Sub Command1_Click()

Dim Pi As Double
Pi = 3.14159265
loaddata  'Verilerin xx serisine atýlmasý XX[ ]'Data array
Dim N As Integer
N = C 'N is the number of points in XX[ ]
ReDim REX(C / 2)
ReDim IMX(C / 2)
ReDim Frekans(C / 2)
'For K = 0 To C / 2 'Zero REX[ ] & IMX[ ] so they can be used as accumulators
'REX(K) = 0
'IMX(K) = 0
'Frekans(K) = 0
'Next K

Open App.Path & "\result.txt" For Output As #1 ' kayýt adilecek dosya açýlýyor
For K = 0 To N / 2 'K loops through each sample in REX[ ] and IMX[ ]
For i = 0 To N - 1 'I loops through each sample in XX[ ]
REX(K) = REX(K) + XX(i) * Cos(2 * Pi * K * i / N) 'reel
IMX(K) = IMX(K) - XX(i) * Sin(2 * Pi * K * i / N) 'Imaginary
Next i
'''Frekans hesaplama'''frequency calculation"
Frekans(K) = Sqr(REX(K) ^ 2 + IMX(K) ^ 2) '
'========================================
'''dosyaya kayýt'''save to file"
'For i = 0 To C / 2
Print #1, Frekans(K)
'Next i
Next K
Close ' ***** DOSYA KAPANIÞI *****
'===========================================
End Sub

Sub loaddata()
Open App.Path & "\data.txt" For Input As #1
Dim e As Double
        C = 0
      ' For say = 1 To 20
       ' Line Input #1, e
       ' Next
        Do
        
        ReDim Preserve XX(C)
        ' ReDim Preserve b(C)
          'ReDim Preserve Value(C)
        Input #1, e ', K
         XX(C) = e
        ' b(C) = K
         'On Error Resume Next
        '  Value(C) = Value(C - 1) + (A(C) - A(C - 1)) * (b(C - 1) + b(C)) / 2
         C = C + 1
        Loop While Not EOF(1)
        MsgBox C, , "Number of Data"
 Close ' ***** DOSYA KAPANIÞI *****
          
End Sub

