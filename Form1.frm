VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPcbdraw 
   BackColor       =   &H00000000&
   Caption         =   "E-Laserpcb4"
   ClientHeight    =   9720
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   648
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   848
   Begin VB.OptionButton Doubleside 
      Caption         =   "Double side"
      Height          =   255
      Left            =   10800
      TabIndex        =   32
      Top             =   8100
      Width           =   1215
   End
   Begin VB.OptionButton Singleside 
      Caption         =   "Single side"
      Height          =   315
      Left            =   10800
      TabIndex        =   31
      Top             =   7800
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox pctSlave 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   10320
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton firmware 
      Caption         =   "new firmware"
      Height          =   495
      Left            =   11040
      TabIndex        =   22
      Top             =   7200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   13680
      Top             =   8280
   End
   Begin VB.TextBox speedFactor 
      Height          =   285
      Left            =   8280
      TabIndex        =   12
      Text            =   "22"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox yfactor 
      Height          =   285
      Left            =   8280
      TabIndex        =   11
      Text            =   "1.210"
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox xfactor 
      Height          =   285
      Left            =   8280
      TabIndex        =   10
      Text            =   "0.960"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox smooth 
      Height          =   285
      Left            =   10095
      TabIndex        =   14
      Text            =   "0"
      Top             =   7320
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   "Factors"
      Height          =   2415
      Left            =   7320
      TabIndex        =   15
      Top             =   7080
      Width           =   4815
      Begin VB.TextBox Xleft 
         Height          =   285
         Left            =   2760
         TabIndex        =   44
         Text            =   "10"
         Top             =   800
         Width           =   375
      End
      Begin VB.CheckBox xref 
         Caption         =   "X-reference"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3480
         TabIndex        =   42
         Top             =   1325
         Width           =   1200
      End
      Begin VB.TextBox config 
         Height          =   285
         Left            =   720
         TabIndex        =   41
         Text            =   "Laserpcb"
         Top             =   2000
         Width           =   1455
      End
      Begin VB.CheckBox swapXY 
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox flipX 
         Height          =   195
         Left            =   2880
         TabIndex        =   36
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton SerMonitor 
         Caption         =   "Ser.Monitor"
         Height          =   375
         Left            =   3360
         TabIndex        =   35
         Top             =   1565
         Width           =   1095
      End
      Begin VB.TextBox HoleSize 
         Height          =   285
         Left            =   2760
         TabIndex        =   34
         Text            =   "20"
         Top             =   1680
         Width           =   375
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox trailing 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "40"
         Top             =   1400
         Width           =   375
      End
      Begin VB.TextBox leading 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "10"
         Top             =   1100
         Width           =   375
      End
      Begin VB.CheckBox reverse 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   500
         Width           =   255
      End
      Begin VB.CommandButton test 
         Caption         =   "test"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "X leftmargin"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   800
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Config:"
         Height          =   255
         Left            =   200
         TabIndex        =   40
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Swap X-Y"
         Height          =   255
         Left            =   3360
         TabIndex        =   39
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Flip X "
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Hole size (0: gerber size)"
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label trailing_label 
         Caption         =   "Trailing burn"
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   1400
         Width           =   900
      End
      Begin VB.Label leading_label 
         Caption         =   "Y leadmargin"
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   1100
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Positive resist"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   500
         Width           =   975
      End
      Begin VB.Label rowCtr 
         Height          =   255
         Left            =   800
         TabIndex        =   21
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Speed"
         Height          =   255
         Left            =   470
         TabIndex        =   19
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   " Y"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   640
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   " X"
         Height          =   255
         Left            =   730
         TabIndex        =   17
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Reduction"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   4080
      Left            =   12240
      TabIndex        =   9
      Top             =   4680
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7197
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Download 
      Caption         =   "Download"
      Height          =   1215
      Left            =   5040
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton NewLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   7440
      Width           =   975
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   13680
      TabIndex        =   6
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   8880
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   30
      Min             =   1
      TabIndex        =   0
      Top             =   6840
      Value           =   1
      Width           =   13815
   End
   Begin VB.PictureBox pctMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5061
      Left            =   65
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   660
      TabIndex        =   1
      Top             =   120
      Width           =   9936
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Speed"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape shpCurr 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Menu Comm 
      Caption         =   "Communication"
      Begin VB.Menu Comm_1 
         Caption         =   "COM1"
      End
      Begin VB.Menu Comm_2 
         Caption         =   "COM2"
      End
      Begin VB.Menu Comm_3 
         Caption         =   "COM3"
      End
      Begin VB.Menu Comm_4 
         Caption         =   "COM4"
      End
      Begin VB.Menu LowBaud 
         Caption         =   "57600 baud"
         Checked         =   -1  'True
      End
      Begin VB.Menu StdBaud 
         Caption         =   "115200 baud"
      End
      Begin VB.Menu HighBaud 
         Caption         =   "230400 baud"
      End
   End
End
Attribute VB_Name = "frmPcbdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEFAULT_PARAM = "laserpcb3"
'  Const TMPFILE = "tmpbmp.bmp"
Const MAXFLASH = &H55000 * 2 '  4096
Const APPFLASH = &H10000 * 2 '  4096

Const FLASH_BLK_LEN = 128
Const FLASH_BIN_WORD_BYTES = 4
'maxvalidaddress  0x0557FF
'senza boot       0x054FFF
'con dati         0x0122FF
Public SerBuffer As String
Public settingName As String
Public filename As String       ' nome puro, senza directory nè estensione
Public filenameBotTmp As String
Public filenameTopTmp As String

Dim Lmaxflash As Long

Dim Points(0 To 10) As POINTAPI
Dim CurrentPoint As POINTAPI
Dim contour As Integer
Dim inPtr As Long
Dim currentRow As Integer
Dim dataString(32000) As String ' formato bitmap non compresso - ogni bit un punto
Dim dataZip0(32000) As String   ' formato compresso di datastring  coppie di bytes ZUZUZU...  (Z: numero di bits a zero   U: numero di bits a uno)
Dim dataZip1(32000) As String   ' formato compresso di datastring XOR precedente   coppie di bytes ZUZUZU...  (Z: numero di bits a zero   U: numero di bits a uno)
Dim dataZip2(32000) As String   ' formato compresso di datastring XOR precedente   coppie di bytes ZZUUZZUU...  (Z: numero di bits a zero   U: numero di bits a uno)
Dim dataBot(32000) As String    ' formato compresso di datastring migliore
Dim xdatabot As Integer
Dim ydatabot As Integer
Dim dataTop(32000) As String    ' formato compresso di datastring migliore
Dim xdatatop As Integer
Dim ydatatop As Integer
Dim dataZipN As String
Dim dataVideo As String
Dim num0 As Integer
Dim num1 As Integer
Dim totaleN As Long
Dim totale0 As Long
Dim totale1 As Long
Dim totale2 As Long
Dim ds As Byte
Dim pctDisplay As Byte
Dim verbose As Byte

Dim macroMax As Integer
Dim macroInd As Integer
Dim macroName(100) As String
Dim macroPrimitive(100) As Integer
Dim macroParam(100, 10) As String

Dim Xscale As Single, Yscale As Single  ' fattori di conversione da mils a pixel modificabili da box
Dim XscaleGraph As Single, YscaleGraph As Single

Dim XgerberScale As Integer
Dim YgerberScale As Integer
Dim Xspeed As Integer
Dim Xhole As Integer
Dim Xborder As Integer
Dim Yborder As Integer
Dim Xstart As Integer
Dim Ystart As Integer
Dim smtpoint As Integer
Dim CurrentPointMin As Long
Dim CurrentPointMax As Long
Dim units As String
Dim iCom As Integer      ' comm serial port
Dim sCom As String
Dim baudCom As Long
Dim bRS232 As Boolean    ' true = usa rs232 vera
'Dim connRS232 As Boolean ' true = rs232 connessa
'Dim v As Double, i As Integer
Dim lngStatus As Long
Dim strError  As String

Dim lineTxt, lineChr, lineParm As String
Dim lineLen  As Integer
Dim linePtr  As Integer
Dim Xtemp  As Long
Dim Ytemp  As Long
Dim Xcurr  As Integer
Dim Ycurr  As Integer
Dim Xprev  As Integer
Dim Yprev As Integer
Dim xMin  As Integer    ' coordinata X più bassa nel gerber
Dim xMax  As Integer    ' coordinata X più alta nel gerber
Dim yMin  As Integer    ' coordinata Y più bassa nel gerber
Dim yMax  As Integer    ' coordinata Y più alta nel gerber
Dim Xoffset As Integer  ' offset X a cui inizia il gerber: vi corrisponde la coordinata X=0 sullo schermo
Dim Yoffset As Integer  ' offset Y a cui inizia il gerber: vi corrisponde la coordinata Y=0 sullo schermo
Dim Xpoints As Integer  ' coordinata X dello schermo più alta - larghezza pcb in scala - compreso margine dx - multiplo di 8 pt
Dim Ypoints As Integer  ' coordinata Y dello schermo più alta - altezza pcb in scala - compreso margine dx
Dim Xbytes As Integer
Dim sFlash As String
Dim sAdr, fAdr, sBlock, sBlock3, sChk As String
Dim smUDP, smFLASH, iAck, myTimer, nrDevices As Integer
Dim sMachine As Integer ' 0=nothing   1=riga dati inviata    2=ultima riga inviata     9=fine lavoro    10=serial monitor     20=fw update in corso
Dim row As Integer
Dim sResp As String
Dim resp  As Integer
Dim retry As Integer
Dim hByte As Long
Dim lByte As Long
Dim nRepeat As Integer
Dim kByte As Long
Dim col As Integer
Dim lRow As Integer
Dim rowEffective As Integer
Dim zType As Integer

Dim apertureCode, apertureType, polarity As String
Dim apertureCodn  As Integer
Dim apertureN As Integer
Dim apertureM As Integer
Dim trackSize As Integer

Dim tabApertureType(200) As String
Dim tabApertureN(200) As Integer
Dim tabApertureM(200) As Integer

Dim Dcode, Gcode, Mcode As String
Dim DcodN  As Integer
Dim GcodN  As Integer
Dim McodN  As Integer

Dim bitVal(8) As Byte

Dim File As String
Dim sFileLogName As String
Dim dFile As Integer
Dim FileExtensionPos As Integer
Dim dAdr, dAdrNxt, dSize

Dim dFileLog As Long
Dim dFileL1 As Long
Dim dFileL2 As Long
Dim dFileL3 As Long

Dim iExtension As Integer
Dim startFileDir As String

Private Declare Function SetPixelV Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal crColor As Long) As Long
    
Private Declare Function GetPixel Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
    
Private Declare Function MoveToEx Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, _
    ByRef lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
    
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As _
    Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    
Private Declare Function Rectangle Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal x2 As Long, ByVal y2 As Long) As Long
    
Private Declare Function Ellipse Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal x2 As Long, ByVal y2 As Long) As Long
'X1 = Centre.X - radius
'Y1 = Centre.Y - radius
'X2 = Centre.X + radius
'X3 = Centre.Y + radius

Private Declare Function TextOut Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal lpString As String, ByVal nCount As Long) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Dim tabHoles(300) As POINTAPI
Dim holeW As Integer
Dim holeR As Integer
' tabHoles(holeW).X = ...
' tabHoles(holeW).Y = ...



' ==========================BOTTONE CLEAR - PULISCE LO SCHERMO e le variabili salvate======================
Private Sub Clear_Click()
    pctMain.DrawWidth = 1
    pctMain.Cls
    pctMain.Picture = Nothing
    pctMain.ForeColor = RGB(50, 90, 50)
    ClearSaved
    End Sub

Private Sub cmdAbout_Click()
    MsgBox "Versione 6.0" & vbCr & "Made by G_Pagani" & vbCr & "Copyright (C) 2018" & vbCr & "Usare con firmware da V32.xx", vbOK + vbInformation
End Sub

Private Sub cmdExit_Click()
    End
End Sub
Function MakeRow(lVersion)
    Dim sRow As String
    Dim iC, iX
    
    If (lVersion = 4) Then

' riga zippata di tipo 1:
'             zNLDEDEDECK  z=tipo riga "z"
'                          N=chr numero ripetizioni (max 0F: 15)
'                          L=chr length riga (max FF: 255)
'                          D=numero di bits a 0 (max FF)
'                          E=numero di bits a 1 (max FF)
'                          C=check byte low
'                          K=check byte high
' riga zippata di tipo >1:
'             zNLDDDDDDCK  z=tipo riga "z"
'                          N=chr numero ripetizioni + 0x10 (max 1F: 15)
'                          L=chr length riga (max FF: 255)
'                          D=numero di bits a 0 (max FF) in XOR rispetto a riga precedente
'                          E=numero di bits a 1 (max FF) in XOR rispetto a riga precedente
'                          C=check byte low
'                          K=check byte high
' riga vuota:
'             vNLD         v=tipo riga "v"
'                          N=chr numero ripetizioni (max FF: 255)
'                          L=chr length riga (1)
'                          D=dato ripetuto (0)
    
'r
        If zType = 0 Then
            sRow = "r"
'n
            nRepeat = 1
            While dataString(row) = dataString(row + nRepeat) And (row + nRepeat) < Ypoints
                nRepeat = nRepeat + 1
            Wend
            If nRepeat > 255 Then LogPrint ("ERRORE - RIPETIZIONI > 255")
        
            kByte = nRepeat
            sRow = sRow + Chr(nRepeat)
            For col = 1 To Xbytes
                kByte = kByte + Asc(Mid(dataString(row), col, 1))
            Next
'data
            sRow = sRow + Left(dataString(row), Xbytes)
        Else
            sRow = "z"
'n*zType
            nRepeat = 1
            
            While dataZip0(row) = dataZip0(row + nRepeat) And (row + nRepeat) < Ypoints
                nRepeat = nRepeat + 1
            Wend
            
            If dataString(row) = String(Xbytes, &H0) Then
                If nRepeat > 255 Then nRepeat = 255
                sRow = "v" + Chr(nRepeat) + Chr(1) + Chr(0)
                MakeRow = sRow
'               LogPrint ("v row: " + CStr(row) + " n. " + CStr(nRepeat))
                Exit Function
            End If
            If nRepeat > 15 Then nRepeat = 15
            
' prova se e' conveniente usare zType = 1
            If (Len(dataZip1(row)) = 0) Or (Len(dataZip0(row)) <= Len(dataZip1(row))) Then
                kByte = nRepeat
                sRow = sRow + Chr(nRepeat)
    'l
                lRow = Len(dataZip0(row))
                hByte = Int(lRow / 256)
                lByte = lRow - hByte * 256
                sRow = sRow + Chr(lByte)
                If lRow > 253 Then LogPrint ("ERRORE - RIGA ZIP0 LUNGA > 255")
                kByte = kByte + lByte
                
                For col = 1 To lRow
                    kByte = kByte + Asc(Mid(dataZip0(row), col, 1))
                Next
    'data                   z3
                sRow = sRow + dataZip0(row)
            Else
                kByte = nRepeat + (1 * 16)
                sRow = sRow + Chr(kByte)
    'l
                lRow = Len(dataZip1(row))
                hByte = Int(lRow / 256)
                lByte = lRow - hByte * 256
                sRow = sRow + Chr(lByte)
                If lRow > 253 Then LogPrint ("ERRORE - RIGA ZIP1 LUNGA > 255")
                kByte = kByte + lByte
                
                For col = 1 To lRow
                    kByte = kByte + Asc(Mid(dataZip1(row), col, 1))
                Next
    'data                   z3
                sRow = sRow + dataZip1(row)
            End If
        End If
'ck
        hByte = Int(kByte / 256)
        lByte = kByte - hByte * 256
        If hByte > 255 Then hByte = 255
        sRow = sRow + Chr(lByte) + Chr(hByte)
    End If



    If (lVersion = 41) Then     ' double side - BOTTOM SIDE file write
        sRow = Mid(dataBot(ydatabot), 1, 1) ' "z"
        If ydatabot < xdatabot Then
            nRepeat = Asc(Mid(dataBot(ydatabot), 2, 1))
            If nRepeat > 15 And sRow <> "v" Then nRepeat = nRepeat - 16
'            If (sRow = "v") Then
'                    LogPrint ("v row: " + CStr(ydatabot) + " n. " + CStr(nRepeat))
'            End If
            sRow = dataBot(ydatabot)
        Else
            nRepeat = 0
            sRow = ""
'            sRow = dataBot(xdatabot - 1)
        End If
        ydatabot = ydatabot + 1
    End If

    If (lVersion = 42) Then     ' double side - TOP SIDE file write
        sRow = Mid(dataTop(ydatatop), 1, 1) '"z"
        If ydatatop < xdatatop Then
            nRepeat = Asc(Mid(dataTop(ydatatop), 2, 1))
            If nRepeat > 15 And sRow <> "v" Then nRepeat = nRepeat - 16
'            If (sRow = "v") Then
'                    LogPrint ("v row: " + CStr(ydatatop) + " n. " + CStr(nRepeat))
'            End If
            sRow = dataTop(ydatatop)
        Else
            nRepeat = 0
            sRow = ""
'           sRow = dataTop(xdatatop - 1)
        End If
        ydatatop = ydatatop + 1
    End If
    

    MakeRow = sRow
End Function
Function StringInt(iToString)
    hByte = Int(iToString / 256)
    lByte = iToString - hByte * 256
    If hByte > 255 Then hByte = 255
    kByte = kByte + hByte + lByte
    StringInt = Chr(lByte) + Chr(hByte)
End Function
Function MakeHeader(lVersion) '// versione 4 or 5
    Dim sHeader As String
    Dim iC, iX
    Dim filenamepad As String
    
    sHeader = "h"
    
'name[30]               ' nome del pcb
        kByte = 0
        filenamepad = Left(filename + String(30, " "), 30)
        sHeader = sHeader + filenamepad
        For iC = 1 To 30
            kByte = kByte + Asc(Mid(filenamepad, iC, 1))
        Next
'll                     ' numero di bytes per riga non compressa
        sHeader = sHeader + StringInt(Xbytes)
'rr                     ' numero di righe effettive (non compresse)
        sHeader = sHeader + StringInt(Ypoints)
's                      ' velocita in cm/sec
        kByte = kByte + Xspeed
        sHeader = sHeader + Chr(Xspeed)
'o                      ' opzione (1=stampa rovesciata - fotosensibile negativo)
        lByte = 0
        If reverse.Value = 1 Then lByte = &H1
        kByte = kByte + lByte
        sHeader = sHeader + Chr(lByte)
'leading lines          ' numero di righe aggiuntive iniziali (se stampa rovesciata)
        If reverse.Value = 1 Then
           lByte = Val(leading.Text)
        Else
           lByte = 0
        End If
        kByte = kByte + lByte
        sHeader = sHeader + Chr(lByte)
'trailing lines          ' numero di righe aggiuntive finali (se stampa rovesciata)
'       If reverse.Value = 1 Then
           lByte = Val(trailing.Text)
'        Else
'           lByte = 0
'        End If
        kByte = kByte + lByte
        sHeader = sHeader + Chr(lByte)
'xx                     ' scala X
        iX = Int(Xscale * 1000)
        sHeader = sHeader + StringInt(iX)
'yy                     ' scala Y
        iX = Int(Yscale * 1000)
        sHeader = sHeader + StringInt(iX)
        
' X margine  in mils
        lByte = Int((Xborder / Xscale) + 0.49)
        kByte = kByte + lByte
        sHeader = sHeader + Chr(lByte)
        
' Y margine  in mils
        lByte = Int((Yborder / Yscale) + 0.49)
        kByte = kByte + lByte
        sHeader = sHeader + Chr(lByte)
        
'ck
        sHeader = sHeader + StringInt(kByte)
        
        If (lVersion > 4) Then
        ' versione 5.1 - dati aggiuntivi - non entrano nel calcolo di ck
                sHeader = sHeader + StringInt(Xbytes * 8)
                sHeader = sHeader + StringInt(Xbytes * 8)
                sHeader = sHeader + StringInt(CurrentPointMin)
                sHeader = sHeader + StringInt(CurrentPointMax)
        End If

    MakeHeader = sHeader
End Function
Private Sub response_display()
    Dim iC As Integer
    Dim sHex 'As String
    
    sHex = ""
    sResp = ReadWait(16, 5)
    If sResp <> "" Then
        LogPrint "    response: " & Len(sResp) & " bytes : "
        For iC = 1 To Len(sResp)
            sHex = sHex & HexString(Asc(Mid(sResp, iC, 1)), 2)
        Next
        LogPrint (sHex)
        sHex = ""
    Else
        LogPrint "    response: <null>"
    End If

End Sub
Private Function write_display(sWrite)
    Dim iC As Integer
    Dim sHex 'As String
    
'    iC = Asc(Left(sWrite, 1)) + 1
'    sWrite = Left(sWrite + String(16, Chr(0)), iC)
    
    sWrite = Chr(Len(sWrite) + 2) + Chr(10) + Chr(7) + sWrite
    
    sHex = ""
    For iC = 1 To Len(sWrite)
        sHex = sHex & HexString(Asc(Mid(sWrite, iC, 1)), 2)
    Next
    LogPrint (sHex)
    WriteBuf (sWrite)
End Function
Private Function write_firmware(sWrite)
    If Len(sWrite) < 254 Then
        WriteBuf (Chr(Len(sWrite) + 2) + Chr(10) + Chr(7) + sWrite)
    Else
        WriteBuf (Chr(0) + sWrite)
    End If
End Function

Private Sub test_fw_Click()
    Dim iC, iD, iRc, iChk
    
    resp = TryConnect(baudCom)

    If resp = 0 Then
        LogPrint "send firmware update request "
        WriteBuf ("@" + Chr(&H11))
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        LogPrint "send no-write request "
        ' invia il comando 0x03
        sChk = BinString2(iChk)
        write_display (Chr(3) + sChk)
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        CommonDialog1.filename = ""
        CommonDialog1.DialogTitle = "Open  FLASH   bin file"
        CommonDialog1.DefaultExt = "hex"
        CommonDialog1.Filter = "*.hex"
        CommonDialog1.ShowOpen
        If CommonDialog1.filename <> "" Then
            
            HexFlashFileRead (CommonDialog1.filename)
            Open "temp.bin" For Binary As dFile
            dSize = LOF(dFile)
            sBlock = ReadBin(64)
            Close dFile

            iChk = 0
            For iC = 1 To 64
                iChk = iChk + Asc(Mid(sBlock, iC, 1))
            Next
            sChk = BinString2(iChk)
            sAdr = BinString3(0)

' invia il comando 0x01 - new flash block binary mode
            write_display (Chr(11) + sAdr + Chr(64))  'block length,   source device, pdu format, data=0x01(newblock), addressL, addressH, dataLength
            response_display
' invia 8 blocchi da 8 bytes - flash binary data
            write_display (Mid(sBlock, 1, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 9, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 17, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 25, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 33, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 41, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 49, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
            write_display (Mid(sBlock, 57, 8))  'block length,   source device, pdu format, binary data 1-8
            response_display
' invia il comando 0x03 - end flash block binary mode - no write
            write_display (Chr(3) + sChk)  'block length,   source device, pdu format, data=0x02(endblock), checkL, checkH
            response_display
        End If
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
        
        LogPrint "send query request "
        ' invia il comando 0x00
        write_display (Chr(0))
        response_display
    
        Disconnect
    End If
    
End Sub


Private Sub config_LostFocus()
   settingName = config.Text
   Call Get_config
   SaveSetting DEFAULT_PARAM, "InitValues", "configName", config.Text
End Sub

Private Sub Doubleside_Click()
    xref.Enabled = True
End Sub

Private Sub firmware_Click()
    Dim iC, iD, iRc
    Dim sE 'As String
    Dim startFirmDir As String
    
'    test_fw_Click
'    Exit Sub
    
    startFirmDir = GetSetting(settingName, "InitValues", "FirmDir")

    CommonDialog1.filename = ""

    CommonDialog1.InitDir = startFirmDir
    CommonDialog1.DialogTitle = "Open  FLASH   bin file"
    CommonDialog1.DefaultExt = "hex"
    CommonDialog1.Filter = "*.hex"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename <> "" Then
         sE = CommonDialog1.filename
         
         startFirmDir = sE
         iC = Len(sE)
         While iC > 0 And Mid(startFirmDir, iC, 1) <> "\"
            iC = iC - 1
         Wend
         If iC > 2 Then
            startFirmDir = Left(startFirmDir, iC - 1)
         Else
            startFirmDir = Left(startFirmDir, 3)
         End If
         SaveSetting settingName, "InitValues", "FirmDir", startFirmDir
         
         Lmaxflash = MAXFLASH                ' risposta SI
  
'         iRc = MsgBox("Aggiorno anche la zona SERVICES ?", 4, "Flash memory update")
'         If iRc = 6 Then Lmaxflash = MAXFLASH                ' risposta SI
'         If iRc = 7 Then Lmaxflash = APPFLASH                ' risposta NO
         
         smFLASH = 0  ' initial operations
         HexFlashFileRead (CommonDialog1.filename)
         CommonDialog1.filename = "temp.bin"

'        Command1.Enabled = False
'        Command3.Enabled = False
'        Command4.Enabled = False
'        Command5.Enabled = False
'        Command6.Enabled = False
'        Command7.Enabled = False
'        Command8.Enabled = False
'        Combo1.Enabled = False

        frmPcbdraw.Refresh
    
        If bRS232 = False Then
            resp = 0
        Else
            resp = TryConnect(baudCom)
        End If
        
        If resp = 0 Then
            LogPrint "send firmware update request"
            Flush
            BinFlashFileRead (CommonDialog1.filename)
            End If
        End If
End Sub
Function HexFlashFileRead(sFileName)
    Dim sZero, sDati, numLine, a
    Dim sLine, sCol, sTrk, sWork, sChex, sResp
    Dim iZ, iLen, iAdr, iExtAdr, iWork, iPtr, iEOF
    Dim ixf As Long
    
    sFlash = String(Lmaxflash, Chr(255))
   '
   ' Check if the file exists
   '
    dFile = FreeFile()
    Open sFileName For Input As dFile
    LogPrint "read & convert file..." & sFileName
    
    iExtAdr = 0
    numLine = 0
    Do
        Line Input #dFile, sLine
        numLine = numLine + 1
        sLine = UCase(sLine)
                
            ' destrutturazione del record
            ':LLAAAATTxxxxxxxxxxxxxxxxxxxxxxxxxCK
        sCol = Left(sLine, 1)
        iLen = DecByte(Mid(sLine, 2, 2))
        iAdr = DecWord(Mid(sLine, 4, 4))
        sTrk = Mid(sLine, 8, 2)
        sDati = Mid(sLine, 10, iLen * 2)

' trk 04: extended linear address record
        If sCol = ":" And sTrk = "04" And iLen > 0 Then
            iExtAdr = DecWord(Mid(sLine, 10, 4)) * 256 * 256
        End If
' trk 00: data record
        If sCol = ":" And sTrk = "00" And iLen > 0 Then
            iAdr = iAdr + iExtAdr
                        
' ============================== FLASH =======================================
            If iAdr >= 0 And iAdr < (Lmaxflash - iLen) Then
                iPtr = iAdr + 1
                For iWork = 0 To iLen - 1
                'NON scambia LSB MSB (little endian) - LLMM LLMM LLMM
                     sChex = Mid(sDati, 1 + (iWork * 2), 2)
                     Mid(sFlash, iPtr, 1) = Chr(DecByte(sChex))
                     iPtr = iPtr + 1
                Next
'            Else
'             iAdr = iAdr
            End If
        End If
    iEOF = iEOF + 1
' trk 01: fine file
    Loop Until EOF(dFile) Or (Len(sLine) < 3) Or (sCol <> ":") Or sTrk = "01" ' Or iEOF > 50000
    Close #dFile
    
    ixf = Lmaxflash
    While Mid(sFlash, ixf, 1) = Chr(&HFF)
        ixf = ixf - 1
    Wend
    ixf = ixf / 4096
    ixf = ixf + 1
    ixf = ixf * 4096
    
    Open "temp.bin" For Binary Access Write As dFile
    
    Put dFile, 1, sFlash
'    Put dFile, 1, Left(sFlash, ixf)
    Close #dFile

    HexFlashFileRead = 1
End Function
Function BinFlashFileRead(sFileName)
    Dim i As Integer
    Dim iChk, iChk3, iVhk As Long
    
    If smFLASH = 0 Then
        dAdr = 0
        dAdrNxt = 0

        dFile = FreeFile()
    
    'On Error GoTo BinNotExist
        Open sFileName For Binary As dFile
        dSize = LOF(dFile)
        If dSize > Lmaxflash Then dSize = Lmaxflash
        
    'On Error GoTo BinSocketError
        LogPrint "Open socket..."
        LogPrint "Wait for ack..."
   
        WriteBuf ("@" + Chr(&H11))
        
        iAck = 0
        smFLASH = 1  ' wait for first ack
        BinFlashFileRead = 0
        
        sMachine = 20
        Timer1.Interval = 1
        Timer1.Enabled = True
        
        Exit Function
    End If
    
    If smFLASH = 1 Then
        LogPrint "Ack received..."
        LogPrint "read & send bin file..." & sFileName
        LogPrint " "
        smFLASH = 2  ' ack received - send first block
    End If
    
    If smFLASH = 3 And Mid(sResp, 4, 1) > Chr(&HEF) Then  ' ack received
        smFLASH = 4 ' nack
    End If
    
    If (smFLASH = 2 Or smFLASH = 3) And (dAdrNxt >= dSize) Then  ' test EOF
        Close dFile
        LogPrint " "
        LogPrint "End..."
        smFLASH = 0  ' wait for first ack
        sAdr = ""
        sMachine = 0
        Timer1.Interval = 0
        Timer1.Enabled = False
' invia il comando 0x80 - end flash
        write_firmware (Chr(&H80)) 'block length,   source device, pdu format, data=0x80(end)
        If bRS232 = True Then Disconnect

        BinFlashFileRead = 0
        Exit Function
    End If
    
    If smFLASH = 3 Then   ' ack received
        LogPrintCont ("k")
        smFLASH = 2
    End If
    
    If smFLASH = 4 Then   ' Nack received
         LogPrintCont ("r")
'        Winsock.SendData (sAdr + sBlock + sChk)  ' REINVIA lo stesso blocco ----->
'        smFLASH = 3  ' block sended - wait for ack
        
        smFLASH = 20 ' resend same block
        BinFlashFileRead = 0
'        Exit Function
    End If
    
    If smFLASH = 2 Then
        sBlock = ReadBin(FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES) ' 128*4 = 512
        sBlock3 = ""
        For i = 1 To (FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES) - 3 Step FLASH_BIN_WORD_BYTES
            sBlock3 = sBlock3 + Mid(sBlock, i, 3)
        Next
        smFLASH = 20
    End If
    
    If smFLASH = 20 Then
        iChk = 0
        iChk3 = 0
        iVhk = (FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES)
        iVhk = iVhk * &HFF
'        iVhk = (FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES * &HFF)
        
        For i = 1 To (FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES)
            iChk = iChk + Asc(Mid(sBlock, i, 1))
        Next
        For i = 1 To (Len(sBlock3))
            iChk3 = iChk3 + Asc(Mid(sBlock3, i, 1))
        Next

        If iChk <> iVhk Then  ' diverso da tutti 0xFF
            sAdr = BinString3(dAdr)
            fAdr = BinString3((dAdr / FLASH_BIN_WORD_BYTES) * 2) '512 bytes = 128 istruzioni da 4 bytes = 256 indirizzi
            LogPrintCont (".")
            
            
' invia il comando 0x01 - new flash block binary mode
'            write_firmware (Chr(1) + fAdr + Chr(64)) 'block length,   source device, pdu format, data=0x01(newblock), addressL, addressH, dataLength
            
'            sChk = BinString3(iChk)
' invia il comando 0x21 - new flash block binary mode
'            write_firmware (Chr(&H21) + fAdr + BinString2(FLASH_BLK_LEN * FLASH_BIN_WORD_BYTES)) '0x21(newblock), addressL, addressH, addressU, dataLength
' invia 1 blocco da 512 bytes - flash binary data
'            write_firmware (sBlock) 'block length = 0,  binary data 1-512

            sChk = BinString3(iChk3)
''''                sResp = ReadNoWait(1)
' invia il comando 0x22 - new flash block binary mode
            write_firmware (Chr(&H22) + fAdr + BinString2(Len(sBlock3))) '0x22(newblock), addressL, addressH, addressU, dataLength
' invia 1 blocco da 384 bytes - flash binary data
            
            write_firmware (sBlock3) 'block length = 0,  binary data 1-512

' invia il comando 0x02 - end flash block binary mode   oppure 0x03 (end flash block - test only, don't write)
            write_firmware (Chr(2) + sChk) 'data=0x02(endblock - write), checkL, checkH, checkU
'       TEST   write_firmware (Chr(3) + sChk) 'data=0x03(endblock - check), checkL, checkH, checkU
            
            smFLASH = 3  ' block sended - wait for ack
            BinFlashFileRead = 0
        Else
'            sAdr = BinString3(dAdr)
'            fAdr = BinString3((dAdr / FLASH_BIN_WORD_BYTES) * 2) '512 bytes = 128 istruzioni da 4 bytes = 256 indirizzi
'            LogPrintCont ("#")
            smFLASH = 2  ' all 0xFF - block NOT sended - wait for loop
            BinFlashFileRead = 1  ' loop request
        End If
        Exit Function
    End If
    
    BinFlashFileRead = 0
    End Function
Function ReadBin(lgth As Long)
    Dim sChar As String
    
    sChar = Space(lgth)  ' space(lof(1))
    Get dFile, , sChar
    dAdr = dAdrNxt
    dAdrNxt = dAdrNxt + lgth
    ReadBin = sChar
End Function

Function BinString2(ThisNumber)
    Dim iUnumber, iHnumber, iLnumber As Long
    '
    ' Convert a integer to a BIN string
    '
    If ThisNumber > 65535 Then
        iUnumber = ThisNumber \ 65536
        ThisNumber = ThisNumber - (iUnumber * 65536)
    End If
    iHnumber = ThisNumber \ 256
    iLnumber = ThisNumber - (iHnumber * 256)
    BinString2 = Chr(iLnumber) + Chr(iHnumber)
End Function

Function BinString3(ThisNumber)
    Dim iUnumber, iHLnumber As Long
    iUnumber = ThisNumber \ 65536
        
    iHLnumber = ThisNumber - (iUnumber * 65536)
    
    BinString3 = BinString2(iHLnumber) + Chr(iUnumber)
End Function

Private Sub flipX_Click()
            If flipX.Value = 1 Then
               LogPrint ("flipped X ")
            Else
               LogPrint ("normal X ")
            End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Close dFileLog
'    Close dFileL1
'    Close dFileL2
'    Close dFileL3
End Sub

Private Sub HoleSize_Change()
'     Xspeed = Int(Val(HoleSize.Text))
'     If Xspeed > 65 Then Xspeed = 65
'     If Xspeed > 0 Then speedFactor.Text = CStr(Xspeed)
'     SaveSetting settingName, "InitValues", "HoleSize", HoleSize.Text
End Sub




Private Sub reverse_Click()
            If reverse.Value = 1 Then
                leading.Enabled = True
                trailing.Enabled = True
                leading_label.Caption = "Leading burn"
            Else
                leading.Enabled = True
'                trailing.Enabled = False
                trailing.Enabled = True
                leading_label.Caption = "Y leadmargin"
            End If
End Sub

Private Sub Rich_KeyPress(KeyAscii As Integer)
    Dim sChar As String
    
    If sMachine = 10 Then       ' serial monitor
'        LogPrint (">" + Chr(KeyAscii))
        sChar = Chr(KeyAscii)
        WriteBuf (sChar)
'       sResp = ReadWait(1024, 5)
'       LogPrint ("response: " & sResp)
    End If
End Sub

Private Sub SerMonitor_Click()
    If sMachine = 0 Then
          ' serial monitor  ON
        Timer1.Enabled = False
        resp = TryConnect(baudCom)
        
        If resp <> 0 Then
            Flush
            Disconnect
        Else
            sMachine = 10        ' serial monitor ON
            firmware.Visible = False
            NewLoad.Visible = False
            test.Visible = False

'            Openb.Visible = False
            LogPrint "START serial ascii monitor mode"
            
'            WriteBuf ("@MA")
'            LogPrintContSer ("@MA")
'            sResp = ReadWait(1024, 5)
'            LogPrint (sResp)

            Timer1.Interval = 2
            Timer1.Enabled = True
        End If
    Else
    If sMachine = 10 Then       ' serial monitor
        Timer1.Enabled = False
        Disconnect
            sMachine = 0        ' serial monitor OFF
        firmware.Visible = True
        NewLoad.Visible = True
        test.Visible = True
        
        LogPrint "END serial monitor mode"
    End If
    End If

End Sub

Private Sub Singleside_Click()
    xref.Enabled = False
End Sub

Private Sub swapXY_Click()
            If swapXY.Value = 1 Then
               LogPrint ("swapped X-Y ")
            Else
               LogPrint ("normal X-Y ")
            End If
End Sub

Private Sub Timer1_Timer()
    Dim lMsg As Integer
    

    If sMachine = 10 Then       ' serial monitor
        sResp = ReadNoWait(16)
       
        If sResp <> "" Then
            LogPrintContSer (sResp)
        End If
    End If
        
    
    If sMachine = 1 Then
        If row <= Ypoints Then
        
            If bRS232 = False Then
                sResp = "a"
            Else
                sResp = ReadNoWait(1)
            End If
            
            LogPrintCont (sResp)
            
            If sResp = "a" Then
                sResp = "n"
                rowCtr.Caption = CStr(row)
                While sResp = "n"
                    LogPrintCont "."
                    WriteBuf (MakeRow(4))
                    If bRS232 = False Then
                        sResp = "k"
                    Else
                        sResp = ReadWait(1, 5)
                        If sResp = "" Then
                            LogPrintCont "t"
                        Else
                            LogPrintCont sResp
                        End If
                        frmPcbdraw.Refresh
                    End If
                Wend
                
                row = row + nRepeat
                rowEffective = rowEffective + 1
            End If
            
            If sResp = "c" Then
                LogPrint " CANCELLED"
                MsgBox "job CANCELLED "
                sMachine = 9
            End If
            If sResp = "b" Then
                LogPrint "end requested"
                If zType = 1 Then
                    WriteBuf ("@E")
                Else
                    WriteBuf ("@e")
                End If
                sMachine = 9
            End If
            
        Else
            sMachine = 2
            LogPrint "end of rows"
        End If
    End If
        
    If sMachine = 2 Then
        ' aspettare fine riga printer !!!!
        If bRS232 = False Then
            sResp = "b"
        Else
            sResp = ReadNoWait(1)
        End If
        If sResp <> "" Then
            LogPrint "end"
                If zType = 1 Then
                    WriteBuf ("@E")
                Else
                    WriteBuf ("@e")
                End If
            sMachine = 9
        End If
    End If
       
    If sMachine = 20 Then
        ' FIRMWARE update in corso !!!!
        ' riceve   0xLL  destination   pduformat    data...
        
        If bRS232 = False Then
            sResp = Chr(1)
        Else
            sResp = ReadNoWait(1)
        End If
        
        If sResp <> "" Then
            lMsg = Asc(sResp)
            If bRS232 = False Then
                sResp = String(lMsg, Chr(0))
            Else
                sResp = ReadWait(lMsg, 5)
            End If
            Do
                Loop While (BinFlashFileRead("") = 1)

        End If
    End If

End Sub
Private Sub VideoPlotRow(rowString)
Dim r As Integer
Dim l As Integer
Dim t As String

    t = Left(rowString, 1)
    r = Asc(Mid(rowString, 2, 1))
    l = Asc(Mid(rowString, 3, 1))
    
    If t = "r" Then
       VideoPlot "r", 9, r, Xbytes, Mid(rowString, 3, Xbytes)
    End If
    
    If t = "v" Then
       VideoPlot "v", 0, r, 1, Mid(rowString, 3, 1)
    End If
    
    If t = "z" Then
        If r > 15 Then
                r = r - 16
                VideoPlot "z", 1, r, l, Mid(rowString, 4, l)
        Else
                VideoPlot "z", 0, r, l, Mid(rowString, 4, l)
        End If
    End If
End Sub
Private Function VideoPlot(vMode, zMode, nDup, sLen, sData)
Dim column As Integer
Dim Xb As Integer
Dim iBit As Integer
Dim iC As Integer
Dim iR As Integer
Dim bi As Byte
Dim ctr As Integer
Dim cB As Byte

    If vMode = "r" Or sLen < Len(sData) Then
        dataVideo = sData
    Else
        If zMode = 0 Then dataVideo = String(Xbytes, Chr(0))
        column = 0
        Xb = 1
        cB = Asc(Mid(dataVideo, Xb, 1))
        iBit = 1
        bi = 0
        iC = 1
        While (iC <= sLen)
            ' un-zip  data stream
                ctr = Asc(Mid(sData, iC, 1))
                While (ctr)
                    If (bi > 0) Then
                        If (zMode = 0) Then cB = (cB Or bitVal(iBit))
                        If (zMode = 1) Then cB = (cB Xor bitVal(iBit))
                    End If
                    iBit = iBit + 1
                    ctr = ctr - 1
                    If (iBit > 8) Then
                        Mid(dataVideo, Xb, 1) = Chr(cB)
                        Xb = Xb + 1
                        If Xb > Xbytes Then Xb = Xbytes
                        cB = Asc(Mid(dataVideo, Xb, 1))
                        iBit = 1
                    End If
                Wend
                        
                bi = (bi Xor 1)
                iC = iC + 1
        Wend
        If Xb <= Xbytes Then Mid(dataVideo, Xb, 1) = Chr(cB)

    End If
    
    ' dataVideo deve coincidere con  dataString(row)
    
    For iR = 1 To nDup
        column = 0
        
        If dataVideo <> dataString(row) Then
            LogPrint ("ERR riga " & CStr(row))
        Else
            iBit = 1
        End If
        
            iBit = 1
            While Xbytes > Len(dataVideo)
                dataVideo = dataVideo & Chr(0)
            Wend
            For Xb = 1 To Xbytes
                For iBit = 1 To 8
                    If (Asc(Mid(dataVideo, Xb, 1)) And bitVal(iBit)) <> 0 Then
                        iC = SetPixelV(pctMain.hdc, column, row, vbGreen)
                    End If
                    column = column + 1
                Next
            Next
        ' End If
        
        row = row + 1
    Next
    pctMain.Refresh
End Function
Private Function Download_on_file_ask()
    Dim iC As Integer
    Dim iD As Integer
    
    startFileDir = GetSetting(settingName, "InitValues", "RowDir")
    CommonDialog1.InitDir = startFileDir
    CommonDialog1.filename = filename
    CommonDialog1.DialogTitle = "Export ROW file"
    CommonDialog1.DefaultExt = "row"
    CommonDialog1.Filter = "*.row"
    CommonDialog1.CancelError = True
    On Error GoTo DownLoad_KO
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename <> "" Then
        File = CommonDialog1.filename
        startFileDir = File
        FileExtensionPos = InStr(File, ".row")
        If FileExtensionPos = 0 Then
            GoTo DownLoad_KO
        End If
        iC = Len(startFileDir)
        iD = iC
        While iC > 0 And Mid(startFileDir, iC, 1) <> "\"
           iC = iC - 1
        Wend
        While iD > 0 And Mid(startFileDir, iD, 1) <> "."
           iD = iD - 1
        Wend
        If iC > 2 Then
           filename = Mid(startFileDir, iC + 1, iD - iC - 1)
           startFileDir = Left(startFileDir, iC - 1)
        Else
           filename = Mid(startFileDir, 5, iD - iC - 1)
           startFileDir = Left(startFileDir, 3)
        End If
        SaveSetting settingName, "InitValues", "RowDir", startFileDir
        Download_on_file_ask = 1
    Else
        Download_on_file_ask = 0
    End If
    Exit Function
    
DownLoad_KO:
        Download_on_file_ask = 0
        On Error GoTo 0
End Function

Private Sub Download_on_file()
    Dim myHeader As String
    Dim myRow As String
    Dim xS As Integer
    
    zType = 1
    rowEffective = 0
    row = 0
    
    LogPrint " "
    LogPrint "********************************************************************* "
    LogPrint "Export file..." & File
    dFile = FreeFile()
    Open File For Output As dFile

    myHeader = "X" + Chr(6) + "H" + MakeHeader(5)
        
    myHeader = myHeader + String(250, Chr(255))
    Print #dFile, Left(myHeader, 256);

'    While row < Ypoints
'        LogPrintCont "."
'        Print #dFile, MakeRow(4);
'        row = row + nRepeat
'        rowEffective = rowEffective + 1
'    Wend

' righe BOTTOM --------------------------------------------------
    row = 0
    ydatabot = 0
    While row < Ypoints
        LogPrintCont "."
        Print #dFile, MakeRow(41);
        row = row + nRepeat
        rowEffective = rowEffective + 1
    Wend

    Close dFile
    LogPrint " "
    LogPrint "********************************************************************* "
End Sub

Private Sub Download_double_on_file()
    Dim myHeader As String
    Dim myRow As String
    Dim xS As Integer
    Dim x As Integer
    
    zType = 1
    rowEffective = 0
    row = 0
    
    LogPrint " "
    LogPrint "********************************************************************* "
    LogPrint "Export file..." & File
    dFile = FreeFile()
    Open File For Output As dFile

    LogPrint "BOTTOM"
' header BOTTOM -------------------------------------------------
    myHeader = "X" + Chr(6) + "1" + MakeHeader(5)
        
    myHeader = myHeader + String(250, Chr(255))
    Print #dFile, Left(myHeader, 256);

' righe BOTTOM --------------------------------------------------
    row = 0
    ydatabot = 0
    While row < Ypoints
        LogPrintCont "."
        Print #dFile, MakeRow(41);
        row = row + nRepeat
        rowEffective = rowEffective + 1
    Wend
    Close dFile




    LogPrint " "
    LogPrint "********************************************************************* "
    Mid(File, FileExtensionPos, 4) = ".rot"
    LogPrint "Export file..." & File
    dFile = FreeFile()
    Open File For Output As dFile

    LogPrint "TOP"
' header TOP ----------------------------------------------------
    myHeader = "X" + Chr(6) + "2" + MakeHeader(5)
            
    myHeader = myHeader + String(250, Chr(255))
    Print #dFile, Left(myHeader, 256);

' righe TOP --------------------------------------------------
    row = 0
    ydatatop = 0
    While row < Ypoints
        LogPrintCont "."
        Print #dFile, MakeRow(42);
        row = row + nRepeat
        rowEffective = rowEffective + 1
    Wend

    Close dFile
    LogPrint " "
    LogPrint "********************************************************************* "
End Sub

Private Sub Download_holes_on_file()
    Dim myVar As Integer
    
    LogPrint " "
    LogPrint "********************************************************************* "
    LogPrint "*********************** H O L E S *********************************** "
    Mid(File, FileExtensionPos, 4) = ".roh"
    LogPrint "Export file..." & File

    dFile = FreeFile()
    
    Open File For Output As dFile   ' just for kill it
    Close dFile                     ' just for kill it
    
    Open File For Binary Access Write As dFile

    Put #dFile, , holeW

    For holeR = 1 To holeW
            LogPrintCont "X "
            LogPrintCont tabHoles(holeR).x
        myVar = tabHoles(holeR).x
        Put #dFile, , myVar
            LogPrintCont "   Y "
            LogPrintCont tabHoles(holeR).Y
        myVar = tabHoles(holeR).Y
        Put #dFile, , myVar
            LogPrint " "
    Next
' tabHoles(holeW).X = ...

    Close dFile
    LogPrint " "
    LogPrint "********************************************************************* "
End Sub
Private Sub test_Download_Click()
Dim rType As String
Dim iC As Integer
Dim rRow As String

    totaleN = 0
    totale0 = 0
    totale1 = 0
    totale2 = 0
    zType = 1
    rowEffective = 0
    row = 0
    dataVideo = String(Xbytes, Chr(0))
    
    pctMain.Cls
    pctMain.Picture = Nothing
    ydatabot = 0
    ydatatop = 0
    
    While row < Ypoints
        LogPrintCont "."
        
        If Doubleside.Value = True Then
            rRow = MakeRow(42)
        Else
            rRow = MakeRow(41)
        End If
        
'       Print #dFileL3, rRow;
        VideoPlotRow (rRow)
        
        rowEffective = rowEffective + 1
    Wend
End Sub
Private Sub test_Reload_Click()
Dim rType As String
Dim rLen As Integer
Dim iC As Integer
Dim dataPic As String

    totaleN = 0
    totale0 = 0
    totale1 = 0
    totale2 = 0
    zType = 1
    row = 0
    dataVideo = String(Xbytes, Chr(0))
    If bRS232 = False Then
        resp = 0
    Else
        resp = TryConnect(baudCom)
    End If
    
    If resp = 0 Then
        LogPrint "send burn simulation request"
        sResp = ""
        Flush
        WriteBuf ("@b")
        If bRS232 = False Then
            sResp = "k"
        Else
            sResp = ReadWait(1, 5)
        End If
        If Left(sResp, 1) <> "k" Then
            LogPrint sResp
            LogPrint "K.O. - request failed..."
            Exit Sub
        Else
            LogPrint Mid(sResp, 2)
        End If
    
        pctMain.Cls
        pctMain.Picture = Nothing
    
        While row < Ypoints
            LogPrintCont "."
            rType = ReadWait(1, 50)
            If rType <> "r" Then
                LogPrint "ROW TYPE ERROR!!!"
                MsgBox "Reload error - end"
                Disconnect
                Exit Sub
            End If
            nRepeat = Asc(ReadWait(1, 5))
            rLen = Asc(ReadWait(1, 5))
            dataPic = ReadWait(rLen, 5)
            iC = VideoPlot(rType, 9, nRepeat, Xbytes, Left(dataPic, Xbytes))
'                row = Ypoints + 1
            
        Wend
    End If
    Disconnect
End Sub
Private Sub Download_Click()
    
    Download.Visible = False
    totaleN = 0
    totale0 = 0
    totale1 = 0
    totale2 = 0
    zType = 1
    If Download_on_file_ask Then
        If Doubleside.Value = True Then
            Download_double_on_file
            Download_holes_on_file
        Else
            Download_on_file
            Download_holes_on_file
        End If
    Else
        MsgBox "Download error - video burn simulation now"
        test_Download_Click
    End If
    
    Download.Visible = True
End Sub

Private Sub pctMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim rc
    
    If Button = 10 Then
          LogPrint ("Row " & CStr(Y) + "  col " & CStr(x) + " clr " + CStr(GetPixel(pctMain.hdc, x * XscaleGraph, Y * YscaleGraph)))
          pctMain.ForeColor = vbYellow

          pctMain.FillStyle = 0
          pctMain.ForeColor = vbRed
          pctMain.FillColor = vbBlue
          pctMain.DrawWidth = 5
          rc = MoveToEx(pctMain.hdc, 1, 1, CurrentPoint)
          rc = LineTo(pctMain.hdc, (x * XscaleGraph), (Y * YscaleGraph))
       '  pctMain.ForeColor = vbYellow
       '  rc = LineTo(pctMain.hdc, (X * XscaleGraphR / Screen.TwipsPerPixelX), (Y * YscaleGraphR / Screen.TwipsPerPixelY))

          pctMain.Refresh
    End If
End Sub

Private Sub Form_Load()
    Dim t As Integer
    Dim x As Long
    Dim tmp
    
    settingName = DEFAULT_PARAM
'   settingName = config.Text
    verbose = 0
    tmp = GetSetting(settingName, "InitValues", "configName")
    If tmp <> "" Then
        config.Text = tmp
        settingName = tmp
    Else
        config.Text = settingName
    End If

    bRS232 = False
    bRS232 = True
    
    sMachine = 0
    
    bitVal(1) = &H80
    bitVal(2) = &H40
    bitVal(3) = &H20
    bitVal(4) = &H10
    bitVal(5) = &H8
    bitVal(6) = &H4
    bitVal(7) = &H2
    bitVal(8) = &H1

    pctDisplay = 1
    

    dFileLog = FreeFile()
    sFileLogName = "HnetTRACE_" + CStr(Year(Now)) + CStr(Month(Now)) + CStr(Day(Now)) + "_" + CStr(Hour(Now)) + CStr(Minute(Now)) + ".txt"
'    Open sFileLogName For Output As dFileLog

'    dFileL1 = FreeFile()
'    Open "d:\fileL1.log" For Output As dFileL1
'    dFileL2 = FreeFile()
'    Open "d:\fileL2.log" For Output As dFileL2
'    dFileL3 = FreeFile()
'    Open "d:\fileL3.log" For Output As dFileL3


    t = 0
    ' Make the picture box bigger than the form:
'         pctMain.Move 0, 0, 1.4 * ScaleWidth, 1.2 * ScaleHeight
    pctMain.Move 0, 0, 1.4 * ScaleWidth, 0.8 * ScaleHeight
    ' Position and size the first TextBox:
    Text2(0).Move 0, 0, pctMain.width / 2, pctMain.Height / 20
    ' Place some sample controls in the picture box:
    For t = 1 To 20
       Load Text2(t)
       Text2(t).Visible = False
       Text2(t).Left = t * pctMain.Height / 20
       Text2(t).Top = Text2(t).Left
    Next
    frmPcbdraw.Move 10, 20
    frmPcbdraw.WindowState = vbMaximized

    pctMain.width = 9000
    pctMain.Height = 9000
'   Xscale = 0.18
    t = Screen.TwipsPerPixelY
    
    XscaleGraph = pctMain.width / pctMain.ScaleWidth
    YscaleGraph = pctMain.Height / pctMain.ScaleHeight
 
'   Xscale = 0.48     ' ogni riga corrisponde ad una riga laser (2 mils)
'   Yscale = 0.605   ' ogni riga corrisponde ad una riga laser (0.4 mils x 4step = 1,6mils)
    Xscale = 0.96     ' ogni riga corrisponde ad una riga laser (1 mils)
    Yscale = 1.21     ' ogni riga corrisponde ad una riga laser (0.4 mils x 2step = 0,8mils)
    
    Call Get_config
    
    For t = 1 To 1000
'        pctMain.DrawWidth = t ' trackSize
'        pctMain.Line (10, 10 + t * 30)-(2000, 10 + t * 30), vbYellow
        pctMain.PSet (t, t), vbRed
        x = pctMain.Point(t, t)
        x = pctMain.Point(t + 2, t)
    Next
    
    x = DrawPolygon(500, 250, 50, 8, 22)
    x = DrawCircle(800, 350, 50)
    x = DrawEllipse(1500, 1000, 200, 100)

    pctMain.FillStyle = 0
    pctMain.ForeColor = vbGreen
    pctMain.DrawWidth = 1
    For t = 85 To 91
        x = MoveToEx(pctMain.hdc, 130, t, CurrentPoint)
        x = LineTo(pctMain.hdc, 310, t)
        x = GetPixel(pctMain.hdc, 220, t)
        If x <> 65280 Then
            x = 0
        End If
    Next
    
    pctMain.ForeColor = vbYellow
    x = DrawLine(0, 500, 800, 500, 1)
    
End Sub
Private Sub Get_config()
    Dim tmp

    LogPrint "get parameters: " + config.Text

    iCom = 5
    sCom = GetSetting(settingName, "InitValues", "ComPort")
    
    tmp = GetSetting(settingName, "InitValues", "Xscale")
    If tmp <> "" Then xfactor.Text = tmp
    
    tmp = GetSetting(settingName, "InitValues", "Yscale")
    If tmp <> "" Then yfactor.Text = tmp
    
    tmp = GetSetting(settingName, "InitValues", "SpeedFactor")
    If tmp <> "" Then speedFactor.Text = tmp
    
    xfactor_Change
    yfactor_Change
    speedFactor_Change
    smooth.Text = "0"
    smooth_Change
    
    sCom = GetSetting(settingName, "InitValues", "ComPort")
    Select Case sCom
        Case "1"
            Comm_1_Click
        Case "2"
            Comm_2_Click
        Case "3"
            Comm_3_Click
        Case "4"
            Comm_4_Click
    End Select
    
    tmp = GetSetting(settingName, "InitValues", "Baud", "std")
    If tmp = "low" Then LowBaud_Click
    If tmp = "std" Then StdBaud_Click
    If tmp = "high" Then HighBaud_Click
    
End Sub
Private Sub Comm_1_Click()
    iCom = 1
    LogPrint ("communication on COM1:")
    Comm_1.Checked = True
    Comm_2.Checked = False
    Comm_3.Checked = False
    Comm_4.Checked = False
    SaveSetting settingName, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_2_Click()
    iCom = 2
    LogPrint ("communication on COM2:")
    Comm_2.Checked = True
    Comm_1.Checked = False
    Comm_3.Checked = False
    Comm_4.Checked = False
    SaveSetting settingName, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_3_Click()
    iCom = 3
    LogPrint ("communication on COM3:")
    Comm_3.Checked = True
    Comm_2.Checked = False
    Comm_1.Checked = False
    Comm_4.Checked = False
    SaveSetting settingName, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_4_Click()
    iCom = 4
    LogPrint ("communication on COM4:")
    Comm_4.Checked = True
    Comm_2.Checked = False
    Comm_3.Checked = False
    Comm_1.Checked = False
    SaveSetting settingName, "InitValues", "ComPort", iCom
End Sub
Private Sub LowBaud_Click()
    LowBaud.Checked = True
    StdBaud.Checked = False
    HighBaud.Checked = False
    
    LogPrint ("communication baud: 57600")
    baudCom = 57600
    SaveSetting settingName, "InitValues", "Baud", "low"
End Sub
Private Sub StdBaud_Click()
    LowBaud.Checked = False
    StdBaud.Checked = True
    HighBaud.Checked = False
    
    LogPrint ("communication baud: 115200")
    baudCom = 115200
    SaveSetting settingName, "InitValues", "Baud", "std"
End Sub
Private Sub HighBaud_Click()
    LowBaud.Checked = False
    StdBaud.Checked = False
    HighBaud.Checked = True
    
    LogPrint ("communication baud: 230400")
    baudCom = 230400
    SaveSetting settingName, "InitValues", "Baud", "high"
End Sub
Function TryConnect(iBaud)
    Dim sResp
    LogPrint ("COM port initialize")
    On Error Resume Next        ' Abilito l'intercettazione degli errori
    
    ' Initialize Communications
    lngStatus = CommClose(iCom)
'    lngStatus = CommOpen(iCom, "COM" & CStr(iCom), "baud=" & iBaud & " parity=N data=8 stop=1")
    lngStatus = CommOpen(iCom, "COM" & CStr(iCom), "baud=" & baudCom & " parity=N data=8 stop=1")
    If lngStatus <> 0 Then
    ' Handle error.
        lngStatus = CommGetError(strError)
        MsgBox "COM Error: " & strError
        TryConnect = 9
        Exit Function
    End If
    
    On Error GoTo 0
    
'    SetBaudRate MSComm1, baudCom
    
    Flush
    LogPrint "TRYING communication"
    ' verifico se risponde a query come programmatore
    WriteBuf ("@q")
    sResp = ReadWait(1024, 5)
    LogPrint ("response: " & sResp)
    'response_display
    If Len(sResp > 1) And Left(sResp, 2) = "qk" Then
'        bEcho = True
        sResp = Mid(sResp, 2)
    End If
    If sResp = "" Or Left(sResp, 1) <> "k" Then
        LogPrint "K.O. - communication failed...", vbRed
        Flush
        TryConnect = 1 ' ko - retry
    Else
        If Len(sResp) > 1 Then
            LogPrintCont "OK!! - firmware version is "
            LogPrint Mid(sResp, 2), vbRed
        Else
            LogPrint "OK!! ", vbRed
        End If
        TryConnect = 0 ' ok
    End If

End Function

Private Sub Disconnect()
    Dim iC
    On Error Resume Next    ' Abilito l'intercettazione degli errori
    lngStatus = CommClose(iCom)
    LogPrint "SERIAL CHANNEL CLOSED"
End Sub
Function Write_synch(sString)
    Dim sResp, k, Y
    WriteBuf sString
'    For k = 1 To 1000000: y = y + 1 / 7: Next
    sResp = ReadWait(1, 250)
    Select Case sResp
    Case "k"
       Write_synch = 1
    Case ""
       LogPrint "PicLaser non risponde al comando (" & Left(sString, 1) & ")... [" & sResp & "]"
       MsgBox ("PicLaser non risponde al comando (" & Left(sString, 1) & ")...")
       Write_synch = 0
    Case Else
       LogPrint "PicLaser non riconosce il comando (" & Left(sString, 1) & ")... [" & sResp & "]"
       MsgBox ("PicLaser non riconosce il comando (" & Left(sString, 1) & ")...")
       Write_synch = 0
    End Select
End Function

Function Wait_synch()
    Dim sResp
    sResp = ReadWait(1, 1)
    If (sResp = "" Or sResp <> "k") Then
       LogPrint "PicLaser non riceve ack comando precedente- rx: " & sResp
       MsgBox ("PicLaser non riceve ack comando precedente")
       Wait_synch = 0
    Else
       Wait_synch = 1
    End If
End Function

Function WriteBuf(sString)
    Dim iP, iTim

    For iP = 1 To Len(sString)
        If bRS232 = True Then
            lngStatus = CommWrite(iCom, Mid(sString, iP, 1))
            If verbose = 1 Then LogPrintCont (HexString(Asc(Mid(sString, iP, 1)), 2))
        End If
    Next
    If verbose = 1 Then LogPrint (" ")
    totaleN = totaleN + Len(sString)
End Function

Function ReadWait(length, timeout)
    Dim tim
    Dim Chread As String
    Dim nrbytes
    Dim lngSize As Long
    Dim nrb As Integer
    ReadWait = ""
    Chread = " "
    tim = 0
    lngSize = 1
'    ShapeCom.FillColor = &HFF&
'    ShapeCom.Refresh
    Do
       'Timer1_Timer
        lngStatus = CommRead(iCom, Chread, lngSize)
'        If Chread = "" Then
        If lngStatus <= 0 Then
            tim = tim + 1
        Else
            If verbose = 1 Then
                For nrb = 1 To Len(Chread)
                    LogPrintCont (HexString(Asc(Mid(Chread, nrb, 1)), 2))
                Next
            End If
            SerBuffer = SerBuffer & Chread
'           Text1.Text = Text1.Text & "."
'           Text1.Text = Text1.Text & chread
'           Text1.DataChanged = True
  ''          Text2.Text = Len(SerBuffer)
        End If
    Loop Until (Len(SerBuffer) >= length Or tim >= 10000 * timeout) ' 5 secondi timeout
    
    If Len(SerBuffer) > length Then
        ReadWait = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadWait = SerBuffer
        SerBuffer = ""
    End If
    If verbose = 1 Then
       LogPrint (" ")
    End If
'    ShapeCom.FillColor = &HFF00&
'    ShapeCom.Refresh
End Function

Function ReadNoWait(length)
    Dim tim
    Dim Chread As String
    Dim nrbytes
    Dim lngSize As Long
    Dim nrb
    ReadNoWait = ""
    tim = 0
    lngSize = 1
    Chread = " "
'    ShapeCom.FillColor = &HFF&
'    ShapeCom.Refresh
       
       'Timer1_Timer
        lngStatus = CommRead(iCom, Chread, lngSize)
        If lngStatus <= 0 Then
            tim = tim + 1
        Else
            SerBuffer = SerBuffer & Chread
            If verbose = 1 Then
                For nrb = 1 To Len(Chread)
                    LogPrintCont (HexString(Asc(Mid(Chread, nrb, 1)), 2))
                Next
            End If
'           Text1.Text = Text1.Text & "."
'           Text1.Text = Text1.Text & chread
'           Text1.DataChanged = True
  ''          Text2.Text = Len(SerBuffer)
        End If
    
    If Len(SerBuffer) > length Then
        If verbose = 1 Then LogPrint (SerBuffer)
        ReadNoWait = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadNoWait = SerBuffer
        SerBuffer = ""
    End If
'    ShapeCom.FillColor = &HFF00&
'    ShapeCom.Refresh
End Function

Function Flush()
    Dim iCol, Chread, SerBufferHex
    If bRS232 = False Then Exit Function

    lngStatus = CommFlush(iCom)
    If Len(SerBuffer) > 0 Then
        LogPrint Len(SerBuffer) & " bytes flushed: "
        For iCol = 1 To Len(SerBuffer)
            SerBufferHex = SerBufferHex & HexString(Asc(Mid(SerBuffer, iCol, 1)), 2)
        Next
        LogPrint (SerBufferHex)
        SerBuffer = ""
    End If
End Function
Function HexString(ThisNumber, length)
    Dim RetVal
    Dim CurLen
    '
    ' Convert a integer to a hex string and
    ' pad it with the desired number of zeros
    '
    RetVal = Hex(ThisNumber)
    CurLen = Len(RetVal)

    If CurLen < length Then
        RetVal = String(length - CurLen, "0") & RetVal
    End If

    HexString = RetVal
End Function


Sub Form_Resize()
         ' Position the scroll bars:
         HScroll1.Left = 0
         VScroll1.Top = 0
         If pctMain.width > ScaleWidth Then
            HScroll1.Top = ScaleHeight - HScroll1.Height
         Else
            HScroll1.Top = ScaleHeight
         End If
         If pctMain.Height > HScroll1.Top Then
            VScroll1.Left = ScaleWidth - VScroll1.width
            If pctMain.width > VScroll1.Left Then
               HScroll1.Top = ScaleHeight - HScroll1.Height
            End If
         Else
            VScroll1.Left = ScaleWidth
         End If
         HScroll1.width = ScaleWidth
         If HScroll1.Top > 0 Then VScroll1.Height = HScroll1.Top
         ' Set the scroll bar ranges
         HScroll1.Max = (pctMain.width - VScroll1.Left) / 8
         VScroll1.Max = pctMain.Height - HScroll1.Top
'         HScroll1.SmallChange = Abs(HScroll1.Max \ 16) + 1
'         HScroll1.LargeChange = Abs(HScroll1.Max \ 4) + 1
'         VScroll1.SmallChange = Abs(VScroll1.Max \ 16) + 1
'         VScroll1.LargeChange = Abs(VScroll1.Max \ 4) + 1
         HScroll1.SmallChange = Abs(HScroll1.Max \ 128) + 1
         HScroll1.LargeChange = Abs(HScroll1.Max \ 16) + 1
         VScroll1.SmallChange = Abs(VScroll1.Max \ 128) + 1
         VScroll1.LargeChange = Abs(VScroll1.Max \ 16) + 1
         HScroll1.ZOrder 0
         VScroll1.ZOrder 0
    End Sub
      Sub HScroll1_Change()
         Dim cmd As Long
         cmd = -HScroll1.Value
         cmd = cmd * 8
         pctMain.Left = cmd
'         pctSlave.Left = cmd
      End Sub

Private Sub test_Click()
Dim sResp As String
Dim resp  As Integer

 ' test_Reload_Click
 ' Exit Sub
  
  
' protocollo di comunicazione
' @x          test X speed

    If bRS232 = False Then
        resp = 0
    Else
        resp = TryConnect(baudCom)
    End If
    
    If resp = 0 Then
    
        LogPrint "send X test request"
        sResp = ""
        Flush
        WriteBuf ("@x")
    's
        WriteBuf (Chr(Xspeed))
        Flush
    '    Disconnect
     End If
     Disconnect
End Sub

Sub VScroll1_Change()
    pctMain.Top = -VScroll1.Value
'    pctSlave.Top = -VScroll1.Value
End Sub
Function DecWord(String4c)
    Dim Ret1
    Dim Ret2
    '
    Ret1 = DecByte(Left(String4c, 2))
    Ret2 = DecByte(Right(String4c, 2))

    DecWord = Ret2 + (Ret1 * 256)

End Function
Function DecByte(String2c)
    Dim RetVal, hexval
    hexval = "0123456789ABCDEF"
    RetVal = 16 * (InStr(hexval, Left(String2c, 1)) - 1)

    DecByte = RetVal + (InStr(hexval, Right(String2c, 1)) - 1)

End Function
Private Sub NewLoad_Click()
    Dim iC As Integer
    Dim iD As Integer

    Xborder = 10
    Xborder = Val(Xleft.Text)
    Yborder = 10
    Yborder = Val(leading.Text)
    totale0 = 0
    totale1 = 0
    totale2 = 0
    totaleN = 0
    
    Download.Visible = False
        
    startFileDir = GetSetting(settingName, "InitValues", "PcbDir")
    CommonDialog1.InitDir = startFileDir
    CommonDialog1.filename = ""
    CommonDialog1.DialogTitle = "Import pcb file"
    CommonDialog1.DefaultExt = "gbl"
    CommonDialog1.Filter = "*.gbl"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename <> "" Then
        LogPrint ("clearscreen")
         
        pctMain.DrawWidth = 1
        pctMain.Cls
        pctMain.Picture = Nothing
        pctMain.ForeColor = RGB(50, 90, 50)
             
        File = UCase(CommonDialog1.filename)
        iExtension = Len(File)
        While Mid(File, iExtension, 1) <> "."
            iExtension = iExtension - 1
        Wend
        
        startFileDir = File
        iC = Len(startFileDir)
        iD = iC
        While iC > 0 And Mid(startFileDir, iC, 1) <> "\"
           iC = iC - 1
        Wend
        While iD > 0 And Mid(startFileDir, iD, 1) <> "."
           iD = iD - 1
        Wend
        If iC > 2 Then
           filename = Mid(startFileDir, iC + 1, iD - iC - 1)
           startFileDir = Left(startFileDir, iC - 1)
        Else
           filename = Mid(startFileDir, 5, iD - iC - 1)
           startFileDir = Left(startFileDir, 3)
        End If
        
        SaveSetting settingName, "InitValues", "PcbDir", startFileDir
        
        filenameBotTmp = filename + "Bot.bmp"
        filenameTopTmp = filename + "Top.bmp"
        
        frmPcbdraw.Caption = "E-Laserpcb 5.2 " + File + "                   " + filename
        
        LogPrint " "
        LogPrint "============================================================== "
        LogPrint "read & convert file..." & File
        
        
        units = "IN"
        
        If (Mid(File, iExtension, 4) = ".BMP") Then
           Call BMP_Load
                
    ' crea una immagine ritagliata alle dimensioni richieste in pctSlave
            pctSlave.width = Xpoints + 4
            pctSlave.Height = Ypoints + 4
            
            pctMain.Picture = pctMain.Image
            pctSlave.AutoRedraw = True
            pctSlave.PaintPicture pctMain.Picture, 0, 0, Xpoints + 4, Ypoints + 4, 0, 0, Xpoints, Ypoints
            pctSlave.Picture = pctSlave.Image
'            pctMain.Picture = Nothing
    ' la salva su disco
            LogPrint ("Save temp file")
'            SavePicture pctSlave.Image, TMPFILE
            SavePicture pctSlave.Image, filenameBotTmp
    ' scansiona il file bitmap
'           currentRow = 0
'           LogPrint ("Scan temp file")
            CurrentPointMin = 60000
            CurrentPointMax = 0
            Call ScanBitmap(filenameBotTmp, 0)
'           ScanScreen (1)  ' 1: sfondo bianco
        Else
           Call GBL_Load(0) ' preparazione
           Call GBL_Load(1) ' load bottom
                
    ' crea una immagine ritagliata alle dimensioni richieste in pctSlave
            pctSlave.width = Xpoints + 4
            pctSlave.Height = Ypoints + 4
            
            pctMain.Picture = pctMain.Image
            pctSlave.AutoRedraw = True
            pctSlave.PaintPicture pctMain.Picture, 0, 0, Xpoints + 4, Ypoints, 0, 0, Xpoints + 4, Ypoints
            pctSlave.Picture = pctSlave.Image
'           pctMain.Picture = Nothing

    ' la salva su disco
            LogPrint ("Save temp file")
'            SavePicture pctSlave.Image, TMPFILE
            SavePicture pctSlave.Image, filenameBotTmp
    
    ' scansiona il file bitmap
'           LogPrint ("Scan temp file")
'           currentRow = 0
            CurrentPointMin = 60000
            CurrentPointMax = 0
            Call ScanBitmap(filenameBotTmp, 0)
'           ScanScreen (0)  ' 0: sfondo nero

    ' genera vettore dataBot()
            row = 0
            xdatabot = 0
            zType = 1
            While row < Ypoints
                dataBot(xdatabot) = MakeRow(4)
                
                row = row + nRepeat
                xdatabot = xdatabot + 1
            Wend





' * richiesto doppia faccia?  carica ed elabora lato TOP
            If Doubleside.Value = True Then
            
'''''''                flipX.Value = flipX.Value Xor 1

                Call GBL_Load(2) ' load top
        ' crea una immagine ritagliata alle dimensioni richieste in pctSlave
                pctSlave.width = Xpoints + 4
                pctSlave.Height = Ypoints + 4
                pctMain.Picture = pctMain.Image
                pctSlave.AutoRedraw = True
                pctSlave.PaintPicture pctMain.Picture, 0, 0, Xpoints + 4, Ypoints, 0, 0, Xpoints + 4, Ypoints
                pctSlave.Picture = pctSlave.Image
        ' la salva su disco
                LogPrint ("Save temp file")
'                SavePicture pctSlave.Image, TMPFILE
                SavePicture pctSlave.Image, filenameTopTmp
        ' scansiona il file bitmap
                totale0 = 0
                totale1 = 0
                totale2 = 0
                totaleN = 0

    '           LogPrint ("Scan temp file")
'               currentRow = 0
                CurrentPointMin = 60000
                CurrentPointMax = 0
                Call ScanBitmap(filenameTopTmp, 0)
    
        ' genera vettore dataTop()
                row = 0
                xdatatop = 0
                While row < Ypoints
                    dataTop(xdatatop) = MakeRow(4)
                
'                    Print #dFileL2, dataTop(xdatatop);
                    
                    row = row + nRepeat
                    xdatatop = xdatatop + 1
                Wend
            
'''''''                flipX.Value = flipX.Value Xor 1

            End If

    End If
    LogPrint ("END OF SCAN - DATA READY")
       
    Download.Visible = True
    frmPcbdraw.Refresh
End If
End Sub
Private Function ScanBitmap(filenamebmp, parmColor)
' --------------------------------------------------------------------------------------------------------
' legge bitmap e costruisce le strutture     dataString()    dataZip0()    dataZip1()    DataZip2()
' --------------------------------------------------------------------------------------------------------
    Dim iD As Integer
    Dim iX As Integer
    Dim iY As Integer
    Dim Ypt As Integer
    Dim iBit As Integer
    Dim iColor As Long
    Dim xorByte As Byte
    Dim xorString As String
    Dim backColor As Double
    Dim Xmils As Integer
    Dim Ymils As Integer
    
    Dim iOffset As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim lAdr As Long
    
    Dim iBitPixel As Integer
    Dim iBytePixel As Integer
    Dim xPixelM As Long
    Dim yPixelM As Long
    Dim iRowDim As Integer
    Dim iRowFinal As Integer
    Dim sPixVal As String
    Dim sRowVal As String
    Dim nWidth, nHeight
    Dim rc
    Dim rgbback As String
    Dim rgbtop  As String
    Dim CurrentPoint As Long

    frmPcbdraw.MousePointer = vbHourglass
    ProgressBar1.Max = Ypoints
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    currentRow = 0

    If parmColor = 0 Then
        rgbback = String$(3, Chr(0))
    Else
        rgbback = String$(3, Chr(255))
    End If
    
    Xmils = Xpoints / Xscale
    Ymils = Ypoints / Yscale
    LogPrint ("dimension mils " + CStr(Xmils) + " x " + CStr(Ymils))
    LogPrint ("dimension mm " + CStr(Int(Xmils * 0.0254 + 0.9)) + " x " + CStr(Int(Ymils * 0.0254 + 0.9)) + " <================")
    frmPcbdraw.Refresh

    
    LogPrint ("scanning temp file ")
    
    
    frmPcbdraw.Refresh
    

' ============================================= legge le caratteristiche del file BMP ============================================
    dFile = FreeFile()
'    Open TMPFILE For Binary As dFile
    Open filenamebmp For Binary As dFile
    iOffset = BMP_GetLong(&HA)
'    LogPrint ("bitmap start at offset " + CStr(iOffset))
    iWidth = BMP_GetLong(&H12)
    iHeight = BMP_GetLong(&H16)
'    LogPrint ("bitmap dimension (pixel): " + CStr(iWidth) + " x " + CStr(iHeight))
    iBitPixel = BMP_GetInt(&H1C)
    iBytePixel = iBitPixel / 8
'    LogPrint ("bits per pixel : " + CStr(iBitPixel))

    dAdr = iOffset ' 54
    rgbback = BMP_GetStr(dAdr, iBytePixel)
    
    If parmColor = 0 Then
        rgbtop = String$(3, Chr(255))
    Else
        rgbtop = String$(3, Chr(0))
    End If

'    dAdr = iOffset - 3
'    sRowVal = BMP_GetStr(dAdr, 3)
    iRowDim = iBytePixel * iWidth
    iRowFinal = (iRowDim Mod 4) + iRowDim '6732

' l'immagine è registrata in ordine Y rovesciato
' calcolo indirizzo ultima riga e poi si retrocede ad ogni riga
    lAdr = Ypoints + 1
    lAdr = lAdr * iRowFinal
    lAdr = lAdr + iOffset '17058942
    ' currentRow = indice di partenza e di riempimento
    ' iY = indice di riempimento locale
    ' Ypt = coordinata Y in lavorazione
    iY = currentRow
    
    For Ypt = 0 To Ypoints                    ' ciclo per ogni RIGA 2533
        ProgressBar1.Value = Ypt
        dataString(iY) = String(Xbytes, &H0)
        dataZip0(iY) = ""
        dataZip1(iY) = ""
        dataZip2(iY) = ""
        dataZipN = ""
        iX = 0
        num0 = 0
        num1 = 0

'        sRowVal = BMP_GetStr(0, iRowFinal) ' scansione in avanti
'        LogPrintCont (".")

    ' scansione in indietro
        sRowVal = BMP_GetStr(lAdr, iRowFinal)
        lAdr = lAdr - iRowFinal
        
        iD = 0
        iX = 1
        For iD = 1 To Xbytes
            ds = &H0
            For iBit = 1 To 8
                CurrentPoint = ((iD - 1) * 8) + iBit
                sPixVal = Mid(sRowVal, iX, 3)
                If smtpoint < 3 Then
                    If sPixVal = rgbback Then
                        If num1 > 0 Then Zip0_1 (iY)
                        num0 = num0 + 1
                    Else
                        If num0 > 0 Then Zip0_0 (iY)
                        ds = ds Or bitVal(iBit)
                        If (CurrentPoint < CurrentPointMin) Then CurrentPointMin = CurrentPoint
                        If (CurrentPoint > CurrentPointMax) Then CurrentPointMax = CurrentPoint
                        num1 = num1 + 1
                    End If
                Else
                    If sPixVal <> rgbtop Then
                        If num1 > 0 Then Zip0_1 (iY)
                        num0 = num0 + 1
                    Else
                        If num0 > 0 Then Zip0_0 (iY)
                        ds = ds Or bitVal(iBit)
                        If (CurrentPoint < CurrentPointMin) Then CurrentPointMin = CurrentPoint
                        If (CurrentPoint > CurrentPointMax) Then CurrentPointMax = CurrentPoint
                        num1 = num1 + 1
                    End If
                End If
                iX = iX + 3
            Next
            Mid(dataString(iY), iD, 1) = Chr(ds)
        Next
        If num0 > 0 Then Zip0_0 (iY)
        If num1 > 0 Then Zip0_1 (iY)
        
        If iY > 0 Then
            num0 = 0
            num1 = 0
            For iD = 1 To Xbytes
                xorByte = Asc(Mid(dataString(iY), iD, 1)) Xor Asc(Mid(dataString(iY - 1), iD, 1))
                For iBit = 1 To 8
                    iColor = xorByte And bitVal(iBit)
                    If iColor = 0 Then
                        If num1 > 0 Then Zip1_1 (iY)
                        num0 = num0 + 1
                    Else
                        If num0 > 0 Then Zip1_0 (iY)
                        num1 = num1 + 1
                    End If
                Next
            Next
            If num0 > 255 Then
                num0 = 255
                Zip1_0 (iY)
            End If
            
            If num1 > 0 Then
                Zip1_1 (iY)
            End If
            
            If dataZip0(iY) <> dataZip0(iY - 1) Then
                totale0 = totale0 + Len(dataZip0(iY))
                totale1 = totale1 + Len(dataZip1(iY))

                If Len(dataZip2(iY)) < Len(dataZip1(iY)) Then
                    totale2 = totale2 + Len(dataZip2(iY))
                Else
                    totale2 = totale2 + Len(dataZip1(iY))
                End If
                
                If Len(dataZip1(iY)) < Len(dataZip0(iY)) Then
                    If Len(dataZip2(iY)) < Len(dataZip1(iY)) Then
                        totaleN = totaleN + Len(dataZip2(iY))
                    Else
                        totaleN = totaleN + Len(dataZip1(iY))
                    End If
                Else
                    totaleN = totaleN + Len(dataZip0(iY))
                End If
            Else
            
            End If
        Else
            totale0 = totale0 + Len(dataZip0(iY))
            totale1 = totale1 + Len(dataZip0(iY))
            totale2 = totale2 + Len(dataZip0(iY))
        End If
        iY = iY + 1
    Next
    currentRow = iY
    Close dFile
    pctMain.Refresh
         
    LogPrint ("row data ready ")
    
    totale0 = totale0 + (Ypoints * 2)
    totale1 = totale1 + (Ypoints * 2)
    totale2 = totale2 + (Ypoints * 2)
    totaleN = totaleN + (Ypoints * 2)
    
    LogPrint ("number of rows " + CStr(Ypoints))
'    LogPrint ("zip 0 bytes " + CStr(totale0))
'    LogPrint ("zip 1 bytes " + CStr(totale1))
'    LogPrint ("zip 2 bytes " + CStr(totale2))
'    LogPrint ("zip M bytes " + CStr(totaleN))
     LogPrint ("zipped bytes " + CStr(totaleN))
             
    totaleN = Ypoints
    totaleN = totaleN * Xbytes
    LogPrint ("unzip bytes " + CStr(totaleN))
    LogPrint ("limit points " + CStr(CurrentPointMin) + " - " + CStr(CurrentPointMax))
    
    ProgressBar1.Visible = False
    frmPcbdraw.MousePointer = vbDefault
End Function
Private Function ScanScreen(parmColor)
    Dim iD As Integer
    Dim iX As Integer
    Dim iY As Integer
    Dim iBit As Integer
    Dim iColor As Long
    Dim xorByte As Byte
    Dim xorString As String
    Dim backColor As Double
    Dim Xmils As Integer
    Dim Ymils As Integer

         frmPcbdraw.MousePointer = vbHourglass
         
         Xmils = Xpoints / Xscale
         Ymils = Ypoints / Yscale
         LogPrint ("dimension mils " + CStr(Xmils) + " x " + CStr(Ymils))
         LogPrint ("dimension mm " + CStr(Int(Xmils * 0.0254 + 0.9)) + " x " + CStr(Int(Ymils * 0.0254 + 0.9)) + " <================")
         frmPcbdraw.Refresh
 
         LogPrint ("scanning screen ")
         
         frmPcbdraw.Refresh

         backColor = GetPixel(pctMain.hdc, 1, 1)
         num0 = 0
         num1 = 0
         
         For iY = 0 To Ypoints                    ' ciclo per ogni RIGA
            dataString(iY) = String(Xbytes, &H0)
            dataZip0(iY) = ""
            dataZip1(iY) = ""
            dataZip2(iY) = ""
            dataZipN = ""
            iX = 0
            num0 = 0
            num1 = 0
            For iD = 1 To Xbytes
                ds = &H0
                For iBit = 1 To 8
                    iColor = GetPixel(pctMain.hdc, iX, iY)
                    If iColor = backColor Then
                        If num1 > 0 Then Zip0_1 (iY)
                        num0 = num0 + 1
                    Else
                        If num0 > 0 Then Zip0_0 (iY)
                        ds = ds Or bitVal(iBit)
                        num1 = num1 + 1
                    End If
                    iX = iX + 1
                Next
                Mid(dataString(iY), iD, 1) = Chr(ds)
            Next
            If num0 > 0 Then Zip0_0 (iY)
            If num1 > 0 Then Zip0_1 (iY)
            
            If iY > 0 Then
                num0 = 0
                num1 = 0
                For iD = 1 To Xbytes
                    xorByte = Asc(Mid(dataString(iY), iD, 1)) Xor Asc(Mid(dataString(iY - 1), iD, 1))
'                    dataZipN = dataZipN & Chr(xorByte)
                    For iBit = 1 To 8
                        iColor = xorByte And bitVal(iBit)
                        If iColor = 0 Then
                            If num1 > 0 Then Zip1_1 (iY)
                            num0 = num0 + 1
                        Else
                            If num0 > 0 Then Zip1_0 (iY)
                            num1 = num1 + 1
                        End If
                    Next
                Next
                If num0 > 255 Then
                    num0 = 255
                    Zip1_0 (iY)
                End If
                
                If num1 > 0 Then
                    Zip1_1 (iY)
                End If
                If dataZip0(iY) <> dataZip0(iY - 1) Then
                    totale0 = totale0 + Len(dataZip0(iY))
                    totale1 = totale1 + Len(dataZip1(iY))

                    If Len(dataZip2(iY)) < Len(dataZip1(iY)) Then
                        totale2 = totale2 + Len(dataZip2(iY))
                    Else
                        totale2 = totale2 + Len(dataZip1(iY))
                    End If
                    
                    If Len(dataZip1(iY)) < Len(dataZip0(iY)) Then
                        If Len(dataZip2(iY)) < Len(dataZip1(iY)) Then
                            totaleN = totaleN + Len(dataZip2(iY))
                        Else
                            totaleN = totaleN + Len(dataZip1(iY))
                        End If
                    Else
                        totaleN = totaleN + Len(dataZip0(iY))
                    End If
                End If
            Else
                totale0 = totale0 + Len(dataZip0(iY))
                totale1 = totale1 + Len(dataZip0(iY))
                totale2 = totale2 + Len(dataZip0(iY))
            End If
         Next
         pctMain.Refresh
         
         LogPrint ("row data ready ")
        
         totale0 = totale0 + (Ypoints * 2)
         totale1 = totale1 + (Ypoints * 2)
         totale2 = totale2 + (Ypoints * 2)
         totaleN = totaleN + (Ypoints * 2)
         
         LogPrint ("number of rows " + CStr(Ypoints))
         LogPrint ("zip 0 bytes " + CStr(totale0))
         LogPrint ("zip 1 bytes " + CStr(totale1))
         LogPrint ("zip 2 bytes " + CStr(totale2))
         LogPrint ("zip M bytes " + CStr(totaleN))
                  
         totaleN = Ypoints
         totaleN = totaleN * Xbytes
         LogPrint ("unzip bytes " + CStr(totaleN))
         
         frmPcbdraw.MousePointer = vbDefault
End Function
Private Function BMP_GetLong(address)
Dim dRead As Long
dRead = 0
Get dFile, address + 1, dRead
BMP_GetLong = dRead
End Function
Private Function BMP_GetInt(address)
Dim dRead As Integer
dRead = 0
Get dFile, address + 1, dRead
BMP_GetInt = dRead
End Function
Private Function BMP_GetStr(address, iLen)
Dim dRead As String
dRead = Space$(iLen)
If address > 0 Then
    Get dFile, address + 1, dRead
Else
    Get dFile, , dRead
End If
BMP_GetStr = dRead
End Function
Private Sub BMP_Load()
    Dim iC As Integer
    Dim iX As Integer
    Dim iY As Integer
    Dim iXnew As Single, iYnew As Single
    
    Dim iColor As Long
    Dim iOffset As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iBitPixel As Integer
    Dim iBytePixel As Integer
    Dim xPixelM As Long
    Dim yPixelM As Long
    Dim xDpi As Long, yDpi As Long
    Dim iL As Integer
    
    Dim sPixVal As String
    
    Dim pic As StdPicture
    Dim sgnRatioX As Double, sgnRatioY As Double
    Dim nWidth, nHeight
    Dim rc

    swapXY.Value = 0
    flipX.Value = 0

    pctMain.DrawWidth = 1
    pctMain.Picture = LoadPicture("")
    pctMain.Cls
    pctMain.ForeColor = RGB(50, 90, 50)
    pctMain.Refresh


' ============================================= legge le caratteristiche del file BMP ============================================
    dFile = FreeFile()
    Open File For Binary As dFile
    iOffset = BMP_GetLong(&HA)
    LogPrint ("bitmap start at offset " + CStr(iOffset))
    iWidth = BMP_GetLong(&H12)
    iHeight = BMP_GetLong(&H16)
    LogPrint ("bitmap dimension (pixel): " + CStr(iWidth) + " x " + CStr(iHeight))
    iBitPixel = BMP_GetInt(&H1C)
    iBytePixel = iBitPixel / 8
    LogPrint ("bits per pixel : " + CStr(iBitPixel))
    xPixelM = BMP_GetLong(&H26)
    yPixelM = BMP_GetLong(&H2A)
    If xPixelM = 0 Then xPixelM = 23622 'pixel per metro
    If yPixelM = 0 Then yPixelM = 23622 'pixel per metro
'    xDpi = (xPixelM * 254) / 10000
'    yDpi = (yPixelM * 254) / 10000
    xDpi = (xPixelM / (10000 / 254))
    yDpi = (yPixelM / (10000 / 254))
    LogPrint ("xPixel: " + CStr(xPixelM) + "  y: " + CStr(yPixelM))
    LogPrint ("bitmap dpi : " + CStr(xDpi) + " x " + CStr(yDpi))

'        dAdr = iOffset
'        sPixVal = BMP_GetStr(dAdr, iBytePixel)
'
'        For iY = 1 To iHigh
'            iL = iL + 1
'            If iL > 499 Then
'                iL = 0
'                LogPrint (CStr(iY))
'            End If
'            sPixVal = BMP_GetStr(0, iBytePixel * iWidth)
'            For iX = 1 To iWidth
'                sPixVal = BMP_GetStr(0, iBytePixel)
'            Next
'        Next
    Close dFile

    
' ================================ carica disegno BMP su oggetto "pic"  =======================================================
'    pctMain.Visible = False
'    pctSlave.Visible = True
    pctMain.Refresh
'    pctSlave.Refresh
    pctDisplay = 2
    
    LogPrint ("load bmp")
    
    Set pic = LoadPicture(File)
    
    
' ================================ rappresenta disegno BMP su pctMain in scala ridotta ========================================
    iWidth = ScaleX(pic.width, vbHimetric, vbPixels)
    iHeight = ScaleY(pic.Height, vbHimetric, vbPixels)
    
    sgnRatioX = (pctMain.ScaleWidth / iWidth) / 8
    sgnRatioY = (pctMain.ScaleHeight / iHeight) / 8
    
    nWidth = iWidth * sgnRatioX
    nHeight = iHeight * sgnRatioY
    
    pctMain.AutoRedraw = True
    pctMain.PaintPicture pic, 0, 0, nWidth, nHeight
       
    frmPcbdraw.Refresh
    pctMain.Refresh
    
    
' ================================ scansiona disegno BMP in scala ridotta (pctMain) per cercare i margini ============================
    
'    iC = MsgBox("Click on righ down angle to proceed", vbExclamation, "Scanning")
    
    LogPrint ("scan for margins...")
    Ypoints = 0
    Xpoints = 0
    Xstart = 9999
    Ystart = 9999
    
    Xbytes = 100
    For iY = 2 To nHeight - 50 ' 600dpi: 8 inches : 20cm high max
        For iX = 2 To nWidth - 5 ' 600dpi: 8 inches : 20cm high max
            iColor = GetPixel(pctMain.hdc, iX, iY)
            If iColor = &HFFFFFF Then
                iC = SetPixelV(pctMain.hdc, iX, iY, vbBlack)
            Else
                iC = SetPixelV(pctMain.hdc, iX, iY, vbRed)
                If iX > Xpoints Then Xpoints = iX
                If iX < Xstart Then Xstart = iX
                If iY > Ypoints Then Ypoints = iY
                If iY < Ystart Then Ystart = iY
            End If
        Next
        pctMain.Refresh
     Next
     Xpoints = Xpoints + 1
     Ypoints = Ypoints + 1
     Xstart = Xstart - 1
     Ystart = Ystart - 1
     
     Xpoints = Xpoints / sgnRatioX
     Xstart = Xstart / sgnRatioX
     Ypoints = Ypoints / sgnRatioY
     Ystart = Ystart / sgnRatioY

     LogPrint ("X min: " + CStr(Xstart) + " max: " + CStr(Xpoints))
     LogPrint ("Y min: " + CStr(Ystart) + " max: " + CStr(Ypoints))

' ================================ calcola la scala opportuna per rappresentare il BMP a 1 pixel x 2mils ============================
     
     sgnRatioX = pctMain.ScaleWidth / pctMain.width
     sgnRatioY = pctMain.ScaleHeight / pctMain.Height

     sgnRatioX = Xscale * 1000 / xDpi  ' 600 pixel x pollice
     sgnRatioY = Yscale * 1000 / yDpi  ' 600 pixel x pollice

     Xpoints = (Xpoints - Xstart + 1) * sgnRatioX
     Ypoints = (Ypoints - Ystart + 2) * sgnRatioY
     Xstart = Xstart * sgnRatioX
     Ystart = Ystart * sgnRatioY
          
     nWidth = iWidth * sgnRatioX
     nHeight = iHeight * sgnRatioY
     
' ================================ rappresenta disegno BMP su pctMain in scala corretta e offset giusti =============================

     pctMain.Cls
     pctMain.Refresh
     pctMain.PaintPicture pic, -Xstart, -Ystart, nWidth, nHeight
      
' ================================ righe verdi di separazione =============================
    Xbytes = (Xpoints / 8) + 1
    If (Xbytes / 4) <> (Xbytes \ 4) Then Xbytes = (Xbytes \ 4 + 1) * 4 ' allineamento DWORD  v.5
    
    Xpoints = Xbytes * 8
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbGreen
    pctMain.DrawWidth = 1
    iColor = MoveToEx(pctMain.hdc, (Xpoints + 20), 0, CurrentPoint)
    iColor = LineTo(pctMain.hdc, (Xpoints + 20), (Ypoints + 10))
    iColor = LineTo(pctMain.hdc, 0, (Ypoints + 10))

    pctMain.Refresh
    pctDisplay = 1
    iC = 0
End Sub

Private Sub GBL_Load(loadType)
Dim iC As Integer
Dim iRc As Integer
Dim iX As Integer
Dim iY As Integer
Dim iBit As Integer
Dim iColor As Long
Dim iTemp As Integer

Dim xorByte As Byte
Dim xorString As String
    
         
If loadType = 0 Then    ' preparazione
' ============================= prepare =============================================
    XgerberScale = 1
    YgerberScale = 1
    xMin = 30000
    yMin = 30000
    xMax = -30000
    yMax = -30000
    If Doubleside.Value = True And Mid(File, iExtension, 4) <> ".GBL" Then ' double side valid only if .GBL selected
        Singleside.Value = True
        Doubleside.Value = False
    End If
    If Doubleside.Value = True And Dir(File) = "" Then                     ' double side valid only if .GBL exists
        Singleside.Value = True
        Doubleside.Value = False
    End If
    If Doubleside.Value = True And Dir(Left(File, iExtension) + "GTL") = "" Then ' double side valid only if .GTL exists
        Singleside.Value = True
        Doubleside.Value = False
    End If
    
' =============================compute dimensions & offsets=============================================
    LogPrint ("compute layer dimension")

    iRc = ScanFile(File, 1)
         
    If Doubleside.Value = True Then
        LogPrint ("compute layer .GTL")
        iRc = ScanFile(Left(File, iExtension) + "GTL", 1)
    End If

    If swapXY.Value = 1 Then
        iTemp = Yborder
        Yborder = Xborder
        Xborder = iTemp
    Else
    End If
    
    Xoffset = xMin - Xborder
    Yoffset = yMin - Yborder
    Xpoints = xMax - xMin + Xborder + 10
    Ypoints = yMax - yMin + Yborder + 10
'         LogPrint ("dimension mils " + CStr(Xpoints) + " x " + CStr(Ypoints))
'         LogPrint ("dimension mm " + CStr(Int(Xpoints * 0.0254 + 0.9)) + " x " + CStr(Int(Ypoints * 0.0254 + 0.9)))
'         frmPcbdraw.Refresh

    If Doubleside.Value = True And xref.Value = 1 Then
    '   SE NECESSARIO AUMENTARE L AREA
        Xpoints = Xpoints + 50
        Ypoints = Ypoints + 50
        Xoffset = Xoffset - 25
        Yoffset = Yoffset - 25
    End If
    
    If swapXY.Value = 1 Then
        iTemp = Ypoints
        Ypoints = Xpoints
        Xpoints = iTemp
    Else
    End If
    
    Xpoints = Xpoints * Xscale
    Ypoints = Ypoints * Yscale
    Xbytes = (Xpoints \ 8) + 1
    If (Xbytes / 4) <> (Xbytes \ 4) Then Xbytes = (Xbytes \ 4 + 1) * 4 ' allineamento DWORD  v.5
    Xpoints = Xbytes * 8
End If



If loadType > 0 Then    ' draw bottom
    pctMain.Visible = True
'    pctSlave.Visible = False
    pctDisplay = 1
    pctMain.Picture = LoadPicture("")
' ===================================draw frame=================================================
    LogPrint ("drawing layer")
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbGreen
    pctMain.DrawWidth = 1
    
'    If swapXY.Value = 1 Then
'        iColor = MoveToEx(pctMain.hdc, Ypoints + 5, 0, CurrentPoint)
'        iColor = LineTo(pctMain.hdc, Ypoints + 5, Xpoints + 5)
'        iColor = LineTo(pctMain.hdc, 0, Xpoints + 5)
'    Else
        iColor = MoveToEx(pctMain.hdc, Xpoints + 5, 0, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xpoints + 5, Ypoints + 5)
        iColor = LineTo(pctMain.hdc, 0, Ypoints + 5)
'    End If


    If Doubleside.Value = True And xref.Value = 1 Then
        pctMain.FillStyle = 0
        pctMain.ForeColor = vbYellow
        pctMain.DrawWidth = 6
' alto dx
        iColor = MoveToEx(pctMain.hdc, Xpoints - 10, 1, CurrentPoint) ' |
        iColor = LineTo(pctMain.hdc, Xpoints - 10, 40)                ' |
        iColor = MoveToEx(pctMain.hdc, Xpoints - 30, 20, CurrentPoint) ' --
        iColor = LineTo(pctMain.hdc, Xpoints + 10, 20)                 '
' basso dx
        iColor = MoveToEx(pctMain.hdc, Xpoints - 10, Ypoints - 30, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xpoints - 10, Ypoints + 10)
        iColor = MoveToEx(pctMain.hdc, Xpoints - 30, Ypoints - 10, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xpoints + 10, Ypoints - 10)
' alto sx
        iColor = MoveToEx(pctMain.hdc, Xborder - 20, 1, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xborder - 20, 40)
        iColor = MoveToEx(pctMain.hdc, Xborder - 35, 20, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xborder + 0, 20)
' basso sx
        iColor = MoveToEx(pctMain.hdc, Xborder - 20, Ypoints - 30, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xborder - 20, Ypoints + 10)
        iColor = MoveToEx(pctMain.hdc, Xborder - 35, Ypoints - 10, CurrentPoint)
        iColor = LineTo(pctMain.hdc, Xborder + 0, Ypoints - 10)

    Else
    End If


    If loadType = 1 Then    ' draw bottom
        iRc = ScanFile(File, 2)
    End If
    If loadType = 2 Then    ' draw top
'    If Doubleside.Value = True Then
        LogPrint ("draw layer .GTL")
        iRc = ScanFile(Left(File, iExtension) + "GTL", 2)
    End If
         
    pctMain.Refresh
    frmPcbdraw.Refresh
    
' ====================================draw holes===================================================
    holeW = 0
'
' new: se esiste il file GDG (drill guide) usare quello !
'
    iRc = ScanFile(Left(File, iExtension) + "GDG", 3)
    If iRc = 1 Then
        LogPrint ("draw holes from .GDG")
    End If
    If iRc = 0 And Mid(File, iExtension, 4) = ".GBL" Then
        LogPrint ("draw holes from .GTL")
        iRc = ScanFile(Left(File, iExtension) + "GTL", 4)
    End If
    If iRc = 0 And Mid(File, iExtension, 4) = ".GTL" Then
        LogPrint ("draw holes from .GBL")
        iRc = ScanFile(Left(File, iExtension) + "GBL", 4)
    End If
        
    pctMain.Refresh
End If


End Sub
Private Function ScanFile(filename, scanType)
    polarity = "D"
    Dcode = ""
    DcodN = 0
    Gcode = ""
    GcodN = 0
    Mcode = ""
    McodN = 0
    
    Xtemp = 0
    Xcurr = 0
    Xprev = 0

    Ytemp = 0
    Ycurr = 0
    Yprev = 0
    trackSize = 1
    macroMax = 0
    Select Case scanType
        Case 3
            Xhole = Int(Val(HoleSize.Text))
        Case 4
            Xhole = Int(Val(HoleSize.Text))
            If Xhole = 0 Then Xhole = 20
    End Select

    On Error GoTo ErrScan
    ScanFile = 0
    dFile = FreeFile()
    Open filename For Input As dFile
    On Error GoTo 0

    While Not EOF(dFile)
        ReadLine
'''        LogPrint (lineTxt)
        While linePtr <= lineLen
            lineChr = Mid(lineTxt, linePtr, 1)
            If (lineChr = "*") Then
                Select Case scanType
                    Case 1              ' calcolo dimensioni e posizione
                       execCmd_Compute
                    Case 2
                       execCmd_Draw          ' disegno
                    Case 3
                       execCmd_HoleGDG       ' foratura da formato GDG
                    Case 4
                       execCmd_HoleGTL_GBL   ' foratura da formato GBL o GTL
                    Case Else
                       LogPrint ("Scanfile process error!!!" + CStr(scanType))
                End Select
            
            ElseIf (lineChr = "%") Then
                getParam
            ElseIf (lineChr = "X") Then
                getXcoord
            ElseIf (lineChr = "Y") Then
                getYcoord
            ElseIf (lineChr = "D") Then
                commandD
            ElseIf (lineChr = "M") Then
                commandM
            ElseIf (lineChr = "G") Then
                commandG
            Else:
                unknownCmd
                linePtr = linePtr + 1
            End If
        Wend
    Wend
    Close dFile
    ScanFile = 1
    Exit Function
ErrScan:
    ScanFile = 0
End Function
Private Sub Zip0_0(iY)  ' costruisce zip0 - aggiunge alla riga DATAZIP0 il numero di bits a zero
    While num0 > 0
        If num0 > 255 Then
            dataZip0(iY) = dataZip0(iY) & Chr(255) & Chr(0)
            num0 = num0 - 255
        Else
            dataZip0(iY) = dataZip0(iY) & Chr(num0)
            num0 = 0
        End If
    Wend
    num0 = 0
End Sub
Private Sub Zip0_1(iY)  ' costruisce zip0 - aggiunge alla riga DATAZIP0 il numero di bits a uno
    While num1 > 0
        If num1 > 255 Then
            dataZip0(iY) = dataZip0(iY) & Chr(255) & Chr(0)
            num1 = num1 - 255
        Else
            dataZip0(iY) = dataZip0(iY) & Chr(num1)
            num1 = 0
        End If
    Wend
    num1 = 0
End Sub
Private Sub Zip1_0(iY)  ' costruisce zip1 - aggiunge alla riga DATAZIP1 il numero di bits a zero
    While num0 > 0
        If num0 > 255 Then
            dataZip1(iY) = dataZip1(iY) & Chr(255) & Chr(0)
            num0 = num0 - 255
        Else
            dataZip1(iY) = dataZip1(iY) & Chr(num0)
            num0 = 0
        End If
    Wend
    num0 = 0
End Sub
Private Sub Zip1_1(iY)  ' costruisce zip1 - aggiunge alla riga DATAZIP1 il numero di bits a uno
Dim h0 As Integer
Dim l0 As Integer
Dim h1 As Integer
Dim l1 As Integer

    h0 = Int(num0 / 256)
    l0 = num0 - (h0 * 256)
    h1 = Int(num1 / 256)
    l1 = num1 - (h1 * 256)
    dataZip2(iY) = dataZip2(iY) & Chr(h0 * 16 + h1) & Chr(l0) & Chr(l1) ' costruisce zip2 - aggiunge alla riga DATAZIP1 il numero di bits a uno
    While num1 > 0
        If num1 > 255 Then
            dataZip1(iY) = dataZip1(iY) & Chr(255) & Chr(0)
            num1 = num1 - 255
        Else
            dataZip1(iY) = dataZip1(iY) & Chr(num1)
            num1 = 0
        End If
    Wend
    num1 = 0
End Sub

Function f1_testX(ByVal x As Integer, ByVal delta As Integer)
    If (x - delta) < xMin Then xMin = x - delta
    If (x + delta) > xMax Then xMax = x + delta
End Function
Function f1_testY(ByVal Y As Integer, ByVal delta As Integer)
    If (Y - delta) < yMin Then yMin = Y - delta
    If (Y + delta) > yMax Then yMax = Y + delta
End Function
Sub execCmd_Compute()
Dim aspect, radius, width As Single
Dim i As Integer
Dim s As Single


    If DcodN = 1 Then
    ' draw
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        width = apertureN ' trackSize
        s = f1_testX(Xprev, apertureN / 2)
        s = f1_testX(Xcurr, apertureN / 2)
        s = f1_testY(Yprev, apertureN / 2)
        s = f1_testY(Ycurr, apertureN / 2)
    ElseIf DcodN = 2 Then
    ' set point
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
    ElseIf DcodN = 3 Then
    ' flash
    
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        pctMain.DrawWidth = 1 ' trackSize

        If apertureType = "R" Then  ' rettangolo
            s = f1_testX(Xcurr, apertureN / 2)
            s = f1_testY(Ycurr, apertureM / 2)
        ElseIf apertureType = "P" Then ' poligono
            s = f1_testX(Xcurr, apertureN / 2)
            s = f1_testY(Ycurr, apertureN / 2)
        ElseIf apertureType = "C" Then ' cerchio
            s = f1_testX(Xcurr, apertureN / 2)
            s = f1_testY(Ycurr, apertureN / 2)
       ElseIf apertureType = "O" Then ' ovale
            s = f1_testX(Xcurr, apertureN / 2)
            s = f1_testY(Ycurr, apertureM / 2)
    End If

    ElseIf DcodN > 9 Then
        ' set current aperture
        
        apertureCodn = DcodN
        apertureType = tabApertureType(DcodN)
        
        apertureN = tabApertureN(DcodN)
        apertureM = tabApertureM(DcodN)
        
        trackSize = apertureN
        
    End If
    
    DcodN = 0
    linePtr = linePtr + 1
End Sub
Sub execCmd_HoleGTL_GBL()
Dim aspect, radius, width As Single
Dim i As Integer
Dim s As Single


    If DcodN = 1 Then
    ' draw
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        width = apertureN ' trackSize
    ElseIf DcodN = 2 Then
    ' set point
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
    ElseIf DcodN = 3 Then
    ' flash
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        pctMain.DrawWidth = 1 ' trackSize

        If apertureType = "R" Then ' rettangolo
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, Xhole)
        ElseIf apertureType = "P" Then ' poligono
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, Xhole)
        ElseIf apertureType = "C" Then ' cerchio
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, Xhole)
        ElseIf apertureType = "O" Then ' ovale
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, Xhole)
        Else   'macro ?
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, Xhole)
        End If

    ElseIf DcodN > 9 Then
        ' set current aperture
        
        apertureCodn = DcodN
        apertureType = tabApertureType(DcodN)
        
        apertureN = tabApertureN(DcodN)
        apertureM = tabApertureM(DcodN)
        
        trackSize = apertureN
        
    End If
    
    DcodN = 0
    linePtr = linePtr + 1
End Sub
Sub execCmd_HoleGDG()
Dim aspect, radius, width As Single
Dim i As Integer
Dim s As Single


    If DcodN = 1 Then
    ' draw
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        width = apertureN ' trackSize
    ElseIf DcodN = 2 Then
    ' set point
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
    ElseIf DcodN = 3 Then
    ' flash
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        pctMain.DrawWidth = 1 ' trackSize

        If apertureType = "C" Then ' cerchio
                If Xhole > 0 Then apertureN = Xhole
                s = DrawHole(Xcurr - Xoffset, Ycurr - Yoffset, apertureN)
        End If

    ElseIf DcodN > 9 Then
        ' set current aperture
        
        apertureCodn = DcodN
        apertureType = tabApertureType(DcodN)
        
        apertureN = tabApertureN(DcodN)
        apertureM = tabApertureM(DcodN)
        
        trackSize = apertureN
        
    End If
    
    DcodN = 0
    linePtr = linePtr + 1
End Sub
Sub execCmd_Draw()
Dim aspect, radius, width As Single
Dim i As Integer
Dim s As Single
Dim xini As Integer
Dim xend As Integer
Dim yini As Integer
Dim yend As Integer

    If GcodN = 36 And (DcodN = 1 Or DcodN = 2) Then
'        If DcodN = 1 Then
'        ElseIf DcodN = 2 Then
'        ElseIf DcodN = 3 Then
'        End If

        If swapXY.Value = 1 Then
            Points(contour).x = Ytemp - Xoffset
            Points(contour).Y = Xtemp - Yoffset
        Else
            Points(contour).x = Xtemp - Xoffset
            Points(contour).Y = Ytemp - Yoffset
        End If
        If flipX.Value = 1 Then
            Points(contour).x = Xpoints - Points(contour).x
        End If
        Points(contour).x = Points(contour).x * Xscale
        Points(contour).Y = Points(contour).Y * Yscale
        contour = contour + 1
        
    ElseIf GcodN = 37 Then
    
        pctMain.FillStyle = 0
        pctMain.ForeColor = vbBlue
        pctMain.FillColor = vbBlue
        pctMain.DrawWidth = 1
        s = Polygon(pctMain.hdc, Points(0), contour)
    
        GcodN = 0
        contour = 0
        
    ElseIf DcodN = 1 Then
    ' draw
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        width = apertureN ' trackSize
 '       If width > 1 Then
 '           pctMain.DrawWidth = apertureN * Xscale ' trackSize
 '        Else
 '           pctMain.DrawWidth = 1 ' trackSize
 '       End If
 
        If apertureType = "R" Then  ' rettangolo
            pctMain.DrawWidth = 1 ' trackSize
            If polarity = "C" Then
'                pctMain.Line (Xcurr - Xoffset - apertureN / 2, Ycurr - Yoffset - apertureM / 2)-(Xcurr - Xoffset + apertureN / 2, Ycurr - Yoffset + apertureM / 2), vbBlack, BF
            Else
                s = DrawRectangle(Xcurr - Xoffset - apertureN / 2, Ycurr - Yoffset - apertureM / 2, Xcurr - Xoffset + apertureN / 2, Ycurr - Yoffset + apertureM / 2, vbRed)
                s = DrawRectangle(Xprev - Xoffset - apertureN / 2, Yprev - Yoffset - apertureM / 2, Xprev - Xoffset + apertureN / 2, Yprev - Yoffset + apertureM / 2, vbRed)
                If (Xcurr - Xprev) > apertureN Then
                    xini = Xprev + apertureN / 2
                    xend = Xcurr - apertureN / 2
                Else
                    xini = Xprev
                    xend = Xcurr
                End If
                
                If (Xprev - Xcurr) > apertureN Then
                    xini = Xcurr + apertureN / 2
                    xend = Xprev - apertureN / 2
                End If
                
                If (Ycurr - Yprev) > apertureN Then
                    yini = Yprev + apertureN / 2
                    yend = Ycurr - apertureN / 2
                Else
                    yini = Yprev
                    yend = Ycurr
                End If
                
                If (Yprev - Ycurr) > apertureN Then
                    yini = Ycurr + apertureN / 2
                    yend = Yprev - apertureN / 2
                End If
'                s = DrawLine(Xprev - Xoffset, Yprev - Yoffset, Xcurr - Xoffset, Ycurr - Yoffset, apertureN)
                 s = DrawLine(xini - Xoffset, yini - Yoffset, xend - Xoffset, yend - Yoffset, apertureN)
            End If
        Else
            If polarity = "C" Then
    '            pctMain.Line (Xprev - Xoffset, Yprev - Yoffset)-(Xcurr - Xoffset, Ycurr - Yoffset), vbBlack
            Else
                If Xprev = Xcurr And Yprev = Ycurr Then
                Else
                    s = DrawLine(Xprev - Xoffset, Yprev - Yoffset, Xcurr - Xoffset, Ycurr - Yoffset, apertureN)
                End If
            End If
        End If
        
    ElseIf DcodN = 2 Then
    ' set point
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
    ElseIf DcodN = 3 Then
    ' flash
    
        Xprev = Xcurr
        Yprev = Ycurr
        Xcurr = Xtemp
        Ycurr = Ytemp
        pctMain.DrawWidth = 1 ' trackSize

        If apertureType = "R" Then  ' rettangolo
            pctMain.DrawWidth = 1 ' trackSize
            If polarity = "C" Then
'                pctMain.Line (Xcurr - Xoffset - apertureN / 2, Ycurr - Yoffset - apertureM / 2)-(Xcurr - Xoffset + apertureN / 2, Ycurr - Yoffset + apertureM / 2), vbBlack, BF
            Else
                s = DrawRectangle(Xcurr - Xoffset - apertureN / 2, Ycurr - Yoffset - apertureM / 2, Xcurr - Xoffset + apertureN / 2, Ycurr - Yoffset + apertureM / 2, vbBlue)
            End If
        ElseIf apertureType = "P" Then ' poligono
'            lato = apotema / numero fisso (1,207 per ottagono)

             s = DrawPolygon(Xcurr - Xoffset, Ycurr - Yoffset, (apertureN / 2) / 1.207, apertureM, 22)
        
        ElseIf apertureType = "C" Then ' cerchio
            If polarity = "C" Then
                pctMain.FillColor = vbBlack
                pctMain.FillStyle = 0
'               pctMain.Circle (Xcurr - Xoffset, Ycurr - Yoffset), apertureN / 2 - 1, vbBlack
            Else
                pctMain.FillColor = vbRed
                pctMain.FillStyle = 0
                s = DrawCircle(Xcurr - Xoffset, Ycurr - Yoffset, apertureN)
            End If
   
       ElseIf apertureType = "O" Then ' ovale
            aspect = apertureM / apertureN
            If aspect <= 1 Then
               radius = apertureN / 2
            Else
               radius = apertureM / 2
            End If
           If polarity = "C" Then
                pctMain.FillColor = vbBlack
               pctMain.FillStyle = 0
'               pctMain.Circle (Xcurr - Xoffset, Ycurr - Yoffset), radius, vbBlack, , , aspect
           Else
               pctMain.FillColor = vbRed
               pctMain.FillStyle = 0
               s = DrawEllipse(Xcurr - Xoffset, Ycurr - Yoffset, apertureN, apertureM)
           End If

       Else
            ' is a macro !
            macroInd = 1
            While macroInd <= macroMax
                If macroName(macroInd) = apertureType Then
                    If macroPrimitive(macroInd) = 5 Then
                                ' poligono regolare
                        apertureM = macroParam(macroInd, 2)
                        If InStr(macroParam(macroInd, 5), "$1") > 0 Then
                            s = DrawPolygon(Xcurr - Xoffset, Ycurr - Yoffset, (apertureN / 2) / 1.207, apertureM, 22.5)
'                           s = DrawPolygon(Xcurr - Xoffset, Ycurr - Yoffset, (apertureN / 2) / 1.207, apertureM, macroParam(macroInd, 7))
                        End If
                    Else
                        LogPrint ("unknown primitive code: " + CStr(macroPrimitive(macroInd)))
                    End If
                    macroInd = macroMax
                End If
                macroInd = macroInd + 1
            Wend
        End If

    ElseIf DcodN > 9 Then
    ' set current aperture
    
        apertureCodn = DcodN
        apertureType = tabApertureType(DcodN)
        
        apertureN = tabApertureN(DcodN)
        apertureM = tabApertureM(DcodN)
        
        trackSize = apertureN
    
    End If
    
    DcodN = 0
    linePtr = linePtr + 1
End Sub
Function getParam()   ' % start

    linePtr = linePtr + 1
    Do
        lineParm = Mid(lineTxt, linePtr, 2)
        linePtr = linePtr + 2
        
        If (lineParm = "AD") Then
            getParamAperture
        ElseIf (lineParm = "FS") Then getParamFormat
        ElseIf (lineParm = "LP") Then getParamPolarity
        ElseIf (lineParm = "AM") Then getParamAmacro
        ElseIf (lineParm = "MO") Then getParamUnit
        Else:   skipParam
        End If
    Loop Until Mid(lineTxt, linePtr, 1) = "%"
    linePtr = linePtr + 1
 End Function
Function skipParam()
    If (Mid(lineTxt, linePtr, 1) = "%") Then Exit Function
    While (Mid(lineTxt, linePtr, 1) <> "*")
           linePtr = linePtr + 1
           If linePtr > lineLen Then
                ReadLine
           End If
    Wend
    linePtr = linePtr + 1
'    If (Mid(lineTxt, linePtr, 1) = "%") Then linePtr = linePtr + 1
End Function
Function skipParamOld()  ' VECCHIA VERSIONE
    While (Mid(lineTxt, linePtr, 1) <> "%")
           linePtr = linePtr + 1
           If linePtr > lineLen Then
                ReadLine
           End If
    Wend
    linePtr = linePtr + 1
'    If linePtr > lineLen Then
'         Line Input #1, lineTxt
'         lineLen = Len(lineTxt)
'         linePtr = 1
'    End If
End Function
Function getStringValue()
    getStringValue = ""
    If linePtr > lineLen Then
          ReadLine
    End If
'    While (Mid(lineTxt, linePtr, 1) <> ",") And (Mid(lineTxt, linePtr, 1) <> "*" And (Mid(lineTxt, linePtr, 1) <> "%") And (Mid(lineTxt, linePtr, 1) <> "X"))
    While (Mid(lineTxt, linePtr, 1) <> ",") And (Mid(lineTxt, linePtr, 1) <> "*" And (Mid(lineTxt, linePtr, 1) <> "%"))
           getStringValue = getStringValue + Mid(lineTxt, linePtr, 1)
           linePtr = linePtr + 1
           If linePtr > lineLen Then
                ReadLine
           End If
    Wend
    If (Mid(lineTxt, linePtr, 1) <> "%") Then
        linePtr = linePtr + 1
    End If
    
'    If linePtr > lineLen Then
'         Line Input #1, lineTxt
'         lineLen = Len(lineTxt)
'         linePtr = 1
'    End If
End Function
Function getMilsValue()
Dim MilsInt As String
Dim MilsDec As String
Dim pd As Byte
    pd = 0
    getMilsValue = 0
    MilsInt = ""
    MilsDec = ""
    If linePtr > lineLen Then
          ReadLine
    End If
    If Mid(lineTxt, linePtr, 1) = " " Then Mid(lineTxt, linePtr, 1) = "0"
    While (((Mid(lineTxt, linePtr, 1) >= "0") And (Mid(lineTxt, linePtr, 1) <= "9")) Or (Mid(lineTxt, linePtr, 1) = "."))
           If pd = 0 Then
               MilsInt = MilsInt + Mid(lineTxt, linePtr, 1)
           Else
               MilsDec = MilsDec + Mid(lineTxt, linePtr, 1)
           End If
           If Mid(lineTxt, linePtr, 1) = "." Then pd = 1
           linePtr = linePtr + 1
           If linePtr > lineLen Then
                ReadLine
           End If
    Wend
    getMilsValue = Val(MilsInt) * 1000
    While Len(MilsDec) < 3
          MilsDec = MilsDec + "0"
    Wend
    If Len(MilsDec) > 3 Then
          MilsDec = Left(MilsDec, 3)
    End If
    getMilsValue = getMilsValue + Val(MilsDec)
    If (units = "MM") Then getMilsValue = getMilsValue / 25.4     ' da MM a INCHES *******************************************************************

End Function
Private Sub ReadLine()
    Do
        Line Input #dFile, lineTxt
        lineLen = Len(lineTxt)
        linePtr = 1
    Loop Until EOF(dFile) Or Left(lineTxt, 2) <> "0 "
End Sub

Function skipValue() As Integer
Dim i, j As Integer
    i = linePtr
    j = 0
    While (Mid(lineTxt, i, 1) >= "0") And (Mid(lineTxt, i, 1) <= "9")
           i = i + 1
           j = j + 1
    Wend
    skipValue = j
End Function
Function getParamFormat()
Dim i, j As Integer
    linePtr = linePtr + 2
    If Mid(lineTxt, linePtr, 1) = "X" Then
        linePtr = linePtr + 2
        i = Val(Mid(lineTxt, linePtr, 1))
        If i = 4 Then
            XgerberScale = 10 ' perche si lavora in mils (i=3)
        Else
        If i = 5 Then
            XgerberScale = 100 ' perche si lavora in mils (i=3)
        Else
        If i = 6 Then
            XgerberScale = 1000 ' perche si lavora in mils (i=3)
        Else
            XgerberScale = 1
        End If
        End If
        End If
        linePtr = linePtr + 1
    End If
    If Mid(lineTxt, linePtr, 1) = "Y" Then
        linePtr = linePtr + 2
        i = Val(Mid(lineTxt, linePtr, 1))
        If i = 4 Then
            YgerberScale = 10 ' perche si lavora in mils (i=3)
        Else
        If i = 5 Then
            YgerberScale = 100 ' perche si lavora in mils (i=3)
        Else
        If i = 6 Then
            YgerberScale = 1000 ' perche si lavora in mils (i=3)
        Else
            YgerberScale = 1
        End If
        End If
        End If
        linePtr = linePtr + 1
    End If

    skipParam
End Function
Function getParamPolarity()
    polarity = Mid(lineTxt, linePtr, 1)
    skipParam
End Function
Function getParamUnit()
    units = Mid(lineTxt, linePtr, 2)
    skipParam
End Function
Function findValue(position) As Integer
Dim i, j As Integer
    i = position
    j = 0
    While (Mid(lineTxt, i, 1) >= "0") And (Mid(lineTxt, i, 1) <= "9")
           i = i + 1
           j = j + 1
    Wend
    findValue = j
End Function
Function getParamAperture()
    Dim l As Integer
    l = findValue(linePtr + 1)
    apertureCode = Mid(lineTxt, linePtr, l + 1)
    apertureCodn = Val(Mid(lineTxt, linePtr + 1, l))
    linePtr = linePtr + l + 1
    
    apertureType = getStringValue
    
    apertureN = getMilsValue                                       '================== NEW =====================
    
    If Mid(lineTxt, linePtr, 1) <> "X" Then linePtr = linePtr + 1
    If apertureType = "R" And Mid(lineTxt, linePtr, 1) = "X" Then
         linePtr = linePtr + 1
         apertureM = getMilsValue                                   '================== NEW =====================
    ElseIf apertureType = "O" And Mid(lineTxt, linePtr, 1) = "X" Then
         linePtr = linePtr + 1
         apertureM = getMilsValue                                   '================== NEW =====================
    ElseIf apertureType = "P" And Mid(lineTxt, linePtr, 1) = "X" Then
        apertureM = CInt(Mid(lineTxt, linePtr + 1, 1))
    Else
        apertureM = 0
    End If
    tabApertureType(apertureCodn) = apertureType
    tabApertureN(apertureCodn) = apertureN
    tabApertureM(apertureCodn) = apertureM
    
    skipParam
End Function
Function getParamAmacro()
Dim mParm As Integer
    macroMax = macroMax + 1
    macroName(macroMax) = getStringValue
    
    If Mid(lineTxt, linePtr, 1) = "0" Then getStringValue
    macroPrimitive(macroMax) = Val(getStringValue)
    mParm = 0
    Do
        mParm = mParm + 1
        macroParam(macroMax, mParm) = getStringValue
'        If InStr(macroParam(macroMax, mParm), "$") > 0 Then
'            mParm = mParm
'        End If
    Loop While macroParam(macroMax, mParm) <> ""
'    skipParam
End Function
Function getXcoord()
Dim ll As Integer
Dim neg As Byte
    neg = 0
    linePtr = linePtr + 1
    If Mid(lineTxt, linePtr, 1) = "-" Then
        neg = 1
        linePtr = linePtr + 1
    End If

    ll = skipValue
    Xtemp = Val(Mid(lineTxt, linePtr, ll))
    
    Xtemp = Xtemp / XgerberScale
    If (units = "MM") Then Xtemp = Xtemp / 25.4     ' da MM a INCHES *******************************************************************
    If neg = 1 Then Xtemp = Xtemp * -1
    
    linePtr = linePtr + ll
End Function
Function getYcoord()
Dim ll As Integer
Dim neg As Byte
    neg = 0
    linePtr = linePtr + 1
    If Mid(lineTxt, linePtr, 1) = "-" Then
        neg = 1
        linePtr = linePtr + 1
    End If
    
    ll = skipValue
    Ytemp = Val(Mid(lineTxt, linePtr, ll))
    
    Ytemp = Ytemp / YgerberScale
    If (units = "MM") Then Ytemp = Ytemp / 25.4     ' da MM a INCHES *******************************************************************
    If neg = 1 Then Ytemp = Ytemp * -1
    
    linePtr = linePtr + ll
End Function
Function commandD()
    Dim l As Integer
    l = findValue(linePtr + 1)
    Dcode = Mid(lineTxt, linePtr, l + 1)
    DcodN = Val(Mid(lineTxt, linePtr + 1, l))
    linePtr = linePtr + l + 1
'    Dcode = Mid(lineTxt, linePtr, 3)
'    DcodN = Val(Mid(lineTxt, linePtr + 1, 2))
'    linePtr = linePtr + 3
End Function
Function commandM()
    Dim l As Integer
    l = findValue(linePtr + 1)
    Mcode = Mid(lineTxt, linePtr, l + 1)
    McodN = Val(Mid(lineTxt, linePtr + 1, l))
    linePtr = linePtr + l + 1
'    Mcode = Mid(lineTxt, linePtr, 3)
'    McodN = Val(Mid(lineTxt, linePtr + 1, 2))
'    linePtr = linePtr + 3
End Function
Function commandG()
 '   Gcode = Mid(lineTxt, linePtr, 3)
 '   GcodN = Val(Mid(lineTxt, linePtr + 1, 2))
 '   linePtr = linePtr + 3
    Dim l As Integer
    l = findValue(linePtr + 1)
    Gcode = Mid(lineTxt, linePtr, l + 1)
    GcodN = Val(Mid(lineTxt, linePtr + 1, l))
    If GcodN = 36 Then
        contour = 0
    End If
    linePtr = linePtr + l + 1
End Function
Function unknownCmd()
    linePtr = lineLen + 1
End Function
Private Sub Swap(ByRef x As Single, ByRef Y As Single)
   Dim temp As Single
   temp = x
   x = Y
   Y = temp
End Sub

Function DrawLine(ByVal x1 As Single, ByVal y1 As Single, _
    ByVal x2 As Single, ByVal y2 As Single, ByVal width As Single) ''', Optional Xreverse As Integer = 0)
' Draw a line
' CTRL can be either a form or a PictureBox control.

' disegna sullo schermo una linea - coordinate e dimensioni assolute in pixel

    Dim RetVal As Long
    
    If swapXY.Value = 1 Then
        Call Swap(x1, y1)
        Call Swap(x2, y2)
    Else
    End If
    
    width = width - smooth
    If Abs(x1 - x2) >= Abs(y1 - y2) Then
        width = width * Yscale
    Else
        width = width * Xscale
    End If
    If width < 1 Then width = 1
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbRed
    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = Int(width + 0.5)
    
    x1 = x1 * Xscale
    x2 = x2 * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
    Else
    End If
        
    RetVal = MoveToEx(pctMain.hdc, x1, y1 * Yscale, CurrentPoint)
    RetVal = LineTo(pctMain.hdc, x2, y2 * Yscale)
End Function

Function DrawRectangle(ByVal x1 As Single, ByVal y1 As Single, _
    ByVal x2 As Single, ByVal y2 As Single, ByVal colore)
' Draw a line
' CTRL can be either a form or a PictureBox control.

' disegna sullo schermo un rettangolo - coordinate e dimensioni assolute in pixel

    Dim RetVal As Long
    
    If swapXY.Value = 1 Then
        Call Swap(x1, y1)
        Call Swap(x2, y2)
    Else
    End If
    
    If x2 > x1 Then
        x1 = x1 + smooth / 2
        x2 = x2 - smooth / 2
    Else
        x1 = x1 - smooth / 2
        x2 = x2 + smooth / 2
    End If
    If y2 > y1 Then
        y1 = y1 + smooth / 2
        y2 = y2 - smooth / 2
    Else
        y1 = y1 - smooth / 2
        y2 = y2 + smooth / 2
    End If
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlue
    pctMain.FillColor = vbBlue
    pctMain.ForeColor = colore
    pctMain.FillColor = colore
    pctMain.DrawWidth = 1
        
    x1 = x1 * Xscale
    x2 = x2 * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
    Else
    End If
    
    RetVal = Rectangle(pctMain.hdc, x1, y1 * Yscale, x2, y2 * Yscale)
End Function
Function DrawCircle(ByVal xc As Single, ByVal yc As Single, _
    ByVal diam As Single)
' Draw a circle, given its center and diameter
' CTRL can be either a form or a PictureBox control.

' disegna sullo schermo un cerchio - coordinate e dimensioni assolute in pixel

    Dim RetVal As Long
    
    Dim r As Single
    Dim x1 As Single
    Dim x2 As Single
    
    If swapXY.Value = 1 Then
        Call Swap(xc, yc)
    Else
    End If

    diam = diam - smooth
    r = diam / 2
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlue
    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
           
    x1 = (xc - r) * Xscale
    x2 = (xc + r) * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
    Else
    End If
    
    RetVal = Ellipse(pctMain.hdc, x1, (yc - r) * Yscale, x2, (yc + r) * Yscale)
'X1 = Centre.X - radius        'Y1 = Centre.Y - radius
'X2 = Centre.X + radius        'X3 = Centre.Y + radius
End Function
Function DrawEllipse(ByVal xc As Single, ByVal yc As Single, _
    ByVal diamX As Single, ByVal diamY As Single)
    
    Dim RetVal As Long
    Dim r As Single
    Dim i As Integer
    
    If swapXY.Value = 1 Then
        Call Swap(xc, yc)
        Call Swap(diamX, diamY)
    Else
    End If
    
    diamX = diamX - smooth
    diamY = diamY - smooth
    If diamX > diamY Then
        r = diamY / 4
    Else
        r = diamX / 4
    End If
' rettangolo con i bordi stondati - coordinate e dimensioni assolute in pixel

        Points(0).x = (xc - diamX / 2) * Xscale
        Points(1).x = (xc - diamX / 2 + r) * Xscale
        Points(2).x = (xc + diamX / 2 - r) * Xscale
        Points(3).x = (xc + diamX / 2) * Xscale
        Points(4).x = (xc + diamX / 2) * Xscale
        Points(5).x = (xc + diamX / 2 - r) * Xscale
        Points(6).x = (xc - diamX / 2 + r) * Xscale
        Points(7).x = (xc - diamX / 2) * Xscale
    
        If flipX.Value = 1 Then
            For i = 0 To 7
                Points(i).x = Xpoints - Points(i).x
            Next
        Else
        End If

        Points(0).Y = (yc - diamY / 2 + r) * Yscale
        Points(1).Y = (yc - diamY / 2) * Yscale
        Points(2).Y = (yc - diamY / 2) * Yscale
        Points(3).Y = (yc - diamY / 2 + r) * Yscale
        Points(4).Y = (yc + diamY / 2 - r) * Yscale
        Points(5).Y = (yc + diamY / 2) * Yscale
        Points(6).Y = (yc + diamY / 2) * Yscale
        Points(7).Y = (yc + diamY / 2 - r) * Yscale

    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlue
    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
    RetVal = Polygon(pctMain.hdc, Points(0), 8)
End Function

Function DrawEllipseTrue(ByVal xc As Single, ByVal yc As Single, _
    ByVal diamX As Single, ByVal diamY As Single)
' Draw a circle, given its center and diameter
' CTRL can be either a form or a PictureBox control.

' disegna sullo schermo un ellisse - coordinate e dimensioni assolute in pixel

    Dim RetVal As Long
    
    Dim rx As Single
    Dim ry As Single
    Dim x2 As Single
    Dim x1 As Single
    
    If swapXY.Value = 1 Then
        Call Swap(xc, yc)
        Call Swap(diamX, diamY)
    Else
    End If
    
    diamX = diamX - smooth
    diamY = diamY - smooth
    
    rx = diamX / 2
    ry = diamY / 2
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlue
    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
           
    x1 = (xc - rx) * Xscale
    x2 = (xc + rx) * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
    Else
    End If
        
    RetVal = Ellipse(pctMain.hdc, x1, (yc - ry) * Yscale, x2, (yc + ry) * Yscale)
'X1 = Centre.X - radius           'Y1 = Centre.Y - radius
'X2 = Centre.X + radius           'X3 = Centre.Y + radius
End Function

Function DrawPolygon(ByVal xc As Single, ByVal yc As Single, _
    ByVal side As Single, ByVal numSides As Integer, _
    Optional ByVal angle As Single)
' Draw a polygon, given its center and number of sides
' CTRL can be either a form or a PictureBox control.
' ANGLE is an optional angle in degrees

' disegna sullo schermo un poligono - coordinate e dimensioni assolute in pixel

    Dim deltaAngle As Single
    Dim RetVal As Long
    
    Dim r As Single
    Dim i As Integer
        
    If swapXY.Value = 1 Then
        Call Swap(xc, yc)
    Else
    End If
    
    side = side - smooth / 2
    ' evaluate the angle formed by each side (in radians)
    deltaAngle = 3.14159265358979 * 2 / numSides
    ' evaluate the radius of the circle that wraps the polygon.
    r = side / Sin(deltaAngle)
    ' convert the initial angle in radians
    angle = angle / 57.2957795130824
    
    ' Draw individual sides
    For i = 1 To numSides
        Points(i - 1).x = (xc + r * Sin(angle)) * Xscale
        Points(i - 1).Y = (yc + r * Cos(angle)) * Yscale
        angle = angle + deltaAngle
    Next
    
    If flipX.Value = 1 Then
        For i = 1 To numSides
            Points(i - 1).x = Xpoints - Points(i - 1).x
        Next
    Else
    End If
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlue
    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
    RetVal = Polygon(pctMain.hdc, Points(0), numSides)
End Function
Function DrawHoleOld(ByVal x As Single, ByVal Y As Single, ByVal diam As Integer)
Dim i, r As Integer
Dim x1 As Single
Dim x2 As Single
Dim xCenter As Single
Dim yCenter As Single

Dim RetVal As Long
    
    If swapXY.Value = 1 Then
        Call Swap(x, Y)
    Else
    End If
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlack
    pctMain.FillColor = vbBlack
    pctMain.ForeColor = vbYellow
'    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
    r = diam / 2
           
    ' la coordinata del centro foro è X*Xscale, Y*Yscale
    
    xCenter = x * Xscale
    yCenter = Y * Yscale
    
    x1 = (x - r) * Xscale
    x2 = (x + r) * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
        xCenter = Xpoints - xCenter
    Else
    End If
    
    RetVal = Ellipse(pctMain.hdc, x1, (Y - r) * Yscale, x2, (Y + r) * Yscale)
    
    holeW = holeW + 1
    tabHoles(holeW).x = xCenter
    tabHoles(holeW).Y = yCenter

End Function
Function DrawHole(ByVal x As Single, ByVal Y As Single, ByVal diam As Integer)
Dim i, r As Integer
Dim x1 As Single
Dim x2 As Single
Dim xCenter As Single
Dim yCenter As Single

Dim RetVal As Long
    
    If swapXY.Value = 1 Then
        Call Swap(x, Y)
    Else
    End If
    
    pctMain.FillStyle = 0
    pctMain.ForeColor = vbBlack
    pctMain.FillColor = vbBlack
    pctMain.ForeColor = vbYellow
'    pctMain.FillColor = vbBlue
    pctMain.DrawWidth = 1
    r = diam / 2
           
    ' la coordinata del centro foro è X*Xscale, Y*Yscale
    
    xCenter = x '* Xscale
    yCenter = Y '* Yscale
    
    x1 = (x - r) * Xscale
    x2 = (x + r) * Xscale
    If flipX.Value = 1 Then
        x1 = Xpoints - x1
        x2 = Xpoints - x2
        xCenter = Xpoints - xCenter '''''''''''''''''''''''' E R R A T O ''''''''''''''''''''''''''''''''
    Else
    End If
    
    RetVal = Ellipse(pctMain.hdc, x1, (Y - r) * Yscale, x2, (Y + r) * Yscale)
    
    holeW = holeW + 1
    tabHoles(holeW).x = xCenter
    tabHoles(holeW).Y = yCenter

End Function
Public Sub ShowData(The_Data As String, the_option As String)

If the_option = "" Then the_option = Chr(13) + Chr(10)
If the_option = "@" Then the_option = ""
inPtr = Len(Rich.Text)
'   Rich.Text = Rich.Text + Chr(10) + The_Data
Rich.Text = Rich.Text + the_option + The_Data

'Print #dFileLog, the_option + The_Data;

inPtr = Len(Rich.Text) - 1
Rich.SelStart = inPtr
End Sub
Function LogPrint(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.Text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = The_Data & vbCrLf
        .SelColor = vbBlack
    End With
    Rich.Refresh
End Function
Function LogPrintCont(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.Text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = The_Data
        .SelColor = vbBlack
    End With
    Rich.Refresh
End Function
Public Sub LogPrintContSer(The_Data As String)
    Rich.Text = Rich.Text + The_Data
'    Print #dFileLog, The_Data;

    inPtr = Len(Rich.Text)
    Rich.SelStart = inPtr
    Rich.Refresh
End Sub
Private Sub xfactor_Change()
    Xscale = Val(xfactor.Text)
    SaveSetting settingName, "InitValues", "Xscale", xfactor.Text
End Sub

Private Sub xref_Click()
  If Val(Xleft.Text) < 40 Then Xleft.Text = 40
  If Val(leading.Text) < 20 Then leading.Text = 20
End Sub

Private Sub ClearSaved()
    SaveSetting settingName, "InitValues", "Xscale", ""
    SaveSetting settingName, "InitValues", "Yscale", ""
    SaveSetting settingName, "InitValues", "SpeedFactor", ""
End Sub

Private Sub yfactor_Change()
    Yscale = Val(yfactor.Text)
    SaveSetting settingName, "InitValues", "Yscale", yfactor.Text
End Sub
Private Sub speedFactor_Change()
     Xspeed = Int(Val(speedFactor.Text))
     If Xspeed > 95 Then Xspeed = 95
     If Xspeed > 0 Then speedFactor.Text = CStr(Xspeed)
     SaveSetting settingName, "InitValues", "SpeedFactor", speedFactor.Text
End Sub
Private Sub smooth_Change()
     smtpoint = Val(smooth.Text)
End Sub
