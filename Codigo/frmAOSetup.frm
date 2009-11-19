VERSION 5.00
Begin VB.Form frmAOSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum Online Setup"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmAOSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Noticias del clan"
      Height          =   735
      Left            =   120
      TabIndex        =   33
      Top             =   5760
      Width           =   6615
      Begin VB.OptionButton optNoMostrar 
         Caption         =   "No mostrarlas"
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   315
         Width           =   1335
      End
      Begin VB.OptionButton optMostrarNoticias 
         Caption         =   "Mostrar noticias al conectarse"
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   315
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo de Arboles"
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   5040
      Width           =   6615
      Begin VB.OptionButton optBig 
         Caption         =   "Grandes"
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptAverage 
         Caption         =   "Medianos"
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSmall 
         Caption         =   "Pequeños"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pruebas de DirectX"
      Height          =   3270
      Left            =   2640
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
      Begin VB.PictureBox DirectDrawTest 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1310
         Left            =   2520
         ScaleHeight     =   1245
         ScaleWidth      =   1395
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         Begin VB.Timer Timer1 
            Left            =   120
            Top             =   600
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "DirectX 7"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "DirectDraw"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "DirectSound"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox fondoVersion 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   3795
         TabIndex        =   8
         Top             =   2715
         Width           =   3855
         Begin VB.Label lVersionFondo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Versión detectada:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   45
            Width           =   1335
         End
         Begin VB.Label lDirectX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1500
            TabIndex        =   9
            Top             =   45
            Width           =   135
         End
      End
      Begin AOSetup.chameleonButton bProbarSonido 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "S&onido"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAOSetup.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin AOSetup.chameleonButton bProbarVideo 
         Height          =   495
         Left            =   2160
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Video"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAOSetup.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame4 
         Caption         =   "Probar"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblDX 
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblDD 
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblDS 
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   2280
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   2280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   2280
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.CheckBox cEjecutar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Ejecutar el juego al Aceptar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   7200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Sonido"
      Height          =   1110
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2415
      Begin VB.CheckBox chkEfectos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Efectos de sonido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   810
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMusica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Música Activada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   525
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkSonido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Sonido Activado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Video"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
      Begin VB.CheckBox chkPantallaCompleta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Pantalla Completa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkUserVideo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Usar M&emoria de Video"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDinamico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Carga &Dinámica"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin AOSetup.PBarY pMemoria 
         CausesValidation=   0   'False
         Height          =   135
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   238
         Value           =   40
         Min             =   4
         Max             =   40
         BackColor       =   0
         FillColor       =   8421631
         BorderColor     =   16777215
         BorderStyle     =   3
         EnabledSlider   =   0   'False
         MousePointer    =   0
         picForeColor    =   12632256
         picFillColor    =   8421504
         Style           =   1
      End
      Begin VB.Label lCuantoVideo 
         Alignment       =   2  'Center
         Caption         =   "Usar 16 Mb de Memoria"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
   End
   Begin AOSetup.chameleonButton bCancelar 
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton bAceptar 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632319
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421631
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":0496
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cLibrerias 
      Height          =   375
      Left            =   105
      TabIndex        =   21
      Top             =   6540
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Verificar &Librerias"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":04B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cCreditos 
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      ToolTipText     =   "Creditos"
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "?"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":04CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6720
      Y1              =   6975
      Y2              =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6720
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   120
      Picture         =   "frmAOSetup.frx":04EA
      Top             =   120
      Width           =   6675
   End
End
Attribute VB_Name = "frmAOSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

' sonido
Dim m_dsBuffer As DirectSoundBuffer
Dim m_bLoaded As Boolean
' video
Private Const SW_SHOWNORMAL = 1
Dim Primary As DirectDrawSurface7
Dim BackBuffer As DirectDrawSurface7
Dim Clipper As DirectDrawClipper
Dim ddsCharacter As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2
Dim ddsdback As DDSURFACEDESC2
Dim destRect As RECT
Dim srcRect As RECT
Dim chanRect As RECT
Dim CharWidth As Integer
Dim CharHight As Integer
Dim PostionX As Integer
Dim postionY As Integer
Dim running As Boolean

Private Sub bAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 11/19/09
'11/19/09: Pato - Now is optional show the frmGuildNews form in the client
'*************************************************
    Dim sFile As String
    
    setupMod.bNoSound = Not CBool(Me.chkSonido.Value)
    
    setupMod.bNoMusic = Not CBool(Me.chkMusica.Value)
    
    setupMod.bNoRes = Not CBool(Me.chkPantallaCompleta.Value) ' 24/06/2006 - ^[GS]^
    
    setupMod.bUseVideo = CBool(Me.chkUserVideo.Value)
    
    setupMod.bDinamic = Me.chkDinamico.Value
    
    setupMod.byMemory = CByte(Me.pMemoria.Value)
    
    setupMod.bNoSoundEffects = Not CBool(Me.chkEfectos.Value)
    
    If optBig.Value Then
        sFile = "Graficos3.ind"
    ElseIf OptAverage.Value Then
        sFile = "Graficos2.ind"
    Else
        sFile = "Graficos1.ind"
    End If
    
    setupMod.sGraficos = sFile
    
    setupMod.bGuildNews = Not setupMod.bGuildNews
    
    DoEvents
    
    Dim handle As Integer
    handle = FreeFile
    Open App.Path & "\Init\AO.DAT" For Binary As handle
        Put handle, , setupMod
    Close handle
    DoEvents
    
    If cEjecutar.Value = 1 Then
        If FileExist(App.Path & "\Argentum.exe", vbArchive) = True Then _
            Call Shell(App.Path & "\Argentum.exe")
        DoEvents
    End If
    
    Unload Me
End Sub

Private Sub bCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    Unload Me
End Sub

Private Sub bProbarSonido_Click()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 24/06/06
'10/03/06: ^[GS]^ - Agregue una revision de la existencia del archivo de sonido
'24/06/06: ^[GS]^ - Una correccion mas mejorada.
'*************************************************
On Error Resume Next
    
    If bProbarSonido.Value = True Then
        ' [GS]
        Dim sonido As String
        sonido = App.Path & "\wav\18.wav"
        
        If FileExist(sonido, vbArchive) = False Then
            MsgBox "No se puede probar el sonido porque falta el archivo de pruebas.", vbCritical
            bProbarSonido.Value = False ' 24/06/06 - ^[GS]^
            Exit Sub
        End If
        ' [/GS]
        
        DirectSound.SetCooperativeLevel Me.hwnd, DSSCL_NORMAL
        
        If m_bLoaded = False Then
            m_bLoaded = True
            LoadWave 0, sonido
        End If
        Dim flag As Long
        flag = 0
        m_dsBuffer.Play flag
        
        If Err.Number <> 0 Then
            MsgBox "Problemas de DirectSound, Reinstale DIRECTX.", vbOKOnly, "Argentum Online Setup"
        End If
    Else
        If m_dsBuffer Is Nothing Then Exit Sub
        m_dsBuffer.Stop
        m_dsBuffer.SetCurrentPosition 0
    End If
End Sub

Sub LoadWave(i As Integer, sFile As String)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 10/03/06
'10/03/06: ^[GS]^ - Borre un codigo al final que no se utilizaba
'*************************************************

    Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
    Dim waveFormat As WAVEFORMATEX
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2    '2 channels
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set m_dsBuffer = DirectSound.CreateSoundBufferFromFile(sFile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        MsgBox "Error en " + sFile
        End
    End If
End Sub

Private Sub bProbarVideo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************

If bProbarVideo.Value = True Then
    DirectDrawTest.Visible = True
    Call DirectDrawTestStart
Else
    DirectDrawTest.Visible = False
    Timer1.Enabled = False
    running = False
End If
End Sub

Private Sub cCreditos_Click()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 15/03/06
'*************************************************
    frmAbout.Show vbModal, Me
End Sub

Private Sub chkDinamico_Click()
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 10/03/06
'*************************************************
    If chkDinamico.Value Then
        lCuantoVideo.ForeColor = vbBlack
        pMemoria.EnabledSlider = True
        pMemoria.picFillColor = &H8080FF
        pMemoria.picForeColor = &H80FF80
    Else
        lCuantoVideo.ForeColor = &H808080
        pMemoria.EnabledSlider = False
        pMemoria.picFillColor = &H808080
        pMemoria.picForeColor = &HC0C0C0
    End If
End Sub

Private Sub cLibrerias_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
frmLibrerias.Show
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
On Error Resume Next
    Me.Show
    
    DoEvents
    
    Call LeerSetup
    
    Call mod_DirectX.ProbarDirectX
    Call mod_DirectX.VersionDirectX
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last modified: 11/03/06
'*************************************************
    If FileExist("C:\DXTest.txt", vbArchive) Then _
        Kill "C:\DXTest.txt"
    
    End
End Sub

Private Sub optMostrarNoticias_Click()
    setupMod.bGuildNews = True
End Sub

Private Sub optNoMostrar_Click()
setupMod.bGuildNews = False
End Sub

Private Sub pMemoria_ChangeValue(NewValue As Long, OldValue As Long)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    lCuantoVideo.Caption = "Usar " & CStr(NewValue) & " Mb de Memoria"
End Sub

Private Sub Timer1_Timer()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
    DoEvents
    PostionX = PostionX + 1
    If PostionX > 3 Then PostionX = 1
    setCharacter PostionX, postionY
    Draw 50, 50
End Sub

Private Sub Draw(ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'10/03/06: Maraxus - Puse los parámetros ByVal
'*************************************************

    directx.GetWindowRect DirectDrawTest.hwnd, destRect
    
    ddsCharacter.GetSurfaceDesc ddsd
    
    chanRect.Left = destRect.Left
    chanRect.Top = destRect.Top
    chanRect.Right = chanRect.Left + 96
    chanRect.Bottom = chanRect.Top + 85
    
    Primary.Blt chanRect, ddsCharacter, srcRect, DDBLT_WAIT
End Sub

Private Sub setCharacter(ByVal charcol As Integer, ByVal charRow As Integer)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'10/03/06: Maraxus - Puse los parámetros ByVal
'*************************************************
  srcRect.Left = (charcol - 1) * CharWidth
  srcRect.Right = charcol * CharWidth
  srcRect.Top = (charRow - 1) * CharHight
  srcRect.Bottom = charRow * CharHight
End Sub

Private Sub DirectDrawTest_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
    If KeyCode = &H28 Then 'abajo
        PostionX = 1
        postionY = 3
    ElseIf KeyCode = &H26 Then 'arriba
        PostionX = 1
        postionY = 1
    ElseIf KeyCode = &H25 Then 'izquierda
        PostionX = 1
        postionY = 4
    ElseIf KeyCode = &H27 Then 'derecha
        PostionX = 1
        postionY = 2
    End If
End Sub

Public Sub DirectDrawTestStart()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 24/06/06
'24/06/06: ^[GS]^ - Realize una correccion que evita que se realize cualquier accion sin la imagen imagen.
'*************************************************
    
    ' [GS] 24/06/06 - ^[GS]^
    If FileExist(App.Path & "\Graficos\test.bmp", vbArchive) = False Then
        DirectDrawTest.Visible = False
        Timer1.Enabled = False
        running = False
        MsgBox "No se puede probar el video porque falta el archivo de pruebas.", vbCritical
        bProbarVideo.Value = False
        Exit Sub
    End If
    ' [/GS]
    
    If lblDD.ForeColor <> &H8000& Then
        DirectDrawTest.Visible = False
        Exit Sub
    End If

    Init
    
    running = True
  
    PostionX = 0
    postionY = 3
    
    Timer1.Interval = 150
    Timer1.Enabled = True
End Sub

Private Sub Init()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 24/06/06
'10/03/06: ^[GS]^ - Hice una correccion para que revise la existencia de la imágen
'14/03/06: Maraxus - Evité que intente crear otro primary buffer que tiraba un automation error,
'así como cualquier otro objeto (memory leak)
'24/06/06: ^[GS]^ - Desplace una linea de codigo a DirectDrawTestStart
'*************************************************


    ScaleMode = 2
    
    'Create DirectDraw object
    DirectDraw.SetCooperativeLevel DirectDrawTest.hwnd, DDSCL_NORMAL
    
    'Create primary surface
    If Primary Is Nothing Then
        ddsd.lFlags = DDSD_CAPS
        ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        Set Primary = DirectDraw.CreateSurface(ddsd)
    End If
    
    'Create and attach clipper
    If Clipper Is Nothing Then
        Set Clipper = DirectDraw.CreateClipper(0)
        Clipper.SetHWnd DirectDrawTest.hwnd
        Primary.SetClipper Clipper
    End If
    
    'Create char surface
    If ddsCharacter Is Nothing Then
        Dim ddsdChar As DDSURFACEDESC2
        ddsdChar.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        ddsdChar.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set ddsCharacter = DirectDraw.CreateSurfaceFromFile(App.Path & "\Graficos\test.bmp", ddsdChar)
    End If
    
    'Store char size
    ddsCharacter.GetSurfaceDesc ddsdChar
    CharWidth = ddsdChar.lWidth / 3
    CharHight = ddsdChar.lHeight / 4
End Sub
