VERSION 5.00
Begin VB.Form frmPrograma 
   BackColor       =   &H00000000&
   Caption         =   "Martin Virtual Rockola 1.0"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "frmDiscoVinilo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14895
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Lista_Virtual 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   13440
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox pxld 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   10335
      Left            =   0
      ScaleHeight     =   10335
      ScaleWidth      =   19215
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   12360
         Picture         =   "frmDiscoVinilo.frx":57E2
         ScaleHeight     =   1170
         ScaleWidth      =   1170
         TabIndex        =   20
         Top             =   3600
         Width           =   1170
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   2880
         ScaleHeight     =   6015
         ScaleWidth      =   7455
         TabIndex        =   16
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "GUARDAR DISCO...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   10560
         TabIndex        =   19
         Top             =   2880
         Width           =   4830
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PANEL DE VIDEO PRINCIPAL ROKOLA 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   14250
      End
   End
   Begin VB.CommandButton versoloDisco 
      Caption         =   "VIDEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   19320
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar Scroll_Vertical_X 
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   7080
      Width           =   18615
   End
   Begin VB.PictureBox imagen1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   3240
      ScaleHeight     =   5655
      ScaleWidth      =   5640
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   5640
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6495
      ScaleWidth      =   19095
      TabIndex        =   7
      Top             =   600
      Width           =   19095
      Begin VB.PictureBox Panel 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Left            =   0
         ScaleHeight     =   6495
         ScaleWidth      =   5775
         TabIndex        =   8
         Top             =   360
         Width           =   5775
         Begin VB.PictureBox ptx 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   5775
            Index           =   0
            Left            =   0
            ScaleHeight     =   5775
            ScaleWidth      =   5775
            TabIndex        =   11
            Top             =   120
            Width           =   5775
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CANTANTE: DESCONOCIDO..."
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   10200
         TabIndex        =   13
         Top             =   0
         Width           =   6255
      End
      Begin VB.Label labCanciones 
         BackStyle       =   0  'Transparent
         Caption         =   "Panel: de Canciónes"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   10095
      End
   End
   Begin VB.CommandButton cmdDetener 
      Caption         =   "Detener"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer temporizador 
      Interval        =   1
      Left            =   1080
      Top             =   3480
   End
   Begin VB.HScrollBar Scroll_Horizontal 
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   4
      Top             =   165
      Width           =   3855
   End
   Begin VB.PictureBox imagen2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   9480
      Picture         =   "frmDiscoVinilo.frx":A00C
      ScaleHeight     =   5655
      ScaleWidth      =   5640
      TabIndex        =   3
      Top             =   7800
      Width           =   5640
   End
   Begin VB.HScrollBar Scroll_Horizontal 
      Height          =   255
      Index           =   2
      Left            =   10560
      TabIndex        =   2
      Top             =   164
      Width           =   2055
   End
   Begin VB.HScrollBar Scroll_Horizontal 
      Height          =   255
      Index           =   1
      Left            =   15240
      TabIndex        =   1
      Top             =   165
      Width           =   3975
   End
   Begin VB.HScrollBar Scroll_Horizontal 
      Height          =   255
      Index           =   0
      Left            =   1000
      TabIndex        =   0
      Top             =   165
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aumentar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9360
      TabIndex        =   6
      Top             =   165
      Width           =   870
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'                                                              '                                                        '
' copyright © : Martinsoft 2006 a 2015                         '
' programa escrito para Martin Virtual Rockola 1.0             '
' parecido al disco  de Virtual DJ                             '
' código fuente para Disco de Vinilo Giratorio en Windows      '
' autor del programa en BASIC                                  '
' by: Martin Grasso Castrillo.                                 '
' address: Canelones/Tala/Uruguay                              '
'                                                              '
'/''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'
'
'
'
'
Option Explicit 'declare qua este módulo de código es explicito.
'               'solo se ejecuta desde este módulo
'
'---------------------------------------------------------------
' Declaraciónes: son parecidas a las declarasiones de amor     -
' pero lo unico es que ejecutan procesos en memoria segun      -
' a las librerias que se declaran a querer                     -
' en este caso es gdit32.dll libreria de procesamiento grafico -
' del sistema operativo Windows                                -
'---------------------------------------------------------------
'
'
'---------------------------------------------------------------
'El SetStretchBltMode función establece el modo de mapa de bits-
'que se extiende en el contexto de dispositivo especificado.   -
'---------------------------------------------------------------
' Ayuda en: https://msdn.microsoft.com/en-us/library/windows/desktop/dd145089(v=vs.85).aspx

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal c_Imagen As Long, ByVal modo_Estirar As Long) As Long

'-----------------------------------------------------------------
' El PlgBlt función realiza una transferencia de bits de bloques -
' de los bits de datos de color de la rectángulo especificado en -
' el contexto de dispositivo de origen a la paralelogramo        -
' especificado en el contexto de dispositivo de destino. Si el   -
' mango máscara de bits                                          -
' dado identifica un mapa de bits monocromo válida, la función   -
' utiliza                                                        -
' este mapa de bits para enmascarar los bits de datos de color   -
' del                                                            -
' rectángulo de origen.                                          -
'-----------------------------------------------------------------
' Ayuda en: https://msdn.microsoft.com/en-us/library/windows/desktop/dd162804%28v=vs.85%29.aspx

Private Declare Function PlgBlt Lib "gdi32.dll" (ByVal prueba_HDC As Long, imagen_1 As imagen_xy, ByVal c_imagen_src As Long, ByVal c_imagen_src_xy As Long, ByVal c_imagen_src_xy_cx As Long, ByVal c_imagen_src_xy_cx_ancho As Long, ByVal c_imagen_src_xy_cx_alto As Long, ByVal mascara_HBM As Long, ByVal mascara_en_X As Long, ByVal mascara_en_Y As Long) As Long

'--------------------------------------------------------------
' constantes de Memoria                                       -
'--------------------------------------------------------------

Const PI = 3.14159265358979         'determina una constante de el número PI
Const GRADOS_180 = 180              'determina una pocicion de grados de rotación
Const GRADOS_90 = 90                'no modificable a no ser en el código de la constante
Const PUNTERO = 2                   'Puntero determinado por un número entero
'--------------------------------------------------------------
' contantes de Memoria en Scroll                              -
'--------------------------------------------------------------
Const Scroll_IDC = 0                'inicio de recorrido de controles
Const Scroll_CDC = 2                'maxima cantidad de Controles a ser recorridos
Const Scroll_LC = 1000              'largo vertical del control de desplazamiento
Const Scroll_MIN_012 = -180         'determina el minimo numero de desplazamiento ref: control 0 , 1 ,2
Const Scroll_MAX_012 = 180          'determina el maximo numero de desplazamiento ref: control 0 , 1 ,2
Const Scroll_IDC3 = 3               'inicio de control                            ref: control 3
Const Scroll_MIN_3 = 1              'determina el minimo numero de desplazamiento ref: control 3
Const Scroll_MAX_3 = 86             'determina el maximo numero de desplazamiento ref: control 3
Const Scroll_VALOR3 = 45            'determina el valor de inicio del Disco de Vinilo ref: control 3
'--------------------------------------------------------------
' variables de Memoria                                        -
'--------------------------------------------------------------

Private variable_memoria_entera As Long ' determina un elemento de memoria en formato entero. de longitud amplia de acceso privado
Private disco_virtual As Byte           ' determina cada cantidad de discos virtuales en memoria
Private i_imagen As Long
Private discos_cargados As Long         ' determina los discos cargados en la Memoria del Sistema
'--------------------------------------------------------------
' define un tipo de graficos tanto en X como en Y             -
'--------------------------------------------------------------

Private Type imagen_xy
 X As Long 'coordenada en X
 Y As Long 'coordenada en Y
End Type
'------------------------------------------------------------------
'determina un vector de  paneles virtuales de graficos en Memoria -
'------------------------------------------------------------------
Dim imagen_Lista(2) As imagen_xy

                                
Private Sub cmdDetener_Click()
If cmdDetener.Caption = "Detener" Then
   temporizador.Enabled = False
   cmdDetener.Caption = "Iniciar"
 ElseIf cmdDetener.Caption = "Iniciar" Then
        temporizador.Enabled = True
        cmdDetener.Caption = "Detener"
 End If
End Sub

Private Sub crear_PanelDiscos()
Dim contador As Byte
For contador = 0 To 5
If Not (disco_virtual = 5) Then

     disco_virtual = disco_virtual + 1
With ptx(disco_virtual)
     Load ptx(disco_virtual)
      ptx(disco_virtual).Visible = True
      ptx(disco_virtual).Left = 5700 * disco_virtual
      Panel.Width = ptx(0).Width * disco_virtual
      
      If disco_virtual = 5 Then
      Unload ptx(5)
      End If
      
End With
     Scroll_Vertical_X.Max = disco_virtual
     Scroll_Vertical_X.Min = 0
     
    End If
    Next contador
End Sub

Private Sub Form_Load()
    Call definir_propiedades    ' define la propiedades de los controles a utilizar
    Call crear_PanelDiscos      ' crea el panel de discos visuales de + - infinito
    Call Redibujar_Disco        ' redibuja los discos de vinilo visuales del + - infinito
    Call cargar_listaVirtual    ' cargar numeros de tiempo en lista
    Lista_Virtual.ListIndex = 3 ' arrancas con 30 de rpm
End Sub

Private Function PI_Dividio_180_grados()
'le pasamos cómo parametro set a la PI_Dividio_180_grados
PI_Dividio_180_grados = PI / GRADOS_180
End Function

Private Sub Redibujar_Disco()
 '
 Dim X            As Single
 '
 Dim Nueva_X      As Integer
 Dim Nueva_Y      As Integer
 '
 Dim Sin_Angulo_1 As Single
 Dim Cos_Angulo_1 As Single
 Dim Sin_Angulo_2 As Single
 Dim Sin_Angulo_3 As Single
 Dim Ampliar      As Single
 '-----------------------------------------------------------------------------------------------
 'Restablece punteros de Listas                                                                 -
 '-----------------------------------------------------------------------------------------------
 imagen_Lista(0).X = -(imagen2.ScaleWidth / PUNTERO)
 imagen_Lista(0).Y = -(imagen2.ScaleHeight / PUNTERO)
 imagen_Lista(1).X = (imagen2.ScaleWidth / PUNTERO)
 imagen_Lista(1).Y = -(imagen2.ScaleHeight / PUNTERO)
 imagen_Lista(2).X = -(imagen2.ScaleWidth / PUNTERO)
 imagen_Lista(2).Y = (imagen2.ScaleHeight / PUNTERO)
 '----------------------------------------------------------------------------------------------
 'Ampliar grafico del Disco de Vinilo                                                          -
 '----------------------------------------------------------------------------------------------
 Ampliar = Tan(Scroll_Horizontal(3).Value * PI_Dividio_180_grados)
 Sin_Angulo_1 = Sin((Scroll_Horizontal(0).Value + GRADOS_90) * PI_Dividio_180_grados)
 Cos_Angulo_1 = Cos((Scroll_Horizontal(0).Value + GRADOS_90) * PI_Dividio_180_grados)
 Sin_Angulo_2 = Sin((Scroll_Horizontal(1).Value + GRADOS_90) * PI_Dividio_180_grados) * Ampliar
 Sin_Angulo_3 = Sin((Scroll_Horizontal(2).Value + GRADOS_90) * PI_Dividio_180_grados) * Ampliar
 For X = 0 To 2
 Nueva_X = (imagen_Lista(X).X * Sin_Angulo_1 + imagen_Lista(X).Y * Cos_Angulo_1) * Sin_Angulo_2
 Nueva_Y = (imagen_Lista(X).Y * Sin_Angulo_1 - imagen_Lista(X).X * Cos_Angulo_1) * Sin_Angulo_3
 imagen_Lista(X).X = Nueva_X + (imagen1.ScaleWidth / PUNTERO)
 imagen_Lista(X).Y = Nueva_Y + (imagen1.ScaleHeight / PUNTERO)
 Next
 imagen1.Cls 'clase de imagen
 '--------------------------------------------------------------------------------
 ' Suaviza el Acabado de la Imagen de Mapa de Byts cuando este en pocicion recta -
 '--------------------------------------------------------------------------------
 
 SetStretchBltMode imagen1.hDC, vbPaletteModeNone
 'API de Windows que rota la imagen
 Call PlgBlt(imagen1.hDC, imagen_Lista(0), imagen2.hDC, 0, 0, imagen2.ScaleWidth, imagen2.ScaleHeight, 0, 0, 0)
 imagen1.Refresh

End Sub


'
' define el estado tamaño y longitud de la barra de Desplazamiento
' y configuraciónes graficas
'

Private Sub definir_propiedades()
'
'
Dim contador_virtual As Byte
'
'
    For contador_virtual = Scroll_IDC To Scroll_CDC
        With Scroll_Horizontal(contador_virtual)
            .LargeChange = Scroll_LC
            .Max = 180
            .Min = -180
        End With
    Next
    
    With Scroll_Horizontal(Scroll_IDC3)
        .LargeChange = Scroll_LC
        .Max = Scroll_MAX_3
        .Min = Scroll_MIN_3
        .Value = Scroll_VALOR3
    End With
    '/----------------------------------------------------------------------/
    '/  Define las propiedades de las Imagenes del control                 -/
    '/----------------------------------------------------------------------/
    With imagen2
        .AutoSize = True
        .Visible = False
        .AutoRedraw = True
        .ScaleMode = vbPixels
End With

    With imagen1
    .AutoRedraw = True
    .ScaleMode = vbPixels
    .AutoSize = True
End With

End Sub
'--------------------------------------------------------------
' este código inferior hace que cuando se produsca un cambio  -
' de pocicion del Scroll se redibuje el Disco de Vinilo       -
'--------------------------------------------------------------
Private Sub Scroll_Horizontal_Change(Index As Integer)
 Call Redibujar_Disco
     Label1.Caption = Scroll_Horizontal(3).Value & "%"
End Sub
'--------------------------------------------------------------
' este código inferior hace que cuando se produsca un         -
' movimiento  de pocicion del Scroll se redibuje el Disco de  -
' Vinilo                                                      -
'--------------------------------------------------------------
Private Sub Scroll_Horizontal_Scroll(Index As Integer)
Scroll_Horizontal_Change Index
End Sub

Private Sub Scroll_Vertical_X_Change()
 Panel.Left = -Scroll_Vertical_X.Value * 3000
  labCanciones.Caption = "DISCO :" & discos_cargados & " TOTAL DE DISCOS CARGADOS: 2800000"
 If Scroll_Vertical_X.Max = Scroll_Vertical_X.Value Then
    Panel.Left = Scroll_Vertical_X.Min
    Scroll_Vertical_X = Scroll_Vertical_X.Min + 1
     discos_cargados = discos_cargados + 1
 End If
 If Scroll_Vertical_X.Min = Scroll_Vertical_X.Value Then
    Panel.Left = Scroll_Vertical_X.Max
    Scroll_Vertical_X = Scroll_Vertical_X.Max - 1
     discos_cargados = discos_cargados - 1
 End If
End Sub

Private Sub Scroll_Vertical_X_Scroll()
Scroll_Vertical_X_Change
End Sub

Private Sub temporizador_Timer()
  Scroll_Horizontal.Item(0).Value = variable_memoria_entera
  variable_memoria_entera = CInt(variable_memoria_entera + Lista_Virtual.List(Lista_Virtual.ListIndex))
 If Scroll_Horizontal.Item(0).Value >= 180 Then
   variable_memoria_entera = -180
   Scroll_Horizontal.Item(0).Value = -180
End If

For i_imagen = 0 To ptx.Count - 1
    ptx(i_imagen).Picture = imagen1.Image
    Picture1.Picture = ptx(i_imagen).Picture
Next
End Sub

Private Sub versoloDisco_Click()
If pxld.Visible = False Then
   pxld.Visible = True
   ElseIf pxld.Visible = True Then
   pxld.Visible = False
End If
End Sub
'agrega numeros de tiempo
Private Sub cargar_listaVirtual()
      With Lista_Virtual
          .AddItem CByte(0)
          .AddItem CByte(1)
          .AddItem CByte(2)
          .AddItem CInt(30)
      End With
End Sub
