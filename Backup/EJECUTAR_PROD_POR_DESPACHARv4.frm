VERSION 5.00
Object = "{5E7F0291-A395-497B-9712-CE79A6B288EC}#3.0#0"; "ctDropDate.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EJECUTAR_PROD_POR_DESPACHARv4 
   BackColor       =   &H8000000A&
   Caption         =   "Ejecutar Productos / Despachar ( v 4.00 )"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MyIcoWait 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6480
      Picture         =   "EJECUTAR_PROD_POR_DESPACHARv4.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   480
      Top             =   6480
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin CTDROPDATELib.ctDropDate AL_FECHA 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropPicture     =   "EJECUTAR_PROD_POR_DESPACHARv4.frx":030A
      FormatType      =   1
      Text            =   "__-__-____"
      DateSepChar     =   "-"
   End
   Begin VB.CommandButton BOTON_SALIR 
      Caption         =   "SALIR /CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      MouseIcon       =   "EJECUTAR_PROD_POR_DESPACHARv4.frx":0326
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton BOTON_EJECUTAR 
      Caption         =   "EJECUTAR REPORTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      MouseIcon       =   "EJECUTAR_PROD_POR_DESPACHARv4.frx":0630
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NOTA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      HelpContextID   =   1
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      Begin VB.Label Label2 
         Caption         =   "Haga ""click"" en el Boton ""Ejecutar Reporte""  y espere mientras  se prepara toda la informacion,...               "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Este modulo reunira en una tabla los datos necesarios para la ejecucion de este Reporte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   6855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Al fecha ( DD-MM-AAAA) : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Por favor sea paciente, ..."
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "EJECUTAR_PROD_POR_DESPACHARv4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'*              * EJECUTAR_PROD_POR_DESPACHARv4 *
'*  Autor        : Henry J. Pulgar B.
'*  Creado       : 22 de Noviembre 2006.
'*  Actualizado  : 30 de Noviembre 2006.
'*
'*  NOTA         : Opcion de Reportes "Productos por Despachar vs Existencia
'*                 vs Prod. Por Producir" del sistema de Control de Pedidos.
'**************************************************************************
Dim CurrentDir As String
Dim CurrentUser As String
Dim MyCursor As StdPicture

'**************************************************************************
Private Sub PAUSAR(TiempoPausar As Integer)
   Iniciar = Timer
   Do While Timer < (Iniciar + TiempoPausar)
     'Do nothing for now
   Loop
End Sub


'**************************************************************************
Private Sub CHECK_DATE_VALUE()
     If Not IsNull(AL_FECHA.Text) Then
        MsgBox "Fecha seleccionada = " + AL_FECHA.Text
     Else
        Beep
        MsgBox "ERROR: fecha no puede se nulo."
     End If
End Sub
'***************************************************************************
Private Function VALIDAR_DATE_VALUE_OK(Fecha As String) As Boolean
    If IsNull(AL_FECHA) Or (Fecha = "__-__-____") Then
       ' Fecha es nula
       VALIDAR_DATE_VALUE_OK = False
    Else
       'MsgBox "Fecha no es nula : (" + Fecha + ")" + Str(Len(Fecha))
       VALIDAR_DATE_VALUE_OK = True
    End If
End Function

'***********************************************************************
'* Autor: Henry J. Pulgar B.
'* Creado el 23-11-2006.
'* PROPOSITO:  Establecer enlace a la B.D. utilizando ODBC
'*             - Estructura de datos "ADODB" del grupo de componentes
'*               "Microsoft ADO Data Control 6.0 (OLEDB )"
'***********************************************************************
Private Function CREAR_DATOS_TEMPORAL()
   '* Dimensionar Variables para establecer el enlace a la B.D.
   Dim MyConn As ADODB.Connection   ' cmd para mant. de la B.D.
   Dim CadenaSql As String
   '
   '* Set Vars.
   Set MyConn = New ADODB.Connection  'cmd para mant. de la B.D.
   MyConn.Open "DSN=DESICA806;UID=OPS$DESINV01;PWD=OPS$DESINV01", "", ""
   '
   ' CUERPO DE LA CONEXION
   'MsgBox " My Conn state = " & ConnRec.State
   '************************** H.B.B: 23-11-2006.**************************
   'MyConn.Execute "insert INTO TEMP_PEDIDOSv4_DAT " & _
   '                "select C4_CODIGO, C4_NO_PEDIDO, nvl( SUM( C4_CANTIDAD ),0 ) " & _
   '                "from   VEND04_DAT, VEND03_DAT " & _
   '                "where  C4_NO_PEDIDO = C3_NO_PEDIDO " & _
   '                "and    C3_FECHA_PEDIDO <= '" + AL_FECHA.Text + "' " & _
   '                "group  by C4_CODIGO, C4_NO_PEDIDO"
   '***********************************************************************
   '********* I Parte: limpiar tabla temporal : TEMP_PEDIDOSv4_DAT:
   MyConn.Execute "delete from TEMP_PEDIDOSv4_DAT"
   '
   MyConn.BeginTrans
   'MyConn.CommandTimeout = 30   ' Configurada/Deshabilitada en la Cadena de Conexion ODBC ( DESICA806 ).
   '********** II Parte : Incluuir Pedidos:
   CadenaSql = "insert INTO TEMP_PEDIDOSv4_DAT " & _
               "select C4_CODIGO, C4_NO_PEDIDO, nvl( SUM( C4_CANTIDAD ),0 ) " & _
               "from   VEND04_DAT, VEND03_DAT " & _
               "where  C4_NO_PEDIDO = C3_NO_PEDIDO " & _
               "and    C3_FECHA_PEDIDO <= '" + AL_FECHA.Text + "' " & _
               "group  by C4_CODIGO, C4_NO_PEDIDO"
   MyConn.Execute CadenaSql
   '
   '********** III Parte : Actualizar despachos:
   CadenaSql = "update TEMP_PEDIDOSv4_DAT " & _
               "set    CANTIDAD = ( select CANTIDAD - nvl( SUM( C2_UNIDADES ),0 ) " & _
                                   "from   GUIAS01_DAT, GUIAS02_DAT " & _
                                   "where  C1_GUIA  = C2_GUIA " & _
                                   "and    C2_CODIGO = CODIGO " & _
                                   "and    C2_NO_PEDIDO = NO_PEDIDO " & _
                                   "and    C1_FECHA_GUIA <= '" + AL_FECHA.Text + "') "
   MyConn.Execute CadenaSql
   
   '********** IV Parte : Ajustar Pedidos del material **********************
   CadenaSql = "update TEMP_PEDIDOSv4_DAT " & _
                "set    CANTIDAD = ( select CANTIDAD + nvl( SUM( C5_CANTIDAD ),0 ) " & _
                                    "from   VEND05_DAT " & _
                                    "where  C5_NO_PEDIDO = NO_PEDIDO " & _
                                    "and    C5_CODIGO    = CODIGO " & _
                                    "and    C5_FECHA     <= '" + AL_FECHA.Text + "' )"
   MyConn.Execute CadenaSql
   '
   'MsgBox "INSERCION TERMINADA" & " " & Str(MyConn.Errors.Count)
   MyConn.CommitTrans
   '
   If (MyConn.Errors.Count = 0) Then
       MyConn.Close
       CREAR_DATOS_TEMPORAL = True
   Else
       MyConn.Close
       CREAR_DATOS_TEMPORAL = False
   End If
End Function

'***********************************************************************
Private Sub AL_FECHA_GotFocus()
  AL_FECHA.BackColor = &H80000005   'Fondo en blanco.
End Sub

'***********************************************************************
Private Sub Form_Load()
    CurrentDir = ""                                'RUN TIME Period.
    'CurrentDir = "C:\Vb6\Proyectos\Pedidos\"      'DEVELOPMENT Period.
    CurrentUser = "OPS$DESINV01/OPS$DESINV01@BD806"
End Sub
'************************************************************************
Private Sub SET_THE_CURSOR(TipoCursor As Integer)
   If TipoCursor = 0 Then
        'Set MyCursor => {ninguno}
        MyIcoWait.Visible = False
   Else
        'Set MyCursor = LoadPicture(CurrentDir + "WAIT02.cur")
         Set MyCursor = LoadPicture(CurrentDir + "WAIT07.cur")
         MyIcoWait.Visible = True
   End If
   EJECUTAR_PROD_POR_DESPACHARv4.MousePointer = TipoCursor  'Definido por el programador.
   EJECUTAR_PROD_POR_DESPACHARv4.MouseIcon = MyCursor
   Refresh
   PAUSAR (2)   'Estara funcionando ???????????
End Sub
'***********************************************************************
'* Autor: Henry J. Pulgar B.
'* Creado el 23-11-2006.
'* NOTA:  No se activa la el iconon de la funcion SET_THE_CURSOR quizas por el
'         consumo de memoria y el tiempo de ejecucion entre el procesador y
'         la memoria de Video.
'***********************************************************************
Private Sub BOTON_EJECUTAR_Click()
  Dim Comando As String
  Dim DefinidoPorUsuario As Integer
  '
  DefinidoPorUsuario = 99
  '
  'CHECK_DATE_VALUE
  If Not VALIDAR_DATE_VALUE_OK(AL_FECHA.Text) Then
     Beep
     AL_FECHA.BackColor = &HFF& 'Fondo en color rojo.
     MsgBox "Fecha para la ejucion de este reporte no puede ser nula.", vbCritical, "ATENCION"
  Else 'VALIDAR_DATE_VALUE
      SET_THE_CURSOR (DefinidoPorUsuario)
      If CREAR_DATOS_TEMPORAL Then
         Comando = "rwrun60 report=" & CurrentDir & "Prod_Por_Despachar_v4.rdf P_AL_FECHA=" & AL_FECHA.Text & " userid=" & CurrentUser
         ExeComando = Shell(Comando, vbNormalFocus)
         'MsgBox "REPORTE EJECUTADO"
         'If ExeComando <> 0 Then  ???????????????'
         '   MsgBox "Se encontraron errores al tratar de ejecutar este reporte. Comsulte al Administrador del Sistema.", vbCritical, "ATENCION"
         'End If
      Else
         Beep
         MsgBox "Error al generar tabla de datos temporales", vbCritical, "ATENCION"
      End If
      SET_THE_CURSOR (0)
      Unload Me
   End If
End Sub

Private Sub BOTON_SALIR_Click()
     Unload Me
End Sub

' Uso de la funcion LoadResPicture() codigo oculto del ejemplo \Visual Studio Sample\...\Vb98\Samples\...\Atm\.
Sub Cursor_Initialize()
    Set curSelect = LoadResPicture(1, vbResCursor)
End Sub
'*****************************EOF(EJECUTAR_PROD_POR_DESPACHARv4.vb)********************
