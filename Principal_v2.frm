VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Menu Principal"
   ClientHeight    =   7830
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   9225
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5404.405
   ScaleMode       =   0  'User
   ScaleWidth      =   8662.754
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      Picture         =   "Principal_v2.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   0
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Info. del sistema..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   2
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "CONSULTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3390
      Left            =   0
      Picture         =   "Principal_v2.frx":030A
      Top             =   2640
      Width           =   14520
   End
   Begin VB.Label lblDescription 
      Caption         =   "CONTROL DE COMPRAS"
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "DERIVADOS SIDERURGICOS, C.A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   5085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   7549.978
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión 1.2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   4125
   End
   Begin VB.Menu Proveedores 
      Caption         =   "&Proveedores"
      Begin VB.Menu Expediente 
         Caption         =   "&Expediente"
      End
      Begin VB.Menu EVALUACION01 
         Caption         =   "Evaluacion &Individual"
      End
      Begin VB.Menu EVALUACION02 
         Caption         =   "Evaluacion &Anual"
      End
   End
   Begin VB.Menu Planilla 
      Caption         =   "P&lanilla Retencion IVA"
   End
   Begin VB.Menu Ordenes 
      Caption         =   "&Ordenes de Compras"
   End
   Begin VB.Menu RequisitosLegales 
      Caption         =   "&Requisitos Legales"
      Begin VB.Menu GrupoServicios 
         Caption         =   "Grupo &Servicios"
      End
      Begin VB.Menu ProductoServicio 
         Caption         =   "Producto &Servicio"
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu Imprimir_Orden_Compra 
         Caption         =   "Imprimir Orden de Compra"
      End
      Begin VB.Menu Comprobante_RET_IVA 
         Caption         =   "Comprobante de RET IVA"
      End
      Begin VB.Menu MEMORANDUM 
         Caption         =   "Emitir &Memorandum"
      End
      Begin VB.Menu Maestro_Proveedoes 
         Caption         =   "&Listado Proveedores"
      End
      Begin VB.Menu Proveedores_Evaluacion 
         Caption         =   "&Proveedores/Evaluacion"
      End
      Begin VB.Menu Proveedor_Rif_Vencido 
         Caption         =   "Proveedor(es) Rif Vencido"
      End
      Begin VB.Menu Evaluar_Tiempo_Entrega 
         Caption         =   "&Evaluar Tiempo de Entrega"
      End
      Begin VB.Menu Consultar_Ord_Prov 
         Caption         =   "Relacion Ordenes x Periodo"
      End
      Begin VB.Menu Consultar_Item_Orden 
         Caption         =   "Relacion Items/Orden"
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'* Proyecto de Visual Basic 6.0
'  Control de Compras
'  ++++++++++++++++++
'
'* Nombre
'    * Logico: Principal
'    * Fisico: Principal
'
' Autor: Henry J. Pulgar B.
' Creado el 18 de Julio del año 2003.
' Actualizado el 18 de Septiembre del año 2003.
'***************************************************
'
Dim CurrentDir As String
Dim CurrentUser As String
Dim Comando As String
Dim ExeComando As String
Option Explicit

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Comprobante_RET_IVA_Click()
  Comando = "rwrun60 report=" & CurrentDir & "COMPROB_RET_IVAv01.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub
'****************************************************************
Private Sub Consultar_Item_Orden_Click()
  Comando = "rwrun60 report=" & CurrentDir & "RELACION_ITEMS.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub
'*****************************************************************
Private Sub Consultar_Ord_Prov_Click()
  Comando = "rwrun60 report=" & CurrentDir & "RELACION_ORDENES_v3.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

Private Sub Evaluar_Tiempo_Entrega_Click()
 Comando = "rwrun60 report=" & CurrentDir & "TIEMPO_ENTREGA_PROV_v2.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

'******************************************************************
Private Sub Form_Load()
  'Me.Caption = "Acerca de " & App.Title
  lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
  'lblTitle.Caption = App.Title
  '--- Init Variables -----
  CurrentDir = ""    'Sera nulo para efectos de instalacion.
  'CurrentUser = "OPS$DESCOM02/OPS$DESCOM02@bd816"
  'CurrentDir = "f:\Reports6\Compras\"
  ' **
  ' ** Relacion de Usuarios y Privilegios:
  ' **
  'CurrentUser = "OPS$contab/OPS$contab@bd733"
  CurrentUser = "OPS$DESCOM02/OPS$DESCOM02@bd806" 'Usuario con privilegio de Actualizacion.
  'CurrentUser = "OPS$DESCOM03/OPS$DESCOM03@bd806"  'Usuario con privilegio de Solo Consulta.
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

Private Sub GrupoServicios_Click()
  Comando = "ifrun60 " & CurrentDir & "GRUPO " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

'************************************************************************
Private Sub Maestro_Proveedoes_Click()
  Comando = "rwrun60 " & CurrentDir & "MAESTRO_PROV " & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

Private Sub MEMORANDUM_Click()
  Comando = "rwrun60 " & CurrentDir & "MEMORANDUM " & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

'*************************************************************************
Private Sub Ordenes_Click()
  'Comando = "ifrun60 " & CurrentDir & "ORDENES_COMPRA_v3 " & CurrentUser
  Comando = "ifrun60 " & CurrentDir & "ORDENES_COMPRA_v4 " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

'******************************************************************+
Private Sub Planilla_Click()
  COMPROB_RET.Show
End Sub

'******************************************************************+
Private Sub Expediente_Click()
  Comando = "ifrun60 " & CurrentDir & "PROVEEDORESv3 " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

'******************************************************************+
Private Sub Evaluacion01_Click()
  Comando = "ifrun60 " & CurrentDir & "EVALUACION01 " & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

'****************************************************************
Private Sub EVALUACION02_Click()
  Comando = "ifrun60 " & CurrentDir & "EVALUACION02 " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

Private Sub ProductoServicio_Click()
  Comando = "ifrun60 " & CurrentDir & "REQ_PRODUCTO_SERV " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

Private Sub Proveedor_Rif_Vencido_Click()
  Comando = "rwrun60 report=" & CurrentDir & "ProvRifVencidos.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

'***************************************
'* Reporte:
'***************************************
Private Sub Proveedores_Evaluacion_Click()
  'Comando = "rwrun60 report=" & CurrentDir & "EVALUACION.rdf userid=" & CurrentUser
  Comando = "rwrun60 report=" & CurrentDir & "EVALUACIONv3.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub


'*******************************************
'*
'*******************************************
Private Sub Imprimir_Orden_Compra_Click()
  'PRINT_ORDEN_COMPRA.Show
  PRINT_ORDEN_COMPRAv2.Show
End Sub
'********************************************
Private Sub Salir_Click()
  Unload Me
End Sub
