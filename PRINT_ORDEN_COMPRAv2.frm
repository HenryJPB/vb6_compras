VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PRINT_ORDEN_COMPRAv2 
   Caption         =   "IMPRIMIR ORDEN COMPRA"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   5295
      Begin VB.ComboBox Num_Orden_Hasta 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox Num_Orden_Desde 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "2. Hasta el Numero de Orden ?:  "
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "1. Desde el Numero de Orden ?:"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CB_Cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton CB_Aceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Picture         =   "PRINT_ORDEN_COMPRAv2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Parametros requeridos para la emision de Ordenes de Compra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "PRINT_ORDEN_COMPRAv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'* GRUPO de programas: CONTROL de PROVEEDORES
'* B.D. Oracle 8.1.0.6.
'* Nombre Logico: IMPRIMIR ORDEN DE COMPRA
'* Nombre Fisico: PRINT_ORDEN_COMPRA
'* Autor:         Henry Jose Pulgar B.
'* Creado : el 28 de Julio de 2003.
'* Actualizado : el 01 de Febrero 2006.
'*********************************************************
Public CurrentUser As String
Public OrigenDatos As String
Public UserName As String
Public Clave    As String
Public Nit_Empresa As String
Public Tope_Lineas_Detail As Integer
Public Coneccion As ADODB.Connection
Public Reg As ADODB.Recordset
'*1 *Mina de Datos publicos **
Public FormatoNroOrden
Public Nombre
Public Fecha_Orden
Public Direccion1
Public Orden_Servicio
Public Direccion2
Public Direccion3
Public Orden_Compra
Public Rif
Public Nit
Public Cod_Prov
Public Req_No1
Public Fecha_Requisicion1
Public Cond_Pago
Public Requerimiento
Public Req_No2
Public Fecha_Requisicion2
Public Anticipo
'*2

Private Sub Form_Load()
    'Fecha_Desde.Text = Format(Now, "DD-MM-YYYY")
    'Fecha_Hasta.Text = Format(Now, "DD-MM-YYYY")
    '********************************
    '** Definido originalmente asi:**
    'OrigenDatos = "DESICA816"
    'UserName = "OPS$DESCOM02"
    'Clave = "OPS$DESCOM02"
    '********************************
    'Usuario temporal:
    'OrigenDatos = "DESICA733"
    'UserName = "OPS$contab"
    'Clave = "OPS$contab"
    '**
    Nit_Empresa = "0002139855"    'NIT DESICA
    OrigenDatos = "desica806"
    UserName = "OPS$DESCOM02"
    Clave = "OPS$DESCOM02"
    Tope_Lineas_Detail = 20     '<- originalmente asi.
    'Tope_Lineas_Detail = 22     '
    'Tope_Lineas_Detail = 21     '
    FormatoNroOrden = "00000"    '*  5's digitos ceros.
    LOAD_ORDENES_DIFERIDAS
End Sub

'*-------------------------------------------------------
Private Sub CB_Aceptar_Click()
    IMPRIMIR_ROUTINE   '<- En stand by esta instruccion. Ver nota adjunta. ????
    'NOTA: de implantarse la instruccion anterior
    '      se debe eliminar las 3 instrucciones siguientes.
    'Printer.Font.Name = "Draft"
    'Printer.Font.Bold = False
    'Imprimir_Orden_Compra_vEPSON_LX810
End Sub

'*--------------------------------------------------------
Private Sub CB_Cancelar_Click()
   Unload Me
End Sub
'*********************************************************
'*********************************************************
Private Function OPEN_DATABASE() As Boolean
 '-------
  'Parte I: establecer la connecion via ODBC.
  '-------
  Set Coneccion = New ADODB.Connection
  '* Via coneccion ODBC ...
  '* Activar ORACLE conneccion - via ODBC
  Coneccion.ConnectionString = "DSN=" & OrigenDatos & _
                               ";UID=" & UserName & _
                               ";PWD=" & Clave
  Coneccion.Open
  '--------
  'PARTE II: Ajustar valores para accesar los registros
  '--------
  Set Reg = New ADODB.Recordset
  SQL1 = "select    C1_NOMBRE," & _
                   "C1_DIRECCION1," & _
                   "C1_DIRECCION2," & _
                   "C1_DIRECCION3," & _
                   "C1_RIF," & _
                   "C1_NIT," & _
                   "C1_TELEFONO1," & _
                   "C1_FAX1," & _
                   "C2_FECHA_ORDEN," & _
                   "DECODE( C2_TIPO_ORDEN,'C','X',' ') Orden_Compra," & _
                   "DECODE( C2_TIPO_ORDEN,'S','X',' ') Orden_Servicio,"
  '
  SQL2 = "C1_CODIGO_PROV," & _
                   "C2_REQUISICION_NO1," & _
                   "C2_REQUISICION_NO2," & _
                   "C2_FECHA_REQUISICION1," & _
                   "C2_FECHA_REQUISICION2," & _
                   "C2_CONDICION_PAGO," & _
                   "nvl(C2_ANTICIPO,0) as C2_ANTICIPO," & _
                   "nvl(C2_PLAZO,0) as C2_PLAZO," & _
                   "C2_REQUERIMIENTO," & _
                   "C2_NUMERO_ORDEN," & _
                   "C2_STATUS," & _
                   "C2_CON_TOTAL," & _
                   "C2_MONEDA_EXTRANJERA," & _
                   "TO_NUMBER( C3_CANTIDAD ) C3_CANTIDAD," & _
                   "C3_CODIGO_ITEM," & _
                   "C3_DESCRIPCION," & _
                   "TO_NUMBER( C3_PRECIO_UNID ) C3_PRECIO_UNID," & _
                   "TO_NUMBER( C3_TOTAL_ITEM ) C3_TOTAL_ITEM "
  '
  SQL3 = "From COMPRAS01_DAT, COMPRAS02_DAT, COMPRAS03_DAT " & _
           "Where C2_CODIGO_PROV = C1_CODIGO_PROV " & _
           "and   C2_NUMERO_ORDEN between " & Num_Orden_Desde.Text & " and " & Num_Orden_Hasta.Text & " " & _
           "and   C3_NUMERO_ORDEN = C2_NUMERO_ORDEN " & _
           "order by C2_NUMERO_ORDEN, C2_FECHA_ORDEN, C3_CODIGO_ITEM "
  SQL = SQL1 + SQL2 + SQL3
  'Reg.Open SQL, Coneccion, adOpenStatic, adLockOptimistic
  Reg.Open SQL, Coneccion, adOpenForwardOnly, adLockReadOnly
  If Reg.EOF Then
     MsgBox "Error al abrir B.D.: Numero de Orden no registrada o fuera de rango.", vbCritical, "ATENCION"
     OPEN_DATABASE = False
  Else
     OPEN_DATABASE = True
  End If
End Function   'OPEN_DATABASE

'*****************************************************
'*****************************************************
Private Sub CLOSE_DATABASE()
'Cerrar
 Reg.Close
 Coneccion.Close
End Sub

'****************************************************
'****************************************************
Private Sub Num_Orden_Desde_LostFocus()
   If Len(Num_Orden_Hasta) = 0 Then
      Num_Orden_Hasta.Text = Num_Orden_Desde.Text
   End If
End Sub

'****************************************************
'****************************************************
Private Sub LOAD_ORDENES_DIFERIDAS()
    Dim Conn As ADODB.Connection
    Dim ConnRec As ADODB.Recordset
    Set Conn = New ADODB.Connection
    Set ConnRec = New ADODB.Recordset
    Conn.ConnectionString = "DSN=" & OrigenDatos & _
                            ";UID=" & UserName & _
                            ";PWD=" & Clave
    Conn.Open
    ConnRec.Open "SELECT C2_NUMERO_ORDEN FROM COMPRAS02_DAT WHERE C2_STATUS = 'D' ORDER BY C2_NUMERO_ORDEN", Conn, adOpenForwardOnly, adLockReadOnly
    While Not ConnRec.EOF
          Num_Orden_Desde.AddItem (ConnRec("C2_NUMERO_ORDEN"))
          Num_Orden_Hasta.AddItem (ConnRec("C2_NUMERO_ORDEN"))
          ConnRec.MoveNext
    Wend
    ConnRec.Close
    Conn.Close
End Sub 'LOAD_ORDENES_DIFERIDAS
'****************************************************
'****************************************************
Private Sub LOAD_REQUISITOS_LEGALES(Numero_Orden As Variant, ByRef Requisitos() As String)
    Dim Conn As ADODB.Connection
    Dim ConnRec As ADODB.Recordset
    Set Conn = New ADODB.Connection
    Set ConnRec = New ADODB.Recordset
    Conn.ConnectionString = "DSN=" & OrigenDatos & _
                            ";UID=" & UserName & _
                            ";PWD=" & Clave
    Conn.Open
    SqlTexto = "select c3a_cod_req, c10_descripcion as RequisitoLegal " & _
               "from   COMPRAS03A_DAT, COMPRAS10_DAT " & _
               "where  c3a_cod_req = c10_cod_req " & _
               "and    c3a_nro_orden = '" & Numero_Orden & "'"
    'ConnRec.Open "SELECT C2_NUMERO_ORDEN FROM COMPRAS02_DAT WHERE C2_STATUS = 'D' ORDER BY C2_NUMERO_ORDEN", Conn, adOpenForwardOnly, adLockReadOnly
    ConnRec.Open SqlTexto, Conn, adOpenForwardOnly, adLockReadOnly
    i = -1
    While Not ConnRec.EOF
          'Num_Orden_Desde.AddItem (ConnRec("C2_NUMERO_ORDEN"))
          'Num_Orden_Hasta.AddItem (ConnRec("C2_NUMERO_ORDEN"))
          RequisitoLegal = ConnRec("RequisitoLegal")
          i = i + 1
          Requisitos(i) = RequisitoLegal
          ConnRec.MoveNext
    Wend
    ConnRec.Close
    Conn.Close
    'MsgBox "Tama�o de mi array=" + Str(UBound(Requisitos) - LBound(Requisitos) + 1)  ' Error.
    'For i = LBound(Requisitos) To UBound(Requisitos)
    '    MsgBox Requisitos(i)
    'Next
    'i = -1
    'Termine = False
    'While (Not Termine And i <= UBound(Requisitos))
    '   i = i + 1
    '   If (Not IsNull(Requisitos(i)) And Mid(Requisitos(i), 1, 1) <> "") Then
    '      MsgBox Requisitos(i)
    '   Else
    '      Termine = True
    '   End If
    'Wend
End Sub 'LOAD_REQUISITOS LEGALES

'**************************************************************************************************************
'**************************************************************************************************************
Private Function POSEE_REQ_LEGALES(ByVal Numero_Orden As String) As Boolean
    Dim cuantosReq As Integer
    Dim Conn As ADODB.Connection
    Dim ConnRec As ADODB.Recordset
    Set Conn = New ADODB.Connection
    Set ConnRec = New ADODB.Recordset
    Conn.ConnectionString = "DSN=" & OrigenDatos & _
                            ";UID=" & UserName & _
                            ";PWD=" & Clave
    Conn.Open
    SqlTexto = "select Count( c3a_cod_req ) as CuantosReq " & _
               "from   COMPRAS03A_DAT " & _
               "where  c3a_nro_orden = '" & Numero_Orden & "'"
    'ConnRec.Open SqlTexto, Conn, adOpenForwardOnly, adLockReadOnly
    ConnRec.Open SqlTexto, Conn, adOpenStatic, adLockOptimistic
    cuantosReq = 0
    If Not ConnRec.EOF Then
          'MsgBox "Aquicaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa..............." + Str(ConnRec("CuantosReq")) + "Nro orden=" + Numero_Orden
          cuantosReq = ConnRec("CuantosReq")
    End If
    ConnRec.Close
    Conn.Close
    If (cuantosReq > 0) Then
        POSEE_REQ_LEGALES = True
    Else
        POSEE_REQ_LEGALES = False
    End If
    '* +++++++++++ Aside ++++++++++++++
    'MsgBox "Tama�o de mi array=" + Str(UBound(Requisitos) - LBound(Requisitos) + 1)  ' Error.
    'For i = LBound(Requisitos) To UBound(Requisitos)
    '    MsgBox Requisitos(i)
    'Next
    'i = -1
    'Termine = False
    'While (Not Termine And i <= UBound(Requisitos))
    '   i = i + 1
    '   If (Not IsNull(Requisitos(i)) And Mid(Requisitos(i), 1, 1) <> "") Then
    '      MsgBox Requisitos(i)
    '   Else
    '      Termine = True
    '   End If
    'Wend
End Function   '* Posee Req Legales ?.  -True/False.


'****************************************************
'****************************************************
Private Sub ACTUALIZAR_STATUS(Numero_Orden As Variant)
    Dim Conn2 As ADODB.Connection
    Dim ConnRec2 As ADODB.Recordset
    Set Conn2 = New ADODB.Connection
    Set ConnRec2 = New ADODB.Recordset
    Conn2.ConnectionString = "DSN=" & OrigenDatos & _
                            ";UID=" & UserName & _
                            ";PWD=" & Clave
    Conn2.Open
    'La siguiente instruccion sin el campo 'C2_NUMERO_ORDEN' ejerce un error indeseable en el
    'comportamiento del programa:
    'ConnRec.Open "SELECT C2_STATUS FROM COMPRAS02_DAT WHERE C2_NUMERO_ORDEN = " & Numero_Orden, Conn, adOpenStatic, adLockOptimistic
    '--------
    'Solucion:
    'ConnRec2.Open "SELECT C2_NUMERO_ORDEN, C2_STATUS FROM COMPRAS02_DAT WHERE C2_NUMERO_ORDEN = " & Numero_Orden, Conn2, adOpenStatic, adLockBatchOptimistic
    Cadena_SQL = "SELECT C2_NUMERO_ORDEN, C2_STATUS FROM COMPRAS02_DAT WHERE C2_NUMERO_ORDEN = '" & Numero_Orden & "'"
    ConnRec2.Open Cadena_SQL, Conn2, adOpenStatic, adLockOptimistic
    If Not ConnRec2.EOF Then
       ConnRec2("C2_STATUS") = "I"  'Orden I)mpresa; D)iferida, A)ctualizada.
       ConnRec2.UpdateBatch adAffectAll
    End If
    ConnRec2.Close
    Conn2.Close
End Sub 'ACTUALIZAR_STATUS( ...

'******************************************************
'*
'******************************************************
Private Function SIMBOLO_MONEDA_EXT(Simbolo As Integer) As String
    Select Case Simbolo
           Case 1
                SIMBOLO_MONEDA_EXT = "$"
           Case 2
                SIMBOLO_MONEDA_EXT = "EU"
           Case 3
                SIMBOLO_MONEDA_EXT = "CHF"
           Case Else
                SIMBOLO_MONEDA_EXT = ""
    End Select
End Function

'****************************************************
'*
'*****************************************************
Private Sub AVANZAR_LINEAS(ByVal Contador As Integer, ByVal ContadorRequisitos As Integer)
    'Tope_Lineas_Detail = 22  '<- originalmente asi.=== Atributo publico y constante.
    'Tope_Lineas_Detail = 18  '<- accidentalmente asi.
    'Printer.Font.Size = 10    '*
    If (ContadorRequisitos <= 0) Then
        Ajuste = 1
    Else
        Ajuste = Int(0.2 * ContadorRequisitos)   ' 20% del contador requisitos.
    End If
    For i = 1 To (Tope_Lineas_Detail - Contador + Ajuste)
        Printer.Print
    Next i
End Sub

'****************************************************
'*
'****************************************************
Private Sub IMPRIMIR_REQ_LEGALES(ByRef Requisitos() As String, ByRef Contador As Integer, ByRef ContadorRequisitos As Integer)
   Dim K1 As Integer
   Dim K2 As Integer
   Dim FormFeed As String
   FormFeed = Chr(12)
   K1 = 5
   K2 = 4
   Printer.Print
   Printer.Font.Size = 8
   Printer.Font.Bold = True
   Printer.Print
   Printer.Font.Underline = True
   Printer.Print Tab(K1 + 13); "REQUISITO(S) EXIGIDO(S):"
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Contador = Contador + 1
   '**********************************************************************************
   If (Contador = (Tope_Lineas_Detail - 1)) Then
       'MsgBox "(dentro de IMPRIMIR_REQ_LEGALES)Avanzar a la siguiente pagina; reimprimir encabezado & on,..."
       Printer.NewPage
       Contador = 0
       ContadorRequisitos = 0
       IMPRIMIR_ENCABEZADO K1, K2
       Printer.Font.Size = 8
       Printer.Font.Bold = True
       Printer.Print
       Printer.Font.Underline = True
       Printer.Print Tab(K1 + 13); "REQUISITO(S) EXIGIDO(S):"
       Printer.Font.Bold = False
       Printer.Font.Underline = False
   End If
   '**********************************************************************************
   i = -1
   Termine = False
   While (Not Termine And i <= UBound(Requisitos))
         i = i + 1
         If (Not IsNull(Requisitos(i)) And Mid(Requisitos(i), 1, 1) <> "") Then
            ' MsgBox Requisitos(i)
            Printer.Print Tab(K1 + 13); Requisitos(i)
            Contador = Contador + 1
            ContadorRequisitos = ContadorRequisitos + 1
        Else
            Termine = True
        End If
   Wend
   Printer.Font.Size = 10  '* Ver IMPRIMIR_ROUTINE.
   Printer.Font.Bold = False
   '*
End Sub
'****************************************************
'*
'****************************************************
Private Sub IMPRIMIR_TOTALES(ByVal K2 As Integer, ByVal Simbolo_Moneda As String, ByVal Total_Orden As Double)
    Total_OrdenS = ""
    If Not IsNull(Total_Orden) Then
       Total_OrdenS = Simbolo_Moneda & " " & Format(Total_Orden, "###,###,##0.00")
    End If
    'Printer.Font.Bold = True
    'Printer.Print Tab(K2 + 115); Spc(14 - Len(Total_OrdenS)); Total_OrdenS
    '***
    'Printer.Font.Size = 9
    'Printer.Font.Bold = True
    '*****
    Printer.Font.Size = 10
    'Printer.Font.Bold = True  ' Remind: Font bold solo funcina con Font Size <= 9........
    Printer.Print Tab(K2 + 113); Spc(14 - Len(Total_OrdenS)); Total_OrdenS
    Printer.Font.Bold = False
End Sub

'****************************************************
'*
'****************************************************
Private Sub NEXT_FORM()
  'null
End Sub

'************************************************************
'************************************************************
Private Sub Imprimir_Orden_Compra_vEPSON_LX300_old()
 'No implatada: Windows 95 y el driver de esta
 '              impresora no se interrelacionan
 '              correctamente. ?????
End Sub  'IMPRIMIR_ORDEN_COMPRA_vEPSON_LX300
'****************************************************
'*   Imprimir informacion MASTER del O.C.
'****************************************************
Private Sub IMPRIMIR_ENCABEZADO(ByVal K1 As Integer, ByVal K2 As Integer)
   'Dim FormFeed As String
   'FormFeed = Chr(12)
   Printer.Font.Size = 10  '* Ver IMPRIMIR_ROUTINE.
   Printer.Font.Bold = False
   For i = 1 To 5
       Printer.Print
   Next i
   Printer.Print
   Printer.Print
   Printer.Print Tab(K1); Nombre; Tab(K1 + 115); Format(Fecha_Orden, "DD     MM    YYYY")
   Printer.Print Tab(K1); Direccion1; ' Tab(K1 + 85); Orden_Servicio
   Printer.Print Tab(K1); Direccion2; Tab(K1 + 85); Orden_Servicio
   Printer.Print Tab(K1); Direccion3 & " " & Telf1; ' Tab(K1 + 85); Orden_Compra
   Printer.Print Tab(K1); "Rif: " & Rif; "   Nit: " & Nit; Tab(K1 + 85); Orden_Compra
   Printer.Print
   Printer.Print
   If Cond_Pago = "CREDITO" Then
      Printer.Print Tab(K1 + 2); Cod_Prov; Tab(K1 + 30); Req_No1; Tab(K1 + 57); Fecha_Requisicion1; Tab(K1 + 78); Cond_Pago; _
                                           Tab(K1 + 89); Str(Plazo) + " DIAS"; Tab(K1 + 110); Requerimiento
   Else
      Printer.Print Tab(K1 + 2); Cod_Prov; Tab(K1 + 30); Req_No1; Tab(K1 + 57); Fecha_Requisicion1; Tab(K1 + 82); Cond_Pago; _
                                           Tab(K1 + 110); Requerimiento
  End If
  'Printer.Print Tab(K1 + 30); Req_No2; Tab(K1 + 57); Fecha_Requisicion2
  Printer.Print Tab(K1 + 30); Req_No2; Tab(K1 + 57); Fecha_Requisicion2; Tab(K1 + 79); Anticipo
End Sub
'****************************************************
'****************************************************
Private Sub ACTUALIZAR_HASTA_NEXT_ORDEN(ByVal Numero_Orden As String, ByVal NextNroFormulario As String)
    Dim Conn As ADODB.Connection
    Dim ConnRec As ADODB.Recordset
    Set Conn = New ADODB.Connection
    Set ConnRec = New ADODB.Recordset
    Conn.ConnectionString = "DSN=" & OrigenDatos & _
                            ";UID=" & UserName & _
                            ";PWD=" & Clave
    Conn.Open
    Cadena_SQL = "select C2_NUMERO_ORDEN, C2_HASTA_NRO_ORDEN from COMPRAS02_DAT WHERE C2_NUMERO_ORDEN = '" & Numero_Orden & "'"
    ConnRec.Open Cadena_SQL, Conn, adOpenStatic, adLockOptimistic
    If Not ConnRec.EOF Then
       ConnRec("C2_HASTA_NRO_ORDEN") = NextNroFormulario
       ConnRec.UpdateBatch adAffectAll
    End If
    ConnRec.Close
    Conn.Close
End Sub

'****************************************************
'*     Imprimir_Orden_Compra_vEPSON_LX300()
'****************************************************
Private Sub Imprimir_Orden_Compra_vEPSON_LX300()
 Dim ContFormularios As Integer
 Dim NextNroOrden As String
 Dim Ajustar_Saltos As Boolean
 Dim Termine As Boolean
 Dim Imprime_Total As Boolean
 Dim Contador As Integer
 Dim ContadorRequisitos As Integer
 Dim Cont_Renglon As Integer
 Dim Total_Orden As Double
 Dim FormFeed As String
 FormFeed = Chr(12)
 MaxReqLegales = 10
 Dim Requisitos(40) As String
 K1 = 5
 K2 = 4
 ContFormularios = 0
 Contador = 0
 ContadorRequisitos = 0
 Cont_Renglon = 0
 If OPEN_DATABASE Then
  'MsgBox "AQUICA dentro del metodo Anticipo=Imprimir_Orden_Compra_vEPSON_LX300()"
  Ajustar_Saltos = False
  'MsgBox "B.D. Ok!"
  While Not Reg.EOF
     'Get Master Data
     Nombre = Reg("C1_NOMBRE")
     Direccion1 = Reg("C1_DIRECCION1")
     If IsNull(Direccion1) Then
        Direccion1 = " "
     End If
     Direccion2 = Reg("C1_DIRECCION2")
     If IsNull(Direccion2) Then
        Direccion2 = " "
     End If
     Direccion3 = Reg("C1_DIRECCION3")
     If IsNull(Direccion3) Then
        Direccion3 = " "
     End If
     Rif = Reg("C1_RIF")
     If IsNull(Rif) Then
        Rif = " "
     End If
     Nit = Reg("C1_NIT")
     If IsNull(Nit) Then
        Nit = " "
     End If
     Telf1 = Reg("C1_TELEFONO1")
     If IsNull(Telf1) Then
        Telf1 = " "
     Else
        Telf1 = ". Telf.: " & Telf1
     End If
     Fax1 = Reg("C1_FAX1")
     Fecha_Orden = Reg("C2_FECHA_ORDEN")
     Orden_Compra = Reg("Orden_Compra")
     Orden_Servicio = Reg("Orden_Servicio")
     Cod_Prov = Reg("C1_CODIGO_PROV")
     Req_No1 = Reg("C2_REQUISICION_NO1")
     Req_No2 = Reg("C2_REQUISICION_NO2")
     If IsNull(Req_No2) Then
        Req_No2 = " "
     End If
     Fecha_Requisicion1 = Reg("C2_FECHA_REQUISICION1")
     Fecha_Requisicion2 = Reg("C2_FECHA_REQUISICION2")
     If IsNull(Fecha_Requisicion2) Then
        Fecha_Requisicion2 = " "
     End If
     Cond_Pago = Reg("C2_CONDICION_PAGO")
     If Cond_Pago = "CREDITO" Then
        Plazo = Reg("C2_PLAZO")
     Else
        Anticipo = Reg("C2_ANTICIPO")
        'MsgBox "Anticipo antes=" + Str(Anticipo) + "****************************************"
        If Not IsNull(Str(Anticipo)) And Anticipo > 0# Then
           Anticipo = "ANTICIPO: " + Str(Anticipo) + "%"
        Else
           Anticipo = ""
        End If
     End If
     'MsgBox "Anticipo despues=" + Str(Anticipo) + "****************************************"
     Requerimiento = Reg("C2_REQUERIMIENTO")
     Numero_Orden = Reg("C2_NUMERO_ORDEN")
     Simbolo_Moneda = SIMBOLO_MONEDA_EXT(Reg("C2_MONEDA_EXTRANJERA"))
     Imprime_Total = False
     If Reg("C2_CON_TOTAL") = "S" Then
        Imprime_Total = True
     End If
     '** Imprimir:
     '   --------
     If Reg("C2_STATUS") = "I" Then
           If MsgBox("Orden No. " + Numero_Orden + ", fue impresa, Deseas Continuar ?", vbYesNo + vbQuestion + vbDefaultButton1, "ATENCION") = vbNo Then
              CLOSE_DATABASE
              Exit Sub
              CB_Cancelar_Click '<-Exit Sub
            End If '... MsgBix
        End If '... Status = I
        'For i = 1 To 6
        For i = 1 To 5
            Printer.Print
        Next i
        'Printer.Print "      RC-CO4.6-02"; Tab(K1 + 85); "N.I.T. "; Nit_Empresa
        Printer.Print
        Printer.Print
        Printer.Print Tab(K1); Nombre; Tab(K1 + 115); Format(Fecha_Orden, "DD     MM    YYYY")
        Printer.Print Tab(K1); Direccion1; ' Tab(K1 + 85); Orden_Servicio
        Printer.Print Tab(K1); Direccion2; Tab(K1 + 85); Orden_Servicio
        Printer.Print Tab(K1); Direccion3 & " " & Telf1; ' Tab(K1 + 85); Orden_Compra
        Printer.Print Tab(K1); "Rif: " & Rif; "   Nit: " & Nit; Tab(K1 + 85); Orden_Compra
        Printer.Print
        Printer.Print
        If Cond_Pago = "CREDITO" Then
           Printer.Print Tab(K1 + 2); Cod_Prov; Tab(K1 + 30); Req_No1; Tab(K1 + 57); Fecha_Requisicion1; Tab(K1 + 78); Cond_Pago; _
                         Tab(K1 + 89); Str(Plazo) + " DIAS"; Tab(K1 + 110); Requerimiento
        Else
           Printer.Print Tab(K1 + 2); Cod_Prov; Tab(K1 + 30); Req_No1; Tab(K1 + 57); Fecha_Requisicion1; Tab(K1 + 82); Cond_Pago; _
                         Tab(K1 + 110); Requerimiento
        End If
        'MsgBox "Anticipo Nodo Cabeza=" + Str(Anticipo) + "****************************************"
        Printer.Print Tab(K1 + 30); Req_No2; Tab(K1 + 57); Fecha_Requisicion2; Tab(K1 + 79); Anticipo
        '*** Modificar status *****
        'Reg("C2_STATUS") = "I"
        'Reg.UpdateBatch adAffectAll
        ACTUALIZAR_STATUS (Numero_Orden)
        '**************************
       'Get DATA detail:
        Printer.Print
        Termine = False
        Total_Orden = 0
        While (Not Reg.EOF) And (Not Termine)
            Cantidad = Reg("C3_CANTIDAD")
            Diferencia = Cantidad - Fix(Cantidad)
            If Diferencia = 0 Then
               CantidadS = Str(Cantidad)
            Else
               CantidadS = Format(Cantidad, "###,#0.00")
            End If
            Cod_Item = Reg("C3_CODIGO_ITEM")
            If IsNull(Cod_Item) Then
               Cod_Item = " "
            End If
            Descripcion = Reg("C3_DESCRIPCION")
            If IsNull(Descripcion) Then
               Descripcion = ""
            End If
            Precio_Unid = Reg("C3_PRECIO_UNID")
            Precio_UnidS = ""
            If Not IsNull(Precio_Unid) Then
                Precio_UnidS = Simbolo_Moneda & " " & Format(Precio_Unid, "###,###,##0.00")
            End If
            Total_Item = Reg("C3_TOTAL_ITEM")
            Total_ItemS = ""
            '**
            If Not IsNull(Total_Item) Then
                Total_ItemS = Simbolo_Moneda & " " & Format(Total_Item, "###,###,##0.00")
                Cont_Renglon = Cont_Renglon + 1
                Total_Orden = Total_Orden + Total_Item
                Printer.Print Tab(K2); Format(Cont_Renglon, "00"); Tab(K2 + 4); Spc(6 - Len(CantidadS)); CantidadS; Tab(K2 + 14); Descripcion; _
                              Tab(K2 + 94); Spc(14 - Len(Precio_UnidS)); Precio_UnidS; Tab(K2 + 114); Spc(14 - Len(Total_ItemS)); Total_ItemS
            Else '* no imprime el No. de renglon
                Printer.Print Tab(K2 + 4); Spc(6 - Len(CantidadS)); CantidadS; Tab(K2 + 14); Descripcion; _
                              Tab(K2 + 94); Spc(14 - Len(Precio_UnidS)); Precio_UnidS; Tab(K2 + 114); Spc(14 - Len(Total_ItemS)); Total_ItemS
            End If  ' No es nulo Total_Item
            '**
            '*****************(-*-)AVANZAR AL SIGUIENTE FORMULARIO(-*-)**************************************
            Contador = Contador + 1
            'If (Contador = (Tope_Lineas_Detail - 2)) Then
            If (Contador = Tope_Lineas_Detail) Then
                MsgBox " (Dentro del metodo IMPRIMIR_OC_LX300+) Avanzar a la siguiente pagina; reimprimir encabezado & on,..."
                Printer.Print
                SubTotalS = Simbolo_Moneda & " " & Format(Total_Orden, "###,###,##0.00")
                Printer.Font.Size = 10
                'Printer.Font.Bold = True
                Printer.Print Tab(K2 + 90); " VAN: "; SubTotalS + " ..."
                Printer.Font.Size = 10
                Printer.Font.Bold = False
                Printer.NewPage                           ' *-))
                '* Actualizar campo: C2_HASTA_NRO_ORDEN <- C2_NUMERO_ORDEN + 1
                'MsgBox "Next Nro Orden=" + Format(Val(Numero_Orden) + 1, FormatoNroOrden)
                ContFormularios = ContFormaularios + 1
                NextNroOrden = Format(Val(Numero_Orden) + ContFormularios, FormatoNroOrden)
                ACTUALIZAR_HASTA_NEXT_ORDEN Numero_Orden, NextNroOrden
                IMPRIMIR_ENCABEZADO K1, K2
                Printer.Print
                Printer.Print
                Printer.Font.Size = 10
                'Printer.Font.Bold = True
                Printer.Print Tab(K2 + 14); "... VIENEN: "; SubTotalS
                Printer.Print
                Printer.Font.Size = 10
                Printer.Font.Bold = False
                Contador = 4
                ContadorRequisitos = 0
            End If
            '*************************************************************
            Reg.MoveNext
            If (Not Reg.EOF) Then
                If (Reg("C2_NUMERO_ORDEN") <> Numero_Orden) Then
                     If (POSEE_REQ_LEGALES(Numero_Orden)) Then
                         LOAD_REQUISITOS_LEGALES Numero_Orden, Requisitos
                         IMPRIMIR_REQ_LEGALES Requisitos, Contador, ContadorRequisitos
                         'Contador = Contador + 2   '* Ajuste
                         AVANZAR_LINEAS Contador, ContadorRequisitos   '* Revisar Contador
                     Else
                         'Contador = Contador + 2   '* Ajuste
                         AVANZAR_LINEAS Contador, ContadorRequisitos
                     End If '  Posee Req Legeles.
                     If Imprime_Total Then
                        IMPRIMIR_TOTALES K2, Simbolo_Moneda, Total_Orden
                     Else
                        Printer.Print
                     End If
                     For Saltos = 1 To 8
                         Printer.Print
                     Next Saltos
                     Contador = 0
                     Cont_Renglon = 0
                     Termine = True
                End If 'If interno: NUMERO_ORDEN ...
            Else
                If (POSEE_REQ_LEGALES(Numero_Orden)) Then
                    LOAD_REQUISITOS_LEGALES Numero_Orden, Requisitos
                    IMPRIMIR_REQ_LEGALES Requisitos, Contador, ContadorRequisitos
                    'Contador = Contador + 2  '* Ajuste
                    AVANZAR_LINEAS Contador, ContadorRequisitos
                Else
                    AVANZAR_LINEAS Contador, ContadorRequisitos
                End If
                If Imprime_Total Then
                   IMPRIMIR_TOTALES K2, Simbolo_Moneda, Total_Orden
                Else
                   Printer.Print
                End If
            End If  'If Not Reg.EOF ...
        Wend 'Get Detail Data
        '**********************************************************************************************************
        ' * Gestionar Impresion REQUISITOS LEGALES DE LA ORDEN: ***************************************************
        'LOAD_REQUISITOS_LEGALES Numero_Orden, Requisitos
        'IMPRIMIR_REQ_LEGALES Requisitos
        'MsgBox "Array Length Req Legales = " + Requisitos.Length
        'Printer.Print
        'Printer.Font.Size = 7
        'Printer.Font.Bold = True
        'i = -1
        'Termine = False
        'While (Not Termine And i <= UBound(Requisitos))
        '   i = i + 1
        '   If (Not IsNull(Requisitos(i)) And Mid(Requisitos(i), 1, 1) <> "") Then
        '      'MsgBox Requisitos(i)
        '      Printer.Print Tab(K1 + 10); Requisitos(i)
        '  Else
        '      Termine = True
        '  End If
       'Wend
       'Printer.Font.Size = 10  '* Ver IMPRIMIR_ROUTINE.
       'Printer.Font.Bold = False
       '*
       '**********************************************************************************************************
  Wend ' Externo
 End If
 Printer.EndDoc
 CLOSE_DATABASE
 Unload Me
End Sub  'IMPRIMIR_ORDEN_COMPRA_vEPSON_LX300+


'****************************************************
'} End CUERPO PRINCIPAL
'****************************************************

'*-----------------------------------------------------------------------
'* Caja de dialogo estandard para rutinas de impresion
'* Microsoft.
'* NOTA: esta rutina de impresion esta suprimida
'*       para Windows 95. Su uso genera un efecto indeseable
'*       en la linea de comando: "CommonDialog1.ShowPrinter " ???? ....
'*       12 de Agosto del a�o 2003.
'*-----------------------------------------------------------------------
Private Sub IMPRIMIR_ROUTINE()
  Dim J As Integer
  Dim Desde As Integer
  ' Valores de impresi�n
  Dim PrimeraPag, �ltimaPag, NumCopias, ImpArchivo, i, T
  ' Si ocurre un error ejecutar ManipularErrorImprimir
  On Error GoTo ManipularErrorImprimir
  ' Generar un error cuando se pulse Cancelar
  CommonDialog1.CancelError = True
  ' Visualizar la caja de di�logo
  CommonDialog1.ShowPrinter
  ' Obtener las propiedades de impresi�n
  PrimeraPag = CommonDialog1.FromPage
  �ltimaPag = CommonDialog1.ToPage
  NumCopias = CommonDialog1.Copies   '<- Esta instruccion no esta funcionando ????
  ImpArchivo = CommonDialog1.Flags And cdlPDPrintToFile
  ' Imprimir
  If ImpArchivo Then
    For i = 1 To NumCopias
      ' Escriba el c�digo para enviar datos a un archivo
      'GENERAR_ARCHIVO
    Next i
  Else  'Dirigir salida a la impresora
    T = NumCopias
    For i = 1 To NumCopias
            Printer.Font.Name = "Draft"
            'Printer.Font.Bold = True
            Printer.Font.Size = 10
            'Imprimir_Orden_Compra_vEPSON_LX810  '**
            Imprimir_Orden_Compra_vEPSON_LX300
    Next i
  End If ' For i ...

SalirImprimir:
  Exit Sub
  
ManipularErrorImprimir:
  ' Manipular el error
  If Err.Number = cdlCancel Then Exit Sub
  MsgBox Err.Description
  Resume SalirImprimir
End Sub 'IMPRIMIR_ROUTINE

'********************EOF ( Print_Orden_Compra )************************



