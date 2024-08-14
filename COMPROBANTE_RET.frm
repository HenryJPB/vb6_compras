VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form COMPROB_RET 
   Caption         =   "COMPROBANTE DE RETENCION DEL I.V.A."
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Boton_Catalog_Prov 
      Height          =   375
      Left            =   6000
      Picture         =   "COMPROBANTE_RET.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Cmd_Modificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
      Height          =   420
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton Cmd_Delete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eliminar"
      Height          =   420
      Left            =   3720
      TabIndex        =   30
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton Cmd_Buscar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar"
      Height          =   420
      Left            =   1920
      TabIndex        =   29
      Top             =   120
      Width           =   945
   End
   Begin VB.CommandButton Cmd_Agregar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
      Height          =   420
      Left            =   1080
      TabIndex        =   28
      Top             =   120
      Width           =   825
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   27
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Actualizar 
      Caption         =   "Actualiza"
      Height          =   495
      Left            =   2520
      TabIndex        =   26
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Rif_Suj 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Nombre_Suj 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Periodo_Fiscal 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Fecha_Comp 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Num_Comp 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Periodo_Comp 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Modificar"
      Height          =   300
      Left            =   2280
      TabIndex        =   19
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Ag&regar"
      Height          =   300
      Left            =   1200
      TabIndex        =   18
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_Cerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cerrar"
      Height          =   420
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Reno&var"
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   3360
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   3600
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "A&ctualizar"
      Height          =   300
      Left            =   2520
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   5760
      Picture         =   "COMPROBANTE_RET.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   5400
      Picture         =   "COMPROBANTE_RET.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   720
      Picture         =   "COMPROBANTE_RET.frx":0AC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   360
      Picture         =   "COMPROBANTE_RET.frx":0E08
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ForeColor       =   4194368
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Detalle Comprobante de Retencion"
      Height          =   255
      Left            =   1800
      TabIndex        =   32
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Rif Contribuyente?:"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre Sujeto Retencion?: "
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Periodo Fiscal?:"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha del Comprobante?:"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Numero del Comprobante?:"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo del comprobante?:"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   3240
      Y2              =   5400
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   120
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "COMPROB_RET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'*--------------------------------------------------*
'*  SISTEMA DE COMPRAS                              *
'*  Modulo RETENCION IVA                            *
'*  Autor: Henry J. Pulgar B.                       *
'*  Creado: 14 de Abril del año 2003.               *
'*  Actualizado: 12 de Mayo del año 2003.           *
'*--------------------------------------------------*
'****************************************************
Dim Conecc_Maestro As Connection
Public Reg_Maestro As Recordset
Public Factor_Iva As Double
Public Cont_Filas_Grid As Long
Dim AddNew_MasterFlag As Boolean
Dim Edit_MasterFlag As Boolean
'.......................................
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Boton_Catalog_Prov_Click()
   If AddNew_MasterFlag Or Edit_MasterFlag Then
      CATALOGO_PROV.Show
   End If
End Sub

'*********** RUTINAS GENERALES : *****************
Private Sub Form_Load()
  ' Iniciar Variables Globales:
  ' ---------------------------
  Factor_Iva = 16#
  AddNew_MasterFlag = False
  Edit_MasterFlag = False
  '...........
  Set_Main_Buttons (True)
  Set Conecc_Maestro = New Connection
  Conecc_Maestro.CursorLocation = adUseClient
  Conecc_Maestro.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=OPS$DESCOM02;pwd=OPS$DESCOM02;"
  Set Reg_Maestro = New Recordset
  Reg_Maestro.Open "select C4_PERIODO_COMP," & _
                           "C4_NUMERO_COMP," & _
                           "C4_FECHA_COMP," & _
                           "C4_PERIODO_FISCAL," & _
                           "C4_NOMBRE_SUJ," & _
                           "C4_RIF_SUJ " & _
                           "from COMPRAS04_DAT " & _
                           "order by C4_PERIODO_COMP, C4_NUMERO_COMP", Conecc_Maestro, adOpenStatic, adLockOptimistic
  If Not Reg_Maestro.EOF Then
     LOAD_DATOS_USUARIO
  End If
  LOAD_GRID
End Sub

Private Sub LOAD_GRID()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDASQL;dsn=DESICA806;uid=OPS$DESCOM02;pwd=OPS$DESCOM02;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select C5_PERIODO_COMP Periodo_Comp," & _
                            "C5_NUMERO_COMP No_Comp," & _
                            "C5_OPER_NO Oper_No," & _
                            "C5_FECHA_FACT Fecha_Fact," & _
                            "C5_NO_FACT No_Factura," & _
                            "C5_NO_CONTROL_FACT No_Control_Fact," & _
                            "C5_NO_ND No_Nota_Db," & _
                            "C5_NO_NC No_Nota_Cr," & _
                            "C5_TIPO_TRANS Tipo_Trans," & _
                            "C5_NO_FACT_AFECT Fact_Afectada," & _
                            "C5_MONTO_COMPRA1 Monto_Compra," & _
                            "C5_MONTO_COMPRA2 Compra_SDer_Cr_Fiscal," & _
                            "C5_BASE_IMP Base_Imp," & _
                            "C5_ALICUOTA Alicuota," & _
                            "C5_IVA Iva," & _
                            "C5_IVA_RET Iva_Retenido from COMPRAS05_DAT " & _
                            "where C5_NUMERO_COMP = '" & Num_Comp.Text & "' " & _
                            "order by C5_OPER_NO", db, adOpenStatic, adLockOptimistic
  Set grdDataGrid.DataSource = adoPrimaryRS
  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  'cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Set_Main_Buttons(bVal As Boolean)
  Cmd_Agregar.Visible = bVal
  Cmd_Buscar.Visible = bVal
  Cmd_Actualizar.Visible = Not bVal
  Cmd_Cancel.Visible = Not bVal
  Cmd_Modificar.Visible = bVal
  Cmd_Delete.Visible = bVal
  Cmd_Cerrar.Visible = bVal
  cmdFirst.Visible = bVal
  cmdNext.Visible = bVal
  cmdPrevious.Visible = bVal
  cmdLast.Visible = bVal
End Sub

Private Sub SetGrdButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = bVal
  cmdCancel.Visible = bVal
  cmdDelete.Visible = bVal
  'cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  'cmdNext.Enabled = bVal
  'cmdFirst.Enabled = bVal
  'cmdLast.Enabled = bVal
  'cmdPrevious.Enabled = bVal
End Sub

Function ACEPTAR(MsgTxt As String)
  Dim Mensaje, Botones, Titulo, Respuesta
  Mensaje = MsgTxt
  Botones = vbYesNo + vbQuestion + vbDefaultButton2
  Titulo = "Confirmar Eliminar"
  Respuesta = MsgBox(Mensaje, Botones, Titulo)
  If Respuesta = vbYes Then
     ACEPTAR = True
  Else
     ACEPTAR = False
  End If
End Function

'********************************************************
'* Conjunto de rutinas para controlar datos Regs. Maestro
'********************************************************
'
Private Sub Nombre_Suj_Validate(Cancel As Boolean)
  Nombre_Suj.Text = UCase(Nombre_Suj.Text)
End Sub

Private Sub Periodo_Comp_Validate(Cancel As Boolean)
  Periodo_Comp.Text = Format(Periodo_Comp, "YYYY-MM")
End Sub

Private Sub Periodo_Fiscal_Validate(Cancel As Boolean)
  Periodo_Fiscal.Text = Format(Periodo_Fiscal, "YYYY-MM")
End Sub

Private Sub INICIAR_MAESTRO_FIELDS()
 Num_Comp.Text = Num_Comp.Text + 1
 Fecha_Comp.Text = Format(Now, "DD-MM-YYYY")
 Periodo_Comp.Text = Format(Fecha_Comp.Text, "YYYY-MM")
 Periodo_Fiscal.Text = Format(Periodo_Comp.Text, "YYYY-MM")
 Nombre_Suj.Text = ""
 Rif_Suj.Text = ""
End Sub

Private Sub BLANK_MAESTRO_FIELDS()
 Num_Comp.Text = ""
 Fecha_Comp.Text = ""
 Periodo_Comp.Text = ""
 Periodo_Fiscal.Text = ""
 Nombre_Suj.Text = ""
 Rif_Suj.Text = ""
End Sub

Private Sub Cmd_Actualizar_Click()
  On Error GoTo Act_Error
  Set_Main_Buttons (True)
  If AddNew_MasterFlag Then
     Reg_Maestro.AddNew
     SET_DATOS_USUARIO
     Reg_Maestro.UpdateBatch adAffectAll
     LOAD_GRID
     AddNew_MasterFlag = False
  ElseIf Edit_MasterFlag Then
     Fecha_Comp.BackColor = &HFFFFFF
     Fecha_Comp.Enabled = True
     Num_Comp.BackColor = &HFFFFFF
     Num_Comp.Enabled = True
     SET_DATOS_USUARIO
     Reg_Maestro.UpdateBatch adAffectAll
     Edit_MasterFlag = False
  End If
  SetGrdButtons (True)
  Exit Sub
Act_Error:
  MsgBox ("ERROR: al actualizar los datos de la tabla")
End Sub

Private Sub SET_DATOS_USUARIO()
 Reg_Maestro("C4_PERIODO_COMP") = Periodo_Comp.Text
 Reg_Maestro("C4_NUMERO_COMP") = Num_Comp.Text
 Reg_Maestro("C4_FECHA_COMP") = Fecha_Comp.Text
 Reg_Maestro("C4_PERIODO_FISCAL") = Periodo_Fiscal.Text
 If Not IsNull(Nombre_Suj.Text) Then
    Reg_Maestro("C4_NOMBRE_SUJ") = Nombre_Suj.Text
 End If
 If Not IsNull(Rif_Suj.Text) Then
    Reg_Maestro("C4_RIF_SUJ") = Rif_Suj.Text
 End If
End Sub

Private Sub LOAD_DATOS_USUARIO()
 Periodo_Comp.Text = Format(Reg_Maestro("C4_PERIODO_COMP"), "YYYY-MM")
 Num_Comp.Text = Reg_Maestro("C4_NUMERO_COMP")
 Fecha_Comp.Text = Reg_Maestro("C4_FECHA_COMP")
 Periodo_Fiscal.Text = Format(Reg_Maestro("C4_PERIODO_FISCAL"), "YYYY-MM")
 If Not IsNull(Reg_Maestro("C4_NOMBRE_SUJ")) Then
    Nombre_Suj.Text = Reg_Maestro("C4_NOMBRE_SUJ")
 Else
    Nombre_Suj.Text = ""
 End If
 If Not IsNull(Reg_Maestro("C4_RIF_SUJ")) Then
    Rif_Suj.Text = Reg_Maestro("C4_RIF_SUJ")
 Else
    Rif_Suj.Text = ""
 End If
End Sub

Private Sub LOAD_DATOS_USUARIO_OLD()
 Periodo_Comp.Text = Format(Reg_Maestro("C4_PERIODO_COMP"), "YYYY-MM")
 Num_Comp.Text = Reg_Maestro("C4_NUMERO_COMP")
 Fecha_Comp.Text = Reg_Maestro("C4_FECHA_COMP")
 Periodo_Fiscal.Text = Reg_Maestro("C4_PERIODO_FISCAL")
 Nombre_Suj.Text = Reg_Maestro("C4_NOMBRE_SUJ")
 Rif_Suj.Text = Reg_Maestro("C4_RIF_SUJ")
End Sub

Private Sub Cmd_Agregar_Click()
 On Error GoTo Add_Error
 '
 AddNew_MasterFlag = True
 Set_Main_Buttons (False)
 SetGrdButtons (False)
 If Not Reg_Maestro.EOF Then
    Reg_Maestro.MoveLast
    'MsgBox "Ado. Status= " & adoPrimaryRS.Status
    'If adoPrimaryRS.State <> 1 Then  ????
    '    adoPrimaryRS.Close ????
    'End If
    'grdDataGrid.ClearFields  ???
    'cmdRefresh_Click ?????
 End If
 LOAD_DATOS_USUARIO
 INICIAR_MAESTRO_FIELDS
 LOAD_GRID
 Periodo_Comp.SetFocus
 Exit Sub
Add_Error:
    Beep
    MsgBox ("ATENCION: B.D. vacia o error al insertar un nuevo registro")
End Sub
Private Sub Cmd_Modificar_Click()
On Error GoTo Add_Error
 '
 Edit_MasterFlag = True
 Fecha_Comp.Enabled = False
 Fecha_Comp.BackColor = &H80FFFF
 Num_Comp.Enabled = False
 Num_Comp.BackColor = &H80FFFF
 Set_Main_Buttons (False)
 SetGrdButtons (False)
 
 'Dessactivar campos claves
Exit Sub
Add_Error:
    Beep
    MsgBox ("ATENCION: Error al modificar el registro.")
End Sub

Private Sub Cmd_Cerrar_Click()
    Reg_Maestro.Close
    Unload Me
End Sub

Private Sub Cmd_Delete_Click()
 If ACEPTAR("Deseas eliminar este registro?") Then
 On Error GoTo DeleteErr
  '--Eliminar registros Detalle--
  While Not adoPrimaryRS.EOF
         With adoPrimaryRS
              .Delete
              If Not .EOF Then
                .MoveNext
              ElseIf .EOF Then
                .MoveLast
              End If
         End With
  Wend
  '--Eliminar registro Maestro
  With Reg_Maestro
    .Delete
    If Not .EOF Then
           .MoveNext
    ElseIf .EOF Then
           .MoveLast
    End If
  End With
  BLANK_MAESTRO_FIELDS
  Exit Sub
DeleteErr:
  MsgBox ("ERROR al Eliminar el registro")
End If 'ACEPTAR ELIMINAR
End Sub

'********************************************************
'* Conjunto de rutinas para controlar la Grid:
'********************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub
Private Sub Cmd_Cancel_Click()
  Set_Main_Buttons (True)
  SetGrdButtons (True)
  BLANK_MAESTRO_FIELDS
  If Not Reg_Maestro.EOF Then
     'MsgBox "B.D. no esta vacia"
     Reg_Maestro.MoveLast
     LOAD_DATOS_USUARIO
     LOAD_GRID
  Else
     MsgBox "B.D. vacia"
  End If
  AddNew_MasterFlag = False
  Edit_MasterFlag = False
  Fecha_Comp.Enabled = True
  Num_Comp.Enabled = True
  Fecha_Comp.BackColor = &HFFFFFF
  Num_Comp.BackColor = &HFFFFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  'lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
  lblStatus.Caption = CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  '
  If Not adoPrimaryRS.EOF Then
    adoPrimaryRS.MoveLast
    Cont_Filas_Grid = adoPrimaryRS("Oper_No")
    Cont_Filas_Grid = Cont_Filas_Grid + 1
  Else
    Cont_Filas_Grid = 1
  End If
  adoPrimaryRS.AddNew
  adoPrimaryRS("PERIODO_COMP") = Periodo_Comp.Text
  adoPrimaryRS("NO_COMP") = Num_Comp.Text
  adoPrimaryRS("ALICUOTA") = Factor_Iva
  adoPrimaryRS("Oper_No") = Cont_Filas_Grid
  grdDataGrid.SetFocus
  cmdEdit_Click
  '
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  If ACEPTAR("Deseas eliminar este registro de la cuadrilla?") = True Then
     On Error GoTo DeleteErr
     With adoPrimaryRS
         .Delete
         If Not .EOF Then
           .MoveNext
         ElseIf .EOF Then
           .MoveLast
         End If
     End With
     Exit Sub
DeleteErr:
     MsgBox Err.Description
 End If ' ACEPTAR
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  
  lblStatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  '
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  ACTUALIZAR_DATOS_GRID
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
  If Not Reg_Maestro.EOF Then
    Reg_Maestro.MoveFirst
    LOAD_DATOS_USUARIO
    LOAD_GRID
  End If
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
   On Error GoTo GoLastError
  If Not Reg_Maestro.EOF Then
    Reg_Maestro.MoveLast
    LOAD_DATOS_USUARIO
    LOAD_GRID
  End If

  mbDataChanged = False
    Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
  '
  If Not Reg_Maestro.EOF Then Reg_Maestro.MoveNext
  If Reg_Maestro.EOF And Reg_Maestro.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    Reg_Maestro.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False
  LOAD_DATOS_USUARIO
  LOAD_GRID
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not Reg_Maestro.BOF Then Reg_Maestro.MovePrevious
  If Reg_Maestro.BOF And Reg_Maestro.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    Reg_Maestro.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False
  LOAD_DATOS_USUARIO
  LOAD_GRID
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub
' Validar datos de la Grid - <<En Periodo de prueba>> ...
Private Sub adoPrimaryRS_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '----
  If Not IsNull(adoPrimaryRS("Tipo_Trans")) Then
     TipoTrans = adoPrimaryRS("Tipo_Trans")
     If (TipoTrans <> "01") And (TipoTrans <> "02") And (TipoTrans <> "03") And (TipoTrans <> "04") Then
         Beep
         Dumy = MsgBox("Tipo de transaccion no definida", , "ATENCION")
     End If
  End If
  '----
  If Not IsNull(adoPrimaryRS("Alicuota")) Then
         Factor_Iva = adoPrimaryRS("Alicuota")
  Else
         adoPrimaryRS("Alicuota") = Factor_Iva
  End If
  If Not IsNull(adoPrimaryRS("Monto_Compra")) And (IsNull(adoPrimaryRS("Iva"))) Then
         MontoCompra = adoPrimaryRS("Monto_Compra")
         BaseImp = (1 / (1 + (Factor_Iva / 100#))) * MontoCompra
         adoPrimaryRS("Base_Imp") = BaseImp
         adoPrimaryRS("Iva") = BaseImp * (Factor_Iva / 100#)
  End If
  '----
  FactorRet = 75   ' 75% / Iva
  If Not IsNull(adoPrimaryRS("Iva")) And IsNull(adoPrimaryRS("Iva_Retenido")) Then
     MontoIva = adoPrimaryRS("Iva")
     adoPrimaryRS("Iva_Retenido") = MontoIva * (FactorRet / 100)
  End If
End Sub
'
 Private Sub ACTUALIZAR_DATOS_GRID()
 '----
  If Not IsNull(adoPrimaryRS("Tipo_Trans")) Then
     TipoTrans = adoPrimaryRS("Tipo_Trans")
     If (TipoTrans <> "01") And (TipoTrans <> "02") And (TipoTrans <> "03") And (TipoTrans <> "04") Then
         Beep
         Dumy = MsgBox("Tipo de transaccion no definida", , "ATENCION")
     End If
  End If
 '----
 'If Not IsNull(adoPrimaryRS("Alicuota")) Then
 '       Factor_Iva = adoPrimaryRS("Alicuota")
 'Else
 '       adoPrimaryRS("Alicuota") = Factor_Iva
 'End If
 'If Not IsNull(adoPrimaryRS("Monto_Compra")) Then
 '       MontoCompra = adoPrimaryRS("Monto_Compra")
 '       BaseImp = (1 / (1 + (Factor_Iva / 100#))) * MontoCompra
 '       adoPrimaryRS("Base_Imp") = BaseImp
 '       adoPrimaryRS("Iva") = BaseImp * (Factor_Iva / 100#)
 'End If
 '----
 ' FactorRet = 75   ' 75% / Base Imponible
 ' If Not IsNull(adoPrimaryRS("Iva")) And IsNull(adoPrimaryRS("Iva_Retenido")) Then
 '    MontoIva = adoPrimaryRS("Iva")
 '    adoPrimaryRS("Iva_Retenido") = MontoIva * (FactorRet / 100)
 ' End If
End Sub
'+++++++++++++++++EOF(COMPROB_RET.frm)++++++++++++++++++++

 
