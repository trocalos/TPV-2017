VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPedido 
   BackColor       =   &H80000013&
   Caption         =   "Venta"
   ClientHeight    =   8988
   ClientLeft      =   192
   ClientTop       =   1116
   ClientWidth     =   13584
   LinkTopic       =   "Form1"
   ScaleHeight     =   8988
   ScaleWidth      =   13584
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNoPuntos 
      Caption         =   "No anotar Puntos"
      Height          =   612
      Left            =   11520
      TabIndex        =   49
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton cmdSiPuntos 
      Caption         =   "Anotar Puntos?"
      Height          =   492
      Left            =   10920
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin MSAdodcLib.Adodc adoIntro1 
      Height          =   375
      Left            =   240
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoIntro1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDescPartePuntos 
      Caption         =   "Descontar puntos"
      Height          =   495
      Left            =   10200
      TabIndex        =   46
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDescTodosPuntos 
      Caption         =   "Descontar todos los puntos"
      Height          =   735
      Left            =   10200
      TabIndex        =   45
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoAnulPuntos 
      Height          =   375
      Left            =   10320
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4466
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Puntos"
      Caption         =   "adoAnulPuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd10 
      BackColor       =   &H0080C0FF&
      Caption         =   "-10%"
      Height          =   735
      Left            =   11040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   39
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoEmail 
      Height          =   495
      Left            =   1440
      Top             =   6360
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5736
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\MasterDIANA.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\MasterDIANA.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "clientas"
      Caption         =   "adoEmail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoIdProfesional 
      Height          =   375
      Left            =   8280
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ClientesProfesionales"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInventario 
      Height          =   495
      Left            =   1560
      Top             =   5400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Inventario"
      Caption         =   "adoInventario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNuevaMarca 
      Caption         =   "Nueva &Marca"
      Height          =   375
      Left            =   4920
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoCBConsulta 
      Height          =   495
      Left            =   240
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6583
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CB"
      Caption         =   "adoCBconsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoCB 
      Height          =   375
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CB"
      Caption         =   "adoCB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtCB 
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "&Nuevo Cliente"
      Height          =   615
      Left            =   1080
      TabIndex        =   27
      Top             =   1200
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoPrecio 
      Height          =   375
      Left            =   8520
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Producto"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoConsVentaProf 
      Height          =   330
      Left            =   9000
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   593
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ClientesProfesionales"
      Caption         =   "adoConsVentaProf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoIntroVentaProf 
      Height          =   330
      Left            =   7080
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4890
      _ExtentY        =   593
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "VentaClienteProfesional"
      Caption         =   "IntroVentaProf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dcomboClienteProf 
      Bindings        =   "frmPedido.frx":0000
      DataField       =   "IdClienProf"
      Height          =   315
      Left            =   5880
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   508
      _Version        =   393216
      ListField       =   "Empresa"
      BoundColumn     =   "IdClienProf"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoConsClienta 
      Height          =   375
      Left            =   480
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "clientas"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Borrar Línea"
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoConsultaPrecio 
      Height          =   375
      Left            =   8520
      Top             =   8280
      Visible         =   0   'False
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoConsultaPrecio"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNuevoProducto 
      Caption         =   "&Nuevo Producto"
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtIDCLIENTA 
      Height          =   285
      Left            =   8640
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoOriseño 
      Height          =   375
      Left            =   4440
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmPedido.frx":001F
      Caption         =   "adoOriseño"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoTotal 
      Height          =   375
      Left            =   3120
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   1
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoTotal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Bindings        =   "frmPedido.frx":00A1
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17378
      _ExtentY        =   5948
      _Version        =   393216
      BackColor       =   -2147483629
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   1
      BackColorFixed  =   -2147483629
      BackColorBkg    =   -2147483629
      BackColorUnpopulated=   -2147483629
      WordWrap        =   -1  'True
      SelectionMode   =   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   10
      _Band(0)._MapCol(0)._Name=   "IdParcial"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "IdSeñorita"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(1)._Alignment=   7
      _Band(0)._MapCol(2)._Name=   "Fecha"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "Hora"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "IdTipo"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Alignment=   7
      _Band(0)._MapCol(5)._Name=   "IdProducto"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(5)._Alignment=   7
      _Band(0)._MapCol(6)._Name=   "IdMarca"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(6)._Alignment=   7
      _Band(0)._MapCol(7)._Name=   "Unidades"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(7)._Alignment=   7
      _Band(0)._MapCol(8)._Name=   "Precio"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(8)._Alignment=   7
      _Band(0)._MapCol(9)._Name=   "IdClienta"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(9)._Alignment=   7
   End
   Begin VB.CommandButton cmdNuevoCliente 
      BackColor       =   &H80000013&
      Caption         =   "&Actualizar"
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdAdelante 
      BackColor       =   &H0000FF00&
      Caption         =   "&Así se habla"
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoIntro 
      Height          =   375
      Left            =   240
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoIntro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoParcial 
      Height          =   375
      Left            =   5400
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4678
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoParcial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPrecio 
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoProducto 
      Height          =   375
      Left            =   -120
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4043
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Producto"
      Caption         =   "adoProducto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dcmbProducto 
      Bindings        =   "frmPedido.frx":00BA
      DataField       =   "IdProducto"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5313
      _ExtentY        =   508
      _Version        =   393216
      ListField       =   "Descripción"
      BoundColumn     =   "IdProducto"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoServicio 
      Height          =   330
      Left            =   2520
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   593
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\TPV 2002\TPV 2002.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\TPV 2002\TPV 2002.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Producto"
      Caption         =   "adoServicio"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoMarca 
      Height          =   375
      Left            =   1560
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Marca"
      Caption         =   "adoMarca"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataList dlstSta 
      Bindings        =   "frmPedido.frx":00D4
      DataField       =   "IdSeñoritas"
      Height          =   1980
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   3302
      _Version        =   393216
      ListField       =   "Nombre"
      BoundColumn     =   "IdSeñoritas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList dlstTipo 
      Bindings        =   "frmPedido.frx":00ED
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1027
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoTipo"
      Height          =   1020
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   1693
      _Version        =   393216
      ListField       =   "Descripción"
      BoundColumn     =   "Descripción"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoTipo 
      Height          =   375
      Left            =   8880
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tipos"
      Caption         =   "adoTipo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtUnidades 
      DataField       =   "Unidades"
      DataSource      =   "datTPV"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Text            =   "1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataListLib.DataCombo dcmbMarca 
      Bindings        =   "frmPedido.frx":0103
      DataField       =   "IdMarca"
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   508
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "Empresa"
      BoundColumn     =   "IdMarca"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoEntradaInventario 
      Height          =   375
      Index           =   0
      Left            =   4800
      Top             =   8040
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5525
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Inventario"
      Caption         =   "adoEntradaInventario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoEntrPuntos 
      Height          =   375
      Index           =   1
      Left            =   10200
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4466
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Puntos"
      Caption         =   "adoEntrPuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoConsultaPuntos 
      Height          =   375
      Left            =   10440
      Top             =   4800
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   656
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tpv 2002\TPV 2002.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Parcial"
      Caption         =   "adoConsultaPuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblAnotaPuntos 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "No se anotan PUNTOS"
      Height          =   372
      Left            =   7200
      TabIndex        =   50
      Top             =   3240
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label lblAvisoNegativo 
      Caption         =   "En NEGATIVO el precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   8040
      TabIndex        =   47
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFechaPuntos 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   10200
      TabIndex        =   44
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblPuntosAcum 
      BackColor       =   &H80000013&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   43
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPuntos 
      BackColor       =   &H80000013&
      Caption         =   "Puntos:"
      Height          =   255
      Left            =   10200
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblTelf 
      BackColor       =   &H80000013&
      Caption         =   "lblTelf"
      Height          =   255
      Left            =   8040
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblMovil 
      BackColor       =   &H80000013&
      Caption         =   "lblMovil"
      Height          =   255
      Left            =   8040
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblAvisoCumple 
      BackColor       =   &H80000013&
      Caption         =   "CUMPLEAÑOS !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3840
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H80000013&
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblIdProfesional 
      BackColor       =   &H80000013&
      Caption         =   "lblIdProfesional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9840
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OLE OLE1 
      Class           =   "Word.Document.8"
      Height          =   2748
      Left            =   11400
      OLEDropAllowed  =   -1  'True
      OleObjectBlob   =   "frmPedido.frx":011A
      SizeMode        =   2  'AutoSize
      SourceDoc       =   "C:\Tpv 2002\cartel.doc"
      TabIndex        =   35
      Top             =   0
      Width           =   10248
   End
   Begin VB.Label lblCumple 
      BackColor       =   &H80000013&
      Caption         =   "Cumpleaños:"
      Height          =   615
      Left            =   3840
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblUnidadesInventario 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2040
      TabIndex        =   33
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblFechaPrecio 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7560
      TabIndex        =   32
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblCB 
      BackColor       =   &H80000013&
      Caption         =   "Código de Barras"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblPorciento 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   6720
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblApellidos 
      BackColor       =   &H80000013&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblNombre 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblPreuProfesional 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7800
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPreu 
      BackColor       =   &H80000013&
      Caption         =   "Precio profesional:"
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblClienta 
      BackColor       =   &H80000013&
      Caption         =   "Clienta nº"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H80000013&
      DataSource      =   "adoParcial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblhora 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDescripción 
      BackColor       =   &H80000013&
      Caption         =   "Unidades"
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDescripción 
      BackColor       =   &H80000013&
      Caption         =   "Precio"
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDescripción 
      BackColor       =   &H80000013&
      Caption         =   "Servicio"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDescripción 
      BackColor       =   &H80000013&
      Caption         =   "Marca"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblDescripción 
      BackColor       =   &H80000013&
      Caption         =   "Producto"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
   Begin VB.Menu mnuListadoS 
      Caption         =   "&Listados"
      Begin VB.Menu mnuCaja 
         Caption         =   "&Caja"
      End
      Begin VB.Menu mnuResumendía 
         Caption         =   "&Resumen del Día"
      End
      Begin VB.Menu mnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentasDía 
         Caption         =   "&Ventas del Día"
      End
      Begin VB.Menu mnuBúsqueda 
         Caption         =   "&Búsqueda de Ventas "
      End
      Begin VB.Menu mnuVisaEfectivo 
         Caption         =   "Ca&mbio de VISA a Efectivo"
      End
      Begin VB.Menu mnulinea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadoSalidaCaja 
         Caption         =   "&Salida de Caja"
      End
   End
   Begin VB.Menu mnuCB 
      Caption         =   "Código de Ba&rras"
      Begin VB.Menu mnuNuevoCB 
         Caption         =   "&Nuevo Código de Barras"
      End
      Begin VB.Menu mnuModificarCB 
         Caption         =   "&Modificar Código de Barras"
      End
   End
   Begin VB.Menu cmdUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnuNuevoProducto 
         Caption         =   "&Nuevo Producto"
      End
      Begin VB.Menu mnuNuevaMarca 
         Caption         =   "Nueva &Marca"
      End
      Begin VB.Menu mnuMatCab 
         Caption         =   "Ma&terial para cabina"
      End
      Begin VB.Menu mnuSalidadeCaja 
         Caption         =   "&Salida de Caja"
      End
      Begin VB.Menu mnulinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReponerTiquet 
         Caption         =   "&Reponer Tiquet"
      End
   End
   Begin VB.Menu mnuClientes 
      Caption         =   "&Clientes"
      Begin VB.Menu mnuParticulares 
         Caption         =   "&Búsqueda de Clientes"
      End
      Begin VB.Menu mnulinea9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComprasClientas 
         Caption         =   "Com&pras realizadas determinado día"
      End
      Begin VB.Menu mnuConsCompraClienta 
         Caption         =   "Compras &realizadas por determinada Clienta"
      End
      Begin VB.Menu mnuLinea10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServicios 
         Caption         =   "&Servicios realizados determinado día"
      End
      Begin VB.Menu mnuServClienta 
         Caption         =   "Servicios realizados a una &Clienta"
      End
      Begin VB.Menu mnuLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAltaClientes 
         Caption         =   "&Altas de clientes"
      End
   End
   Begin VB.Menu mnuClientesProfesionales 
      Caption         =   "Clientes &Profesionales"
      Begin VB.Menu mnuEmpresas 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnuTelClienProf 
         Caption         =   "Búsqueda por &teléfono de C.Prof."
      End
      Begin VB.Menu mnuBusCliProf 
         Caption         =   "&Búsqueda de Clientes Profesionales"
      End
      Begin VB.Menu mnuAlbaranes 
         Caption         =   "&Albaranes pendientes"
      End
      Begin VB.Menu mnuClientasAlbaranVez 
         Caption         =   "Clientes &que se le hizo albarán alguna vez"
      End
   End
   Begin VB.Menu mnuInventario 
      Caption         =   "&Inventario"
      Begin VB.Menu mnuEntradaInventario 
         Caption         =   "Entrada de &Inventario"
      End
      Begin VB.Menu mnuSalidaInventario 
         Caption         =   "&Salida de Inventario"
      End
      Begin VB.Menu mnuActualizaciónInventario 
         Caption         =   "&Actualización de Inventario"
      End
      Begin VB.Menu mnuSubInventario 
         Caption         =   "In&ventario"
      End
      Begin VB.Menu mnuInventarioCasas 
         Caption         =   "Inventario por &casas"
      End
      Begin VB.Menu mnuListadoFaltas 
         Caption         =   "&Listado de faltas"
      End
      Begin VB.Menu mnuline4 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuListAdelantadoClientas 
         Caption         =   "Listado de producto adelantado a clientas"
      End
   End
   Begin VB.Menu mnuPersonal 
      Caption         =   "Persona&l"
      Begin VB.Menu mnuSalidaInventPersonal 
         Caption         =   "Salida de inventario para personal del centro"
      End
      Begin VB.Menu mnuline5 
         Caption         =   "-"
      End
      Begin VB.Menu mnulistAdelantadoSeñoritas 
         Caption         =   "Listado del producto adelantado al personal del centro"
      End
      Begin VB.Menu mnuAlbaTrab 
         Caption         =   "Situación de los albaranes"
      End
      Begin VB.Menu mnuListadoAlbaranTrab 
         Caption         =   "Albaranes pendientes de pago del personal del centro"
      End
   End
   Begin VB.Menu mnublanco3 
      Caption         =   ""
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "&Bloc de Notas"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "&Ayuda"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "A&cerca de  TPV Diana Guill"
      End
   End
   Begin VB.Menu mnuBlanco2 
      Caption         =   ""
   End
   Begin VB.Menu mnuCalculadora 
      Caption         =   "Calcula&dora"
   End
   Begin VB.Menu mnuFlotarium 
      Caption         =   "&Temporizador"
   End
   Begin VB.Menu mnuCodigosInternos 
      Caption         =   "Có&digos Internos"
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resultado As String
Dim Fecha As String
Dim Hora As String
'voy a colocar todo en días para calcular
Dim diahoy As String
Dim diacumple As String









Private Sub cmd10_Click()
If txtPrecio.Text <> "" Then
    lblPreuProfesional.Caption = Format(txtPrecio.Text * 90 / 100, "0.00")
    lblPreu.Visible = True
    lblPreu.Caption = "Precio con descuento:"
    lblPreuProfesional.Visible = True
End If
    

End Sub

Private Sub cmdAdelante_Click()

If dlstTipo.BoundText = "Venta al PROFESIONAL" Then
adoIntroVentaProf.Refresh
adoIntroVentaProf.Recordset.AddNew

adoIntroVentaProf.Recordset.Fields(1).Value = LBLFECHA.Caption
adoIntroVentaProf.Recordset.Fields(2).Value = lblhora.Caption
If dcomboClienteProf.BoundText = "" Then
adoIntroVentaProf.Recordset.Fields(3).Value = 23
Else
adoIntroVentaProf.Recordset.Fields(3).Value = dcomboClienteProf.BoundText
End If
adoIntroVentaProf.Recordset.Update
adoIntroVentaProf.Refresh

End If



'***************************************************
 '22/04/2015 Comentario para intro Puntos (1) por un error (revisar si por ahí lo ha dejado en blanco)


On Error Resume Next
If dlstTipo.BoundText <> "Venta al PROFESIONAL" And txtIDCLIENTA.Text <> "" And lblAnotaPuntos.Caption = "Se anotan puntos" Then
adoEntrPuntos(1).Refresh
adoEntrPuntos(1).Recordset.AddNew

adoEntrPuntos(1).Recordset.Fields(1).Value = txtIDCLIENTA.Text
adoEntrPuntos(1).Recordset.Fields(2).Value = StrConv(LBLFECHA.Caption, vbUpperCase)
adoEntrPuntos(1).Recordset.Fields(3).Value = StrConv(lblhora.Caption, vbUpperCase)
adoEntrPuntos(1).Recordset.Fields(4).Value = Val(Replace(lblTotal.Caption, ",", "."))

adoEntrPuntos(1).Recordset.Fields(5).Value = ((Val(Replace(lblTotal.Caption, ",", ".")) * 0.03) + Val(Replace(lblPuntosAcum.Caption, ",", ".")))


 
 
adoEntrPuntos(1).Recordset.Update
adoEntrPuntos(1).Refresh
  End If
'***************************************************

frmTotal.Show

End Sub



Private Sub cmdDescPartePuntos_Click()
If dlstTipo.BoundText = "Venta al PÚBLICO" Or dlstTipo.BoundText = "Venta de CABINA" Then
txtCB.Text = "9998"

adoProducto.Refresh
      adoProducto.CommandType = adCmdText
      adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción FROM CB INNER JOIN Producto ON CB.IdProducto = Producto.IdProducto WHERE ((CB.CB)=" & txtCB.Text & ");"
      adoProducto.Refresh
      dcmbProducto(0).Text = adoProducto.Recordset.Fields(1).Value

      dcmbProducto(0).Refresh
    
       
       adoMarca.CommandType = adCmdText
       adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa FROM CB INNER JOIN Marca ON CB.IdMarca = Marca.IdMarca WHERE ((CB.CB)=" & txtCB.Text & ");"
       adoMarca.Refresh
       'esto es para que aparezca el primer nombre en la casilla
       
       dcmbMarca(1).Text = adoMarca.Recordset.Fields(1).Value

       dcmbMarca(1).Refresh


'dcmbProducto(0).Text = "PUNTOS"

'dcmbMarca(1).Text = "PUNTOS"

txtPrecio.SetFocus
lblAvisoNegativo.Visible = True
cmdDescTodosPuntos.Enabled = False
End If
If dlstTipo.BoundText = "SERVICIO" Then
cmdDescTodosPuntos.Enabled = False

dcmbProducto(0).Text = "Puntos Servicio"
txtPrecio.SetFocus
lblAvisoNegativo.Visible = True
End If

End Sub

Private Sub cmdDescTodosPuntos_Click()
If dlstTipo.BoundText = "Venta al PÚBLICO" Or dlstTipo.BoundText = "Venta de CABINA" Then
cmdDescTodosPuntos.Enabled = False
txtCB.Text = "9998"


adoProducto.Refresh
      adoProducto.CommandType = adCmdText
      adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción FROM CB INNER JOIN Producto ON CB.IdProducto = Producto.IdProducto WHERE ((CB.CB)=" & txtCB.Text & ");"
      adoProducto.Refresh
      dcmbProducto(0).Text = adoProducto.Recordset.Fields(1).Value

      dcmbProducto(0).Refresh
    
       
       adoMarca.CommandType = adCmdText
       adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa FROM CB INNER JOIN Marca ON CB.IdMarca = Marca.IdMarca WHERE ((CB.CB)=" & txtCB.Text & ");"
       adoMarca.Refresh
       'esto es para que aparezca el primer nombre en la casilla
       
       dcmbMarca(1).Text = adoMarca.Recordset.Fields(1).Value

       dcmbMarca(1).Refresh


txtPrecio.Text = lblPuntosAcum.Caption * -1

cmdOK.SetFocus
End If

If dlstTipo.BoundText = "SERVICIO" Then
cmdDescTodosPuntos.Enabled = False

dcmbProducto(0).Text = "Puntos Servicio"


txtPrecio.Text = lblPuntosAcum.Caption * -1

cmdOK.SetFocus
End If

End Sub

Private Sub cmdEditar_Click()
MousePointer = 11


adoIntro.Refresh
For a = 0 To adoIntro.Recordset.RecordCount - 1
If adoIntro.Recordset.Fields(0).Value = flex.Text Then
adoIntro.Recordset.Delete
Exit For
Else
adoIntro.Recordset.MoveNext
End If
Next a
adoIntro.Refresh
adoIntro.Refresh




  adoParcial.CommandType = adCmdText
 adoParcial.RecordSource = "SELECT  Parcial.idparcial, Señoritas.Nombre, Tipos.Descripción, Parcial.Fecha, Parcial.Hora, Producto.Descripción, Marca.Empresa, Parcial.Unidades, Parcial.Precio, [Parcial]![Unidades]*[Parcial]![Precio] AS Total FROM Tipos INNER JOIN (Señoritas INNER JOIN (Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto) ON Señoritas.IdSeñoritas = Parcial.IdSeñorita) ON Tipos.IdTipos = Parcial.IdTipo WHERE (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)= '" & lblhora.Caption & "' ));"
 adoParcial.Refresh
 
 flex.Refresh
  flex.Visible = True
flex.ColWidth(0) = 0
flex.ColWidth(3) = 0
flex.ColWidth(4) = 0
 flex.Refresh

On Error GoTo errorborradototal

adoTotal.CommandType = adCmdText
adoTotal.RecordSource = "SELECT Sum([Parcial]![Unidades]*[Parcial]![Precio]) AS Subtotal From Parcial GROUP BY Parcial.Fecha, Parcial.Hora HAVING (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)='" & lblhora.Caption & "' ));"
adoTotal.Refresh

lblTotal.Caption = adoTotal.Recordset.Fields(0).Value




dcmbProducto(0).SetFocus
cmdEditar.Visible = False
MousePointer = 1


Exit Sub
errorborradototal:
lblTotal.Caption = ""
 cmdEditar.Visible = False
    cmdAdelante.Visible = False
cmdNuevoCliente.Visible = True
 flex.Visible = False
    mnuSalir.Visible = True
    cmdNuevoCliente.Visible = True
    dcmbProducto(0).SetFocus
    MousePointer = 1




End Sub

Private Sub cmdNoPuntos_Click()
lblAnotaPuntos.Caption = "No se anotan PUNTOS"
lblAnotaPuntos.BackColor = &H80FFFF
cmdNoPuntos.Visible = False
cmdSiPuntos.Visible = True
cmdAdelante.BackColor = &H80FFFF


End Sub

Private Sub cmdNuevaMarca_Click()
frmNuevaMarca.Show
frmNuevaMarca.txtNuevaMarca.SetFocus
End Sub

Private Sub cmdNuevoCliente_Click()
Unload Me
frmPedido.Show


End Sub

Private Sub cmdNuevoProducto_Click()

frmNuevoProducto.Show
frmNuevoProducto.txtNuevoProducto.SetFocus
End Sub

Private Sub cmdOK_Click()


'01/12/15 prueba inventario
On Error Resume Next


If dlstTipo.BoundText <> "SERVICIO" Then
adoEntradaInventario(0).Refresh
        adoEntradaInventario(0).Recordset.AddNew
        adoEntradaInventario(0).Recordset.Fields(1).Value = StrConv(txtCB.Text, vbUpperCase)
            If txtUnidades.Text = "" Then
            adoEntradaInventario(0).Recordset.Fields(2).Value = 0
            Else
            'Aquí pongo el número para que salga negativo
            adoEntradaInventario(0).Recordset.Fields(2).Value = -1 * (StrConv(txtUnidades.Text, vbUpperCase))
            End If
        adoEntradaInventario(0).Recordset.Fields(3).Value = StrConv(LBLFECHA.Caption, vbUpperCase)
        adoEntradaInventario(0).Recordset.Fields(4).Value = StrConv(lblhora.Caption, vbUpperCase)
        adoEntradaInventario(0).Recordset.Fields(5).Value = 5

        adoEntradaInventario(0).Recordset.Update
        adoEntradaInventario(0).Refresh


End If


'************************
 'Modificación para poner puntos a 0
 
 '********02/07/2015 para que no se ponga en negativo
 lblAvisoNegativo.Visible = False
 If (dcmbProducto(0).Text = "PUNTOS" Or dcmbProducto(0).Text = "Puntos Servicio") And Val(txtPrecio.Text) > 0 Then
MsgBox ("el precio no puede ser positivo pues se trata de puntos")
txtPrecio.SetFocus
 MousePointer = 1
Exit Sub
End If
If (dcmbProducto(0).Text = "PUNTOS" Or dcmbProducto(0).Text = "Puntos Servicio") And Val(Replace(txtPrecio.Text, ",", ".")) < (Val(Replace(lblPuntosAcum.Caption, ",", ".")) * -1) Then
MsgBox ("el precio no puede ser más grande que el acumulado pues se trata de puntos")
txtPrecio.SetFocus
 MousePointer = 1
Exit Sub
End If
'*************************************************************
'Modificación para poner puntos a 0

If dlstTipo.BoundText <> "Venta al PROFESIONAL" And txtIDCLIENTA.Text <> "" And (dcmbProducto(0).Text = "PUNTOS" Or dcmbProducto(0).Text = "Puntos Servicio") Then


adoAnulPuntos.Refresh
adoAnulPuntos.Recordset.AddNew

adoAnulPuntos.Recordset.Fields(1).Value = txtIDCLIENTA.Text
adoAnulPuntos.Recordset.Fields(2).Value = StrConv(LBLFECHA.Caption, vbUpperCase)
adoAnulPuntos.Recordset.Fields(3).Value = StrConv(lblhora.Caption, vbUpperCase)
adoAnulPuntos.Recordset.Fields(4).Value = Val(Replace(txtPrecio.Text, ",", "."))
adoAnulPuntos.Recordset.Fields(5).Value = (Val(Replace(txtPrecio.Text, ",", ".")) + Val(Replace(lblPuntosAcum.Caption, ",", ".")))
 
adoAnulPuntos.Recordset.Update
adoAnulPuntos.Refresh

End If
 'fIN Modificación para poner puntos a 0
  '************************

MousePointer = 11
lblFechaPrecio.Caption = ""
lblPreu.Visible = False
lblPreuProfesional.Visible = False
'Tratamiento de errores por falta de datos
On Error GoTo errordato

'Para que no se pueda salir del programa
'si no se ha borrado o acabado todo el ciclo


'Para que solo se pueda hacer un nuevo cliente si está todo limpio


If dcmbProducto(0).BoundText = "" Then
MsgBox ("Introduce el producto")
dcmbProducto(0).SetFocus
 MousePointer = 1
Exit Sub
End If
If dcmbMarca(1).BoundText = "" Then
MsgBox ("Introduce la marca")
dcmbMarca(1).SetFocus
 MousePointer = 1
Exit Sub
End If
If txtUnidades.Text = "" Then
MsgBox ("Introduce las unidades")
txtUnidades.SetFocus
 MousePointer = 1
Exit Sub
End If
If txtPrecio.Text = "" Then
MsgBox ("Introduce el precio")
txtPrecio.SetFocus
 MousePointer = 1
Exit Sub
End If

' ahora le digo que si el código de barras esta vacio que suba todo menos eso
If txtCB.Text = "" Then
adoIntro.Refresh
adoIntro.Refresh
adoIntro.Recordset.AddNew

adoIntro.Recordset.Fields(1).Value = dlstSta.BoundText
adoIntro.Recordset.Fields(2).Value = LBLFECHA.Caption
adoIntro.Recordset.Fields(3).Value = lblhora.Caption
adoIntro.Recordset.Fields(4).Value = dlstTipo.SelectedItem
adoIntro.Recordset.Fields(5).Value = dcmbProducto(0).BoundText

adoIntro.Recordset.Fields(6).Value = dcmbMarca(1).BoundText
adoIntro.Recordset.Fields(7).Value = txtUnidades.Text
    If Val(txtPrecio.Text) = CCur(Format(txtPrecio.Text, "0,00")) Then
    adoIntro.Recordset.Fields(8).Value = Replace((txtPrecio.Text), ",", ".")
    Else
    adoIntro.Recordset.Fields(8).Value = Replace(txtPrecio.Text, ",", ".")
    End If

    If dlstTipo.BoundText = "SERVICIO" Then
        If txtIDCLIENTA.Text = "" Then
        adoIntro.Recordset.Fields(9).Value = "0"
        Else
        adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text
        End If
    End If

' 21/04/2015 que suba el dato de idclienta si es v. público
    If dlstTipo.BoundText = "Venta al PÚBLICO" Then
        If txtIDCLIENTA.Text = "" Then
        adoIntro.Recordset.Fields(9).Value = "0"
        Else
        adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text
        End If
    End If

' hasta aqui 21/04/2015 ------------------------------

' 06/05/2015 que suba el dato de idclienta si es v. CABINA
    If dlstTipo.BoundText = "Venta de CABINA" Then
        If txtIDCLIENTA.Text = "" Then
        adoIntro.Recordset.Fields(9).Value = "0"
        Else
        adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text
        End If
    End If

' hasta aqui 06/05/2015 ------------------------------

adoIntro.Recordset.Update
adoIntro.Refresh
End If
If txtCB.Text <> "" Then
On Error Resume Next
'Si sí tiene CB entonces
adoCB.Refresh
adoCB.Recordset.AddNew
adoCB.Recordset.Fields(0).Value = txtCB.Text
adoCB.Recordset.Fields(1).Value = dcmbProducto(0).BoundText
adoCB.Recordset.Fields(2).Value = dcmbMarca(1).BoundText
adoCB.Recordset.Update
adoCB.Refresh

On erro GoTo error1
adoIntro1.Refresh
adoIntro1.Refresh
adoIntro1.Recordset.AddNew

adoIntro1.Recordset.Fields(1).Value = dlstSta.BoundText
adoIntro1.Recordset.Fields(2).Value = LBLFECHA.Caption
adoIntro1.Recordset.Fields(3).Value = lblhora.Caption
adoIntro1.Recordset.Fields(4).Value = dlstTipo.SelectedItem
adoIntro1.Recordset.Fields(5).Value = dcmbProducto(0).BoundText
adoIntro1.Recordset.Fields(6).Value = dcmbMarca(1).BoundText
adoIntro1.Recordset.Fields(7).Value = txtUnidades.Text
adoIntro1.Recordset.Fields(10).Value = txtCB.Text


    If Val(txtPrecio.Text) = CCur(Format(txtPrecio.Text, "0,00")) Then
    adoIntro1.Recordset.Fields(8).Value = Replace((txtPrecio.Text), ",", ".")
    Else
    adoIntro1.Recordset.Fields(8).Value = Replace(txtPrecio.Text, ",", ".")
    End If

    If dlstTipo.BoundText = "SERVICIO" Then
    If txtIDCLIENTA.Text = "" Then
    adoIntro1.Recordset.Fields(9).Value = "0"
    Else
    adoIntro1.Recordset.Fields(9).Value = txtIDCLIENTA.Text
    End If
    End If

' 21/04/2015 que suba el dato de idclienta si es v. público
    If dlstTipo.BoundText = "Venta al PÚBLICO" Then
        If txtIDCLIENTA.Text = "" Then
        adoIntro1.Recordset.Fields(9).Value = "0"
        Else
        adoIntro1.Recordset.Fields(9).Value = txtIDCLIENTA.Text
        End If
    End If

' hasta aqui 21/04/2015 ------------------------------

' 06/05/2015 que suba el dato de idclienta si es v. CABINA
    If dlstTipo.BoundText = "Venta de CABINA" Then
        If txtIDCLIENTA.Text = "" Then
        adoIntro1.Recordset.Fields(9).Value = "0"
        Else
        adoIntro1.Recordset.Fields(9).Value = txtIDCLIENTA.Text
        End If
    End If

' hasta aqui 06/05/2015 ------------------------------


adoIntro1.Recordset.Update
adoIntro1.Refresh




    

End If


  'y AHORA QUIERO QUE LA PANTALLA ME APAREZCA ASÍ

  '
  On Error Resume Next
  
MousePointer = 1
 
dcmbProducto(0).Text = ""
  
dcmbMarca(1).Text = ""
txtUnidades.Text = "1"
txtPrecio.Text = ""
txtCB.Text = ""


 adoParcial.CommandType = adCmdText
 adoParcial.RecordSource = "SELECT  Parcial.idparcial, Señoritas.Nombre, Tipos.Descripción, Parcial.Fecha, Parcial.Hora, Producto.Descripción, Marca.Empresa, Parcial.Unidades, CCur(Format(Parcial.Precio, '0.00')) AS Precio, ccur(format([Parcial]![Unidades]*[Parcial]![Precio], '0.00')) AS Total FROM Tipos INNER JOIN (Señoritas INNER JOIN (Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto) ON Señoritas.IdSeñoritas = Parcial.IdSeñorita) ON Tipos.IdTipos = Parcial.IdTipo WHERE (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)= '" & lblhora.Caption & "' ));"
 adoParcial.Refresh



 
 flex.Refresh
  flex.Visible = True
flex.ColWidth(0) = 0
flex.ColWidth(1) = 1000
flex.ColWidth(2) = 1600
flex.ColWidth(3) = 0
flex.ColWidth(4) = 0
flex.ColWidth(5) = 3200
flex.ColWidth(6) = 1200
flex.ColWidth(7) = 800



 
 
adoTotal.CommandType = adCmdText
adoTotal.RecordSource = "SELECT Sum([Parcial]![Unidades]*[Parcial]![Precio]) AS Subtotal From Parcial GROUP BY Parcial.Fecha, Parcial.Hora HAVING (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)='" & lblhora.Caption & "' ));"
adoTotal.Refresh


lblTotal.Caption = adoTotal.Recordset.Fields(0).Value



If lblTotal.Caption > "0" Then mnuSalir.Visible = False
If lblTotal.Caption > "0" Then cmdNuevoCliente.Visible = False




 




cmdAdelante.Visible = True


If dlstTipo.BoundText <> "SERVICIO" Then

txtCB.SetFocus
 MousePointer = 1
 Else

 dcmbMarca(1).Text = adoMarca.Recordset.Fields(1).Value
dcmbProducto(0).SetFocus
 End If
 
  '***********************
  'Para refrescar los puntos acumulados
  
  On Error Resume Next
If txtIDCLIENTA.Text <> "" Then
' Lo de arriba es que me da error si el campo de clientas está en blanco
lblPuntos.Visible = True
lblPuntosAcum.Visible = True
adoConsultaPuntos.CommandType = adCmdText
adoConsultaPuntos.RecordSource = "SELECT Last(Puntos.Fecha) AS ÚltimoDeFecha, Last(Puntos.Hora) AS ÚltimoDeHora, Last(Puntos.Acumulado) AS ÚltimoDeAcumulado From Puntos WHERE ((Puntos.IdClienta)=" & txtIDCLIENTA.Text & ");"
adoConsultaPuntos.Refresh
adoConsultaPuntos.Refresh
    If adoConsultaPuntos.Recordset.Fields(2).Value <> "" Then

    lblFechaPuntos.Caption = adoConsultaPuntos.Recordset.Fields(0).Value
 
    diahoy = Format(LBLFECHA.Caption, "dd") + Format(LBLFECHA.Caption, "mm") * 30 + Format(LBLFECHA.Caption, "yy") * 365
    diaPuntos = Format(lblFechaPuntos.Caption, "dd") + Format(lblFechaPuntos.Caption, "mm") * 30 + Format(lblFechaPuntos.Caption, "yy") * 365
        If Abs(diahoy - diaPuntos) < 90 Then
        lblPuntosAcum.Caption = Format(adoConsultaPuntos.Recordset.Fields(2).Value, "0.00")
    
        cmdDescTodosPuntos.Visible = True
        cmdDescPartePuntos.Visible = True
    'If dlstTipo.BoundText = "SERVICIO" Then
    'cmdDescTodosPuntos.Enabled = False
    'cmdDescPartePuntos.Enabled = False
    'End If
    
        Else
        lblPuntosAcum.Caption = "0,00"
        End If
    
    Else
    lblPuntosAcum.Caption = "0,00"
    End If
End If
 




 
 '*******************************************************
Exit Sub
errordato:
    MsgBox "No puedes ser", vbInformation
    dcmbProducto(0).SetFocus
     MousePointer = 1
    Exit Sub
errorformato:
   
    adoIntro.Recordset.Fields(8).Value = txtPrecio.Text
    Resume Next
    Exit Sub
error1:
    MsgBox "No puedes 1 ser", vbInformation
    dcmbProducto(0).SetFocus
     MousePointer = 1
    Exit Sub


End Sub

 




Private Sub cmdSiPuntos_Click()
lblAnotaPuntos.Caption = "Se anotan puntos"
lblAnotaPuntos.BackColor = &HFF00&
cmdNoPuntos.Visible = True
cmdSiPuntos.Visible = False
cmdAdelante.BackColor = &HFF00&

End Sub

Private Sub Command1_Click()
frmNuevaClienta.Show
Unload Me


End Sub

Private Sub dcmbMarca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_RETURN Then txtUnidades.SetFocus
txtUnidades.SelStart = 0
txtUnidades.SelLength = Len(txtUnidades.Text)
End Sub

Private Sub dcmbProducto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errorservicio
If KeyCode = KEY_RETURN Then
If dlstTipo.BoundText = "SERVICIO" Then

adoPrecio.CommandType = adCmdText
adoPrecio.RecordSource = "SELECT Producto.[Servicio?], Producto.Descripción, Producto.Precio From Producto WHERE (((Producto.[Servicio?])=-1) AND ((Producto.Descripción)='" & dcmbProducto(0).Text & "'));"
adoPrecio.Refresh
txtPrecio.Text = adoPrecio.Recordset.Fields(2).Value
txtPrecio.SetFocus
txtPrecio.SelStart = 0
txtPrecio.SelLength = Len(txtPrecio.Text)

Else
dcmbMarca(1).SetFocus
End If
End If

Exit Sub
errorservicio:
If KeyCode = KEY_RETURN Then txtUnidades.SetFocus

End Sub




Private Sub dcomboClienteProf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_RETURN Then
txtCB.SetFocus
adoIdProfesional.CommandType = adCmdText
adoIdProfesional.RecordSource = "SELECT ClientesProfesionales.Empresa, ClientesProfesionales.IdClienProf From ClientesProfesionales WHERE (((ClientesProfesionales.Empresa)='" & dcomboClienteProf.Text & "'));"
adoIdProfesional.Refresh
adoIdProfesional.Refresh
If dcomboClienteProf.Text = "" Then
lblIdProfesional.Caption = 23
lblIdProfesional.Visible = True
Else

lblIdProfesional.Caption = adoIdProfesional.Recordset.Fields(1).Value
lblIdProfesional.Visible = True
End If
End If

End Sub

Private Sub dlstSta_KeyPress(KeyAscii As Integer)
dlstTipo.SetFocus
End Sub





Private Sub dlstTipo_KeyPress(KeyAscii As Integer)
cmdNuevoProducto.Visible = True
cmdNuevaMarca.Visible = True
cmd10.Visible = True

If dlstTipo.BoundText = "Venta al PÚBLICO" Then

'***********************
'02/07/15 Si ve que quitar puntos está visible entonces enable también
' era para que si se ha hecho con servicio se tenga que poner en público
If cmdDescPartePuntos.Visible = True Then cmdDescPartePuntos.Enabled = True
If cmdDescTodosPuntos.Visible = True Then cmdDescTodosPuntos.Enabled = True

'****************************************


adoProducto.CommandType = adCmdText
adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
adoProducto.Refresh


adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca Where (((Marca.[Servicio?]) = 0)) ORDER BY Marca.Empresa;"
adoMarca.Refresh




dcmbProducto(0).Visible = True
dcmbMarca(1).Visible = True
'dcmbServicio(2).Visible = False
txtUnidades.Visible = True
txtPrecio.Visible = True
lblDescripción(0).Visible = True
lblDescripción(1).Visible = True
lblDescripción(2).Visible = False
lblDescripción(3).Visible = True
lblDescripción(4).Visible = True
lblCB.Visible = True
txtCB.Visible = True
'cmdOK.Visible = True
lblClienta.Visible = True
txtIDCLIENTA.Visible = True



txtCB.SetFocus


End If


If dlstTipo.BoundText = "Venta al PROFESIONAL" Then
adoProducto.CommandType = adCmdText
adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
adoProducto.Refresh


adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca Where (((Marca.[Servicio?]) = 0)) ORDER BY Marca.Empresa;"
adoMarca.Refresh

adoConsVentaProf.CommandType = adCmdText
adoConsVentaProf.RecordSource = "SELECT ClientesProfesionales.IdClienProf, ClientesProfesionales.Empresa From ClientesProfesionales GROUP BY ClientesProfesionales.IdClienProf, ClientesProfesionales.Empresa ORDER BY ClientesProfesionales.Empresa;"
adoConsVentaProf.Refresh


dcomboClienteProf.Visible = True

dcmbProducto(0).Visible = True
dcmbMarca(1).Visible = True
'dcmbServicio(2).Visible = False
txtUnidades.Visible = True
txtPrecio.Visible = True
lblDescripción(0).Visible = True
lblDescripción(1).Visible = True
lblDescripción(2).Visible = False
lblDescripción(3).Visible = True
lblDescripción(4).Visible = True
'cmdOK.Visible = True
dcomboClienteProf.SetFocus
lblCB.Visible = True
txtCB.Visible = True

End If


If dlstTipo.BoundText = "SERVICIO" Then
dcmbProducto(0).Visible = True
dcmbMarca(1).Visible = True
dcmbMarca(1).BoundText = "--"


lblClienta.Visible = True
txtIDCLIENTA.Visible = True
dcmbProducto(0).SetFocus


adoProducto.CommandType = adCmdText
adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = -1)) ORDER BY Producto.Descripción;"
adoProducto.Refresh

adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa, Marca.[Servicio?] From Marca Where (((Marca.[Servicio?]) = -1)) ORDER BY Marca.Empresa;"
adoMarca.Refresh
dcmbMarca(1).Text = "SERVICIO"



txtUnidades.Visible = True
txtPrecio.Visible = True
lblDescripción(0).Visible = False
lblDescripción(1).Visible = False
lblDescripción(2).Visible = True
lblDescripción(3).Visible = True
lblDescripción(4).Visible = True
'cmdOK.Visible = True


End If



If dlstTipo.BoundText = "Venta de CABINA" Then

'***********************
'02/07/15 Si ve que quitar puntos está visible entonces enable también
' era para que si se ha hecho con servicio se tenga que poner en público
If cmdDescPartePuntos.Visible = True Then cmdDescPartePuntos.Enabled = True
If cmdDescTodosPuntos.Visible = True Then cmdDescTodosPuntos.Enabled = True

'****************************************

adoProducto.CommandType = adCmdText
adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
adoProducto.Refresh


adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca Where (((Marca.[Servicio?]) = 0)) ORDER BY Marca.Empresa;"
adoMarca.Refresh

dcmbProducto(0).Visible = True
dcmbMarca(1).Visible = True
'dcmbServicio(2).Visible = False
txtUnidades.Visible = True
txtPrecio.Visible = True
lblDescripción(0).Visible = True
lblDescripción(1).Visible = True
lblDescripción(2).Visible = False
lblDescripción(3).Visible = True
lblDescripción(4).Visible = True
'cmdOK.Visible = True

lblCB.Visible = True
txtCB.Visible = True
'cmdOK.Visible = True
txtCB.SetFocus
End If

dlstTipo.Refresh

End Sub


Private Sub flex_Click()

flex.Refresh
flex.Col = 0

If lblTotal.Caption <> "" Then
MsgBox "Para borrar la línea ' " & flex.RowSel & " ' pulsa el botón Borrar Línea, y si borras no se te olvide de introducir el producto en el inventario"
cmdEditar.Visible = True
Else
cmdEditar.Visible = False
End If
If KeyCode = KEY_ESCAPE Then flex.Refresh


End Sub

Private Sub Form_Load()
 MousePointer = 1

LBLFECHA.Caption = Format(Now, "dd/mm/yy")
lblhora.Caption = Format(Now, "hh:mm:ss")
Fecha = LBLFECHA.Caption
Hora = lblhora.Caption

adoOriseño.CommandType = adCmdText
adoOriseño.RecordSource = "SELECT Señoritas.Nombre, Señoritas.IdSeñoritas From Señoritas Where (((Señoritas.Secuencia) > '0') And ((Señoritas.[Activo?]) = -1)) ORDER BY Señoritas.Secuencia;"
'adoOriseño.RecordSource = "SELECT Señoritas.Nombre, Señoritas.IdSeñoritas From Señoritas Where (((Señoritas.[Activo?]) = -1))ORDER BY Señoritas.Nombre;"
adoOriseño.Refresh

adoTipo.CommandType = adCmdText
adoTipo.RecordSource = "SELECT Tipos.*FROM Tipos;"
adoTipo.Refresh

adoProducto.CommandType = adCmdText
adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
adoProducto.Refresh


'adoMarca.CommandType = adCmdText
'adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca ORDER BY Marca.Empresa;"
'adoMarca.Refresh




End Sub


















Private Sub mnuAcerca_Click()
frmAbout.Show

End Sub

Private Sub mnuActualizaciónInventario_Click()
frmActualizaInventario.Show
frmActualizaInventario.Text1.SetFocus
End Sub





Private Sub mnuAlbaranes_Click()
frmListaAdelClientas.Show
frmListaAdelClientas.dcmbClientasAlgo.SetFocus
End Sub

Private Sub mnuAlbaTrab_Click()
frmClienSeñoCredito.Show

End Sub

Private Sub mnuAltaClientes_Click()
frmNuevaClienta.Show

End Sub

Private Sub mnuApellido_Click()
frmApellidos.Show

End Sub



Private Sub mnuBusCliProf_Click()
frmVentaProfesional.Show


End Sub

Private Sub mnuBúsqueda_Click()
frmFecha.Show
frmFecha.txtFecha.SetFocus

End Sub

Private Sub mnuCaja_Click()
MousePointer = 11

frmCaja.Show
MousePointer = 1

frmCaja.cmdSalir.SetFocus
End Sub

Private Sub mnuCalculadora_Click()
frmCalculadora.Show

End Sub

Private Sub mnuClientasAlbaranVez_Click()
frmClienProfCredito.Show

End Sub

Private Sub mnuCodigosInternos_Click()
frmCodigosInternos.Show

End Sub

Private Sub mnuComprasClientas_Click()
frmComprasClientas.Show
End Sub

Private Sub mnuConsCompraClienta_Click()
frmConsCompraClienta.Show
End Sub

Private Sub mnuEditor_Click()
frmEditor.Show

End Sub

Private Sub mnuEmpresas_Click()
frmProfesionales.Show

End Sub

Private Sub mnuNombre_Click()
frmClientasParticulares.Show

End Sub

Private Sub mnuEntradaInventario_Click()
frmEntradaInventario.Show

End Sub

Private Sub mnuFlotarium_Click()
FRMaLARMA.Show

End Sub

Private Sub mnuInventarioCasas_Click()
frmInventarioCasas.Show
frmInventarioCasas.dcmbEmpresas.SetFocus
frmInventarioCasas.adoProveedores.CommandType = adCmdText
frmInventarioCasas.adoProveedores.RecordSource = "SELECT Marca.Empresa from Marca ORDER BY Marca.Empresa;"
frmInventarioCasas.adoProveedores.Refresh

End Sub

Private Sub mnuListAdelantadoClientas_Click()
frmListaAdelClientas.Show
frmListaAdelClientas.dcmbClientasAlgo.SetFocus
End Sub

Private Sub mnulistAdelantadoSeñoritas_Click()
frmListAdelSeño.Show
frmListAdelSeño.dlstSta.SetFocus
End Sub

Private Sub mnuListadoAlbaranTrab_Click()
frmListaAdelTrab.Show

End Sub

Private Sub mnuListadoFaltas_Click()
frmListadoFaltas.Show

End Sub

Private Sub mnuListadoSalidaCaja_Click()
frmListadoSalidaCaja.Show
frmListadoSalidaCaja.txtFecha.SetFocus
End Sub

Private Sub mnuMatCab_Click()
frmMatCab.Show

End Sub

Private Sub mnuModificarCB_Click()
frmModificaCB.Show
frmModificaCB.txtCB.SetFocus
End Sub

Private Sub mnuNuevaMarca_Click()
frmNuevaMarca.Show

End Sub

Private Sub mnuNuevoCB_Click()
frmNuevoCB.Show
frmNuevoCB.txtCB.SetFocus
End Sub

Private Sub mnuNuevoProducto_Click()
frmNuevoProducto.Show

End Sub



Private Sub mnuPagoAlbaranes_Click()
frmPagoAdelClientas.Show
End Sub

Private Sub mnuParticulares_Click()
frmApellidos.Show

End Sub

Private Sub mnuReponerTiquet_Click()
frmReponerTiquet.Show

End Sub

Private Sub mnuResumendía_Click()
frmResumendia.Show
frmResumendia.txtFecha.SetFocus
End Sub

Private Sub mnuSalidadeCaja_Click()
frmSalidadeCaja.Show

End Sub

Private Sub mnuSalidaInventario_Click()
frmSalidaInventario.Show
frmSalidaInventario.lstMotivo.SetFocus

End Sub

Private Sub mnuSalidaInventPersonal_Click()
frmSalidaInventario.Show
frmSalidaInventario.lstMotivo.SetFocus
frmSalidaInventario.lstMotivo.ListIndex = "4"
frmSalidaInventario.txtPrecio.Visible = True
frmSalidaInventario.lbl(6).Visible = True
frmSalidaInventario.dlstSta.Visible = True
frmSalidaInventario.dlstSta.SetFocus
frmSalidaInventario.Command1.Visible = False


End Sub

Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuServClienta_Click()
frmConsServicio.Show

End Sub

Private Sub mnuServicios_Click()
frmServiciosClientas.Show

End Sub

Private Sub mnuSubInventario_Click()
frmInventario.Show

End Sub

Private Sub mnuTelClienProf_Click()
frmTelClienProf.Show

End Sub

Private Sub mnuVentasDía_Click()
frmListadoVentaDia.Show

End Sub



   
    

Private Sub mnuVisaEfectivo_Click()
frmVisaEfectivo.Show

End Sub










Private Sub txtCB_KeyDown(KeyCode As Integer, Shift As Integer)


On Error GoTo caca


If KeyCode = KEY_RETURN Then

'Si el codigo de barras esta vacio rellena con todo

If txtCB.Text = "" Then

 adoProducto.Refresh
   adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
  adoProducto.Refresh
  dcmbProducto(0).Refresh
  dcmbProducto(0).SetFocus
  
  adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca Where (((Marca.[Servicio?]) = 0)) ORDER BY Marca.Empresa;"
adoMarca.Refresh
dcmbMarca(1).Refresh
  


Else
 'si esta lleno entonces depende
 
 ' le hago la consulta de ese campo
 
 adoInventario.CommandType = adCmdText
adoInventario.RecordSource = "SELECT Inventario.CB, Sum(Inventario.Unidades) AS SumaDeUnidades From Inventario GROUP BY Inventario.CB HAVING (((Inventario.CB)=" & txtCB.Text & "));"
adoInventario.Refresh
lblUnidadesInventario.Caption = adoInventario.Recordset.Fields(1).Value
 
  adoProducto.Refresh
      adoProducto.CommandType = adCmdText
      adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción FROM CB INNER JOIN Producto ON CB.IdProducto = Producto.IdProducto WHERE ((CB.CB)=" & txtCB.Text & ");"
      adoProducto.Refresh
      dcmbProducto(0).Text = adoProducto.Recordset.Fields(1).Value

      dcmbProducto(0).Refresh
    
       
       adoMarca.CommandType = adCmdText
       adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa FROM CB INNER JOIN Marca ON CB.IdMarca = Marca.IdMarca WHERE ((CB.CB)=" & txtCB.Text & ");"
       adoMarca.Refresh
       'esto es para que aparezca el primer nombre en la casilla
       
       dcmbMarca(1).Text = adoMarca.Recordset.Fields(1).Value

       dcmbMarca(1).Refresh
          
  'txtUnidades.SetFocus
  'txtUnidades.SelStart = 0
'txtUnidades.SelLength = Len(txtUnidades.Text)
 
 'LO NUEVO
 
' BUSQUEDA DEL PRECIO DEL PRODUCTO

On Error GoTo errorconsultaprecio
' AQUI SE BUSCA

If dlstTipo.BoundText = "SERVICIO" Then
adoPrecio.CommandType = adCmdText
adoPrecio.RecordSource = "SELECT Producto.[Servicio?], Producto.Descripción, Producto.Precio From Producto WHERE (((Producto.[Servicio?])=1) AND ((Producto.Descripción)='" & dcmbProducto(0).Text & "'));"
adoPrecio.Refresh


Else





adoConsultaPrecio.Refresh

adoConsultaPrecio.CommandType = adCmdText

adoConsultaPrecio.RecordSource = "SELECT Producto.Descripción, Marca.Empresa, Last(Parcial.Precio) AS ÚltimoDePrecio, Last(Parcial.Fecha) AS ÚltimoDeFecha, Parcial.IdTipo FROM Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto GROUP BY Producto.Descripción, Marca.Empresa, Parcial.IdTipo HAVING (((Producto.Descripción)='" & dcmbProducto(0).Text & "') AND ((Marca.Empresa)='" & dcmbMarca(1).Text & "') AND ((Parcial.IdTipo)=1));"
'adoConsultaPrecio.RecordSource = "SELECT Last(Parcial.Fecha) AS ÚltimoDeFecha, Parcial.CB, Max(Parcial.Precio) AS MáxDePrecio, Parcial.IdTipo From Parcial GROUP BY Parcial.CB, Parcial.IdTipo HAVING (((Parcial.CB)='" & txtCB.Text & "') AND ((Parcial.IdTipo)=1));"
'lo que pone debajo era lo que había antes y buscaba el valor más alto
'adoConsultaPrecio.RecordSource = "SELECT Producto.Descripción, Marca.Empresa, Max(Parcial.Precio) AS MáxDePrecio FROM Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto GROUP BY Producto.Descripción, Marca.Empresa HAVING (((Producto.Descripción)= '" & dcmbProducto(0).Text & "' ) AND ((Marca.Empresa)='" & dcmbMarca(1).Text & "' ));"
adoConsultaPrecio.Refresh


End If
' AQUI SE ESCRIBE EN LA CAJA DE TEXTO

If KeyCode = KEY_RETURN Then txtPrecio.SetFocus

If dlstTipo.BoundText = "SERVICIO" Then
txtPrecio.Text = adoPrecio.Recordset.Fields(2).Value
Else
txtPrecio.Text = adoConsultaPrecio.Recordset.Fields(2).Value
lblFechaPrecio.Refresh
lblFechaPrecio.Caption = adoConsultaPrecio.Recordset.Fields(3).Value



End If

txtPrecio.SelStart = 0
txtPrecio.SelLength = Len(txtPrecio.Text)
If dlstTipo.BoundText = "Venta al PROFESIONAL" Then
lblPreuProfesional.Caption = Format(adoConsultaPrecio.Recordset.Fields(2).Value * 90 / 100, "0.00")
lblPreu.Visible = True
lblPreuProfesional.Visible = True
End If





Exit Sub

errorconsultaprecio:
    txtPrecio.Text = ""
   
    Exit Sub
    

    



 
 
 
 txtPrecio.SetFocus
 txtPrecio.SelStart = 0
 txtPrecio.SelLength = Len(txtPrecio.Text)
 
 
 
 End If


End If

'arreglo:

'
Exit Sub

caca:

   adoProducto.CommandType = adCmdText
   adoProducto.RecordSource = "SELECT Producto.IdProducto, Producto.Descripción From Producto GROUP BY Producto.IdProducto, Producto.Descripción, Producto.[Servicio?] Having (((Producto.[Servicio?]) = 0)) ORDER BY Producto.Descripción;"
  adoProducto.Refresh
  dcmbProducto(0).Refresh

adoMarca.CommandType = adCmdText
adoMarca.RecordSource = "SELECT Marca.IdMarca, Marca.Empresa From Marca Where (((Marca.[Servicio?]) = 0)) ORDER BY Marca.Empresa;"
adoMarca.Refresh
dcmbMarca(1).Refresh
    
    dcmbProducto(0).SetFocus

Exit Sub




      End Sub




Private Sub txtIDCLIENTA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_RETURN Then
lblAvisoCumple.Visible = False
lblPreuProfesional.Visible = False
lblPreu.Visible = False


If txtIDCLIENTA <> "" Then

'***************************
'2015 mayo. Pongo la etiqueta de puntos si la clienta existe y hace menos de 3 meses que ha comprado
On Error Resume Next

lblPuntos.Visible = True
lblPuntosAcum.Visible = True
adoConsultaPuntos.CommandType = adCmdText
adoConsultaPuntos.RecordSource = "SELECT Last(Puntos.Fecha) AS ÚltimoDeFecha, Last(Puntos.Hora) AS ÚltimoDeHora, Last(Puntos.Acumulado) AS ÚltimoDeAcumulado From Puntos WHERE (((Puntos.IdClienta)=" & txtIDCLIENTA.Text & "));"
adoConsultaPuntos.Refresh
If adoConsultaPuntos.Recordset.Fields(2).Value <> "" Then
'lblPuntosAcum.Caption = adoConsultaPuntos.Recordset.Fields(2).Value
 lblFechaPuntos.Caption = adoConsultaPuntos.Recordset.Fields(0).Value
 
 diahoy = Format(LBLFECHA.Caption, "dd") + Format(LBLFECHA.Caption, "mm") * 30 + Format(LBLFECHA.Caption, "yy") * 365
 diaPuntos = Format(lblFechaPuntos.Caption, "dd") + Format(lblFechaPuntos.Caption, "mm") * 30 + Format(lblFechaPuntos.Caption, "yy") * 365
    If Abs(diahoy - diaPuntos) < 90 Then
    lblPuntosAcum.Caption = Format(adoConsultaPuntos.Recordset.Fields(2).Value, "0.00")
    Else
    lblPuntosAcum.Caption = "0,00"
    End If
Else
lblPuntosAcum.Caption = "0,00"
End If

'***************************
'13/03/2017 para intentar que no se anoten siempre los puntos

lblAnotaPuntos.Visible = True
cmdSiPuntos.Visible = False
lblAnotaPuntos.Caption = "Se anotan puntos"
lblAnotaPuntos.BackColor = &HFF00&
cmdNoPuntos.Visible = True


'***************************

adoEmail.CommandType = adCmdText
adoEmail.RecordSource = "SELECT clientas.email From clientas WHERE (((clientas.IDClientas)=" & txtIDCLIENTA.Text & "));"
adoEmail.Refresh


adoConsClienta.CommandType = adCmdText
adoConsClienta.RecordSource = "SELECT clientas.IDClientas, clientas.Nombre, clientas.Apellidos, clientas.Descuentos, clientas.cumple, clientas.Móvil, clientas.Teléfono From clientas GROUP BY clientas.IDClientas, clientas.Nombre,clientas.Descuentos, clientas.Apellidos, clientas.cumple, clientas.Móvil, clientas.Teléfono HAVING (((clientas.IDClientas)=" & txtIDCLIENTA.Text & "));"
adoConsClienta.Refresh
lblNombre.Visible = True
lblApellidos.Visible = True
lblPorciento.Visible = True
lblEmail.Visible = True



If adoEmail.Recordset.Fields(0).Value <> "" Then
lblEmail.Caption = adoEmail.Recordset.Fields(0).Value
Else
lblEmail.Caption = "sin e-mail"
End If

If adoConsClienta.Recordset.Fields(5).Value = "" Then
lblMovil.Visible = True
lblMovil.Caption = "Sin móvil"
Else
lblMovil.Visible = True
lblMovil.Caption = adoConsClienta.Recordset.Fields(5).Value
End If

If adoConsClienta.Recordset.Fields(6).Value = "" Then
lblTelf.Visible = True
lblTelf.Caption = "Sin teléfono"
Else
lblTelf.Visible = True
lblTelf.Caption = adoConsClienta.Recordset.Fields(6).Value
End If

'24/04/2015 Debido al erro que creaba el haber puesto caja de clienta cuando solo era
'para servicios


If adoConsClienta.Recordset.Fields(3).Value = "5%" Then
'la linea de abajo con su correspondiente if
If dlstTipo.BoundText <> "Venta al PÚBLICO" Or dlstTipo.BoundText <> "SERVICIO" Then
lblPreuProfesional.Visible = True
lblPreuProfesional.Caption = Format(adoPrecio.Recordset.Fields(2).Value * 95 / 100, "0.00")
lblPreuProfesional.Refresh
End If
End If

If adoConsClienta.Recordset.Fields(3).Value = "10%" Then
'la linea de abajo con su correspondiente if
If dlstTipo.BoundText <> "Venta al PÚBLICO" Or dlstTipo.BoundText <> "SERVICIO" Then
lblPreuProfesional.Visible = True
lblPreuProfesional.Caption = Format(adoPrecio.Recordset.Fields(2).Value * 90 / 100, "0.00")
lblPreuProfesional.Refresh
End If
End If

If adoConsClienta.Recordset.Fields(4).Value <> "" Then
lblCumple.Caption = Format(adoConsClienta.Recordset.Fields(4).Value, "dd/mm/yy")
lblCumple.Visible = True
'    If DateAdd("d", 5, Format(lblCumple.Caption, "dd/mm")) - Format(lblFecha.Caption, "dd/mm") < 5 Then lblAvisoCumple.Visible = True
 diahoy = Format(LBLFECHA.Caption, "dd") + Format(LBLFECHA.Caption, "mm") * 30
 diacumple = Format(lblCumple.Caption, "dd") + Format(lblCumple.Caption, "mm") * 30
    If Abs(diahoy - diacumple) < 6 Then
    lblAvisoCumple.Visible = True
    lblPreuProfesional.Caption = Format(txtPrecio.Text * 90 / 100, "0.00")
    lblPreu.Visible = True
    lblPreu.Caption = "Precio con descuento:"
    lblPreuProfesional.Visible = True
    End If
    



Else
lblCumple.Visible = True
lblCumple.Caption = "No tiene fecha de cumpleaños"
End If


If adoConsClienta.Recordset.Fields(1).Value <> "" Then
lblNombre.Caption = adoConsClienta.Recordset.Fields(1).Value
Else
lblNombre.Caption = "---"
End If
If adoConsClienta.Recordset.Fields(2).Value <> "" Then
lblApellidos.Caption = adoConsClienta.Recordset.Fields(2).Value
Else
lblApellidos.Caption = "---"
End If




If adoConsClienta.Recordset.Fields(3).Value <> "" Then
lblPorciento.Caption = adoConsClienta.Recordset.Fields(3).Value
Else
lblPorciento.Caption = ""
End If
End If
cmdOK.SetFocus

End If


End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)




If KeyCode = KEY_RETURN Then

'*********02/07/2015
lblAvisoNegativo.Visible = False
    If (dcmbProducto(0).Text = "PUNTOS" And Val(txtPrecio.Text) > "0") Or (dcmbProducto(0).Text = "Puntos Servicio" And Val(txtPrecio.Text) > "0") Then
    MsgBox ("el precio no puede ser positivo pues se trata de puntos")
    txtPrecio.SetFocus
    MousePointer = 1
    Exit Sub
    End If
    If (dcmbProducto(0).Text = "PUNTOS" Or dcmbProducto(0).Text = "Puntos Servicio") And Val(Replace(txtPrecio.Text, ",", ".")) < (Val(Replace(lblPuntosAcum.Caption, ",", ".")) * -1) Then
    MsgBox ("el precio no puede ser más grande que el acumulado pues se trata de puntos")
    txtPrecio.SetFocus
    MousePointer = 1
    Exit Sub
    End If

'*******************************



'If dlstTipo.BoundText = "SERVICIO" Then
' 31/07/15 Esta linea de abajo es lo único que he puesto para que se vea lo de la clienta
If dlstTipo.BoundText <> "Venta al PROFESIONAL" Then
cmdOK.Visible = True
lblClienta.Visible = True

txtIDCLIENTA.Visible = True

txtIDCLIENTA.SetFocus
Else

'Nuevo de 21/04/2015

    If dlstTipo.BoundText = "Venta al PÚBLICO" Then

    cmdOK.Visible = True
    txtIDCLIENTA.SetFocus
    'cmdOK.Visible = True
    'cmdOK.SetFocus
    Else
    'Hasta aqui lo nuevo de 21/04/2015 y coloco donde vea el final el end if

    'lo nuevo, antes iba en el cmdOK
    MousePointer = 11
    lblFechaPrecio.Caption = ""
    lblPreu.Visible = False
    lblPreuProfesional.Visible = False
    'Tratamiento de errores por falta de datos


    'Para que no se pueda salir del programa
    'si no se ha borrado o acabado todo el ciclo


    'Para que solo se pueda hacer un nuevo cliente si está todo limpio


        If dcmbProducto(0).BoundText = "" Then
        MsgBox ("Introduce el producto")
        dcmbProducto(0).SetFocus
        MousePointer = 1
        Exit Sub
        End If
        If dcmbMarca(1).BoundText = "" Then
        MsgBox ("Introduce la marca")
        dcmbMarca(1).SetFocus
        MousePointer = 1
        Exit Sub
        End If
        If txtUnidades.Text = "" Then
        MsgBox ("Introduce las unidades")
        txtUnidades.SetFocus
        MousePointer = 1
        Exit Sub
        End If
        If txtPrecio.Text = "" Then
        MsgBox ("Introduce el precio")
        txtPrecio.SetFocus
        MousePointer = 1
        Exit Sub
        End If
    On Error GoTo errordato
' ahora le digo que si el código de barras esta vacio que suba todo menos eso
        If txtCB.Text = "" Then

        adoIntro.Refresh
        adoIntro.Recordset.AddNew

        adoIntro.Recordset.Fields(1).Value = dlstSta.BoundText
        adoIntro.Recordset.Fields(2).Value = LBLFECHA.Caption
        adoIntro.Recordset.Fields(3).Value = lblhora.Caption
        adoIntro.Recordset.Fields(4).Value = dlstTipo.SelectedItem
        adoIntro.Recordset.Fields(5).Value = dcmbProducto(0).BoundText

        adoIntro.Recordset.Fields(6).Value = dcmbMarca(1).BoundText
        adoIntro.Recordset.Fields(7).Value = txtUnidades.Text
            If Val(txtPrecio.Text) = CCur(Format(txtPrecio.Text, "0,00")) Then
            adoIntro.Recordset.Fields(8).Value = Replace((txtPrecio.Text), ",", ".")
            Else
            adoIntro.Recordset.Fields(8).Value = Replace(txtPrecio.Text, ",", ".")
            End If
'Modificación, antes estaba = "SERVICIO"  y ahora <> "Venta al PROFESIONAL"
            If dlstTipo.BoundText <> "Venta al PROFESIONAL" Then
                If txtIDCLIENTA.Text = "" Then
                adoIntro.Recordset.Fields(9).Value = "0"
                Else
                adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text
                End If
            End If

' 06/05/2015 que suba el dato de idclienta si es v. CABINA
            If dlstTipo.BoundText = "Venta de CABINA" Then
'*************************************************************
'Modificación 03/07/2015


            If dlstTipo.BoundText <> "Venta al PROFESIONAL" And txtIDCLIENTA.Text <> "" And (dcmbProducto(0).Text = "PUNTOS" Or dcmbProducto(0).Text = "Puntos Servicio") Then


            adoAnulPuntos.Refresh
            adoAnulPuntos.Recordset.AddNew

            adoAnulPuntos.Recordset.Fields(1).Value = txtIDCLIENTA.Text
            adoAnulPuntos.Recordset.Fields(2).Value = StrConv(LBLFECHA.Caption, vbUpperCase)
            adoAnulPuntos.Recordset.Fields(3).Value = StrConv(lblhora.Caption, vbUpperCase)
            adoAnulPuntos.Recordset.Fields(4).Value = Val(Replace(txtPrecio.Text, ",", "."))
            adoAnulPuntos.Recordset.Fields(5).Value = (Val(Replace(txtPrecio.Text, ",", ".")) + Val(Replace(lblPuntosAcum.Caption, ",", ".")))
 
            adoAnulPuntos.Recordset.Update
            adoAnulPuntos.Refresh

            End If

  '************************




            If txtIDCLIENTA.Text = "" Then
            adoIntro.Recordset.Fields(9).Value = "0"
            Else

            adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text






            End If
            End If

' hasta aqui 06/05/2015 ------------------------------

        adoIntro.Recordset.Update
        adoIntro.Refresh

        Else
        On Error Resume Next
        adoCB.Refresh
        adoCB.Recordset.AddNew
        adoCB.Recordset.Fields(0).Value = txtCB.Text
        adoCB.Recordset.Fields(1).Value = dcmbProducto(0).BoundText
        adoCB.Recordset.Fields(2).Value = dcmbMarca(1).BoundText
        adoCB.Recordset.Update
        adoCB.Refresh


        adoIntro.Refresh
        adoIntro.Recordset.AddNew

        adoIntro.Recordset.Fields(1).Value = dlstSta.BoundText
        adoIntro.Recordset.Fields(2).Value = LBLFECHA.Caption
        adoIntro.Recordset.Fields(3).Value = lblhora.Caption
        adoIntro.Recordset.Fields(4).Value = dlstTipo.SelectedItem
        adoIntro.Recordset.Fields(5).Value = dcmbProducto(0).BoundText

        adoIntro.Recordset.Fields(6).Value = dcmbMarca(1).BoundText
        adoIntro.Recordset.Fields(7).Value = txtUnidades.Text
        adoIntro.Recordset.Fields(10).Value = txtCB.Text

            If Val(txtPrecio.Text) = CCur(Format(txtPrecio.Text, "0,00")) Then
            adoIntro.Recordset.Fields(8).Value = Replace((txtPrecio.Text), ",", ".")
            Else
            adoIntro.Recordset.Fields(8).Value = Replace(txtPrecio.Text, ",", ".")
            End If

            If dlstTipo.BoundText = "SERVICIO" Then
                If txtIDCLIENTA.Text = "" Then
                adoIntro.Recordset.Fields(9).Value = "0"
                Else
                adoIntro.Recordset.Fields(9).Value = txtIDCLIENTA.Text
                End If
            End If

        adoIntro.Recordset.Update
        adoIntro.Refresh

  '23/04/2015
  
'Aqui sigue la entrada en el inventario (el (0) es por error
adoEntradaInventario(0).Refresh
        adoEntradaInventario(0).Recordset.AddNew
        adoEntradaInventario(0).Recordset.Fields(1).Value = StrConv(txtCB.Text, vbUpperCase)
            If txtUnidades.Text = "" Then
            adoEntradaInventario(0).Recordset.Fields(2).Value = 0
            Else
            'Aquí pongo el número para que salga negativo
            adoEntradaInventario(0).Recordset.Fields(2).Value = -1 * (StrConv(txtUnidades.Text, vbUpperCase))
            End If
        adoEntradaInventario(0).Recordset.Fields(3).Value = StrConv(LBLFECHA.Caption, vbUpperCase)
        adoEntradaInventario(0).Recordset.Fields(4).Value = StrConv(lblhora.Caption, vbUpperCase)
        adoEntradaInventario(0).Recordset.Fields(5).Value = 5

        adoEntradaInventario(0).Recordset.Update
        adoEntradaInventario(0).Refresh


        Text1.Text = ""
        Text2.Text = ""
        lblProducto.Caption = ""
        lblMarca.Caption = ""
        lblExistencias.Caption = ""

    

        End If



  'y AHORA QUIERO QUE LA PANTALLA ME APAREZCA ASÍ

  '
MousePointer = 1
 
dcmbProducto(0).Text = ""
  
dcmbMarca(1).Text = ""
txtUnidades.Text = "1"
txtPrecio.Text = ""
txtCB.Text = ""


 adoParcial.CommandType = adCmdText
 adoParcial.RecordSource = "SELECT  Parcial.idparcial, Señoritas.Nombre, Tipos.Descripción, Parcial.Fecha, Parcial.Hora, Producto.Descripción, Marca.Empresa, Parcial.Unidades, CCur(Format(Parcial.Precio, '0.00')) as Precio, ccur(format([Parcial]![Unidades]*[Parcial]![Precio], '0.00')) AS Total FROM Tipos INNER JOIN (Señoritas INNER JOIN (Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto) ON Señoritas.IdSeñoritas = Parcial.IdSeñorita) ON Tipos.IdTipos = Parcial.IdTipo WHERE (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)= '" & lblhora.Caption & "' ));"
 adoParcial.Refresh



 
 flex.Refresh
  flex.Visible = True
flex.ColWidth(0) = 0
flex.ColWidth(1) = 1000
flex.ColWidth(2) = 1600
flex.ColWidth(3) = 0
flex.ColWidth(4) = 0
flex.ColWidth(5) = 3200
flex.ColWidth(6) = 1200
flex.ColWidth(7) = 800



 
 
adoTotal.CommandType = adCmdText
adoTotal.RecordSource = "SELECT Sum([Parcial]![Unidades]*[Parcial]![Precio]) AS Subtotal From Parcial GROUP BY Parcial.Fecha, Parcial.Hora HAVING (((Parcial.Fecha)=# " & Format(Now, "mm/dd/yy") & " #) AND ((Parcial.Hora)='" & lblhora.Caption & "' ));"
adoTotal.Refresh


lblTotal.Caption = adoTotal.Recordset.Fields(0).Value



If lblTotal.Caption > "0" Then mnuSalir.Visible = False
If lblTotal.Caption > "0" Then cmdNuevoCliente.Visible = False




 




cmdAdelante.Visible = True


If dlstTipo.BoundText <> "SERVICIO" Then

txtCB.SetFocus
 MousePointer = 1
 Else

 dcmbMarca(1).Text = adoMarca.Recordset.Fields(1).Value

 End If
 
  '***********************
  'Para refrescar los puntos acumulados
  
  On Error Resume Next
If txtIDCLIENTA.Text <> "" Then
lblPuntos.Visible = True
lblPuntosAcum.Visible = True
adoConsultaPuntos.CommandType = adCmdText
adoConsultaPuntos.RecordSource = "SELECT Last(Puntos.Fecha) AS ÚltimoDeFecha, Last(Puntos.Hora) AS ÚltimoDeHora, Last(Puntos.Acumulado) AS ÚltimoDeAcumulado From Puntos WHERE (((Puntos.IdClienta)=" & txtIDCLIENTA.Text & "));"
adoConsultaPuntos.Refresh
adoConsultaPuntos.Refresh
If adoConsultaPuntos.Recordset.Fields(2).Value <> "" Then

 lblFechaPuntos.Caption = adoConsultaPuntos.Recordset.Fields(0).Value
 
 diahoy = Format(LBLFECHA.Caption, "dd") + Format(LBLFECHA.Caption, "mm") * 30 + Format(LBLFECHA.Caption, "yy") * 365
 diaPuntos = Format(lblFechaPuntos.Caption, "dd") + Format(lblFechaPuntos.Caption, "mm") * 30 + Format(lblFechaPuntos.Caption, "yy") * 365
    If Abs(diahoy - diaPuntos) < 90 Then
    lblPuntosAcum.Caption = Format(adoConsultaPuntos.Recordset.Fields(2).Value, "0.00")
    
    cmdDescTodosPuntos.Visible = True
    cmdDescPartePuntos.Visible = True
    'If dlstTipo.BoundText = "SERVICIO" Then
    'cmdDescTodosPuntos.Enabled = False
    'cmdDescPartePuntos.Enabled = False
    'End If
    
    Else
    lblPuntosAcum.Caption = "0,00"
    End If
    
Else
lblPuntosAcum.Caption = "0,00"
End If
End If
 




 
 '******************************************************* '***********************
  'Para refrescar los puntos acumulados
  
  On Error Resume Next
If dlstTipo.BoundText = "Venta de CABINA" And txtIDCLIENTA.Text <> "" Then
lblPuntos.Visible = True
lblPuntosAcum.Visible = True
adoConsultaPuntos.CommandType = adCmdText
adoConsultaPuntos.RecordSource = "SELECT Last(Puntos.Fecha) AS ÚltimoDeFecha, Last(Puntos.Hora) AS ÚltimoDeHora, Last(Puntos.Acumulado) AS ÚltimoDeAcumulado From Puntos WHERE (((Puntos.IdClienta)=" & txtIDCLIENTA.Text & "));"
adoConsultaPuntos.Refresh
adoConsultaPuntos.Refresh
If adoConsultaPuntos.Recordset.Fields(2).Value <> "" Then

 lblFechaPuntos.Caption = adoConsultaPuntos.Recordset.Fields(0).Value
 
 diahoy = Format(LBLFECHA.Caption, "dd") + Format(LBLFECHA.Caption, "mm") * 30 + Format(LBLFECHA.Caption, "yy") * 365
 diaPuntos = Format(lblFechaPuntos.Caption, "dd") + Format(lblFechaPuntos.Caption, "mm") * 30 + Format(lblFechaPuntos.Caption, "yy") * 365
    If Abs(diahoy - diaPuntos) < 90 Then
    lblPuntosAcum.Caption = Format(adoConsultaPuntos.Recordset.Fields(2).Value, "0.00")
    
    cmdDescTodosPuntos.Visible = True
    cmdDescPartePuntos.Visible = True
    'If dlstTipo.BoundText = "SERVICIO" Then
    'cmdDescTodosPuntos.Enabled = False
    'cmdDescPartePuntos.Enabled = False
    'End If
    
    Else
    lblPuntosAcum.Caption = "0,00"
    End If
    
Else
lblPuntosAcum.Caption = "0,00"
End If
End If
 




 
 '*******************************************************
 
 
Exit Sub
errordato:
    MsgBox "No", vbInformation
    dcmbProducto(0).SetFocus
     MousePointer = 1
    Exit Sub
errorformato:
   
    adoIntro.Recordset.Fields(8).Value = txtPrecio.Text
    Resume Next
    Exit Sub




End If
'Pongo aquí el nuevo end if 21/04/2015
End If
End If
End Sub

Private Sub txtUnidades_KeyDown(KeyCode As Integer, Shift As Integer)

' BUSQUEDA DEL PRECIO DEL PRODUCTO

On Error GoTo errorconsultaprecio
' AQUI SE BUSCA

If dlstTipo.BoundText = "SERVICIO" Then
adoPrecio.CommandType = adCmdText
adoPrecio.RecordSource = "SELECT Producto.[Servicio?], Producto.Descripción, Producto.Precio From Producto WHERE (((Producto.[Servicio?])=1) AND ((Producto.Descripción)='" & dcmbProducto(0).Text & "'));"
adoPrecio.Refresh


Else






adoConsultaPrecio.CommandType = adCmdText
adoConsultaPrecio.RecordSource = "SELECT Producto.Descripción, Marca.Empresa, Last(Parcial.Precio) AS ÚltimoDePrecio, Last(Parcial.Fecha) AS ÚltimoDeFecha, Parcial.IdTipo FROM Marca INNER JOIN (Producto INNER JOIN Parcial ON Producto.IdProducto = Parcial.IdProducto) ON Marca.IdMarca = Parcial.IdMarca GROUP BY Producto.Descripción, Marca.Empresa, Parcial.IdTipo HAVING (((Producto.Descripción)= '" & dcmbProducto(0).Text & "' ) AND ((Marca.Empresa)='" & dcmbMarca(1).Text & "') AND ((Parcial.IdTipo)=1));"

'lo que pone debajo era lo que había antes y buscaba el valor más alto
'adoConsultaPrecio.RecordSource = "SELECT Producto.Descripción, Marca.Empresa, Max(Parcial.Precio) AS MáxDePrecio FROM Producto INNER JOIN (Marca INNER JOIN Parcial ON Marca.IdMarca = Parcial.IdMarca) ON Producto.IdProducto = Parcial.IdProducto GROUP BY Producto.Descripción, Marca.Empresa HAVING (((Producto.Descripción)= '" & dcmbProducto(0).Text & "' ) AND ((Marca.Empresa)='" & dcmbMarca(1).Text & "' ));"
adoConsultaPrecio.Refresh


End If
' AQUI SE ESCRIBE EN LA CAJA DE TEXTO

If KeyCode = KEY_RETURN Then txtPrecio.SetFocus

If dlstTipo.BoundText = "SERVICIO" Then
txtPrecio.Text = adoPrecio.Recordset.Fields(2).Value
Else

txtPrecio.Text = adoConsultaPrecio.Recordset.Fields(2).Value
If txtCB.Text = "" Then
lblFechaPrecio.Caption = adoConsultaPrecio.Recordset.Fields(3).Value
Else
lblFechaPrecio.Caption = adoConsultaPrecio.Recordset.Fields(3).Value
End If

End If

txtPrecio.SelStart = 0
txtPrecio.SelLength = Len(txtPrecio.Text)
If dlstTipo.BoundText = "Venta al PROFESIONAL" Then
lblPreuProfesional.Caption = Format(adoConsultaPrecio.Recordset.Fields(2).Value * 90 / 100, "0.00")
lblPreu.Visible = True
lblPreuProfesional.Visible = True
End If





Exit Sub

errorconsultaprecio:
    txtPrecio.Text = ""
   
    Exit Sub
    

    




End Sub
