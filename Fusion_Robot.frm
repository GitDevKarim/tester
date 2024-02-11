VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF18F2A4-8B30-11D3-A95C-00008639BD6E}#1.0#0"; "APToolkit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{14C63D93-4072-4593-88F9-D89858D7A88D}#1.0#0"; "DataMatrix.dll"
Begin VB.Form Fusion_Robot 
   Caption         =   "Robot ""Fusion"""
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   Icon            =   "Fusion_Robot.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   15210
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text_NUM_PREPARATION 
      Height          =   285
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame_Manual_Detail 
      Caption         =   "D�tail de la Fusion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8970
      Left            =   9360
      TabIndex        =   2
      Top             =   960
      Width           =   5655
      Begin VB.CommandButton Cmd_RepriseSMS 
         Caption         =   "Reprise SMS"
         Height          =   255
         Left            =   2400
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Cmd_Flow 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4080
         Picture         =   "Fusion_Robot.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Ouvrir le r�pertoire outpur du ""d�coupe PDF"""
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Text_NUM_FLOW 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text_NUM_RECEPTION 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2040
         Width           =   2175
      End
      Begin DATAMATRIXLibCtl.MW6DataMatrix MW6DataMatrixFusion 
         Height          =   495
         Left            =   4560
         TabIndex        =   44
         Top             =   2640
         Width           =   855
         BackColor       =   16777215
         BarColor        =   0
         BorderStyle     =   0
         Data            =   "12"
         ModuleSize      =   0,07
         Orientation     =   0
         Mode            =   0
         PreferredFormat =   0
         HandleTilde     =   0   'False
         _cx             =   1997407716
         _cy             =   1997407081
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   5400
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text_COMMENTAIRE 
         Height          =   1575
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   6360
         Width           =   3615
      End
      Begin VB.CommandButton CommandFaxTel 
         Caption         =   "Command3"
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CommandA3MIndex 
         Caption         =   "Command3"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CommandNumRecordEmission 
         Caption         =   "Command2"
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CommandUpdatePrice 
         Caption         =   "Command1"
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.Toolbar Toolbar_Manual_Detail 
         Height          =   810
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1429
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "Image_list"
         HotImageList    =   "Image_list"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Fusionner"
               Key             =   "ok"
               Object.ToolTipText     =   "Fusionner l'�l�ment s�lectionn�"
               ImageKey        =   "ok"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Annuler"
               Key             =   "cancel"
               Object.ToolTipText     =   "Annuler la fusion de l'�l�ment s�lectionn�"
               ImageKey        =   "cancel"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Erreur"
               Key             =   "error"
               Object.ToolTipText     =   "Passer l'�l�ment s�lectionn� en erreur"
               ImageKey        =   "cancel"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Arr�ter"
               Key             =   "stop"
               Object.ToolTipText     =   "Arr�ter le traitement"
               ImageKey        =   "stop"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "BAT"
               Key             =   "bat"
               ImageKey        =   "bat"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.TextBox Text_segment 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text_NOMBRE_ENREGISTREMENT_DATA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox Text_MAJ_USERID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   8520
         Width           =   2175
      End
      Begin VB.TextBox Text_MAJ_DATE 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   8160
         Width           =   2175
      End
      Begin VB.TextBox Text_NUM_SOCIETE 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox Text_NOM_FICHIER_INFO 
         Height          =   525
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text_NOM_FICHIER_DATA 
         Height          =   525
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4080
         Width           =   3615
      End
      Begin VB.TextBox Text_STATUT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text_DEBUT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox Text_FIN 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "# Flux"
         Height          =   255
         Left            =   165
         TabIndex        =   48
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label_NUM_RECEPTION 
         Caption         =   "# R�ception"
         Height          =   255
         Left            =   165
         TabIndex        =   46
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label_lot_suite 
         AutoSize        =   -1  'True
         Caption         =   "du lot d�coup� en pr�paration"
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Top             =   5160
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "pr�par�(s)"
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   4800
         Width           =   705
      End
      Begin VB.Label Label_lot 
         AutoSize        =   -1  'True
         Caption         =   "Nb. Enregistrement(s)"
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   5160
         Width           =   1515
      End
      Begin VB.Label Label_NOMBRE_ENREGISTREMENT_DATA 
         AutoSize        =   -1  'True
         Caption         =   "Nb. Enregistrement(s)"
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   4800
         Width           =   1515
      End
      Begin VB.Label Label_MAJ_USERID 
         AutoSize        =   -1  'True
         Caption         =   "Par"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   8565
         Width           =   240
      End
      Begin VB.Label Label_MAJ_DATE 
         AutoSize        =   -1  'True
         Caption         =   "Mise � jour"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   8205
         Width           =   765
      End
      Begin VB.Label Label_DEBUT 
         Caption         =   "Date D�but"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label_FIN 
         Caption         =   "Date Fin"
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label_NUM_SOCIETE 
         Caption         =   "# Client"
         Height          =   255
         Left            =   165
         TabIndex        =   15
         Top             =   2895
         Width           =   1095
      End
      Begin VB.Label Label_NOM_FICHIER_INFO 
         Caption         =   "Fichier info"
         Height          =   255
         Left            =   165
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label_NOM_FICHIER_DATA 
         Caption         =   "Fichier Data"
         Height          =   255
         Left            =   165
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label_STATUT 
         Caption         =   "Statut"
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   1335
         Width           =   1095
      End
      Begin VB.Label Label_NUM_PREPARATION 
         Caption         =   "# Pr�paration"
         Height          =   255
         Left            =   165
         TabIndex        =   11
         Top             =   2415
         Width           =   1095
      End
      Begin VB.Label Label_COMMENTAIRE 
         Caption         =   "Commentaire"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   6360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Fusion_Robot.frx":053E
      Height          =   8970
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   15822
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
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
      Caption         =   "Suivi des Fusions"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "PK_PREPARATION"
         Caption         =   "PK_PREPARATION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "num_preparation"
         Caption         =   "# Pr�paration"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "num_societe"
         Caption         =   "# Client"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PRESTATION_MODEL_NOM"
         Caption         =   "Prestation / Mod�le"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "status"
         Caption         =   "Statut"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "NOMBRE_DATA"
         Caption         =   "Nb"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1695,118
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1695,118
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   3195,213
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar_Manual_Header 
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   1429
      ButtonWidth     =   1614
      ButtonHeight    =   1376
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Image_list"
      HotImageList    =   "Image_list"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sp1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archives"
            Key             =   "archives"
            ImageKey        =   "archives"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Trier"
            Key             =   "sort"
            Object.ToolTipText     =   "D�finir les crit�res de tri"
            ImageKey        =   "sort"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rafraichir"
            Key             =   "refresh"
            Object.ToolTipText     =   "Rechercher les nouveaux fichiers � Pr�parer"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sp2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Les erreurs"
            Key             =   "all_error"
            Object.ToolTipText     =   "Reprendre les erreurs"
            ImageKey        =   "links"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sp3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "En Boucle"
            Key             =   "boucle"
            ImageKey        =   "Boucle"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Stop Boucle"
            Key             =   "stop"
            ImageKey        =   "stop"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fermer"
            Key             =   "closew"
            Object.ToolTipText     =   "Fermer la fen�tre"
            ImageKey        =   "closew"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.OptionButton Option_Type 
         Caption         =   "Que PrintWord"
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
         Height          =   255
         Index           =   6
         Left            =   9720
         TabIndex        =   37
         Top             =   0
         Width           =   1755
      End
      Begin MSAdodcLib.Adodc Adodc_Tool2 
         Height          =   330
         Left            =   12480
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Caption         =   "Adodc_Tool2"
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
      Begin VB.TextBox Text_Systeme_Status_Data 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   25
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox Text_Systeme_Status_Info 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   24
         Top             =   120
         Width           =   5295
      End
      Begin VB.OptionButton Option_Type 
         Caption         =   "Que TrustOffice"
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
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   36
         Top             =   480
         Width           =   1755
      End
      Begin VB.OptionButton Option_Type 
         Caption         =   "Que ePobox"
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
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   35
         Top             =   240
         Width           =   1635
      End
      Begin VB.OptionButton Option_Type 
         Caption         =   "Pas de Word"
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
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   34
         Top             =   0
         Width           =   1515
      End
      Begin MSAdodcLib.Adodc Adodc_Tmp 
         Height          =   330
         Left            =   12360
         Top             =   -120
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Caption         =   "Adodc_preparation"
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
      Begin MSAdodcLib.Adodc Adodc_Tool 
         Height          =   330
         Left            =   12960
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Caption         =   "Adodc_preparation"
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
      Begin VB.OptionButton Option_Type 
         Caption         =   "Sauf e-mail"
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
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   33
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton Option_Type 
         Caption         =   "Que Fax"
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
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   32
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton Option_Type 
         Caption         =   "Tout"
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
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   31
         Top             =   0
         Value           =   -1  'True
         Width           =   1395
      End
      Begin MSAdodcLib.Adodc Adodc_preparation 
         Height          =   330
         Left            =   11280
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Caption         =   "Adodc_preparation"
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
      Begin MSComctlLib.ImageList Image_list 
         Left            =   10680
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":055E
               Key             =   "closew"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":07F2
               Key             =   "links"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":1444
               Key             =   "receive"
               Object.Tag             =   "posteasy_receiver"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":2096
               Key             =   "scan"
               Object.Tag             =   "scan"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":2CE8
               Key             =   "ok"
               Object.Tag             =   "ok"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":2F7A
               Key             =   "cancel"
               Object.Tag             =   "cancel"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":320C
               Key             =   "sort"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":349E
               Key             =   "stop"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":38F0
               Key             =   "refresh"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":3A4A
               Key             =   "Boucle"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":3D64
               Key             =   "bat"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Fusion_Robot.frx":3FF6
               Key             =   "archives"
            EndProperty
         EndProperty
      End
      Begin PETOCXLib.PETOCX PETOCX1 
         Left            =   8160
         Top             =   0
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   0
      End
   End
End
'----------------------------------------------------------------------------------------------------------------
Attribute VB_Name = "Fusion_Robot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim L_tab_Champs_Emis_detail()          As T_Champs_Emis_Detail
    Dim L_tab_Champs_Emis()                 As T_Champs_Emis
    Dim L_tab_Champs_Emis_Regroupement()    As T_Champs_Emis_Regroupement
    Dim L_tab_Champs_Lies()                 As T_Champs_Lies
    
    Dim L_Fichier_Joint()                   As t_Signet
    
    'Dim M_Make_Sequestre                    As Boolean                     'LCI SUPPRESSION SEQUESTRE LE 14/08/2017 => TOUT EST DEJA SEQUESTR�
    
Rem v478
    'Dim M_Make_Xades                        As Boolean                     'LCI SUPPRESSION XADES LE 14/08/2017
Rem v478
Rem v491
    Dim M_Make_PdfS                         As Boolean
    'Dim M_Make_Worm                         As Boolean                     'LCI SUPPRESSION WORM LE 14/08/2017
Rem v491
    
Rem v439
    'Dim M_Make_Archivage                    As Boolean
Rem v439

    Private Type T_ePoBox_FileToXfer
        t_OriginalFile  As String 'Nom du fichier Original
        t_NewFile       As String 'Nouveau nom (id_pe?+Occurence)
        t_Chemin        As String 'Chemin d'origine
    End Type
    
    Private Type T_ePoBox_Data
        t_Field_Fk      As Long
        t_Field_Value   As String
        t_Field_Type    As String 'Valeurs autoris�es: "Data", "File"
    End Type
    
    Dim L_tab_ePoBox_Data()                                   As T_ePoBox_Data
    
    
Rem v430
    'Public Type T_ePoBox_Facture
    '    field_name      As String
    '    field_value     As String
    'End Type

    'Dim L_tab_ePoBox_Facture()                                   As T_ePoBox_Facture
Rem v430
    
    
    
    Rem d�claration des variables
    Dim M_Mode                  As String
    
Rem v423
    Dim M_Fk_Prestation_eMail   As Long
    Dim M_Fk_Prestation_Fax     As Long
    Dim M_Fk_Prestation_ePobox  As Long
Rem v423

Rem v431
    Dim M_Nb                    As Long
Rem v431

Rem v439 Optimisations
    Public P_Fk_Pli_Statut_AAFFC        As Integer 'A affecter
    Public P_Fk_Pli_Statut_ENPRD        As Integer 'En production
    Public P_Fk_Pli_Statut_ENWAI        As Integer 'Encours � valider
    Public P_Fk_Pli_Statut_ECVAL        As Integer 'Encours � valider
    Public P_Fk_Pli_Statut_AGROU        As Integer 'A regrouper
    Public P_Fk_Pli_Statut_AVALI        As Integer 'A valider par le client
Rem v439 Optimisations
Rem v480
    Public P_Fk_Pli_Statut_ARCH         As Integer 'Archiv� (pour prestation de type EDAT)
Rem v480
Rem v521
    Public WithEvents PDFCreator1 As PdfCreator.clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1
    Public pErr As clsPDFCreatorError, opt As clsPDFCreatorOptions
    Public noStart As Boolean, fac As Double, StartTime As Date
'--------------------------------------------------------------------------------------------------
Private Sub Cmd_Flow_Click()

Dim DecoupePDF As String


    On Error Resume Next
    If Me.Text_NUM_FLOW.tag > 0 Then
        Rem Lecture du chemin
        DecoupePDF = Trim(Lire_Un_Champ("concat('\\',  date_format(date_creation, '%Y%m%d'), '\\', num_societe, '\\', num_flow, '\\', jointFileDir)", "flow, societe", "pk_societe = fk_societe and Pk_flow = " & Me.Text_NUM_FLOW.tag))
        If DecoupePDF <> "" Then
            DecoupePDF = G_dir_sequestre & DecoupePDF
            ShellExecute 0, "open", DecoupePDF, "", "", 1
            Exit Sub
        End If
    End If

End Sub

Private Sub Cmd_RepriseSMS_Click()
    Dim L_Smtp                  As SmtpX
    Dim MessageObject           As Message
    Dim L_Result                As String
    Dim L_Chemin_To             As String
    Dim SQL                     As String
    
    Set L_Smtp = New SmtpX
    
    L_Result = Init_Smtp_New(L_Smtp, "Sms", "")
    
    If L_Result <> "Ok" Then
        MsgBox ("Erreur de connexion SMTP (" & L_Result & ")")
        Exit Sub
    End If
    
    G_dir_Web_PDF = Lire_Un_Champ("REPERTOIRE", "REPERTOIRE", "CODE = 'WEB'")
    
    Call Init_Adodc(Me.Adodc_Tmp)
    SQL = " SELECT pl.chemin_pli, pl.id_pe"
    SQL = SQL & " FROM pli pl, prestation_model pm"
    SQL = SQL & " WHERE pl.fk_prestation_model = pm.pk_prestation_model"
    SQL = SQL & " AND pm.fk_type_prestation = 9"
    SQL = SQL & " AND pl.fk_pli_statut = 43"
    SQL = SQL & " AND pl.date_emission >= '2019-04-16'"
    SQL = SQL & " AND pl.date_emission < '2019-04-17'"
    
    Me.Adodc_Tmp.RecordSource = SQL
    Me.Adodc_Tmp.Refresh
    
    While Not Me.Adodc_Tmp.Recordset.EOF
        L_Chemin_To = G_dir_Web_PDF & "\" & Me.Adodc_Tmp.Recordset("CHEMIN_PLI")
        
        If Not FolderExists(L_Chemin_To, True) Then
            MsgBox ("Le r�pertoire dans lequel est le mail � envoyer n'existe pas! (""" & L_Chemin_To & """)")
            Exit Sub
        End If
        
        Set MessageObject = New Message
        
        If FileExists(L_Chemin_To & "\" & Me.Adodc_Tmp.Recordset("ID_PE") & "_mail.eml") Then
            MessageObject.Read L_Chemin_To & "\" & Me.Adodc_Tmp.Recordset("ID_PE") & "_mail.eml"
        Else
            MsgBox ("Le fichier """ & L_Chemin_To & "\" & Me.Adodc_Tmp.Recordset("ID_PE") & "_mail.eml"" est introuvable!")
            Exit Sub
        End If
        
        L_Smtp.SendMessage MessageObject
        Set MessageObject = Nothing
        
        Me.Adodc_Tmp.Recordset.MoveNext
     Wend
     
     MsgBox ("Reprise SMS termin�e")
End Sub

Private Sub Command1_Click()

    Dim nb As Long
    Dim msg As String
    
    msg = "0123456789012345678901234567890123456789012345678901234567890123456789"
    msg = msg & msg & "012345678901234567890"
    
    nb = ReadNbSms(msg)
    

End Sub

Private Sub CommandUpdatePrice_Click()


    Rem Mise � jour des PRIX !!!!
    
    If Date > CDate("24/10/2014") Then
        Me.CommandUpdatePrice.Visible = False
        Exit Sub
    End If
    
    'stop
    Dim IdPe                            As String
    Dim L_SheetCount                    As Long
    Dim L_PageCount                     As Long
    Dim p_Prestation_Model_Pk           As Long
    Dim L_Poids_Pli                     As Double
    Dim L_Pli_Code_pays_Iso3A           As String
    Dim L_Pli_Service_Postal            As String
    Dim L_Pli_Zone_Postale              As String
    Dim L_Pli_Adresse                   As String
    Dim L_FaxSmsNumber                  As String
    Dim L_Service_Transformation_fax    As Boolean
    Dim L_Mnt_Service                   As Double
    Dim L_Mnt_Affranchissement          As Double
    Dim RectoVerso                      As Boolean
    Dim nb                              As Long
    Dim c                               As Long
    Dim typeEnvoi                       As String
    Dim PkPli                           As Long
    Dim SQL                             As String
    Dim L_Mnt_Total                     As Double
    
    
D�but:
    Call Init_Adodc(Me.Adodc_Tmp)
    SQL = " SELECT p.ID_PE, p.pk_pli, pm.pk_prestation_model, p.nombre_pages, p.page_count, p.poids_pli, p.prix_pli, "
    SQL = SQL & " d.dest_adresse, d.dest_cp, d.dest_ville, d.dest_pays_nom, pm.type_impression"
    SQL = SQL & " FROM pli p, destinataire d, prestation_model pm"
    SQL = SQL & " WHERE pk_destinataire = p.fk_destinataire "
    SQL = SQL & " AND p.date_emission >= '2015-01-01' "
    Stop
    SQL = SQL & " AND p.fk_societe = 5659 "
    'Stop
    'SQL = SQL & " AND p.fk_societe in () "
    SQL = SQL & " AND p.fk_prestation_model = pm.pk_prestation_model"
    SQL = SQL & " ORDER BY p.pk_pli "
    
    Rem Original = SELECT * FROM pli, destinataire WHERE pk_destinataire = fk_destinataire AND date_emission > '2010-01-01' AND pli_zone_postale = '' order by pk_pli
    Me.Adodc_Tmp.RecordSource = SQL
    Me.Adodc_Tmp.Refresh
    If Me.Adodc_Tmp.Recordset.EOF Then
        Stop
    End If
    
    nb = Me.Adodc_Tmp.Recordset.RecordCount
    c = 0
    While Not Me.Adodc_Tmp.Recordset.EOF
        'stop
        c = c + 1
        Me.Caption = c & " / " & nb
        DoEvents
        typeEnvoi = Lire_Un_Champ("code", "type_prestation, prestation_model", "fk_type_prestation = pk_type_prestation and pk_prestation_model = " & Me.Adodc_Tmp.Recordset("pk_prestation_model"))
        'If typeEnvoi = "Fax" Or typeEnvoi = "Sms" Or typeEnvoi = "EDA" Or typeEnvoi = "MEL" Then
        'If typeEnvoi = "MEL" Then
        '    GoTo Suiv1
        'End If
        'If typeEnvoi = "Fax" Or typeEnvoi = "Sms" Or typeEnvoi = "EDA" Then
        '    stop
        '    GoTo Suivant
        'End If
        'If typeEnvoi <> "LS" And typeEnvoi <> "LSR" And typeEnvoi <> "LAR" And typeEnvoi <> "EPB" Then
        '    stop
        'End If
Suiv1:
        IdPe = Me.Adodc_Tmp.Recordset("ID_PE")
        L_SheetCount = CLng(Me.Adodc_Tmp.Recordset("nombre_pages"))
        L_PageCount = CLng(Me.Adodc_Tmp.Recordset("page_count"))
        p_Prestation_Model_Pk = Me.Adodc_Tmp.Recordset("pk_prestation_model")
        L_Poids_Pli = Me.Adodc_Tmp.Recordset("poids_pli")
        L_Pli_Adresse = Me.Adodc_Tmp.Recordset("dest_adresse") & vbNewLine & Me.Adodc_Tmp.Recordset("dest_cp") & " " & Me.Adodc_Tmp.Recordset("dest_ville") & vbNewLine & Me.Adodc_Tmp.Recordset("dest_pays_nom")
        RectoVerso = Me.Adodc_Tmp.Recordset("type_impression") <> "Recto"
        L_Mnt_Total = PrixPliGlobal(IdPe, L_SheetCount, _
                                                    p_Prestation_Model_Pk, _
                                                    L_Poids_Pli, _
                                                    L_Mnt_Service, _
                                                    L_Mnt_Affranchissement, _
                                                    L_Pli_Code_pays_Iso3A, _
                                                    L_Pli_Service_Postal, _
                                                    L_Pli_Zone_Postale, _
                                                    L_Pli_Adresse, "", _
                                                    Me, L_PageCount, RectoVerso)
                                                    
                PkPli = Me.Adodc_Tmp.Recordset("pk_pli")
                
                SQL = " update pli set "
                
                'SQL = SQL & " prix_affranchissement = " & Replace(L_Mnt_Affranchissement, ",", ".") & ", "
                'SQL = SQL & " pli_code_pays_iso3A = '" & L_Pli_Code_pays_Iso3A & "', "
                'SQL = SQL & " pli_zone_postale = '" & L_Pli_Zone_Postale & "' "
                SQL = SQL & " prix_pli = " & Replace(L_Mnt_Service, ",", ".")
                SQL = SQL & " where pk_pli = " & PkPli
                If Run_Execute_Sql(SQL) = -1 Then
                    Stop
                End If
                
                SQL = " update pli_prix set "
                SQL = SQL & " service_price_ht = " & Replace(L_Mnt_Service, ",", ".")
                SQL = SQL & " where fk_pli = " & PkPli
                If Run_Execute_Sql(SQL) = -1 Then
                    Stop
                End If
                
Suivant:
        Me.Adodc_Tmp.Recordset.MoveNext
    Wend
    
    MsgBox "Termin� � " & Now & " - Nombre de mise � jour : " & nb
    GoTo D�but
End Sub

Private Sub CommandNumRecordEmission_Click()

    Rem NumRecordEmission
    
    Dim SQL As String
    Dim PkPli As Long
    Dim c As Integer
    Dim nb  As Long
    Dim R   As String
    Dim NbR As Integer
    
          SQL = " select pk_pli, num_record_emission, dest_ref_client, dest_ref_comptable"
    SQL = SQL & " From pli, destinataire"
    SQL = SQL & " Where pli.fk_societe = 1602"
    SQL = SQL & " and pk_destinataire = fk_destinataire"
    SQL = SQL & " order by pk_pli desc"
    SQL = SQL & " limit 1500 "
    Call Init_Adodc(Me.Adodc_Tmp)
    Me.Adodc_Tmp.RecordSource = SQL
     Me.Adodc_Tmp.Refresh
    If Me.Adodc_Tmp.Recordset.EOF Then
        Stop
    End If

    nb = Me.Adodc_Tmp.Recordset.RecordCount
    While Not Me.Adodc_Tmp.Recordset.EOF
        'stop
        c = c + 1
        NbR = 1
        If Me.Adodc_Tmp.Recordset("num_record_emission") = 0 Then
            R = Me.Adodc_Tmp.Recordset("dest_ref_client")
            While InStr(1, R, "|", vbTextCompare) > 0
                R = Mid(R, InStr(1, R, "|", vbTextCompare) + 1)
                NbR = NbR + 1
            Wend
            Rem Mise � jour
            R = Run_Execute_Sql("update pli set num_record_emission = " & NbR & " where pk_pli = " & Me.Adodc_Tmp.Recordset("pk_pli"))
            If R <> "0" Then
                Stop
            End If
        End If
        Me.Caption = c & " / " & nb
        DoEvents
        Me.Adodc_Tmp.Recordset.MoveNext
    Wend
    
    MsgBox "Fini � " & Now & " !"
    Stop
    

End Sub

Private Sub CommandA3MIndex_Click()


Rem SPECIAL A3M Index

Stop
Exit Sub

    Dim L_Fnum  As Integer
    Dim L_Fnom  As String
    Dim nb      As Long
    Dim Line    As String
    Dim TLine() As String
    Dim SqlUpdate   As String
    Dim SlqSearch   As String
    Dim PkDestinataire  As Long
    Dim Tmp         As String
    
    Call Init_Adodc(Me.Adodc_Tmp)
    
    L_Fnum = FreeFile
    L_Fnom = "\\datamaster\data\Reception\033411005551\20131008\A3M.txt"
    Open L_Fnom For Input As #L_Fnum
    While Not EOF(L_Fnum)
NextLine:
        Line Input #L_Fnum, Line
        TLine = Split(Line, ";")
        nb = nb + 1
        If UBound(TLine) <> 22 Then
            Stop
        End If
        If TLine(0) = "SEQ" Then
            GoTo NextLine
        End If
        SlqSearch = " destinataire.fk_societe = 5341 "
        SlqSearch = SlqSearch & " and dest_ref_client = '" & TLine(2) & "' "
        SlqSearch = SlqSearch & " and dest_ref_client2 = '" & TLine(13) & "' "
        SlqSearch = SlqSearch & " and dest_ref_comptable = '" & TLine(3) & "' "
        SlqSearch = SlqSearch & " and destinataire.maj_date > '2013-10-03'"
        'SlqSearch = SlqSearch & " and fk_destinataire = pk_destinataire "
        SlqSearch = SlqSearch & " and destinataire.r1 = ''"
        SlqSearch = SlqSearch & " and pk_destinataire >= 32273265 "

        Tmp = Lire_Un_Champ("count(*)", "destinataire", SlqSearch)
        If Tmp = 0 Then
            GoTo Suivant
        ElseIf Tmp = 1 Then
            PkDestinataire = Lire_Un_Champ("pk_destinataire", "destinataire", SlqSearch)
        Else
            'Stop
            PkDestinataire = Lire_Un_Champ("min(pk_destinataire)", "destinataire", SlqSearch)
        End If
        SqlUpdate = " Update destinataire set "
        SqlUpdate = SqlUpdate & " R1 = '" & TLine(0) & "', "
        If Not IsNumeric(TLine(3)) Then
            SqlUpdate = SqlUpdate & " dest_ref_client = '" & TLine(3) & "',  "
        Else
            SqlUpdate = SqlUpdate & " dest_ref_client = '" & CLng(TLine(3)) & "',  "
        End If
        SqlUpdate = SqlUpdate & " dest_ref_comptable = '" & TLine(2) & "',  "
        SqlUpdate = SqlUpdate & " dest_nom = '" & Valid_Text(TLine(20)) & "',  "
        SqlUpdate = SqlUpdate & " dest_prenom = '" & Valid_Text(TLine(19)) & "',  "
        SqlUpdate = SqlUpdate & " maj_date = now(),  "

        'SqlUpdate = SqlUpdate & " maj_userid = 'prod41' "
        SqlUpdate = SqlUpdate & " maj_userid = 'prod4" & Tmp & "' "
        
        SqlUpdate = SqlUpdate & " where pk_destinataire = " & PkDestinataire
        Tmp = Run_Execute_Sql(SqlUpdate)
        If Tmp <> 0 Then
            Stop
        End If
Suivant:
        Me.Caption = nb - 1
        DoEvents
        
    Wend
    MsgBox "Termin�"
    
    
    
End Sub

Private Sub CommandFaxTel_Click()

    Dim IdPe As String
    Dim L_SheetCount    As Long
    Dim L_PageCount    As Long
    Dim p_Prestation_Model_Pk   As Long
    Dim L_Poids_Pli As Double
    Dim L_Pli_Code_pays_Iso3A As String
    Dim L_Pli_Service_Postal As String
    Dim L_Pli_Zone_Postale As String
    Dim L_Pli_Adresse As String
    Dim L_FaxSmsNumber  As String
    Dim L_Service_Transformation_fax As Boolean
    Dim L_Mnt_Service               As Double
    Dim L_Mnt_Affranchissement      As Double
    Dim RectoVerso                  As Boolean
    
    Dim nb As Long
    Dim c As Long
    Dim typeEnvoi As String
    Dim PkPli As Long
    Dim SQL As String
    Dim L_Mnt_Total As Double
    Dim Tel As String
    Dim PkContact As Long
    Dim Update As Boolean
    Dim Champ As String
    Dim Table   As String
    
    
    
    Rem MISE � jour des coordonn�es t�l�phone + fax
    
    Champ = "tel_gsm"
    Champ = "fax"
    Champ = "tel_bureau"
    
    Table = "contact"
    
    
    Champ = "fax"
    Champ = "telephone"
    Table = "etablissement"
    
    
    Call Init_Adodc(Me.Adodc_Tmp)
    Me.Adodc_Tmp.RecordSource = "select pk_" & Table & ", " & Champ & " from " & Table & " where " & Champ & " is not null order by pk_" & Table
    Me.Adodc_Tmp.Refresh
    If Me.Adodc_Tmp.Recordset.EOF Then
        Stop
    End If
    
    nb = Me.Adodc_Tmp.Recordset.RecordCount
    c = 0
    While Not Me.Adodc_Tmp.Recordset.EOF
        'stop
        c = c + 1
        Me.Caption = c & " / " & nb
        Update = False
        DoEvents
        Tel = Me.Adodc_Tmp.Recordset(Champ)
        
TestTel:
        If InStr(1, Tel, " ", vbTextCompare) > 0 Then
            Tel = Replace(Tel, " ", "", , , vbTextCompare)
            Update = True
        End If
        
        If InStr(1, Tel, "-", vbTextCompare) > 0 Then
            Tel = Replace(Tel, "-", "", , , vbTextCompare)
            Update = True
        End If
        
        If InStr(1, Tel, ".", vbTextCompare) > 0 Then
            Tel = Replace(Tel, ".", "", , , vbTextCompare)
            Update = True
        End If
        
        If InStr(1, Tel, "+", vbTextCompare) > 0 Then
            Tel = Replace(Tel, "+", "", , , vbTextCompare)
            Update = True
        End If
        
        If Tel = "" Then
            Tel = " null "
            GoTo GoUpdate
        End If
        
        If Not IsNumeric(Tel) Then
            If Left(Tel, 5) = "33(0)" Then
                Tel = "33" & Mid(Tel, 6)
                GoTo TestTel
            End If
            If Left(Tel, 3) = "(0)" Then
                Tel = "33" & Mid(Tel, 4)
                GoTo TestTel
            End If
            If Tel = "?" Then
                Tel = " null "
                GoTo GoUpdate
            End If
            If Tel = "dd" Then
                Tel = " null "
                GoTo GoUpdate
            End If
            If Tel = "a" Then
                Tel = " null "
                GoTo GoUpdate
            End If
            Stop
            Tel = " null "
            GoTo GoUpdate
        End If
        
        If Left(Tel, 1) = "0" Then
            Tel = "33" & Mid(Tel, 2)
            Update = True
        End If
        
        If Len(Tel) <> 11 Then
            If Len(Tel) = 12 And Left(Tel, 2) = "33" Then
                GoTo SuiteLen
            End If
            If Len(Tel) = 9 Then
                Tel = "33" & Tel
                Update = True
                GoTo SuiteLen
            End If
            If Len(Tel) < 9 Then
                Tel = " null"
                Update = True
                GoTo SuiteLen
            End If
            If Len(Tel) = 10 And Left(Tel, 2) = "33" Then
                Tel = " null"
                Update = True
                GoTo SuiteLen
            End If
            'Stop
            'Tel = " null"
            'Update = True
            'GoTo SuiteLen
        End If

SuiteLen:
        
        
        If Update Then
GoUpdate:
            SQL = " update " & Table & " set " & Champ & " = " & Tel & " where pk_" & Table & " = " & Me.Adodc_Tmp.Recordset("pk_" & Table)
            If Run_Execute_Sql(SQL) = -1 Then
                Stop
            End If
        End If
Suivant:
        Me.Adodc_Tmp.Recordset.MoveNext
    Wend
    Stop

End Sub

Private Sub Form_Activate()

    If G_AutoStart Then
        G_AutoStart = False
        Call Run_Execute_Sql("update preparation set status = 'Erreur fusion' where status = 'En fusion' and maj_userid = '" & G_User_Id & "'")
        Call Toolbar_Manual_Header_ButtonClick(Me.Toolbar_Manual_Header.Buttons("boucle"))
    End If
    
End Sub

Rem

Private Sub Form_Load()

    On Error GoTo Error:

    M_Mode = vbNullString
    Me.Height = 3125
    
    Rem Set the screen coordinates
    Fusion_Robot.Caption = Fusion_main.Caption
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    Call Init_Adodc(Adodc_preparation)
    Call Init_Data(G_Selection)
    Call Load_Pk_Fusion
    Rem v599
    G_Sae_Demo_Mail_To = Init_Parameter("G_Sae_Demo_Mail_To")
    Rem v599
    Me.BackColor = Fusion_main.BackColor
    noStart = True
    Set PDFCreator1 = New clsPDFCreator
    Set pErr = New clsPDFCreatorError
    With PDFCreator1
        .cVisible = True
        If .cStart("/NoProcessingAtStartup") = False Then
            If .cStart("/NoProcessingAtStartup", True) = False Then
                Exit Sub
            End If
            .cVisible = True
        End If
        Set opt = .cOptions
        .cClearCache
        noStart = False
    End With
    G_RobotNextControlerActivity = DateAdd("n", 7, Now)

Exit Sub

Rem Routine d'erreur
Error:
    G_msg_error_p1 = Err.Number
    G_msg_error_p2 = Err.Description
    G_msg_error_p3 = Me.Name
    G_msg_error_p4 = "Form_Load"
'stop 'Debug AUTO LCI
    G_msg_error_p5 = Adodc_preparation.RecordSource
    G_msg_error_p6 = vbNullString
    error_manager.Show vbModal
    'Call CleanAndQuit
    Unload Me
    'End
    
End Sub

Private Sub Load_Pk_Fusion()

    M_Fk_Prestation_eMail = Lire_Un_Champ("PK_TYPE_PRESTATION", "TYPE_PRESTATION", "CODE = 'MEL'")
    M_Fk_Prestation_Fax = Lire_Un_Champ("PK_TYPE_PRESTATION", "TYPE_PRESTATION", "CODE = 'FAX'")
    M_Fk_Prestation_ePobox = Lire_Un_Champ("PK_TYPE_PRESTATION", "TYPE_PRESTATION", "CODE = 'EPB'")
    P_Fk_Pli_Statut_ARCH = Fk_Pli_Statut("ARCH")
    P_Fk_Pli_Statut_AAFFC = Fk_Pli_Statut("AAFFC")
    P_Fk_Pli_Statut_ENPRD = Fk_Pli_Statut("ENPRD")
    P_Fk_Pli_Statut_ENWAI = Fk_Pli_Statut("ENWAI")
    P_Fk_Pli_Statut_ECVAL = Fk_Pli_Statut("ECVAL")
    P_Fk_Pli_Statut_AGROU = Fk_Pli_Statut("AGROU")
    P_Fk_Pli_Statut_AVALI = Fk_Pli_Statut("AVALI")
        
End Sub

Private Sub Form_Resize()

    'If Me.Height > 10515 Then
    '    Me.Height = 10515
    'End If
    
    DataGrid1.Height = IIf(Me.Height > 1600, Me.Height - 1600, 0)
    Frame_Manual_Detail.Height = DataGrid1.Height
    'Toolbar_Manual_Detail.Top = IIf(Me.Frame_Manual_Detail.Height > 1000, Me.Frame_Manual_Detail.Height - 1000, 0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    PDFCreator1.cClearCache
    If noStart = False Then
        DoEvents
        PDFCreator1.cClose
    End If
    DoEvents
    Set PDFCreator1 = Nothing
    Set pErr = Nothing
    Set opt = Nothing
    Call KillProcessus("PDFCREATOR.EXE")

End Sub

Private Sub Option_Type_Click(Index As Integer)

    Dim Selection As String

    Select Case Index
    Case 0 'TOUT
        G_Selection = ""
    Case 1 'FAX
        G_Selection = "FAX"
    Case 2 'MAIL
        G_Selection = "NOMAIL"
    Case 3 'PAS WORD
        G_Selection = "SANS WORD"
    Case 4 'EPOBOX
        G_Selection = "EPOBOX"
    Case 5 'TRUSTOFFICE
        G_Selection = "TRUST OFFICE"
    Case 6 'IMPRESSION WORD
        G_Selection = "WORD"
    End Select
    Call Init_Data(G_Selection)
    
End Sub

Private Sub Toolbar_Manual_Detail_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim L_tmp       As Long
    Dim L_tmp_pm    As Long
    Dim L_reprise   As Boolean
    Dim L_grid_pos  As Long
    Dim i           As Long
    Dim L_pos       As Long
    
    If Me.Adodc_preparation.Recordset.EOF Then
        Call Toolbar_Set_All_Detail(vbFalse)
        Exit Sub
    End If

    L_tmp = Me.Adodc_preparation.Recordset("PK_PREPARATION")
    L_tmp_pm = Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")
    
    Select Case Button.Key
    'Case "ok", "bat"
    Case "ok"
        'If Button.Key = "bat" Then
        '    G_BAT = True
        'Else
        '    G_BAT = False
        'End If
            
        Screen.MousePointer = vbHourglass
        If Me.Adodc_preparation.Recordset("Status") = "Erreur" _
        Or Me.Adodc_preparation.Recordset("Status") = "Avort�" _
        Or Me.Adodc_preparation.Recordset("Status") = "Avort� fusion" _
        Or Me.Adodc_preparation.Recordset("Status") = "Erreur fusion" Then
            L_reprise = True
        Else
            L_reprise = False
        End If
        
        If "En fusion" = Lire_Un_Champ("STATUS", "PREPARATION", "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION")) Then
            MsgBox "Cette fusion est ex�cut�e par un autre utilisateur", vbInformation, Me.Caption
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        Select Case G_flux_type
        Case "ALL"
            Rem On traite tous les flux... Sauf ceux exclus
            If Is_In_List_To_Exclure(Me.Adodc_preparation.Recordset("FK_SOCIETE"), _
                                     G_flux_customer_pk, _
                                     ",") = True Then
                GoTo Suivant
            End If
            
        Case "REF"
            Rem
            If Not Is_Client_Reference(Me.Adodc_preparation.Recordset("NUM_SOCIETE")) Then
                GoTo Suivant
            End If
        Case "SEL", "SIGMA_EASYLINK"
            If Not Client_Selectionne(Me.Adodc_preparation.Recordset("FK_SOCIETE"), G_flux_customer_pk) Then
                GoTo Suivant
            End If
        End Select

        Rem -------------------------------------------------------------------------------------------
        'Rem Attention v�rification qu'il existe un Service Impression pour le(s) produit(s)
        'Rem Qui compose la Prestation mod�le S�lectionn�e
        If Existence_Service_Pe(L_tmp_pm, G_CONST_SERVICE_FUSION) Then
            If Not PDFCreator_Driver_Ok Then
                Exit Sub
            End If
        End If
        
        Me.Adodc_preparation.Refresh
        If Not Adodc_preparation.Recordset.EOF Then
            Adodc_preparation.Recordset.MoveFirst
            Adodc_preparation.Recordset.Find "PK_PREPARATION = " & L_tmp
        End If
            
        For i = 1 To Me.Toolbar_Manual_Header.Buttons.Count
            Me.Toolbar_Manual_Header.Buttons(i).Enabled = False
        Next
        For i = 1 To Me.Toolbar_Manual_Detail.Buttons.Count
            Me.Toolbar_Manual_Detail.Buttons(i).Enabled = False
        Next


        DataGrid1.Enabled = False
        Rem v641
        If ReserverUnEnregistrement("PREPARATION", "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), "STATUS", "Pr�par�2|Pr�par�|Avort� fusion|Erreur fusion", "En fusion") = False Then
        'If Modifier_un_Statut("PREPARATION", "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), "STATUS", "Pr�par�2|Pr�par�|Avort� fusion|Erreur fusion", "En fusion") = False Then
            GoTo Suivant
        End If
        Rem v641
        Call Init_Adodc(Me.Adodc_Tool)
        Call StartApplication
        Call Fusion_Run(Me.Adodc_preparation.Recordset("NUM_PREPARATION"), _
                        Me.Adodc_preparation.Recordset("FK_RECEPTION"), _
                        Me.Adodc_preparation.Recordset("NUM_SOCIETE"), _
                        L_tmp, _
                        L_reprise, _
                        Me.Adodc_preparation.Recordset("EPOBOX_IN"), _
                        Me, _
                        Me.Adodc_preparation.Recordset("PK_PRESTATION_MODEL"), _
                        Me.Adodc_preparation.Recordset("NOM_FICHIER_DATA"), _
                        Me.Adodc_preparation.Recordset("NOMBRE_DATA"), _
                        Me.Adodc_preparation.Recordset("NUM_RECEPTION"), _
                        Me.Adodc_preparation.Recordset("FK_FLOW"))

        If Me.Adodc_preparation.Recordset("EPOBOX_IN") Then
            Call Init_Adodc(Me.Adodc_Tool)
        End If
        Call stopApplication


Suivant:

        Me.Adodc_preparation.Refresh
        If Not Me.Adodc_preparation.Recordset.EOF Then
            If L_grid_pos > 1 Then
                On Error Resume Next
                Me.DataGrid1.Bookmark = L_grid_pos - 1
            Else
                Me.Adodc_preparation.Recordset.MoveFirst
            End If
        Else
            Call RAZ_Preparation_Detail
        End If
        Call Display_Status("", "", Me)

        For i = 1 To Me.Toolbar_Manual_Header.Buttons.Count
            Me.Toolbar_Manual_Header.Buttons(i).Enabled = True
        Next
        
        DataGrid1.Enabled = True
        Screen.MousePointer = vbNormal
        
    Case "cancel"
        If MsgBox("Attention, �tes-vous sur de vouloir annuler cette fusion?", vbExclamation + vbYesNo, Me.Caption) = vbYes Then
            Screen.MousePointer = vbHourglass
            Rem M�morisation de la position dans la grille*
            L_grid_pos = Me.DataGrid1.Bookmark
            DataGrid1.Enabled = False
            Call Toolbar_Set_All_header(False)
            Call Toolbar_Set_All_Detail(False)
Rem v406
'            Call Fusion_Cancel(Me.Adodc_preparation.Recordset("NUM_SOCIETE"), _
                               Me.Adodc_preparation.Recordset("NUM_RECEPTION"), _
                               L_tmp, Fusion_Robot)
            Call Fusion_Cancel(Me.Adodc_preparation.Recordset("NUM_SOCIETE"), _
                               L_tmp, Fusion_Robot)
Rem v406
            Call Toolbar_Set_All_header(True)
            Me.Adodc_preparation.Refresh
            If Not Me.Adodc_preparation.Recordset.EOF Then
                If L_grid_pos > 1 Then
                    On Error Resume Next
                    Me.DataGrid1.Bookmark = L_grid_pos - 1
                    Err.Clear
                End If
            Else
                Call RAZ_Preparation_Detail
            End If
            Call Display_Status("", "", Me)
            DataGrid1.Enabled = True
            Screen.MousePointer = vbNormal
        End If
        Rem v442 - fusion
        Call KillProcessus("WINWORD.EXE")
        'Call KillProcessus("PDFCREATOR.EXE")
        Rem v442 - fusion
    
    Case "error"
        If MsgBox("Attention, �tes-vous sur de vouloir passer cette fusion en mode Erreur?", vbExclamation + vbYesNo, Me.Caption) = vbYes Then
            Screen.MousePointer = vbHourglass
            DoEvents
            Rem M�morisation de la position dans la grille*
            L_grid_pos = Me.DataGrid1.Bookmark
            DataGrid1.Enabled = False
            Call Toolbar_Set_All_header(False)
            Call Toolbar_Set_All_Detail(False)
            'Call Run_Execute_Sql("UPDATE PREPARATION SET FUSION = 0, STATUT = 'Erreur', STATUS = 'Erreur fusion' WHERE PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"))
            Call Run_Execute_Sql("UPDATE PREPARATION SET STATUS = 'Erreur fusion' WHERE PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"))
            Call Toolbar_Set_All_header(True)
            If Me.Toolbar_Manual_Header.Buttons("archives").Caption = "En cours..." Then
                If M_Nb > 0 Then
                    M_Nb = M_Nb - 1
                End If
                L_pos = InStr(1, Adodc_preparation.RecordSource, "DESC LIMIT", vbTextCompare)
                If L_pos > 0 Then
                    Adodc_preparation.RecordSource = Left(Adodc_preparation.RecordSource, L_pos + Len("DESC LIMIT")) & M_Nb
                End If
            End If
            Me.Adodc_preparation.Refresh
            If Not Me.Adodc_preparation.Recordset.EOF Then
                If L_grid_pos > 1 Then
                    On Error Resume Next
                    Me.DataGrid1.Bookmark = L_grid_pos
                End If
            Else
                Call RAZ_Preparation_Detail
            End If
            Call Display_Status("", "", Me)
            DataGrid1.Enabled = True
            Screen.MousePointer = vbNormal
            If Me.Toolbar_Manual_Header.Buttons("archives").Caption = "En cours..." And M_Nb = 0 Then
                Call Toolbar_Manual_Header_ButtonClick(Toolbar_Manual_Header.Buttons("archives"))
            End If
            
        End If
        
    Case "stop"
        Screen.MousePointer = vbHourglass
        G_Identifiant = Me.Adodc_preparation.Recordset("PK_PREPARATION")
        Me.Adodc_preparation.Refresh
        If Not Me.Adodc_preparation.Recordset.EOF Then
            Me.Adodc_preparation.Recordset.MoveFirst
            Me.Adodc_preparation.Recordset.Find "PK_PREPARATION = " & G_Identifiant
        End If
        Screen.MousePointer = vbNormal
        Select Case Me.Text_STATUT
        Case "En fusion", "En traitement", "?"
            If MsgBox("Etes-vous s�r(e) de vouloir interrompre ce traitement en cours?!", vbExclamation + vbYesNo, Me.Caption) = vbYes Then
                Screen.MousePointer = vbHourglass
                Call Run_Update_Sql(G_Adoconnection, "PREPARATION", " FIN = NOW(), STATUS = 'Avort� fusion'", "PK_PREPARATION = " & G_Identifiant)
                Me.Adodc_preparation.Refresh
                If Not Me.Adodc_preparation.Recordset.EOF Then
                    Me.Adodc_preparation.Recordset.MoveFirst
                    Me.Adodc_preparation.Recordset.Find "PK_PREPARATION = " & G_Identifiant
                Else
                    Call RAZ_Preparation_Detail
                End If
            End If
        End Select
        Screen.MousePointer = vbNormal
        
    End Select

End Sub

Private Sub Toolbar_Manual_Header_ButtonClick(ByVal Button As MSComctlLib.Button)


    'MsgBox CurrentSequestre("033411000123")
    
    Dim L_nb_element        As Long
    Dim L_Nb                As Long
    Dim L_reprise           As Boolean
    Dim L_NbArchive         As Long
    Dim Tmp                 As String
    Dim L_tmp               As Long
    Dim L_Query             As String
    Dim L_RunningRobot      As Boolean
    Dim L_NbErrors          As Long
    Dim L_NextTentative     As Date
    Dim i                   As Integer
    Dim L_FkOptimisationBalancing   As String
    
    
    L_NbErrors = 0
    
    If G_flux_type = "REF" Then
        L_Query = "PREPARATION_NETCO"
    Else
        L_Query = "PREPARATION"
    End If
    L_RunningRobot = False
    L_NextTentative = DateAdd("m", 1, Now)
    
    Select Case Button.Key
                
    Case "archives"
Rem v406
        Call CheckApplication("Archives search")
        G_flux_type_precedent = G_flux_type
        If Me.Toolbar_Manual_Header.Buttons("archives").Caption = "Archives" Then
            G_flux_type = "SEL"
            ReDim G_data_Xmit(99)
            G_data_Xmit(0) = ""
            robot_customer_selection.Show vbModal
            If G_data_Xmit(0) = "" Then
                Exit Sub
            End If
            Tmp = "x"
            While Not IsNumeric(Tmp)
                Tmp = InputBox("Combien d'enregistrements voulez-vous visualiser?", "S�lection", "100")
                If Tmp = "" Or Tmp = "0" Then
                    Exit Sub
                End If
            Wend
            M_Nb = CLng(Tmp)
            If G_data_Xmit(0) = "" Then
'stop 'Debug
                Adodc_preparation.RecordSource = SQL(L_Query, " PREPARATION.STATUS in ('Fusionn�', 'Annul� en fusion') ") & " ORDER BY NUM_PREPARATION DESC LIMIT " & CLng(Tmp)
            Else
                Adodc_preparation.RecordSource = SQL(L_Query, " PREPARATION.STATUS in ('Fusionn�', 'Annul� en fusion') AND PREPARATION.FK_SOCIETE = " & IIf(IsNumeric(G_data_Xmit(0)), G_data_Xmit(0), G_data_Xmit(1)) & " ORDER BY NUM_PREPARATION DESC LIMIT " & CLng(Tmp))
            End If
            G_data_Xmit(0) = ""
            Screen.MousePointer = vbHourglass
            Adodc_preparation.Refresh
            Me.Toolbar_Manual_Header.Buttons("archives").Caption = "En cours..."
            Screen.MousePointer = vbDefault
        Else
            For i = 0 To Option_Type.Count - 1
                If Option_Type(i).Value = True Then
                    Call Option_Type_Click(i)
                    Exit For
                End If
            Next
            'Call Init_Data("")
            Me.Toolbar_Manual_Header.Buttons("archives").Caption = "Archives"
        End If
        G_flux_type = G_flux_type_precedent

    Case "sort"
        'Call Run_Sort(Me.Adodc_preparation, DataGrid1, Me, Sql(L_Query, " PREPARATION.STATUS in ('Pr�par�', 'En fusion', 'Erreur fusion', 'Avort� fusion')"))
        Call Run_Sort(Me.Adodc_preparation, DataGrid1, Me, SQL(L_Query, " PREPARATION.STATUS in ('Pr�par�', 'Pr�par�2', 'En fusion', 'Erreur fusion', 'Avort� fusion')"))
        
    Case "refresh"
        Screen.MousePointer = vbHourglass
        Me.Adodc_preparation.Refresh
        If Me.Adodc_preparation.Recordset.EOF Then
            Call RAZ_Preparation_Detail
        Else
            Call Init_Preparation_Detail
        End If
        Screen.MousePointer = vbNormal
        Exit Sub
    
    Case "stop"
        Me.WindowState = vbNormal
        Call stopApplication
        G_Arret_Boucle = True
        Call Display_Status("Arr�t du robot de Fusion en cours ...", "", Me)
        Exit Sub
        
    Case "boucle"

        Me.WindowState = vbMinimized
        Call StartApplication
        L_RunningRobot = True
        
        'G_BAT = False
        M_Mode = "ROBOT"
        
        Call Toolbar_Set_All_header(False)
        Call Toolbar_Set_All_Detail(False)
        
        Me.Toolbar_Manual_Header.Buttons("stop").Visible = True
        Me.Toolbar_Manual_Header.Buttons("stop").Enabled = True
        Me.Toolbar_Manual_Header.Buttons("boucle").Visible = False
        
        Rem On consid�re qu'il peut y avoir besoin d'une fusion
        Rem Controle syst�matique de la pr�sence du driver
        If Not PDFCreator_Driver_Ok Then
            Exit Sub
        End If
        
        Rem v442 - fusion
        If Not Me.Option_Type(3).Value Then
            Call KillProcessus("WINWORD.EXE")
            'Call KillProcessus("PDFCREATOR.EXE")
        End If
        Rem v442 - fusion
        
        Do
Continue_Robot:
            DoEvents
            PDFCreator1.cClearCache
            DoEvents
            Me.Adodc_preparation.Refresh
            Call CheckApplication("Waiting")
            L_nb_element = Me.Adodc_preparation.Recordset.RecordCount
            If L_nb_element > 0 Then
                Me.Adodc_preparation.Recordset.MoveFirst
            End If
            L_NbErrors = 0
            If Not Me.Adodc_preparation.Recordset.EOF Then
                L_Nb = 1
                DataGrid1.Enabled = False
                Me.Adodc_preparation.Recordset.MoveFirst
                L_FkOptimisationBalancing = ""
                While Not Me.Adodc_preparation.Recordset.EOF
                    If Not Me.Option_Type(3).Value Then
                        Call KillProcessus("WINWORD.EXE")
                    End If
                    
                    Select Case Lire_Un_Champ("STATUS", "PREPARATION", "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"))
                    Case "Erreur fusion"
                        L_NbErrors = L_NbErrors + 1
                        
                    Case "Pr�par�", "Pr�par�2"
                        L_reprise = False
                        
                        L_tmp = Me.Adodc_preparation.Recordset("PK_PREPARATION")
                        Select Case G_flux_type
                        Case "ALL"
                            Rem On traite tous les flux... Sauf ceux exclus
                            If Is_In_List_To_Exclure(Me.Adodc_preparation.Recordset("FK_SOCIETE"), _
                                                     G_flux_customer_pk, _
                                                     ",") = True Then
                                GoTo Suivant_Boucle
                            End If
                        
                        Case "REF"
                            If Not Is_Client_Reference(Me.Adodc_preparation.Recordset("NUM_SOCIETE")) Then
                                GoTo Suivant_Boucle
                            End If
Rem // --------------------------------------------------------------------------------------
Rem // v352
                        Case "SEL", "SIGMA_EASYLINK"
                            Rem Ok. Tested on 9/18 @15:33
Rem v472
                            'If G_flux_customer_pk <> Me.Adodc_preparation.Recordset("FK_SOCIETE") Then
                            If Not Client_Selectionne(Me.Adodc_preparation.Recordset("FK_SOCIETE"), G_flux_customer_pk) Then
                                GoTo Suivant_Boucle
                            End If
Rem v472
Rem // --------------------------------------------------------------------------------------
                        End Select

                        Rem Tout
                        If Not Type_Ok(CLng(Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")), Me.Adodc_preparation.Recordset("PRESTATION_MODEL_NOM")) Then
                            GoTo Suivant_Boucle
                        End If
Rem v436

Rem v576 - Avant de r�server la fusion, je m'assure que toutes les autres fusions ne sont pas r�serv�es au m�me client
                        If L_FkOptimisationBalancing <> "" Then
                            If L_FkOptimisationBalancing = Me.Adodc_preparation.Recordset("FK_SOCIETE") Then
                                GoTo Suivant_Boucle
                            End If
                        End If
                        
                        If LoadBalancing(Me.Adodc_preparation.Recordset("FK_SOCIETE"), G_flux_type, "fusion") <> "Ok" Then
                            L_FkOptimisationBalancing = Me.Adodc_preparation.Recordset("FK_SOCIETE")
                            GoTo Suivant_Boucle
                        End If
Rem v576
                        Call Display_Status("Pr�paration de la fusion...", L_Nb & " / " & L_nb_element, Me)
Rem v641
                        If ReserverUnEnregistrement("PREPARATION", _
                                              "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), _
                                              "STATUS", _
                                              "Pr�par�2|Pr�par�", _
                                              "En fusion") = False Then
                        'If Modifier_un_Statut("PREPARATION", _
                                              "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), _
                                              "STATUS", _
                                              "Pr�par�2|Pr�par�", _
                                              "En fusion") = False Then
Rem v641
                            Rem Le statut a �t� modifi� entre temps
                            GoTo Suivant_Boucle
                        End If
                        Call Fusion_Run(Me.Adodc_preparation.Recordset("NUM_PREPARATION"), _
                                Me.Adodc_preparation.Recordset("FK_RECEPTION"), _
                                Me.Adodc_preparation.Recordset("NUM_SOCIETE"), _
                                L_tmp, L_reprise, _
                                Me.Adodc_preparation.Recordset("EPOBOX_IN"), _
                                Fusion_Robot, _
                                Me.Adodc_preparation.Recordset("PK_PRESTATION_MODEL"), _
                                Me.Adodc_preparation.Recordset("NOM_FICHIER_DATA"), _
                                Me.Adodc_preparation.Recordset("NOMBRE_DATA"), _
                                Me.Adodc_preparation.Recordset("NUM_RECEPTION"), _
                                Me.Adodc_preparation.Recordset("FK_FLOW"))
Rem v406
                    End Select
                               
Rem v350
Suivant_Boucle:
Rem v350
                    If G_Arret_Boucle Then
                        Call Arret_Boucle
                        Exit Sub
                    End If
                    L_Nb = L_Nb + 1
                    Me.Adodc_preparation.Recordset.MoveNext
                Wend
                
                Me.Adodc_preparation.Refresh
                If L_NbErrors > 0 Then
                    If L_NextTentative < Now Then
                        L_NextTentative = DateAdd("m", 1, Now)
                        GoTo All_errors
                    Else
                        L_NextTentative = DateAdd("n", 1, Now)
                    End If
                Else
                    L_NextTentative = DateAdd("m", 1, Now)
                End If
                Call Display_Status("", "", Me)
                DataGrid1.Enabled = True
            End If
            
            If G_Arret_Boucle Then
                Call Arret_Boucle
                Exit Sub
            End If
            
            Call Sleep_And_Events(G_Delay * 1000)
            
            If ServerNow > G_RobotNextControlerActivity Then
                Call RobotControlerChecker
                G_RobotNextControlerActivity = DateAdd("n", 7, Now)
            End If
            'G_RobotNextStopTime = "25/10/2016 14:48:00"
            If ServerNow > G_RobotNextStopTime Then
                Call RobotControlerChecker
                Call Arret_Boucle
                Call Toolbar_Manual_Header_ButtonClick(Toolbar_Manual_Header.Buttons("closew"))
                Exit Sub
            End If
            
            'If 1 Then
            '    G_AutomaticUnLoadMe = True
            '    G_Arret_Boucle = True
            '    Call Arret_Boucle
            '    Call Toolbar_Manual_Header_ButtonClick(Toolbar_Manual_Header.Buttons("stop"))
            '    Call Toolbar_Manual_Header_ButtonClick(Toolbar_Manual_Header.Buttons("closew"))
            '    Exit Sub
            'End If
            DoEvents
            
            If G_Arret_Boucle Then
                Call Arret_Boucle
                Exit Sub
            End If
        Loop
    
        Rem v442 - fusion
        Call KillProcessus("WINWORD.EXE")
        'Call KillProcessus("PDFCREATOR.EXE")
        Rem v442 - fusion
        Call stopApplication

    
    Case "all_error"
        Call StartApplication
        Call CheckApplication("Error running")
        'G_BAT = False
        L_nb_element = Me.Adodc_preparation.Recordset.RecordCount
        
        If L_nb_element > 0 Then
            If Me.Adodc_preparation.Recordset.AbsolutePosition = 1 Then
                If MsgBox("Attention, vous allez reprendre toutes les lignes en erreurs!" & vbNewLine & "Etes-vous sur(e)?", vbYesNo, Me.Caption) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        If Me.Adodc_preparation.Recordset.EOF Then
            MsgBox "Aucune fusion � effectuer!", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        Rem On consid�re qu'il peut y avoir besoin d'une fusion
        Rem Controle syst�matique de la pr�sence du driver
        'If Not PCL_Driver_Ok Then Exit Sub
        If Not PDFCreator_Driver_Ok Then
            Exit Sub
        End If
        L_Nb = 1
        Call Toolbar_Set_All_header(vbFalse)
        Call Toolbar_Set_All_Detail(vbFalse)
        DataGrid1.Enabled = False
        
        Rem v442 - fusion
        Call KillProcessus("WINWORD.EXE")
        'Call KillProcessus("PDFCREATOR.EXE")
        Rem v442 - fusion


All_errors:
        
        While Not Me.Adodc_preparation.Recordset.EOF
            
            Select Case Lire_Un_Champ("STATUS", "PREPARATION", "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"))
            Case "Erreur fusion"
                L_reprise = True
                L_tmp = Me.Adodc_preparation.Recordset("PK_PREPARATION")
                Select Case G_flux_type
                Case "ALL"
                    Rem On traite tous les flux... Sauf ceux exclus
                    If Is_In_List_To_Exclure(Me.Adodc_preparation.Recordset("FK_SOCIETE"), _
                                             G_flux_customer_pk, _
                                             ",") = True Then
                        GoTo Suivant_All_error
                    End If
                    
                Case "REF"
                    If Not Is_Client_Reference(Me.Adodc_preparation.Recordset("NUM_SOCIETE")) Then
                        GoTo Suivant_All_error
                    End If
                Case "SEL", "SIGMA_EASYLINK"
                    If Not Client_Selectionne(Me.Adodc_preparation.Recordset("FK_SOCIETE"), G_flux_customer_pk) Then
                        GoTo Suivant_All_error
                    End If
                End Select
                Call Display_Status("Pr�paration...", L_Nb & " / " & L_nb_element, Me)
Rem v641
                If ReserverUnEnregistrement("PREPARATION", _
                                      "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), _
                                      "STATUS", _
                                      "Erreur fusion", _
                                      "En fusion") = False Then
                'If Modifier_un_Statut("PREPARATION", _
                                      "PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION"), _
                                      "STATUS", _
                                      "Erreur fusion", _
                                      "En fusion") = False Then
Rem v641
                    Rem Le statut a �t� modifi� entre temps
                    GoTo Suivant_All_error
                End If
                
                Call Fusion_Run(Me.Adodc_preparation.Recordset("NUM_PREPARATION"), _
                        Me.Adodc_preparation.Recordset("FK_RECEPTION"), _
                        Me.Adodc_preparation.Recordset("NUM_SOCIETE"), _
                        L_tmp, L_reprise, _
                        Me.Adodc_preparation.Recordset("EPOBOX_IN"), _
                        Fusion_Robot, _
                        Me.Adodc_preparation.Recordset("PK_PRESTATION_MODEL"), _
                        Me.Adodc_preparation.Recordset("NOM_FICHIER_DATA"), _
                        Me.Adodc_preparation.Recordset("NOMBRE_DATA"), _
                        Me.Adodc_preparation.Recordset("NUM_RECEPTION"), Me.Adodc_preparation.Recordset("FK_FLOW"))
Rem // v406
Rem v350
Suivant_All_error:
Rem v350
            End Select
            L_Nb = L_Nb + 1
            Me.Adodc_preparation.Recordset.MoveNext
        
        Wend
        Me.Adodc_preparation.Refresh
        If L_RunningRobot Then
            L_NbErrors = 0
            GoTo Continue_Robot
        End If
        Call CheckApplication("Error terminated")
        Call Toolbar_Set_All_header(True)
        Call Display_Status("", "", Me)
        Rem v442 - fusion
        Call KillProcessus("WINWORD.EXE")
        'Call KillProcessus("PDFCREATOR.EXE")
        Rem v442 - fusion
        Call stopApplication
        DataGrid1.Enabled = True
        Screen.MousePointer = vbNormal
        
    Case "closew"
        Call CheckApplication("Window closed")
        'Call KillProcessus("PDFCREATOR.EXE")
        Unload Me
        Rem COCA
        'End
        Rem Fin
        Exit Sub
        
    End Select

End Sub

Private Sub Init_Preparation_Detail()

Dim L_Data() As String
Dim Tmp As String

    Me.Text_NUM_PREPARATION = do_read_data(Me.Adodc_preparation.Recordset("NUM_PREPARATION"), "Alpha")
    Me.Text_NUM_SOCIETE = do_read_data(Me.Adodc_preparation.Recordset("NUM_SOCIETE"), "Alpha")
    
    Me.Text_NOM_FICHIER_DATA = do_read_data(Me.Adodc_preparation.Recordset("NOM_FICHIER_DATA"), "Alpha")
    
    If M_Mode = "ROBOT" Then
        Me.Text_NOM_FICHIER_INFO = "n/a"
        Me.Text_NOMBRE_ENREGISTREMENT_DATA = "?"
        Me.Text_STATUT = "?"
        Me.Text_COMMENTAIRE = ""
        Me.Text_DEBUT = ""
        Me.Text_FIN = ""
        Me.Text_MAJ_DATE = ""
        Me.Text_MAJ_USERID = ""
        Me.Text_segment = ""
        Toolbar_Manual_Detail.Enabled = False
        Exit Sub
    End If
    
    Tmp = Me.Adodc_preparation.Recordset("FK_FLOW")
    If Tmp = "0" Then
        Me.Text_NUM_FLOW = "Flux non migr�"
        Me.Text_NUM_FLOW.tag = ""
    Else
        Me.Text_NUM_FLOW = Lire_Un_Champ("NUM_FLOW", "FLOW", "PK_FLOW = " & Tmp)
        Me.Text_NUM_FLOW.tag = Tmp
    End If
    
    'L_Data = Split(Lire_Des_Champs("NOM_FICHIER_INFO, NOMBRE_ENREGISTREMENT_DATA, STATUS, COMMENTAIRE, DEBUT, FIN, MAJ_DATE, MAJ_USERID, NOMBRE_DATA, NB_SEGMENT", "preparation", "pk_preparation = " & Me.Adodc_preparation.Recordset("PK_PREPARATION")), "|", , vbTextCompare)
    L_Data = Split(Lire_Des_Champs("RECEPTION.NOM_FICHIER_INFO, PREPARATION.NOMBRE_ENREGISTREMENT_DATA, PREPARATION.STATUS, PREPARATION.NB_SEGMENT, PREPARATION.DEBUT, PREPARATION.FIN, PREPARATION.MAJ_DATE, PREPARATION.MAJ_USERID, PREPARATION.NOMBRE_DATA, PREPARATION.COMMENTAIRE, RECEPTION.NUM_RECEPTION", "reception, preparation", "PK_RECEPTION = FK_RECEPTION AND PK_PREPARATION = " & Me.Adodc_preparation.Recordset("PK_PREPARATION")), "|", , vbTextCompare)
    Me.Text_NOM_FICHIER_INFO = L_Data(0)
    Me.Text_NOMBRE_ENREGISTREMENT_DATA = L_Data(1)
    Me.Text_STATUT = L_Data(2)
    Me.Text_COMMENTAIRE = L_Data(9)
    Me.Text_DEBUT = L_Data(4)
    Me.Text_FIN = L_Data(5)
    Me.Text_MAJ_DATE = L_Data(6)
    Me.Text_MAJ_USERID = L_Data(7)
    Me.Text_segment = L_Data(8)
    Me.Text_NUM_RECEPTION = L_Data(10)
    If L_Data(3) = 0 Then
        Me.Text_segment.Visible = False
        Me.Label_lot.Visible = False
        Me.Text_segment.Locked = True
        Me.Label_lot_suite.Visible = False
    Else
        Me.Text_segment.Visible = True
        Me.Label_lot.Visible = True
        Me.Text_segment.Locked = True
        Me.Label_lot_suite.Visible = True
    End If
    
        
    Toolbar_Manual_Detail.Enabled = True
    Select Case Me.Text_STATUT
    Case "Fusionn�", "Annul�", "Annul� fusion"
        Me.Toolbar_Manual_Detail.Buttons("ok").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("bat").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("cancel").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("error").Enabled = True
        Me.Toolbar_Manual_Detail.Buttons("error").Visible = True
        Me.Toolbar_Manual_Detail.Buttons("stop").Visible = False
    Case "En traitement", "En fusion"
        Me.Toolbar_Manual_Detail.Buttons("ok").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("bat").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("cancel").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("stop").Enabled = True
        Me.Toolbar_Manual_Detail.Buttons("error").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("error").Visible = False
        Me.Toolbar_Manual_Detail.Buttons("stop").Visible = True
    Case Else '"Erreur", "A fusionner"
        Me.Toolbar_Manual_Detail.Buttons("ok").Enabled = True
        Me.Toolbar_Manual_Detail.Buttons("bat").Enabled = True
        Me.Toolbar_Manual_Detail.Buttons("cancel").Enabled = True
        Me.Toolbar_Manual_Detail.Buttons("error").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("error").Visible = False
        Me.Toolbar_Manual_Detail.Buttons("stop").Visible = False
    End Select
    
End Sub

Private Sub RAZ_Preparation_Detail()

    Me.Text_NUM_FLOW = vbNullString
    Me.Text_NUM_RECEPTION = vbNullString
    Me.Text_NUM_PREPARATION = vbNullString
    Me.Text_NUM_SOCIETE = vbNullString
    Me.Text_NOM_FICHIER_INFO = vbNullString
    Me.Text_NOM_FICHIER_DATA = vbNullString
    Me.Text_NOMBRE_ENREGISTREMENT_DATA = vbNullString
    Me.Text_STATUT = vbNullString
    Me.Text_COMMENTAIRE = vbNullString
    Me.Text_DEBUT = vbNullString
    Me.Text_FIN = vbNullString
    Me.Text_MAJ_DATE = vbNullString
    Me.Text_MAJ_USERID = vbNullString
        
    Toolbar_Manual_Detail.Enabled = False
    
    If M_Mode <> "ROBOT" Then
        Me.Toolbar_Manual_Detail.Buttons("ok").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("bat").Enabled = False
        Me.Toolbar_Manual_Detail.Buttons("cancel").Enabled = False
    End If
    
End Sub

Private Sub Toolbar_Set_All_header(p_what As Boolean)
    
    Dim lv_i        As Integer
    
    For lv_i = 1 To Me.Toolbar_Manual_Header.Buttons.Count
        Me.Toolbar_Manual_Header.Buttons(lv_i).Enabled = p_what
    Next lv_i
    'Me.Check_Ignore_eMail.Enabled = p_what
    For lv_i = 0 To Me.Option_Type.Count - 1
        Me.Option_Type(lv_i).Enabled = p_what
    Next
    
End Sub

Private Sub Toolbar_Set_All_Detail(p_what As Boolean)
    
    Dim lv_i        As Integer
    
    For lv_i = 1 To Me.Toolbar_Manual_Detail.Buttons.Count
        Me.Toolbar_Manual_Detail.Buttons(lv_i).Enabled = p_what
    Next lv_i
    
End Sub

Private Sub adodc_preparation_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

    fCancelDisplay = True
    G_msg_error_p1 = ErrorNumber
    G_msg_error_p2 = Description
    G_msg_error_p5 = Adodc_preparation.RecordSource
    G_msg_error_p3 = Me.Name
    G_msg_error_p4 = "adodc_preparation_Error"
    G_msg_error_p6 = Source
    error_manager.Show vbModal
    End
    
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Me.Adodc_preparation.Recordset.EOF Then
        Call Init_Preparation_Detail
    Else
        Call RAZ_Preparation_Detail
    End If
    
End Sub

Private Sub Arret_Boucle()

    Me.Adodc_preparation.Refresh
    M_Mode = vbNullString
    If Me.Adodc_preparation.Recordset.EOF Then
        Call RAZ_Preparation_Detail
    Else
        Call Init_Preparation_Detail
    End If
    Call Toolbar_Set_All_header(True)
    Call Toolbar_Set_All_Detail(True)
    Me.Toolbar_Manual_Header.Buttons("stop").Visible = False
    Me.Toolbar_Manual_Header.Buttons("boucle").Visible = True
    Call Display_Status("", "", Me)
    DataGrid1.Enabled = True
    G_Arret_Boucle = False
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Init_Data(p_Selection As String)
    
    Dim L_Fk As String
    Dim P2 As String
    
    Select Case p_Selection
    Case "", "ALL", "TOUT"
        P2 = ""
    Case "FAX"
        P2 = "fk_type_prestation = 7"
    Case "MAIL"
        P2 = "fk_type_prestation = 6"
    Case "NOMAIL"
        P2 = "fk_type_prestation <> 6"
    Case "SANS WORD"
        P2 = "prestation_model.model_type <> 'Mod�le Word' "
    Case "EPOBOX"
        P2 = "fk_type_prestation = 8"
    Case "TRUST OFFICE"
        P2 = "prestation_model_nom like 'WS_%' "
    Case "WORD"
Rem v578
        'P2 = "prestation_model.model_type = 'Mod�le Word' and PRINT_PDF = 0 and etat_migration <> 'N/A' "
Rem v578
        P2 = "prestation_model.model_type = 'Mod�le Word' "
    End Select
    
    Screen.MousePointer = vbHourglass
    Select Case G_flux_type
    Case "REF"
'stop 'Debug
        Adodc_preparation.RecordSource = SQL("PREPARATION_NETCO", " PREPARATION.STATUS in ('Pr�par�', 'Pr�par�2', 'En fusion', 'Erreur fusion', 'Avort� fusion')", P2) & " ORDER BY NUM_PREPARATION ASC "
        
    Case "SIGMA_EASYLINK", "SEL"
        L_Fk = Trim(G_flux_customer_pk)
        If Right(L_Fk, 1) = "," Then
            L_Fk = Left(L_Fk, Len(L_Fk) - 1)
        End If
'stop 'Debug
        Adodc_preparation.RecordSource = SQL("PREPARATION", " PREPARATION.STATUS in ('Pr�par�','Pr�par�2', 'En fusion', 'Erreur fusion', 'Avort� fusion') and prestation_model.fusion_version = 0", "preparation.fk_societe in (" & L_Fk & ")", P2) & " ORDER BY NUM_PREPARATION ASC "
    Case Else
        Adodc_preparation.RecordSource = SQL("PREPARATION", " PREPARATION.STATUS in ('Pr�par�','Pr�par�2',  'En fusion', 'Erreur fusion', 'Avort� fusion') and prestation_model.fusion_version = 0", P2) & " ORDER BY NUM_PREPARATION ASC "
        'stop
        'Adodc_preparation.RecordSource = Replace(Adodc_preparation.RecordSource, "ORDER BY", "AND PREPARATION.FK_SOCIETE >= 11728 ORDER BY", , , vbTextCompare)
        
    End Select
    
    Adodc_preparation.Refresh
    
    If Not Me.Adodc_preparation.Recordset.EOF Then
        Me.Adodc_preparation.Recordset.MoveFirst
        Call Init_Preparation_Detail
    Else
        Call RAZ_Preparation_Detail
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Function Type_Ok(p_Fk_prestation_model As Long, Optional p_Prestation_Model_Name As String)

    Dim R As String
    
    Type_Ok = True
    If Me.Option_Type(0).Value = True Then
        Rem Pas de restriction
    ElseIf Me.Option_Type(1).Value = True Then
        Rem Tout Sauf e-mail
        If M_Fk_Prestation_eMail = Lire_Un_Champ("FK_TYPE_PRESTATION", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")) Then
            Type_Ok = False
        End If
    'Else 'Que e-mail
    ElseIf Me.Option_Type(2).Value = True Then
        If M_Fk_Prestation_eMail = Lire_Un_Champ("FK_TYPE_PRESTATION", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")) Then
            Type_Ok = False
        End If
    ElseIf Me.Option_Type(3).Value = True Then
        'If M_Fk_Prestation_ePobox = Lire_un_champ("FK_TYPE_PRESTATION", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")) Then
        R = Lire_Un_Champ("model_type", "prestation_model", "PK_PRESTATION_MODEL = " & Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL"))
        If R = "Mod�le Word" Then
            Type_Ok = False
        End If
    ElseIf Me.Option_Type(4).Value = True Then
        If M_Fk_Prestation_ePobox <> Lire_Un_Champ("FK_TYPE_PRESTATION", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & Me.Adodc_preparation.Recordset("FK_PRESTATION_MODEL")) Then
            Type_Ok = False
        End If
    ElseIf Me.Option_Type(5).Value = True Then
        If Left(UCase(p_Prestation_Model_Name), 3) <> "WS_" Then
            Type_Ok = False
        End If
    ElseIf Me.Option_Type(6).Value = True Then
        Rem Pas de restriction
        'stop
    End If


End Function

Private Sub PDFCreator1_eError()
    
    Set pErr = PDFCreator1.cError
    'AddStatus "Error[" & pErr.Number & "]: " & pErr.Description
    Screen.MousePointer = vbNormal

End Sub

Private Sub PDFCreator1_eReady()
    
    'AddStatus """" & PDFCreator1.cOutputFilename & """ a �t� cr��! (" & _
    DateDiff("s", StartTime, Now) & " secondes)"
    PDFCreator1.cPrinterStop = True
    Screen.MousePointer = vbNormal

End Sub
    
Public Sub Fusion_Run(p_num_preparation As String, _
                      p_pk_reception As Long, _
                      p_Customer_Number As String, _
                      p_preparation_pk As Long, _
                      p_Reprise As Boolean, _
                      p_ePoBox_In As String, _
                      ByRef p_Form As Form, _
                      ByVal p_Pk_Prestation_Model As Long, _
                      ByVal p_DataFileName As String, _
                      ByVal p_Nombre_Data As Long, _
                      Optional ByVal pNumReception As String, _
                      Optional ByVal pFkFlow As Long)

    Dim L_dir_from                                  As String
    Dim L_societe_pk                                As String
    Dim L_data_file_name                            As String
    Dim L_Prestation_Model_Pk                       As String
    Dim L_statut                                    As String
    Dim L_commentaire                               As String
    Dim L_resultat_fusion                           As String
    'Dim L_Referencement_Fk                          As Long
    Dim L_liste_fichiers_joints_reception           As String
    Dim L_liste_fichiers_joints_production_locale   As String
    Dim i                                           As Integer
    Dim L_Temporary_Table_To_Delete                 As String
    Dim L_Read()                                    As String
    Dim Path_name, File_name                        As String
    Dim fs                                          As Scripting.FileSystemObject
    Rem v461
    Dim L_Fk_Contact                                As String
    Dim FkPreReception                              As String
    
    Dim FlowCreationDate            As String
    Dim L_DirFlowPreparation        As String
    Dim L_DirFlowFusionTO           As String
    Dim FlowSqlUpdate               As String
    
    Dim rfListePdfRapproches        As String
    Dim rfTab()                     As String
    Dim rfFileName                  As String
    Dim rfDir                       As String
    Dim ListePdfRapproches          As String
    
    
    L_Fk_Contact = Lire_Un_Champ("RECEPTION.FK_CONTACT", "RECEPTION", "PK_RECEPTION = " & p_pk_reception)
    'FkPreReception = Lire_Un_Champ("PK_PRE_RECEPTION", "PRE_RECEPTION", "FK_RECEPTION = " & p_pk_reception)
    
On Error GoTo Error

    Rem Alimentation de la variable pour l'horodatage!
    G_Date_Debut = Now
    L_Temporary_Table_To_Delete = vbNullString


    Err.Clear
    If Not Fusion_Robot.Option_Type(3).Value Then
        Call KillProcessus("WINWORD.EXE")
    End If
    Rem initialisation
    'M_Make_Sequestre = False                 'LCI SUPPRESSION SEQUESTRE LE 14/08/2017
    'M_Make_Xades = False                     'LCI SUPPRESSION WORM LE 14/08/2017
    M_Make_PdfS = False
    'M_Make_Worm = False                      'LCI SUPPRESSION WORM LE 14/08/2017
    ListePdfRapproches = ""
    
    If pFkFlow > 0 Then
        L_dir_from = G_dir_sequestre & Trim(Lire_Un_Champ("concat('\\',  date_format(date_creation, '%Y%m%d'), '\\', num_societe, '\\', num_flow)", "flow, societe", "pk_societe = fk_societe and pk_flow = " & pFkFlow))
        Rem R�pertoire de travail local sur S�questre
        L_dir_from = L_dir_from & "\fusion"
        If Not FolderExists(L_dir_from, False, False) Then
            If FolderExists(Replace(L_dir_from, "\033411000000", p_Customer_Number, , , vbTextCompare), False, False) Then
                L_dir_from = Replace(L_dir_from, "\033411000000", p_Customer_Number, , , vbTextCompare)
            End If
            If Not FolderExists(L_dir_from, False, False) Then
                L_statut = "Erreur fusion"
                L_commentaire = "R�pertoire de fusion non trouv�"
                Call Run_Update_Sql(G_Adoconnection, "PREPARATION", " FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'", "PK_PREPARATION = " & p_preparation_pk)
                Exit Sub
            End If
        End If
        FlowSqlUpdate = " status = 26, maj_date = now(), maj_userid = '" & G_User_Id & "' "
        Call Run_Update_Sql(G_Adoconnection, "FLOW", FlowSqlUpdate, "pk_flow = " & pFkFlow)
        
        pNumReception = Lire_Un_Champ("jointfiledir", "flow", "pk_flow = " & pFkFlow)
        If pNumReception = "" Then
            Rem SI C'EST UN DECOUPE PDF
            pNumReception = Lire_Un_Champ("action", "flow", "pk_flow = " & pFkFlow)
            If Left(pNumReception, 13) = "TraitementLot" Then
                Rem Mise � jour
                FlowSqlUpdate = " jointfiledir = '" & pNumReception & "', maj_date = now(), maj_userid = '" & G_User_Id & "' "
                Call Run_Update_Sql(G_Adoconnection, "FLOW", FlowSqlUpdate, "pk_flow = " & pFkFlow)
                
            Else
                L_commentaire = "Probl�me de fusion!!! jointfiledir non renseign� (FR001)!"
                GoTo Erreur_Fusion
            End If
  
        End If
        pNumReception = Replace(L_dir_from, "fusion", pNumReception, , , vbTextCompare) & "\"
        If Not FolderExists(pNumReception, False, False) Then
            L_commentaire = "Probl�me de fusion!!! jointfiledir inaccessible (FR002)!"
            GoTo Erreur_Fusion
        End If
        
    Else
        L_dir_from = G_dir_production & "\" & p_Customer_Number
    End If
    
    Rem Lire la pk de la societe
    L_societe_pk = Lire_Un_Champ("PK_SOCIETE", "SOCIETE", "NUM_SOCIETE ='" & p_Customer_Number & "'")
    
    Rem Lecture du nom du fichier data � fusionner
    Call Display_Status("Lecture du nom du fichier Data...", "", p_Form)
    
    
    Rem v464 Optiomisation
    L_data_file_name = p_DataFileName
    L_Prestation_Model_Pk = CStr(p_Pk_Prestation_Model)
    
    Rem Verifie que le nom du fichier Data est renseign�
    If L_data_file_name = vbNullString Then
        L_statut = "Erreur fusion"
        L_commentaire = "Fichier data non trouv� (Base)"
        Call Run_Update_Sql(G_Adoconnection, "PREPARATION", " FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'", "PK_PREPARATION = " & p_preparation_pk)
        Exit Sub
    End If
    
    Rem Verifie que le fichier Data est toujours l�
    Call Display_Status("Contr�le de pr�sence du fichier Data", L_data_file_name, p_Form)
    If Not FileExists(L_dir_from & "\" & L_data_file_name) Then
        Rem Quand tout sera en flow, deviendra inutile
        For i = 0 To 20
            If Not FileExists(L_dir_from & "\" & Format(Date - i, "yyyymmdd") & "\" & L_data_file_name) Then
                Rem Mise � jour du statut
                If i = 20 Then
                    L_statut = "Erreur fusion"
                    L_commentaire = Valid_Text("Fichier data non trouv� (R�pertoire)." & vbNewLine & "Fichier manquant: " & L_dir_from & "\" & L_data_file_name & "(" & L_dir_from & "\" & Format(Date, "yyyymmdd") & "\" & L_data_file_name & ")")
                    Call Run_Update_Sql(G_Adoconnection, "PREPARATION", " FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'", "PK_PREPARATION = " & p_preparation_pk)
                    Exit Sub
                End If
            Else
                Rem fichier trouv� � d�placer
                FileCopy L_dir_from & "\" & Format(Date - i, "yyyymmdd") & "\" & L_data_file_name, L_dir_from & "\" & L_data_file_name
                Exit For
            End If
        Next
    End If

    Rem Attention v�rification qu'il existe un Service Impression pour le(s) produit(s)
    Rem Qui compose la Prestation mod�le S�lectionn�e
    M_Make_PdfS = Existence_Service_Pe(CLng(L_Prestation_Model_Pk), G_CONST_SERVICE_PDFS)
    Rem LCI Doit disparaitre !!!
    'LCI SUPPRESSION SEQUESTRE LE 14/08/2017 M_Make_Sequestre = Existence_Service_Pe(CLng(L_Prestation_Model_Pk), G_CONST_SERVICE_SEQUESTRE)
    'LCI SUPPRESSION XADES LE 14/08/2017 M_Make_Xades = Existence_Service_Pe(CLng(L_Prestation_Model_Pk), G_CONST_SERVICE_XADES)
    'LCI SUPPRESSION WORM LE 14/08/2017 M_Make_Worm = Existence_Service_Pe(CLng(L_Prestation_Model_Pk), G_CONST_SERVICE_WORM)
    
    If Not PDFCreator_Driver_Ok Then
        Exit Sub
    End If

    Rem Gestion de la reprise
    If p_Reprise Then
        If p_ePoBox_In = "0" Then
            Call CheckApplication("Retry")
            Call Nettoyage_Fusion("Pr�paration", _
                                  p_preparation_pk, _
                                  p_Customer_Number, _
                                  p_Form)
        Else
            Call CheckApplication("ePobox Retry")
            Call Nettoyage_Fusion_ePoBox("Pr�paration", _
                                         p_preparation_pk, _
                                         p_Customer_Number, _
                                         p_Form)
        End If
    End If
    
    Rem
    If pFkFlow = 0 Then
        If pNumReception <> "" Then
            For i = 0 To 10
                If FolderExists(G_dir_sequestre & "\" & Format(Date - i, "YYYYMMDD") & "\" & p_Customer_Number & "\" & pNumReception) Then
                    pNumReception = G_dir_sequestre & "\" & Format(Date - i, "YYYYMMDD") & "\" & p_Customer_Number & "\" & pNumReception & "\"
                    i = 99
                    Exit For
                End If
            Next
            If i <> 99 Then
                pNumReception = "no"
            End If
        End If
    End If
    Rem Fusion du fichier Data
    Call Display_Status("Pr�paration de la fusion...", "", p_Form)
    
    If p_ePoBox_In = 0 Then
        L_resultat_fusion = Fusion_fichier_Data(L_dir_from, _
                                                L_data_file_name, _
                                                CStr(L_Prestation_Model_Pk), _
                                                p_Customer_Number, _
                                                L_societe_pk, _
                                                p_preparation_pk, _
                                                L_liste_fichiers_joints_reception, _
                                                L_liste_fichiers_joints_production_locale, _
                                                L_Temporary_Table_To_Delete, _
                                                p_Form, L_Fk_Contact, p_Nombre_Data, p_num_preparation, pNumReception, _
                                                pFkFlow, _
                                                ListePdfRapproches)
    Else
        L_resultat_fusion = Fusion_fichier_Data_ePoBox(L_dir_from, _
                                                       L_data_file_name, _
                                                       CStr(L_Prestation_Model_Pk), _
                                                       p_Customer_Number, _
                                                       L_societe_pk, _
                                                       p_preparation_pk, _
                                                       L_liste_fichiers_joints_reception, _
                                                       L_liste_fichiers_joints_production_locale, _
                                                       p_Form, p_Nombre_Data, pNumReception, pFkFlow)
    End If

    If L_Temporary_Table_To_Delete <> vbNullString Then
        Call Run_Execute_Sql("DROP TABLE " & L_Temporary_Table_To_Delete)
    End If

    If L_resultat_fusion = "Ok" Then
    
        Rem v496
        Rem Controle que le nombre de pli g�n�r� n'est pas sup�rieur au nombre de plis � g�n�rer
        Rem si c'est le cas => mettre la fusion en erreur, sinon, c'est bon!
        If p_ePoBox_In = 0 Then
            If Lire_Un_Champ("count(1)", "pli", "fk_preparation = " & p_preparation_pk) <> p_Nombre_Data Then
'stop 'Debug
                If Lire_Un_Champ("count(1)", "pli", "fk_preparation = " & p_preparation_pk) > p_Nombre_Data Then
                    L_commentaire = "Probl�me de fusion!!! trop de plis g�n�r�s! (" & Lire_Un_Champ("count(1)", "pli", "fk_preparation = " & p_preparation_pk) & ")"
                    GoTo Erreur_Fusion
                Else
                    If Lire_Un_Champ("regroupement_page", "Prestation_Model", "pk_prestation_Model = " & p_Pk_Prestation_Model) = "1" Then
                        Rem C'est normal
                        GoTo FusionOk
                    Else
                        L_commentaire = "Probl�me de fusion!!! Pas assez de plis g�n�r�s! (" & Lire_Un_Champ("count(1)", "pli", "fk_preparation = " & p_preparation_pk) & ")"
                    End If
                    GoTo Erreur_Fusion
                End If
            End If
        Else
'stop 'Debug
            If Lire_Un_Champ("count(1)", "pli_epobox_destinataire", "fk_preparation = " & p_preparation_pk) <> p_Nombre_Data Then
                If Lire_Un_Champ("count(1)", "pli_epobox_destinataire", "fk_preparation = " & p_preparation_pk) > p_Nombre_Data Then
                    L_commentaire = "Probl�me de fusion!!! trop de plis g�n�r�s!"
                    GoTo Erreur_Fusion
                Else
                    L_commentaire = "Probl�me de fusion!!! Pas assez de plis g�n�r�s!"
                    GoTo Erreur_Fusion
                End If
            End If
        End If
        If pFkFlow > 0 Then
            Rem Tous les plis sont fusionn�s
            FlowSqlUpdate = " status = 28, maj_date = now(), maj_userid = '" & G_User_Id & "' "
            Call Run_Update_Sql(G_Adoconnection, "FLOW", FlowSqlUpdate, "pk_flow = " & pFkFlow)
        End If
        
        Rem v496
        
FusionOk:
        Rem Mise � jour du statut
        L_statut = "Fusionn�"
        L_commentaire = "Fusionn� le " & Format(Now, "Le dd/mm/yyyy � HH:mm:ss.")
        
        Rem Controle des r�pertoires
        If pFkFlow > 0 Then
            'FlowSqlUpdate = " status = 26, maj_date = now(), maj_userid = '" & G_User_Id & "' "
            'Call Run_Update_Sql(G_Adoconnection, "FLOW", FlowSqlUpdate, "pk_flow = " & pFkFlow)
            If L_liste_fichiers_joints_reception <> "" Then
                Rem � ce stade, s'il existe des fichiers dans la r�ception, on peut les supprimer !!!
                Call FolderExists(G_dir_reception & "\" & p_Customer_Number, True)
                Call Deplacer_les_fichiers(G_dir_reception & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"), _
                                           G_dir_reception, _
                                           L_liste_fichiers_joints_reception)
                
            End If
        Else
            Call Display_Status("Contr�les d'existence du r�pertoire de production Client:", G_dir_production & "\" & p_Customer_Number, p_Form)
            If Not FolderExists(G_dir_production & "\" & p_Customer_Number, True) Then
                L_commentaire = Valid_Text("Impossible d'acc�der au r�pertoire " & G_dir_production & "\" & p_Customer_Number)
                GoTo Erreur_Fusion
            End If
            Call Display_Status("Contr�les d'existence du r�pertoire de r�ception Client / Jour:", G_dir_production & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"), p_Form)
            If Not FolderExists(G_dir_production & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"), True) Then
                L_commentaire = Valid_Text("Impossible d'acc�der au r�pertoire " & G_dir_production & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"))
                GoTo Erreur_Fusion
            End If
            Rem Transfert du fichier data transform� dans le r�pertoire de production / date
            Call Display_Status("Transfert du fichier Data...", "", p_Form)
            Call Move_File(G_dir_production & "\" & p_Customer_Number, G_dir_production & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"), L_data_file_name)
            Rem D�placement des fichiers joints de la r�ception vers la r�ception / client / Jour
            Call Deplacer_les_fichiers(G_dir_reception & "\" & p_Customer_Number & "\" & Format(Date, "YYYYMMDD"), _
                                       G_dir_reception, _
                                       L_liste_fichiers_joints_reception)
            Rem D�placement des fichiers joints de la production_locale vers la r�ception / client / Jour
            Rem LCI suppression du nettoyage inutile !!!
            'If G_dir_local_production <> "" Then
            '    Call Deplacer_les_fichiers(G_dir_reception & "\" & p_Customer_number & "\" & Format(Date, "YYYYMMDD"), _
                                           G_dir_local_production, _
                                           L_liste_fichiers_joints_production_locale)
            'End If
        End If
        Rem D�placement des fichiers rapproch�s
        On Error Resume Next
        If ListePdfRapproches <> "" Then
            rfDir = G_dir_client & "\" & p_Customer_Number & "\" & G_CONST_FICHIERS_RAPPROCHEMENT_FACULTATIF & "\Rapproch�s\"
            If Not FolderExists(rfDir, True) Then
                L_commentaire = Valid_Text("Impossible d'acc�der au r�pertoire " & rfDir)
                GoTo Erreur_Fusion
            End If
            rfDir = G_dir_client & "\" & p_Customer_Number & "\" & G_CONST_FICHIERS_RAPPROCHEMENT_FACULTATIF & "\Rapproch�s\" & Format(Date, "YYYYMMDD") & "\"
            If Not FolderExists(rfDir, True) Then
                L_commentaire = Valid_Text("Impossible d'acc�der au r�pertoire " & rfDir)
                GoTo Erreur_Fusion
            End If
            
            rfTab = Split(ListePdfRapproches, "|")
            For i = 0 To UBound(rfTab)
                If rfTab(i) <> "" Then
                    rfFileName = FileName(rfTab(i))
                    FileCopy rfTab(i), rfDir & rfFileName
                    Kill rfTab(i)
                End If
            Next
        End If
Rem v463
        Rem STATUT : attention, ici, il convient de v�rifier que NetCo et Consulteasy n'utilisent plus le champ STATUT
        Call Run_Update_Sql(G_Adoconnection, _
                            "PREPARATION", _
                            "FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'," & _
                            "MAJ_DATE = now(), MAJ_USERID = '" & G_User_Id & "'", _
                            "PK_PREPARATION = " & p_preparation_pk)
Rem v463
    
    Else
Rem v463
        'L_statut = "Erreur"

        
Rem v463
        L_commentaire = Valid_Text(L_resultat_fusion)

Erreur_Fusion:
        If InStr(1, L_resultat_fusion, "Concat_PDF - Impossible de concat�ner/ouvrir le fichier suivant:", vbTextCompare) > 0 Or InStr(1, L_resultat_fusion, "Le fichier joint est inexploitable", vbTextCompare) > 0 Then
            L_statut = "Avort� fusion"
        Else
            L_statut = "Erreur fusion"
        End If
        
        Call Run_Update_Sql(G_Adoconnection, _
                            "PREPARATION", _
                            " FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'", _
                            "PK_PREPARATION = " & p_preparation_pk)
        Call Word_Close
        
Rem v425 - Suppression des plis d�j� g�n�r�s
Rem Permet notement de ne pas afficher les plis 'A valider sur le web'

        If p_ePoBox_In = "0" Then
            Call Nettoyage_Fusion("Pr�paration", _
                                  p_preparation_pk, _
                                  p_Customer_Number, _
                                  p_Form)
        Else
            Call Nettoyage_Fusion_ePoBox("Pr�paration", _
                                         p_preparation_pk, _
                                         p_Customer_Number, _
                                         p_Form)
        End If
        
        'If InStr(1, L_resultat_fusion, "concat�ner/ouvrir", vbTextCompare) > 0 Then
        Rem v456
        'If p_Customer_Number = "033411000233" Then
'        Select Case p_Customer_Number
'        Case "033411000233", "033411000331"
'            GoTo Maj
'        Case Else
'            GoTo Error_And_File_Open_Close
'        'End If
'        End Select
    End If
    Call Display_Status("", "", p_Form)
    
Exit Sub

Error:
    Rem Mise � jour du statut
Rem v463
    'L_statut = "Erreur"
    L_statut = "Erreur fusion"
Rem v463
    L_commentaire = "Erreur Fusion: " & Valid_Text(Err.Number & " " & Err.Description)
    
Maj:
    Call Run_Update_Sql(G_Adoconnection, _
                        "PREPARATION", _
                        " FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & Valid_Text(L_commentaire) & "', MAJ_DATE = now(), MAJ_USERID = '" & G_User_Id & "'", _
                        "PK_PREPARATION = " & p_preparation_pk)
    If G_wd_opened Then
        Call Word_Close
    End If
    Exit Sub
    

Error_And_File_Open_Close:
    
Dim L_Error_File As String
Dim L_Error_Line As Long
Dim L_Error_Field As String
Dim L_Error_Value As String
Dim L_Error_Msg As String
Dim L_Error_Customer_Number As String
Dim L_Error_Record_In As String
Dim L_Error_Prestation_Model As String
Dim L_Error_Name As String
Dim L_Error_Date As Date
Dim L_error_filename As String

    L_Error_File = Lire_Un_Champ("nom_fichier_info", "RECEPTION", "PK_RECEPTION = '" & p_pk_reception & "'")
    
Rem v425
    If Fusion_Robot.Text_Systeme_Status_Data = vbNullString Then
        L_Error_Line = 1
    Else
        L_Error_Line = 1 + CLng(Left(Fusion_Robot.Text_Systeme_Status_Data, CLng(InStr(1, Fusion_Robot.Text_Systeme_Status_Data, "/", vbTextCompare) - 1)))
    End If
Rem v425
    L_Error_Field = "Nom du fichier info d�pos�"
    L_Error_Value = L_Error_File
    L_Error_Msg = Replace(L_resultat_fusion, G_dir_local_production & "\", "")
    L_Error_Customer_Number = p_Customer_Number
    
    L_Error_Record_In = 0
    L_Error_Prestation_Model = Lire_Un_Champ("PRESTATION_MODEL_NOM", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & L_Prestation_Model_Pk)
    L_Error_Name = "n/a"
    L_Error_Date = Now
    
    Call Error_File("fusion", _
                    L_Error_File, _
                    L_Error_Line, _
                    L_Error_Field, L_Error_Value, _
                    L_Error_Msg, _
                    L_Error_Customer_Number, _
                    "write", _
                    L_Error_Record_In, _
                    L_Error_Prestation_Model, _
                    L_Error_Name, _
                    L_Error_Date, _
                    0)
                    
    L_error_filename = Error_File("fusion", _
                                  L_Error_File, _
                                  L_Error_Line, _
                                  L_Error_Field, L_Error_Value, _
                                  L_Error_Msg, _
                                  L_Error_Customer_Number, _
                                  "close", _
                                  L_Error_Record_In, _
                                  L_Error_Prestation_Model, _
                                  L_Error_Name, _
                                  L_Error_Date, _
                                  0)
                    
    If L_error_filename <> vbNullString Then
        Rem Envoi sur le ftp
        If FTP_Connect_new(G_Ftp_Out, Left$(L_Error_File, 12)) = vbFalse Then
            Call Display_Status(L_error_filename, G_msg_error_p1, Fusion_Robot)
        Else
            Call FTP_SendFile(G_Ftp_Out, L_error_filename)
            Call Ftp_Close(G_Ftp_Out)
        End If
        '
        Set fs = New Scripting.FileSystemObject
        Path_name = fs.GetParentFolderName(L_error_filename)
        File_name = fs.GetFileName(L_error_filename)
        Set fs = Nothing
        '
        Call Sequestre(CStr(Path_name), File_name, p_Customer_Number)
        
    End If
    
End Sub

Rem v406
'Public Sub Fusion_Cancel(p_cust_number As String, _
                         ByVal p_num_reception As String, _
                         ByVal p_preparation_pk As Long, _
                         ByRef p_Form As Form)
                         
Public Sub Fusion_Cancel(p_cust_number As String, _
                         ByVal p_preparation_pk As Long, _
                         ByRef p_Form As Form)

    Dim L_file              As String
    Dim L_statut            As String
    Dim L_commentaire       As String
    Dim L_dir_to            As String
    Dim PkFlow              As String
    
    L_statut = "Annul� fusion"
    L_commentaire = "Fusion annul�e manuellement le " & Format(Now, "dd/mm/yyyy � HH:mm:ss")
    Call Run_Update_Sql(G_Adoconnection, _
                        "PREPARATION", _
                        "FIN = NOW(), STATUS = '" & L_statut & "', COMMENTAIRE = '" & L_commentaire & "'", _
                        "PK_PREPARATION = " & p_preparation_pk)
    
    Call Nettoyage_Fusion("Annulation", p_preparation_pk, p_cust_number, p_Form)
    
    Rem Supression de tous les fichiers Info, Data, et joints (+tard)
    Rem Info
    L_dir_to = G_dir_preparation & "\" & p_cust_number
    Call FolderExists(L_dir_to, True)
    L_dir_to = L_dir_to & "\Erreurs"
    Call FolderExists(L_dir_to, True)
    
    L_file = Lire_Un_Champ("NOM_FICHIER_DATA", "PREPARATION", "PK_PREPARATION = " & p_preparation_pk)
    Rem Data Converti
    Call Move_File(G_dir_preparation & "\" & p_cust_number, L_dir_to, L_file)
    
    Rem Data Transform�
    L_dir_to = G_dir_production & "\" & p_cust_number
    Call FolderExists(L_dir_to, True)
    L_dir_to = L_dir_to & "\Erreurs"
    Call FolderExists(L_dir_to, True)
    Call Move_File(G_dir_production & "\" & p_cust_number, L_dir_to, L_file)
    
    PkFlow = Lire_Un_Champ("fk_flow", "preparation", "pk_preparation = " & p_preparation_pk)
    If PkFlow <> "" Then
        Call Run_Sql_Update(G_Adoconnection, "flow", "status = 46, maj_date = now(), maj_userid = '" & G_User_Id & "'", "pk_flow = " & PkFlow)
        Rem Si il reste encore des fusions non annul�es, je vais directement en CancelEnd:
        L_statut = Lire_Un_Champ("count(*)", "preparation", "status not in ('D�coup�', 'Annul� fusion') and fk_flow = " & PkFlow)
        If L_statut <> "0" Then
            GoTo CancelEnd
        End If
    End If
    
    
    If MsgBox("voulez-vous annuler la r�ception associ�e?", vbExclamation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
        Dim Pk_Reception As String
        Pk_Reception = Lire_Un_Champ("fk_reception", "preparation", "pk_preparation = " & p_preparation_pk)
        If Pk_Reception <> "" Then
Raison:
            Me.Text_COMMENTAIRE = InputBox("Veuillez saisir la raison de l'annulation :", "Raison de l'annulation", Me.Text_COMMENTAIRE)
            If MsgBox("La raison suivante vous satisfait?" & vbNewLine & vbNewLine & Me.Text_COMMENTAIRE, vbQuestion + vbYesNo, Me.Caption) <> vbYes Then
                GoTo Raison
            End If
            Call Run_Sql_Update(G_Adoconnection, "reception", "statut = 'Annul�', info_client = '" & Valid_Text(Me.Text_COMMENTAIRE) & "', commentaire = 'Annul� en fusion le " & Now & "', maj_date = now(), maj_userid = '" & G_User_Id & "'", "pk_reception = " & Pk_Reception)
            
        End If
    End If
    
CancelEnd:
    Rem Joints
    Rem Liste des fichiers joints � faire � partir du fichier data
    Rem Algo
    Rem     Rechercher les noms de balise des champs recus de type Fichier joint ou Image
    Rem     Pour chacune de ces balises, faire une liste des fichiers joints
    Rem     Supprimer le contenu de la liste!!
    
    Rem Remarque: dans le cas ou un fichier data ne peut �tre rattach� � un client
    Rem           pr�voir une proc�dure sp�ciale de nettoyage des fichiers (sur date par exemple)
    Rem
        
   
End Sub

Public Function Fusion_Champs_Lies(ByVal p_Doc As String, _
                                   ByVal p_Table_Lies_Tmp As String, _
                                   ByVal p_Nb_champs_lies As Long, _
                                   ByVal p_Champ_Lie_Jointure_Data As String, _
                                   ByRef p_Connexion As ADODB.Connection, _
                                   ByRef p_Form As Form)


'Dim L_recordset     As ADODB.Recordset
Dim L_SQL           As String
Dim L_field         As Field

'Dim L_recorset2     As Adodc
On Error GoTo Error_Fusion_Champs_Lies
   
Rem v375
    Fusion_Champs_Lies = vbNullString
Rem v375

    Dim L_Array_Format_Cells()  As String  ' List of Cells formats to applicate
    Dim L_Array_Format_Cell()   As String   ' Cell Format to applicate
    Dim L_Array_Merge()         As String         ' Merge to do
    Dim L_Merge_Method          As String
    Dim i                       As Integer
    Dim L_CellRowIndex          As Integer       'Index of row of current cell
    Dim L_Nb_Records_To_Link    As Long
    Dim L_counter               As Long
    Dim L_info                  As String
    Dim L_Data                  As String
    Dim L_Value                 As String
    Dim L_Ubound_Cell           As Integer
    
    L_info = p_Form.Text_Systeme_Status_Info
    L_Data = p_Form.Text_Systeme_Status_Data

    If InStr(1, p_Doc, ".doc") = 0 Then
        p_Doc = p_Doc & ".doc"
    End If
    
    L_CellRowIndex = 2
    L_counter = 0

    L_Nb_Records_To_Link = Lire_Un_Champ("COUNT(*)", p_Table_Lies_Tmp, "CHAMP_LIE_JOINTURE = '" & Valid_Text(p_Champ_Lie_Jointure_Data) & "'")

    G_Wd.Documents(p_Doc).Application.Selection.GoTo What:=wdGoToTable, Which:=wdGoToFirst
    G_Wd.Documents(p_Doc).Application.Selection.GoToNext wdGoToLine
    
            L_SQL = " SELECT * FROM " & p_Table_Lies_Tmp
    L_SQL = L_SQL & " WHERE CHAMP_LIE_JOINTURE = '" & Valid_Text(p_Champ_Lie_Jointure_Data) & "'"
    L_SQL = L_SQL & " ORDER BY PK_CHAMPS_LIES_TMP"
    
'G_Wd.Visible = False
'
'G_Wd.Visible = True
    
    'Set L_recordset2 = Fusion_main.Adodc_tool
    Call Init_Adodc(Fusion_main.Adodc_Tool)
    'Set L_recordset = New ADODB.Recordset
    'L_recordset.Open L_Sql, p_connexion
    Fusion_main.Adodc_Tool.RecordSource = L_SQL
    Fusion_main.Adodc_Tool.Refresh
    
    If Not Fusion_main.Adodc_Tool.Recordset.EOF Then
        Fusion_main.Adodc_Tool.Recordset.MoveFirst
        While Not Fusion_main.Adodc_Tool.Recordset.EOF
            
            For Each L_field In Fusion_main.Adodc_Tool.Recordset.Fields
                L_Value = do_read_data(Fusion_main.Adodc_Tool.Recordset(L_field.Name), "Alpha")
                Select Case LCase(L_field.Name)
                Case "pk_champs_lies_tmp", "champ_lie_jointure"
                    Rem nothing to do
                Case "format_cells"
                    Rem Format cells
Rem v375
                    'L_Data = Do_Read_Data(L_recordset(L_field.Name), "Alpha")
                    L_Ubound_Cell = 0
                    'If Trim$(L_recordset(L_field.Name)) <> vbNullString Then
                    If Trim$(L_Value) <> vbNullString Then
                        'L_Array_Format_Cells = Split(L_recordset(L_field.Name), "|")
                        L_Array_Format_Cells = Split(L_Value, "|")
                            Rem ex Value: "1;TRBL;L;;;VAG Rounded Lt|2;TRBL;L;;;VAG Rounded Lt|3;TRBL;L;;;VAG Rounded Lt|4;TRBL;L;;;VAG Rounded Lt|5;N;;;;"
Rem v375
                        For i = 0 To UBound(L_Array_Format_Cells)
                            L_Array_Format_Cell = Split(L_Array_Format_Cells(i), ";")
                            L_Ubound_Cell = UBound(L_Array_Format_Cell)
                            If L_Ubound_Cell <> -1 Then
                            If Trim$(L_Array_Format_Cell(4)) = vbNullString Then
                                L_Array_Format_Cell(4) = 0
                            End If
                            Call Word_Format_Cell(G_Wd.Documents(p_Doc), _
                                                1, _
                                                L_CellRowIndex, _
                                                CInt(L_Array_Format_Cell(0)), _
                                                L_Array_Format_Cell(1), _
                                                L_Array_Format_Cell(2), _
                                                L_Array_Format_Cell(3), _
                                                CInt(L_Array_Format_Cell(4)), _
                                                L_Array_Format_Cell(5))
                            Else
                                Fusion_Champs_Lies = "Erreur de param�trage des informations de formattage (champs li�s)"
                                Exit Function
                            End If
                        Next
                        G_Wd.Documents(p_Doc).Application.Selection.GoToNext wdGoToLine
                        Rem id if array6 = np then Selection.InsertBreak Type:=wdPageBreak
                        If L_Ubound_Cell > 5 Then
                            If UCase(L_Array_Format_Cell(6)) = "Y" Then
                                G_Wd.Documents(p_Doc).Application.Selection.GoToPrevious wdGoToLine
                                G_Wd.Documents(p_Doc).Application.Selection.InsertBreak Type:=wdPageBreak
                                G_Wd.Documents(p_Doc).Application.Selection.GoToNext wdGoToLine
                                G_Wd.Documents(p_Doc).Application.Selection.GoToNext wdGoToLine
                                If L_Ubound_Cell > 6 Then
                                    If IsNumeric(L_Array_Format_Cell(7)) Then
                                        G_Wd.Documents(p_Doc).Application.Selection.MoveRight Unit:=wdCell, Count:=CLng(L_Array_Format_Cell(7))
                                    End If
                                Else
                                    G_Wd.Documents(p_Doc).Application.Selection.MoveRight Unit:=wdCell, Count:=1
                                End If
                                G_Wd.Documents(p_Doc).Application.Selection.GoToPrevious wdGoToLine
                                L_CellRowIndex = 1
                            End If
                        End If
                    End If
                    
                Case "merge_cells"
                    
                    If Trim$(L_Value) <> vbNullString Then
                        L_Array_Merge = Split(L_Value, "|") 'Ex value = "1-3|O"
                        If UBound(L_Array_Merge) = 0 Then
                            L_Merge_Method = vbNullString
                        Else
                            L_Merge_Method = L_Array_Merge(1)
                        End If
                        Call Word_Merge_Cells(G_Wd.Documents(p_Doc), _
                                              1, L_CellRowIndex, _
                                              L_Array_Merge(0), L_Merge_Method)
                    End If
                               
                Case Else
                    G_Wd.Documents(p_Doc).Application.Selection.TypeText Text:=L_Value
                    G_Wd.Documents(p_Doc).Application.Selection.MoveRight Unit:=wdCell, Count:=1
                End Select
            Next
            Fusion_main.Adodc_Tool.Recordset.MoveNext
            L_counter = L_counter + 1
            
Rem v416
            If L_counter Mod 250 = 0 Then
                G_Wd.Documents(p_Doc).Save
            End If
Rem v416
            Call Display_Status(L_info & " ( D�tail des champs li�s )", L_Data & " ( " & L_counter & " / " & L_Nb_Records_To_Link & " )", p_Form)
            Call CheckApplication("Fusion processing " & L_info & " - Link fields / " & L_counter & " / " & L_Nb_Records_To_Link)
            L_CellRowIndex = L_CellRowIndex + 1
        Wend
        
        G_Wd.Documents(p_Doc).Application.Selection.Rows.Delete
    End If
    
    G_Wd.Documents(p_Doc).Save
    
    Fusion_Champs_Lies = "Ok"

    
Exit Function

Error_Fusion_Champs_Lies:
      Fusion_Champs_Lies = Err.Number & " - " & Err.Description
      Exit Function
      


End Function

Public Function Fusion_fichier_Data(ByVal p_Dir_From As String, _
                                    ByVal p_Data_File_Name As String, _
                                    ByVal p_Prestation_Model_Pk As Long, _
                                    ByVal pCustomerNumber As String, _
                                    ByVal p_Societe_Fk As String, _
                                    ByVal p_Preparation_Fk As Long, _
                                    ByRef p_Liste_Fichiers_Joints_Reception As String, _
                                    ByRef p_liste_fichiers_joints_production_locale As String, _
                                    ByRef p_Liste_Tables_Temporaires_a_Supprimer As String, _
                                    ByRef p_Form As Form, _
                                    ByVal p_Fk_Contact As String, _
                                    ByVal p_Nb_lignes_a_traiter As Long, _
                                    ByVal p_NumPrepa As String, _
                                    ByVal pUnzipDir As String, _
                                    ByVal pFkFlow As Long, _
                                    ByRef pListePdfRapproches As String) _
                                    As String


Rem Declaration des variables locales
    Rem Fichier data en lecture
    Dim L_Fnum                                                  As Long
    Dim L_Fnom                                                  As String
    Dim L_ligne                                                 As String
    Dim L_Nb_champs_emis_detail                                 As Long
    'Dim PatchPonceletNbPages                                    As Long      Rem LCI le 14/08/2017 Suppression de Masque PDF
    Dim NumReception                                            As String
    
    
Rem Requ�tes Prestation Mod�le
    Dim L_array()                                               As String

    Dim L_Valeur_balise_ligne                                   As String
Rem Chaine SQL
    Dim L_SQL                                                   As String

Rem Chaine des champs pour Word
    Dim L_Word_fields                                           As String
Rem Compteur pour le nom des fichiers
    Dim L_Compteur                                              As Integer
    Dim L_Compteur_precedent                                    As Integer
Rem Divers
    Dim L_Nom_datasource_doc                                    As String
    Dim L_Nb_champs_emis_nom                                    As Long
    Dim L_Chemin_nom_datasource_doc                             As String
    Dim L_ModelName                                             As String
    Dim L_New_doc                                               As String
    Dim L_Index                                                 As Long
    
    Rem identifiant Post<e>asy
    Dim L_ID_PE                                                 As String
    Dim L_Sql_Insert                                            As String
    Dim L_Sql_value                                             As String
    Dim L_Fk_destinataire                                       As String
    Dim L_Fk_Pli                                                As String

    Dim L_Nb_pages_dans_pli                                     As Long
    
    Rem Lecture des donn�es
    Dim L_Chaine_header                                         As String
    Dim L_Chaine_data                                           As String
    Dim L_Num_data_source
    Dim L_Bool                                                  As Boolean
            
Rem Suivi des plis
    Dim L_Nb_lignes_traitees                                    As Long
    Rem Num�ro d'enregistrement dans le fichier d'origine
    Dim L_Num_Record                                            As Long

           
Rem Conteneur de la connexion
    Dim L_connexion                                             As ADODB.Connection
    
    'Dim L_Recordset_Tmp                                         As ADODB.Recordset     'LCI SUPPRESSION CONSERVATION LE 14/08/2017

    'Dim L_Reception_fk                                          As String
    
    Dim UpdateAdresse                                           As Boolean
    
    Dim L_Lignes_lisibles                                       As Boolean
    
    Dim L_Pli_statut_pk                                         As Long
    
    'Dim L_Index_Champ                                           As Long                'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    Dim i                                                       As Long
    Dim ii                                                      As Long
    Dim iii                                                     As Integer
    Dim iv                                                      As Integer
    'Dim ENREG                                                   As Long                'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    'Dim HEADER_FLAG                                             As Boolean             'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    'Dim Nom_Champ()                                             As String              'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    Dim l                                                       As String
    'Dim NB_Champ                                                As Long                'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    'Dim L_Index2                                                As Long                'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    
    Dim L_New_pdf                                               As String
    
    Dim L_OF_Auto                                               As Boolean
    Dim L_Nb_Fichier_Joint                                      As Long
    
    Dim L_PageDeGarde                                           As Boolean
    Dim L_Fond_de_Page_P2P                                      As Boolean
    
    Dim L_Type_PDF                                              As Boolean
    Dim L_list_Pdf_Files                                        As String
    
    Dim L_TG                                                    As Boolean
    Dim L_Sequence                                              As Long

    Dim L_Validation_Contenu                                    As Boolean

    Dim L_Referencement                                         As String
    Dim L_Referencement_Fk                                      As Long
    
    Dim L_Liste_Pieces_Jointes                                  As String
    Dim L_sql_mail                                              As String
    
    Dim lMailCC                                                 As String
    Dim lMailCCI                                                As String
    Dim lMailReplyTO                                            As String

Rem // CHAMPS LIES
    Dim L_Champ_Lie                                             As Boolean
    Rem Table temporaire relative aux champs li�s
    Dim L_Table_Lies_Tmp                                        As String
    Rem Nom du champ lie de jointure
    Dim L_Nom_Champ_lie_jointure                                As String
    Rem Lecture du nombre de champs li�s d�finis
    Dim L_Nb_champs_lies                                        As Long
    Rem ????
    Dim L_champ_lie_jointure_data                               As String
    Rem
    Dim L_chemin                                                As String
    Dim L_Type_Fichier_joint                                    As String
    
Rem // Les mails
    Dim L_Email_From                                            As String
    Dim L_Email_xFer                                            As String
    Dim L_Email_ReplyTo                                         As String
    Dim L_Email_to                                              As String
    Dim L_Email_Sql_Pli_Update                                  As String
    Dim L_Email_Piece_Jointe_Fusionnee                          As Boolean
    Dim L_CorpsMailFusionne                                     As Boolean
    Dim L_SujetMailFusionne                                     As Boolean
    Dim L_Url_Dir                                               As String
    
Rem // Liste des statuts (pour �viter les requ�tes � chaque enregistrement)
    Dim L_Statut_A_Valider_Par_Le_Client                        As String
    Dim L_Presence_Statut_Dynamique                             As Boolean
    Dim L_Statut_Telegramme_a_traiter                           As String
    
Rem // Fax
    Dim L_FaxSmsNumber                                          As String
Rem // Sms
    Dim L_eSmsAddressFrom                                       As String
    Dim L_eSmsAddressTo                                         As String

Rem // Les SERVICES
    Dim L_Service_Transformation_fax                            As Boolean
    
    Dim L_Service_Conservation_Data                             As Boolean
    Dim L_Service_Transformation_Mail                           As Boolean
    Dim L_Service_Transformation_Sms                            As Boolean
    
Rem Cr�ation des coffres
    Dim L_Service_Creation_Des_Coffres                          As Boolean
    Dim L_Sql_CreateUserInInsert                                As String
    Dim L_Sql_CreateUserInSelect                                As String
    Dim TabRue()                                                As String
    Dim IndRue                                                  As Integer
    Dim IndRueInsert                                            As Integer
    
Rem // PLI
    Dim L_Sql_InsertPli                                         As String
    Dim L_Sql_ValuesPli                                         As String
Rem // ePoBox
    Dim L_Service_ePoBox                                        As Boolean
    Dim L_ePoBox_Prestation_Model_Nom                           As String
    Dim L_ePoBox_FileToXfer                                     As String
    Dim L_ePoBox_Nb_Fields                                      As Long
    Dim L_Pk_Pli_EpoBox_Emetteur                                As Long

Rem // Variables ePoBox
    Dim L_ePoBox_Destinataire_PlateformId                       As String
    Dim L_ePoBox_Destinataire_ClientID                          As String
    Dim L_ePoBox_Destinataire_Prestation_Model_Nom              As String
    Dim L_ePoBox_Destinataire_ContactId                         As String
    
    Dim L_ePoBox_Destinataire_Reference                         As String
    Dim L_ePoBox_Destinataire_Adresse                           As String
    Dim L_ePobox_Id_Liste_Recapitulative                        As String
    Dim L_ePoBox_Hybrid                                         As Boolean
    
Rem // Variable multi-emploi
    Dim L_Result                                                As String
    
    Dim L_List_Fichiers_Mail_To_Move                            As String
    'Dim L_tmp_Array()                                           As String
    'Dim L_tmp_Array2()                                          As String   REM IMPRESSION DYNAMIQUE SUPPRIMEE LCI LE 14/08/2017
    'Dim L_tmp_Index                                             As Long
    
    Dim L_Document_Size                                         As Double
    Dim L_Fond_de_Page                                          As String
    Dim L_Modele_Merge                                          As Boolean
    Dim L_Pdf_Generated                                         As Boolean
    Dim L_Type_Impression                                       As String

    Dim L_Pk_Of                                                 As Long
    
    Dim L_Mnt_Affranchissement                                  As Double
    Dim L_Mnt_Service                                           As Double
    Dim L_Poids_Pli                                             As Double
    Dim L_Pli_Zone_Postale                                      As String
    Dim L_Pli_Code_pays_Iso3A                                   As String
    Dim L_Pli_Service_Postal                                    As String
    Dim L_Pli_Adresse                                           As String
    Dim L_Pli_Adresse_Rue                                       As String
    Dim L_Pli_Adresse_Cp                                        As String
    Dim L_Pli_Adresse_Ville                                     As String
    Dim L_Pli_Adresse_Pays                                      As String
    Dim L_Mnt_Total                                             As Double
    
    Dim L_Sms_Message                                           As String
    Dim L_Envoi_Automatique                                     As Boolean

    Dim L_List_Pk_Pli                                           As String

    Dim L_Archivage_Only                                        As Boolean

    Dim L_ModelType                                             As String
    Dim L_PageCount                                             As Long
    Dim L_SheetCount                                            As Long

    Dim L_NoWord                                                As Boolean
    Dim L_bcTab()                                               As Long
    Dim L_bcCoverFormat                                         As String

    Dim TmpFields                                               As String
    Dim DirSequestre                                            As String
    Dim L_ShortFile()                                           As String
    Dim L_OriginalFile()                                        As String
    
    Dim BlankPage                                               As String
    
    'Dim Sigma                                                   As Boolean LCI Suppression du cas SIGMA le 14/08/2017
    'Dim SigmaStr                                                As String
    
    Dim NbExemplaireSup                                         As Long
    Dim NbPagesIntermediaire                                    As Long
    Dim SupBlankPage                                            As String
    Dim IndiceExemplaires                                       As Long
    
    Dim DestSociete, DestNom, DestPrenom, DestCivilite          As String
    Dim DestAdresseComplete                                     As String
    
    Dim L_Distinct_Index                                        As Boolean
    
    Dim L_FondDePageSurPJOnly                                   As Boolean
    
    Rem Si premium, on traite l'adresse !!!
    Dim Premium                                                 As Boolean
    Dim whiteCB                                                 As Boolean
    
    Dim AdresseFond                                             As String
    Dim AdresseFondTab()                                        As String
    
    Dim t2c_Profil                                              As String
    Dim t2c_IndexDocument                                       As String 'Index
    Dim t2c_NomDocument                                         As String 'Nom du doc dans le SAE
    Dim t2c_Classement                                          As String 'Classement
    Dim t2c_DateDocument                                        As String 'Date du doc
    Dim L_Service_robot_t2c                                     As Boolean
    Dim t2c_WebService_Actif                                    As Boolean
    Dim t2c_Status_pk_pli_list_go                               As String
    Dim t2c_UserID                                              As String
    
    Dim DatamatriX                                              As Boolean
    Dim DtmxVide(0)                                             As String
    
    Dim PM_eSmsAddressFrom                                      As String

    Dim ListeChamps                                             As String
    Dim IndexCL                                                 As Integer
    Dim TabChamps()                                             As String
    Dim PkCL                                                    As String
    
    Dim ePoBoxR                                                 As String
    Dim TypeePoBox                                              As String
    Dim L_Msg                                                   As String
    Dim WaitSignedPdf                                           As Long
    Dim TryDestinataire                                         As Integer



    Rem v606 - Contr�le de pr�sence d'un flux traitement lot (standart)
    Dim decoupePDFs                                 As String   'Cl� du d�coupe PDF
    Dim decoupePDFsPath                             As String   'Chemin du d�coupe PDF
    Rem v606
    Dim tmpRead1                                    As String
    Dim tmpRead2                                    As String
    
    Dim L_Fso                                       As New FileSystemObject
    Dim PacthAs                                     As Boolean
    
    Rem v630
    Dim NewURLforAffranchissement                   As String
    Rem Je rajoute cette variable ICI afin de pouvoir changer de serveur � la vol�e, sans red�marrer le robot
    Rem Si cette valeur est vide, alors, il n'y aura pas besoin de calculer de l'affranchissement dans ce qui suit
    Rem v630
    
    
    t2c_WebService_Actif = True
    t2c_Status_pk_pli_list_go = "" 'On laisse toujours � Wait, sauf � la fin de la fusion
    
    Rem Initialisation d'un page vierge
    BlankPage = G_dir_Documents & "\BlankPage.pdf"
    If Not FileExists(BlankPage) Then
        Call createPdfBlankPage(BlankPage)
        If Not FileExists(BlankPage) Then
            Fusion_fichier_Data = "Impossible de cr�er une page blanche pour la fusion!!!!"
            Exit Function
        End If
    End If
    
    Rem LCI le 14�08/2017 plus de SIGMA !!!!
    'SigmaStr = "," & Replace(Lire_Un_Champ("valeur", "sys_parameters", "code = 'SIGMA'"), " ", "") & ","
    'If InStr(1, SigmaStr, "," & p_Societe_Fk & ",", vbTextCompare) > 0 Then
    '    Rem Cas d'une soci�t� sigma
    '    Sigma = True
    'Else
    '    Sigma = False
    'End If
    
        
        
    If pFkFlow > 0 Then
        Rem Dans ce cas, on lit le type d'action
        decoupePDFsPath = pUnzipDir
        DirSequestre = p_Dir_From & "\" & G_User_Id & "\"
        If Not FolderExists(DirSequestre, True, False) Then
            Fusion_fichier_Data = "Acc�s impossible au r�pertoire de s�questre !!!! " & DirSequestre
            Exit Function
        End If
        DirSequestre = DirSequestre & p_NumPrepa & "\"
        If Not FolderExists(DirSequestre, True, False) Then
            Fusion_fichier_Data = "Acc�s impossible au r�pertoire de s�questre !!!! " & DirSequestre
            Exit Function
        End If
        
    Else
        tmpRead1 = Lire_Un_Champ("nom_fichier_info", "RECEPTION, preparation", "PK_RECEPTION = fk_reception and pk_preparation = " & p_Preparation_Fk)
        tmpRead2 = InStr(1, tmpRead1, "_info_", vbTextCompare)
        If tmpRead2 > 0 Then
            Rem Lecture de la PK de traitement lot
            tmpRead1 = Mid(tmpRead1, tmpRead2)
            tmpRead1 = Replace(tmpRead1, "_info_", "", , , vbTextCompare)
            tmpRead1 = Replace(tmpRead1, ".xml", "", , , vbTextCompare)
            If IsNumeric(tmpRead1) Then
                Rem Cas du traitement lot
                tmpRead2 = Lire_Un_Champ("count(*)", "traitement_lot", "pk_traitement_lot = " & tmpRead1)
                If tmpRead2 = "1" Then
                    decoupePDFs = tmpRead1
                    Rem Contr�le pour v�rifier que le traitement est RECENT et rattach� au m�me client
                    tmpRead1 = Lire_Un_Champ("fk_societe", "traitement_lot, prestation_model", "fk_prestation_model = pk_prestation_model and pk_traitement_lot = " & decoupePDFs)
                    If tmpRead1 <> p_Societe_Fk Then
                        decoupePDFs = ""
                        decoupePDFsPath = ""
                    Else
                        decoupePDFsPath = Lire_Un_Champ("path", "traitement_lot", "pk_traitement_lot = " & decoupePDFs)
                        If InStr(1, decoupePDFsPath, "input", vbTextCompare) > 0 Then
                            decoupePDFsPath = Replace(decoupePDFsPath, "\input", "\output\", , , vbTextCompare)
                        Else
                            decoupePDFsPath = decoupePDFsPath & "\output\"
                        End If
                    End If
                End If
            End If
        End If
    End If
        
    Call CheckApplication("Initialization")
    
    L_Document_Size = 0

    Rem A supprimer LCI ???
    'L_Url_Dir = "http://www.posteasy.com/marketing"
    L_Url_Dir = ""
    
    Rem Table temporaire
    L_Table_Lies_Tmp = vbNullString
    
    Rem v630 Lire l'URL pour le calcul d'affranchissement si besoin
    If Lire_Produit(p_Prestation_Model_Pk) <> "" Then
        Rem Lire produit renvoi le CODE EDITIQUE, s'il ne renvoi rien, il n'y a pas d'affranchissement
        NewURLforAffranchissement = Lire_Un_Champ("valeur", "sys_parameters", "code = 'SWEB_AFFRA_NEW'")
        If NewURLforAffranchissement = "" Then
            Fusion_fichier_Data = "URL pour l'affranchissement non renseign�e dans la base !!!!"
            Exit Function
        End If
    End If
    Rem fin v630
    
    
    'HEADER_FLAG = False            'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    'ENREG = 1                      'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    L_Nb_Fichier_Joint = 0
    
Rem - Initialisation des statuts et variables
    L_Statut_A_Valider_Par_Le_Client = Fusion_Robot.P_Fk_Pli_Statut_AVALI
    
    Rem BOOLEENS
    L_Presence_Statut_Dynamique = Lire_Un_Champ("count(*)", "CHAMP_EMIS", "CHAMP_EMIS_TYPE = 'Statut dynamique ' AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk)
    L_Email_Piece_Jointe_Fusionnee = Lire_Un_Champ("count(*)", "CHAMP_EMIS", "CHAMP_EMIS_TYPE = 'Pi�ce Mail Fusionn�e' AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk)
    L_CorpsMailFusionne = Lire_Un_Champ("count(*)", "CHAMP_EMIS", "CHAMP_EMIS_TYPE in ('Corps Mail Fusionn�', 'Corps Mail R�pertori� Fusionn�', 'Corps Mail Joint Fusionn�') AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk)
    L_SujetMailFusionne = Lire_Un_Champ("count(*)", "CHAMP_EMIS", "CHAMP_EMIS_TYPE in ('Sujet Mail R�pertori� Fusionn�') AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk)
    
Rem Initialisation des champs li�s
    L_Nb_champs_lies = 0
    L_champ_lie_jointure_data = vbNullString
    L_chemin = vbNullString
    
    On Error Resume Next
    L_TG = False
    L_Statut_Telegramme_a_traiter = vbNullString
    L_Archivage_Only = False
    L_NoWord = False
    
    Select Case Lire_Un_Champ("CODE", "PRESTATION_MODEL,TYPE_PRESTATION", "PK_TYPE_PRESTATION=FK_TYPE_PRESTATION AND PK_PRESTATION_MODEL =" & p_Prestation_Model_Pk)
    Case "TG"
        L_TG = True
        L_Statut_Telegramme_a_traiter = pk_statut_telegramme("PROCESS")
    Case "EDA"
        L_Archivage_Only = True
        L_NoWord = True
    End Select
    
Rem suivi
    L_List_Pk_Pli = ""
    L_Nb_lignes_traitees = 0
Rem Contr�les
    L_Lignes_lisibles = False
    L_Liste_Pieces_Jointes = vbNullString

Rem Liste des fichiers de donn�es
    p_Liste_Fichiers_Joints_Reception = vbNullString
    p_liste_fichiers_joints_production_locale = vbNullString
    
    L_Pli_statut_pk = Fusion_Robot.P_Fk_Pli_Statut_AAFFC
    
    Rem VARIABLES RELATIVES A LA PRESTATION MODELE
    TmpFields = ""
    TmpFields = TmpFields & "REGROUPEMENT_PLI, "            '   L_Regroupement_pli = L_array(0)
    TmpFields = TmpFields & "OF_AUTOMATIQUE, "              '   L_OF_Auto = L_array(1)
    TmpFields = TmpFields & "PAGEDEGARDE, "                 '   L_PageDeGarde = L_array(2)
    TmpFields = TmpFields & "TYPE_PDF, "                    '   L_Type_PDF = L_array(3)
    TmpFields = TmpFields & "VALID_CONTENTS, "              '   L_Valid_Contents = L_array(4)
    TmpFields = TmpFields & "VALIDATION_CONTENU, "          '   L_Validation_Contenu = L_array(5)
    TmpFields = TmpFields & "REGROUPEMENT_PAGE, "           '   L_Regroupement_page = L_array(6)
    TmpFields = TmpFields & "REGROUPEMENT_ENREGISTREMENT, " '   L_Regroupement_enregistrement = L_array(7)
    TmpFields = TmpFields & "MODELE_MERGE, "                '   L_Modele_Merge = L_array(8)
    TmpFields = TmpFields & "MODELE_MERGE_P2P, "            '   L_Fond_de_Page_P2P = L_array(9)
    TmpFields = TmpFields & "TYPE_IMPRESSION, "             '   L_Type_Impression = L_array(10)
    TmpFields = TmpFields & "EPOBOX_2_PRINT, "              '   L_ePoBox_Hybrid = (L_array(11) = "1")
Rem v578 - ATTENTION, SI ON ENLEVE, IL FAUT TOUT REINDEXER
    TmpFields = TmpFields & "distinct_index, "              '   L_array(12) => L_Print_Word

    TmpFields = TmpFields & "ENVOI_AUTOMATIQUE, "           '   L_Envoi_Automatique = L_array(13)
    TmpFields = TmpFields & "MODEL_TYPE, "                  '   L_ModelType = L_array(14)
    TmpFields = TmpFields & "bcCoverFormat, "               '   L_bcCoverFormat = L_array(15)
    TmpFields = TmpFields & "bcX, bcY, bcFontSize, "        '   L_array(16), L_array(17), L_array(18)
    TmpFields = TmpFields & "prestation_model_libelle, "    '   L_ePoBox_Prestation_Model_Nom = L_array(19)
    TmpFields = TmpFields & "CHAMP_LIE, "                   '   L_champ_lie = L_array(20)

    TmpFields = TmpFields & "PRESTATION_MODEL_WORD, "        '   L_ModelName = L_array(21)
    TmpFields = TmpFields & "MODELE_MERGE_PJ, "              '   Applique le fond e page uniquement sur les PJ  = L_array(22)
    TmpFields = TmpFields & "AFFRANCHISSEMENT, "              '   si L_array(23) = M ou O => Premium => Traitement Sp�cial !!!
    TmpFields = TmpFields & "xyhl_fond, "              '   si L_array(24) = M ou O => Premium => Traitement Sp�cial !!!
    TmpFields = TmpFields & "datamatrix, "              '   si L_array(25) = 0 ou 1
    TmpFields = TmpFields & "sms_sender "               '   L_array(26)

    Rem Variables de la prestation Mod�le
    L_array = Split(Lire_Des_Champs(TmpFields, "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & p_Prestation_Model_Pk), "|")
    L_OF_Auto = L_array(1)
    L_PageDeGarde = L_array(2)
    L_Type_PDF = L_array(3)
    'L_Valid_Contents = L_array(4)
    L_Validation_Contenu = L_array(5)
    
    L_Modele_Merge = L_array(8)
    L_Fond_de_Page_P2P = L_array(9)
    L_Type_Impression = L_array(10)
    L_ePoBox_Hybrid = (L_array(11) = "1")
    L_Distinct_Index = L_array(12)
    If L_Validation_Contenu Then
        L_Envoi_Automatique = False
    Else
        L_Envoi_Automatique = L_array(13)
    End If
    L_ModelType = L_array(14)
    L_bcCoverFormat = L_array(15)
    ReDim L_bcTab(3)
    L_bcTab(0) = L_array(16)
    L_bcTab(1) = L_array(17)
    L_bcTab(2) = L_array(18)

    L_ePoBox_Prestation_Model_Nom = L_array(19)
    L_Champ_Lie = L_array(20)
    
    L_FondDePageSurPJOnly = L_array(22)

    Select Case UCase(L_array(23))
    Case "M", "O"
        Premium = True
        whiteCB = True
        UpdateAdresse = True
    Case Else
        Premium = False
        whiteCB = False
    End Select

    DatamatriX = L_array(25)
    PM_eSmsAddressFrom = L_array(26)

    Select Case L_ModelType
    Case "Aucun", "CB Direct", "CB + Adresse"
        L_ModelName = vbNullString
        L_NoWord = True
        UpdateAdresse = True
    Case "Mod�le Word"
        L_ModelName = L_array(21)
        If UCase(L_array(23)) <> "A" Then
            UpdateAdresse = True
        End If
    Case Else
        L_ModelName = L_array(21)
        UpdateAdresse = False
    End Select

    ReDim AdresseFondTab(0)
    Select Case L_ModelType
    Case "CB Direct", "CB + Adresse"
        AdresseFond = Trim(L_array(24))
        If AdresseFond <> "" Then
            AdresseFondTab = Split(AdresseFond, ";")
            If UBound(AdresseFondTab) <> 3 Then
                Fusion_fichier_Data = "Le param�trage du fond pour la prestation est incorrecte"
                Exit Function 'Ici, il n'y a ni connexion, ni fichier ouver, ni word ouver
            End If
        End If
    Case Else
        AdresseFond = ""
    End Select

    pListePdfRapproches = ""
    
    If L_Archivage_Only Then
        L_Pli_statut_pk = Fusion_Robot.P_Fk_Pli_Statut_ARCH
    End If
Rem // Fin Lecture / Initialisation des variables

    Fusion_fichier_Data = "Erreur"
    If pFkFlow = 0 Then
        L_Fnom = G_dir_production & "\" & pCustomerNumber & "\" & p_Data_File_Name
    Else
        L_Fnom = p_Dir_From & "\" & p_Data_File_Name
    End If

Rem **************************************************************************************
Rem Chargement des Services
Rem **************************************************************************************

    
    L_Service_Transformation_fax = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_TRANSFORMATION_FAX)
   
    L_Service_robot_t2c = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_T2C)
    If L_Service_Transformation_fax Then
        Call CheckApplication("Fax initialization")
    End If
    
    L_Service_Transformation_Sms = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_TRANSFORMATION_SMS)
    Rem Param�tres d'envoi SMS
    If L_Service_Transformation_Sms Then
        Call CheckApplication("Sms initialization")
        
        'GRAPHNET
        L_eSmsAddressTo = Lire_Un_Champ("valeur", "emetteur_fax_parametres", "parametre = 'eSmsAddressTo'")
        If InStr(1, L_eSmsAddressTo, "@", vbTextCompare) = 0 Then
            L_eSmsAddressTo = "@" & L_eSmsAddressTo
        End If
        L_eSmsAddressFrom = Lire_Un_Champ("valeur", "emetteur_fax_parametres", "parametre = 'eSmsAddressFrom'")
        L_Email_ReplyTo = L_eSmsAddressFrom
        If PM_eSmsAddressFrom <> "" Then
            L_eSmsAddressFrom = PM_eSmsAddressFrom
        End If
    End If
    'LCI SUPPRESSION CONSERVATION LE 14/08/2017
    'L_Service_Conservation_Data = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_CONSERVATION_DONNEES)
    L_Service_Transformation_Mail = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_TRANSFORMATION_MAIL)
    L_Service_ePoBox = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_EPOBOX)

    L_Service_Creation_Des_Coffres = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_CREATION_COFFRE)

Rem **************************************************************************************
Rem Chargement des Variables relatives aux Services
Rem **************************************************************************************
    If L_Service_Transformation_Mail Then
        Call CheckApplication("Mail initialization")
        L_Result = Mail_Init_Parameters_PM(p_Prestation_Model_Pk, _
                                           L_Email_From, _
                                           L_Email_ReplyTo)
        If L_Result <> "Ok" Then
            Fusion_fichier_Data = L_Result
            Exit Function 'Ici, il n'y a ni connexion, ni fichier ouver, ni word ouver
        End If
    End If
    
    If L_Service_ePoBox Then
        Rem Informations pour le fichier xml � g�n�rer
        Rem Plateforme_Ide
        Call CheckApplication("ePobox initialization")
        Rem SenderName
        Rem SenderReference (customer number?)
        Rem PrestationModelName
        Rem Lire le nombre de champs � envoyer...
        L_ePoBox_Nb_Fields = Lire_Un_Champ("count(*)", "champ_emis", "champ_emis_epobox = 1 and fk_prestation_model = " & p_Prestation_Model_Pk)
    End If

    Rem **************************************************************************************
    Rem Champs Li�s
    Rem **************************************************************************************
    'If L_Champ_Lie Then
    If Lire_Un_Champ("CHAMP_LIE", "PRESTATION_MODEL", "PK_PRESTATION_MODEL = " & p_Prestation_Model_Pk) = "1" Then
        Rem ************************************************************************
        Call CheckApplication("Link fields initialization")
        Fusion_fichier_Data = Init_Champs_Lies(L_Nb_champs_lies, _
                                               p_Prestation_Model_Pk, _
                                               L_Nom_Champ_lie_jointure, _
                                               L_Table_Lies_Tmp, _
                                               L_Fnom, _
                                               p_Form)
        p_Liste_Tables_Temporaires_a_Supprimer = L_Table_Lies_Tmp
        If Fusion_fichier_Data <> "Ok" Then
            Fusion_fichier_Data = "Impossible d'initialiser les champs li�s"
            Exit Function 'Ici, il n'y a ni connexion, ni fichier ouver, ni word ouver
        End If
        Rem // v240
        Rem ************************************************************************
    End If
        
    Rem **************************************************************************************
    Rem Champs Emis
    Rem **************************************************************************************
    Call Init_Tab_Champs_Emis_Detail(L_tab_Champs_Emis, _
                                     L_tab_Champs_Emis_detail, _
                                     p_Prestation_Model_Pk, _
                                     L_Nb_champs_emis_detail, _
                                     L_Nb_champs_emis_nom, _
                                     L_Word_fields)
    
    Rem **************************************************************************************
    Rem Fusion
    Rem **************************************************************************************
    Rem Si pas de fusion, pas de mod�le et pas Word � ouvrir
    '
    Rem Ouverture du mod�le,
    Rem Dans le cas sans regroupement page !!!
    Call CheckApplication("Fusion initialization")
    

    Rem Gestion des diff�rents Mod�les
    If L_Service_Transformation_Sms Then
        GoTo Fusion_Suite
    End If
    If L_Archivage_Only Then
        GoTo Fusion_Suite
    End If
    Select Case L_ModelType
    Case "Aucun", "CB Direct", "CB + Adresse"
        L_NoWord = True
        GoTo Fusion_Suite
    Rem LCI le 14/08/2017 Suppression de Masque PDF 'Case "Masque Pdf", "Mod�le Texte", "Mod�le Html"
    Case "Mod�le Texte", "Mod�le Html"
        L_NoWord = True
    End Select
            
    Call Display_Status("Chargement des propri�t�s de la prestation mod�le...", "", p_Form)
    If Trim(L_ModelName) = vbNullString Then
        If L_Service_Transformation_Mail Then
            L_NoWord = True
            GoTo Fusion_Suite
        Else
            Fusion_fichier_Data = "Le mod�le de la Prestation/Mod�le n'est pas renseign�"
            Exit Function 'Ici, il n'y a ni connexion, ni fichier ouver, ni word ouver
        End If
    End If
    
    Rem LCI le 14/08/2017 Suppression de Masque PDF If L_ModelType = "Masque Pdf" Then
        'L_ModelName = Replace(L_ModelName, ".pdf", "", , , vbTextCompare) & ".pdf"
    'Else
    If L_ModelType = "Mod�le Texte" Then
        L_ModelName = Replace(L_ModelName, ".txt", "", , , vbTextCompare) & ".txt"
    ElseIf L_ModelType = "Mod�le Html" Then
        L_ModelName = Replace(L_ModelName, ".html", "", , , vbTextCompare) & ".html"
    ElseIf L_ModelType = "Mod�le Word" Then
        L_ModelName = Replace(L_ModelName, ".doc", "", , , vbTextCompare) & ".doc"
        Call Display_Status("Ouverture de Word...", "", p_Form)
        Set G_Wd = CreateObject("Word.Application")
        G_wd_opened = True
        G_Wd.Visible = False
    End If

    L_Referencement_Fk = Num_Referencement(pCustomerNumber)
    If L_Referencement_Fk <> 0 Then
        L_Referencement = Lire_Un_Champ("NUM_SOCIETE", "SOCIETE", "PK_SOCIETE=" & L_Referencement_Fk)
    Else
        L_Referencement = vbNullString
    End If
    If L_Referencement = vbNullString Then
        If Not FileExists(G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\" & L_ModelName) Then
            Fusion_fichier_Data = "Le mod�le """ & L_ModelName & """ est introuvable dans le r�pertoire " & vbNewLine & G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\"
            GoTo Clean_Exit_Word
        End If
    Else
        If Not FileExists(G_dir_client & "\" & L_Referencement & "\" & G_CONST_MODELES & "\" & L_ModelName) Then
            Fusion_fichier_Data = "Le mod�le """ & L_ModelName & """ est introuvable dans le r�pertoire " & vbNewLine & G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\"
            GoTo Clean_Exit_Word
        End If
    End If
    
    Call Display_Status("Copie du mod�le...", L_ModelName, p_Form)
    Rem Id�e Ajouter un ID de la fusion + heure pour conserver les mod�les utilis�s?

    If pFkFlow = 0 Then
        DirSequestre = CurrentSequestre(pCustomerNumber, "FUSION", G_User_Id, p_NumPrepa)
        If DirSequestre = "KO" Then
            Fusion_fichier_Data = "Le r�pertoire de s�questre """ & DirSequestre & """ ne peut �tre cr��."
            GoTo Clean_Exit_Word
        End If
    End If

    If Not FolderExists(DirSequestre & G_CONST_MODELES, True) Then
        Fusion_fichier_Data = "Le r�pertoire de Mod�les dans le s�questre """ & DirSequestre & """ ne peut �tre cr��."
        GoTo Clean_Exit_Word
    End If

    Rem LCI : R�pertoire local de productino supprim� le 14/08/2017
    'G_dir_local_production = DirSequestre & G_CONST_MODELES
    Rem Copie dans le r�pertoire local
    If L_Referencement = vbNullString Then
        If Not FileExists(G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\" & L_ModelName) Then
            Fusion_fichier_Data = "Le fichier """ & G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\" & L_ModelName & """ n'existe pas!"
            GoTo Clean_Exit_Word
        End If
    Else
        If Not FileExists(G_dir_client & "\" & L_Referencement & "\" & G_CONST_MODELES & "\" & L_ModelName) Then
            Fusion_fichier_Data = "Le fichier """ & G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\" & L_ModelName & """ n'existe pas!"
            GoTo Clean_Exit_Word
        End If
    End If
    On Error Resume Next
    Rem IMPORTANT : copie du mod�le utilis� pour ce lot de fusion
    If L_Referencement = vbNullString Then
        FileCopy G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_MODELES & "\" & L_ModelName, DirSequestre & G_CONST_MODELES & "\" & L_ModelName
    Else
        FileCopy G_dir_client & "\" & L_Referencement & "\" & G_CONST_MODELES & "\" & L_ModelName, DirSequestre & G_CONST_MODELES & "\" & L_ModelName
    End If
    Select Case Err.Number
    Case 0
        On Error GoTo 0
    Case Else
        Fusion_fichier_Data = "Erreur " & Err.Number & "-" & Err.Description & " lors de la copie dans le r�pertoire"
        GoTo Clean_Exit_Word
    End Select
    
    Select Case L_ModelType
    
    Case "Mod�le Word"
        Rem Dans le cas des mod�les WORD, j'ouvre le mod�le et je me tiens pr�t � publiposter
        Call Display_Status("Ouverture du mod�le...", L_ModelName, p_Form)
        Rem Ouverture (locale) du mod�le
        On Error Resume Next
        G_Wd.Documents.Open DirSequestre & G_CONST_MODELES & "\" & L_ModelName, , , False
        Select Case Err.Number
        Case 0
            On Error GoTo 0
        Case Else
            Fusion_fichier_Data = "Erreur " & Err.Number & "-" & Err.Description & " lors de l'ouverture par Word du document : " & DirSequestre & G_CONST_MODELES & "\" & L_ModelName
            GoTo Clean_Exit_Word
        End Select
        On Error Resume Next

        Call Display_Status("Param�trage du mod�le...", L_ModelName, p_Form)
        G_Wd.Documents(L_ModelName).Application.Options.UpdateFieldsAtPrint = True
        G_Wd.Options.PrintBackground = False

        If InStr(1, G_Wd.Documents(L_ModelName).Application.ActivePrinter, G_Driver_PDFCreator, vbTextCompare) = 0 Then
            Call Display_Status("Affectation du driver d'impression (PdfCreator)...", L_ModelName, p_Form)
            G_Wd.Documents(L_ModelName).Application.ActivePrinter = G_Driver_PDFCreator
        End If
    
        G_Wd.Options.CheckGrammarAsYouType = False
        G_Wd.Options.CheckGrammarWithSpelling = False
        G_Wd.Options.CheckSpellingAsYouType = False
        G_Wd.Options.UpdateFieldsAtPrint = True
        G_Wd.Options.PrintFieldCodes = False
    
    Rem LCI le 14/08/2017 Suppression de Masque PDF Case "Masque Pdf", "Mod�le Texte", "Mod�le Html"
    Case "Mod�le Texte", "Mod�le Html"
        L_ModelName = DirSequestre & G_CONST_MODELES & "\" & L_ModelName
    
    End Select
        
    
Fusion_Suite:
Skip_Coca0:

    L_Compteur = 0
    L_Compteur_precedent = 0
Rem Cr�ation des conteneurs
    Set L_connexion = New ADODB.Connection
   
Rem Ouverture de la connexion
    L_connexion.ConnectionTimeout = 0
    L_connexion.CommandTimeout = 0
    L_connexion.Open (G_Adoconnection)
    
Rem **************************************************************************************
Rem Parcours du fichier de donn�es
Rem **************************************************************************************
    Call Display_Status("Lecture des donn�es...", "", p_Form)
    Rem Parcourir le fichier et lire le tableau
    L_Sequence = 0
    L_Fnum = FreeFile
    Open L_Fnom For Input As #L_Fnum
    While Not EOF(L_Fnum)
        Line Input #L_Fnum, L_ligne
        While InStr(UCase(L_ligne), G_Balise_Data_Line_In) > 0
            Rem La balise <line_in> est pr�sente, tourner jusqu'a </line_in>
            While InStr(UCase(L_ligne), G_Balise_Data_Line_Out) = 0
                Line Input #L_Fnum, L_ligne
            Wend
            Rem Si je sors de la boucle, charger la nouvelle ligne � analyser
            Line Input #L_Fnum, L_ligne
        Wend
        
        Call CheckApplication("Fusion reading file")
        
        Rem Si la balise du fichier correspond � la balise attendue
DecodeUTF8:
        If Existe_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">") Then
            Rem Le num�ro de s�quence est valide!
            L_Bool = False
            L_Valeur_balise_ligne = Valeur_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">", L_Bool)
            While L_Bool
                Line Input #L_Fnum, L_ligne
                L_Valeur_balise_ligne = L_Valeur_balise_ligne & vbNewLine & Valeur_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">", L_Bool)
            Wend

            L_tab_Champs_Emis(L_Sequence).t_champ_emis_data = L_Valeur_balise_ligne
            Rem *****************************
            Rem Incrementation de la s�quence
            Rem *****************************
            Rem En fin de s�quence
            Rem     et on fusionne
            Rem     et on red�marre
            
            If L_Sequence = L_Nb_champs_emis_nom - 1 Then
Rem Analyse
Rem Il suffirait de d�tecter un champ </RECORD_IN>
                Rem A ce stade, tous les champs d'une ligne sont lus
                Rem initialisation des variables
                If L_Service_Creation_Des_Coffres Then
                    L_Sql_CreateUserInInsert = " insert into creation_utilisateur_in ("
                    L_Sql_CreateUserInSelect = " select "
                Else
                    L_Sql_CreateUserInInsert = ""
                    L_Sql_CreateUserInSelect = ""
                End If
                L_Sql_InsertPli = ""
                L_Sql_ValuesPli = ""
                L_Lignes_lisibles = True
                L_Sequence = 0
                L_chemin = vbNullString
                Rem Et dans le tableau Champ_emis!!
                L_ID_PE = vbNullString
                L_Email_Sql_Pli_Update = vbNullString
                L_List_Fichiers_Mail_To_Move = vbNullString
                L_ePobox_Id_Liste_Recapitulative = vbNullString
                
                DestSociete = vbNullString
                DestNom = vbNullString
                DestPrenom = vbNullString
                DestCivilite = vbNullString
                DestAdresseComplete = vbNullString
                
                L_Pli_Adresse = vbNullString
                L_Pli_Adresse_Rue = vbNullString
                L_Pli_Adresse_Cp = vbNullString
                L_Pli_Adresse_Ville = vbNullString
                L_Pli_Adresse_Pays = vbNullString
                
                t2c_Profil = vbNullString
                t2c_IndexDocument = vbNullString                                'Index
                t2c_NomDocument = vbNullString                              'Nom du doc dans le SAE
                t2c_Classement = vbNullString                              'Classement
                t2c_DateDocument = vbNullString                              'Date du doc
                t2c_UserID = vbNullString
                
                NbExemplaireSup = 0
                Rem Ici, si L_Distinct_Index = 1, on "dedoublonne" les champs re�us
                If L_Distinct_Index Then
                    For L_Index = 0 To L_Nb_champs_emis_nom - 1
                        If InStr(1, L_tab_Champs_Emis(L_Index).t_champ_emis_data, "|", vbTextCompare) > 0 Then
                            Rem On d�doublonne, sinon rien
                            L_tab_Champs_Emis(L_Index).t_champ_emis_data = DistinctValues(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "|")
                        End If
                    Next
                End If
                ReDim L_ShortFile(0)
                ReDim L_OriginalFile(0)
                
                If pFkFlow = 0 Then
                    DirSequestre = CurrentSequestre(pCustomerNumber, "FUSION", G_User_Id, p_NumPrepa)
                    If DirSequestre = "KO" Then
                        Fusion_fichier_Data = "Impossible de cr�er le r�pertoire de s�questre pour la fusion"
                        GoTo Clean_Exit
                    End If
                End If
                
                Rem Si pas de regroupement, ou Premier d'un regroupement potentiel
                TryDestinataire = 0
                If L_ID_PE = vbNullString Then
NextTryDestinataire:
                    L_ID_PE = Attribuer_un_numero_New("PLI", "033")
                End If
                
                Rem Si le destinataire n'existe pas encore, cr�ation de la chaine SQL
                L_sql_mail = vbNullString
                    
                If L_Service_ePoBox Then
                    ReDim L_tab_ePoBox_Data(L_ePoBox_Nb_Fields)
                End If

                L_Result = Add_Destinataire(L_Nb_champs_emis_nom, _
                                             p_Prestation_Model_Pk, _
                                             L_Pli_statut_pk, _
                                             L_sql_mail, _
                                             L_champ_lie_jointure_data, _
                                             p_Societe_Fk, _
                                             L_ID_PE, _
                                             L_Email_From, _
                                             L_Email_xFer, _
                                             L_FaxSmsNumber, _
                                             L_Sms_Message, _
                                             UpdateAdresse, _
                                             DirSequestre, _
                                             pUnzipDir, _
                                             L_Sql_CreateUserInInsert, _
                                             L_Sql_CreateUserInSelect, _
                                             L_Sql_InsertPli, _
                                             L_Sql_ValuesPli)
                If L_Result <> "Ok" Then
                    If TryDestinataire < 3 And InStr(1, L_Result, "insert into", vbTextCompare) > 0 Then
                        TryDestinataire = TryDestinataire + 1
                        GoTo NextTryDestinataire
                    Else
                        Fusion_fichier_Data = "Erreur lors de la cr�ation du destinataire (" & L_Result & ")"
                        GoTo Clean_Exit
                    End If
                End If
                
                Rem Lecture de la pk du destinataire
                L_Fk_destinataire = Lire_Un_Champ("PK_DESTINATAIRE", "DESTINATAIRE", "DEST_ID_PE = '" & L_ID_PE & "'")
                
                Rem ajout du d�tail de mail pour le nouveau destinataire
                If L_sql_mail <> "" Then
                    Call Add_Destinataire_detail(L_Fk_destinataire, L_sql_mail)
                End If
                
                If L_Sql_CreateUserInInsert <> "" Then
                    L_Sql_CreateUserInInsert = L_Sql_CreateUserInInsert & " fk_destinataire, "
                    L_Sql_CreateUserInSelect = L_Sql_CreateUserInSelect & L_Fk_destinataire & ", "
                End If
                
                
                L_Nb_Fichier_Joint = 0
                    
                Rem Rappel, Dans ce cas, le document Word est dej� ouvert (quand il y en a un)
                Rem Cas "Normal" de Fusion
                If L_NoWord = True Then
                    'Faire la liste des variables inutiles ici!!!
                Else
                    Rem Nom de la source de donn�es
                    L_Nom_datasource_doc = "Datasource_" & L_ID_PE & ".csv"
                    L_Chemin_nom_datasource_doc = DirSequestre & L_Nom_datasource_doc
                End If
                L_Index = 0
                Rem Lecture des champs et donn�es
                L_Chaine_header = vbNullString
                L_Chaine_data = vbNullString
                
                L_Num_Record = 0
                L_Nb_Fichier_Joint = 0
                L_ePoBox_FileToXfer = vbNullString
                L_Fond_de_Page = vbNullString
                L_Pdf_Generated = False
                lMailCC = vbNullString
                lMailCCI = vbNullString
                lMailReplyTO = vbNullString
                        
                ReDim L_Fichier_Joint(L_Nb_Fichier_Joint)
                Rem LCI le 14/08/2017 suppression de l'impression dynamique
                'ReDim G_Print_Properties(0)
                        
                For L_Index = 0 To L_Nb_champs_emis_nom - 1
                    L_Chaine_header = L_Chaine_header & L_tab_Champs_Emis(L_Index).t_champ_emis_nom & ";"
                    
                    Select Case L_tab_Champs_Emis(L_Index).t_champ_emis_type
                    Case "t2c_Profil" 'Obligatoire si service t2c activ�
                        t2c_Profil = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "t2c_Classement" 'Par d�faut, aucun
                        t2c_Classement = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "t2c_IndexDocument" 'PAr d�faut Aucun
                        t2c_IndexDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "t2c_DateDocument" 'Par D�faut Aucune
                        t2c_DateDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "t2c_NomDocument" 'Par d�faut identique � celui g�n�r�
                        t2c_NomDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "t2c_UserID"      'A quoi cela sert=il?
                        t2c_UserID = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"

                    Case "Fichier joint", "Corps mail joint", "Fichier(s) multiple(s)", "Pi�ce mail jointe"
                            
                        Rem D�termination automatique du type de fichier
                        L_Type_Fichier_joint = Type_Fichier(Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                        If L_Type_Fichier_joint <> "" Then
                            L_Nb_Fichier_Joint = L_Nb_Fichier_Joint + 1
                            ReDim Preserve L_Fichier_Joint(L_Nb_Fichier_Joint)
                            'MsgBox Len(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                            L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                            L_Fichier_Joint(L_Nb_Fichier_Joint).t_Signet = L_tab_Champs_Emis(L_Index).t_champ_emis_nom
                            L_Fichier_Joint(L_Nb_Fichier_Joint).t_type = L_Type_Fichier_joint
                            If pUnzipDir = "flow" Then
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = DirSequestre
                            ElseIf pUnzipDir = "no" Then
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = DirSequestre
                            Else
                                If decoupePDFsPath <> "" Then
                                    L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = decoupePDFsPath
                                Else
                                    L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = pUnzipDir
                                End If
                            End If
                            ReDim Preserve L_ShortFile(L_Nb_Fichier_Joint)
                            ReDim Preserve L_OriginalFile(L_Nb_Fichier_Joint)

                            PacthAs = False
                            If InStr(1, L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, "/", vbTextCompare) > 0 Then
                                PacthAs = True
                                While InStr(1, L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, "/", vbTextCompare) > 0
                                    L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier = Mid(L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, 1 + InStr(1, L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, "/", vbTextCompare))
                                Wend
                            End If
                            If PacthAs = True Then
                                If Not FileExists(L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier) Then
                                Rem Le fichier est l� et a d�j� �t� v�rifi�, on continue!!!
                                    Fusion_fichier_Data = "Fichier " & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier & " manquant! "
                                    GoTo Clean_Exit
                                End If
                            End If
                            If Len(L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier) > 250 Then
                                L_OriginalFile(L_Nb_Fichier_Joint) = L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                L_ShortFile(L_Nb_Fichier_Joint) = Left(L_OriginalFile(L_Nb_Fichier_Joint), 231 - Len(L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir)) & ".pdf"
                            Else
                                L_OriginalFile(L_Nb_Fichier_Joint) = ""
                                L_ShortFile(L_Nb_Fichier_Joint) = ""
                            End If
                                    
Rem optimisations : ne plus les recopier localement, travailler depuis l'original !!!
                            If decoupePDFsPath <> "" Then
                                GoTo ConvertTest
                            End If
                            If pUnzipDir <> "no" Then
                                If FileExists(L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier) Then
                                    If L_ShortFile(L_Nb_Fichier_Joint) <> "" Then
                                        FileCopy L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, DirSequestre & L_ShortFile(L_Nb_Fichier_Joint)
                                    End If
                                    GoTo ConvertTest
                                End If
                                If FileExists(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                    FileCopy DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data, L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                    GoTo ConvertTest
                                End If
                                If FileExists(G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                    FileCopy G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                    GoTo ConvertTest
                                End If
                            End If
                            If Not FileExists(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                If FileExists(G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                    FileCopy G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                Else
                                    If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                        FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                    Else
                                        For iii = 0 To 20
                                Rem dans reception
                                            If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iii, "YYYYMMDD") & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                                If L_ShortFile(L_Nb_Fichier_Joint) <> "" Then
                                                    FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iii, "YYYYMMDD") & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_ShortFile(L_Nb_Fichier_Joint)
                                                Else
                                                    FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iii, "YYYYMMDD") & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                End If
                                                iii = 30
                                            End If
                                 Rem dans sequestre
                                            If FileExists(G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, vbNormal) Then
                                                If L_ShortFile(L_Nb_Fichier_Joint) <> "" Then
                                                    If pUnzipDir <> "no" Then
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, pUnzipDir & L_ShortFile(L_Nb_Fichier_Joint)
                                                    Else
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_dir_reception & "\" & L_ShortFile(L_Nb_Fichier_Joint)
                                                    End If
                                                Else
                                                    If pUnzipDir <> "no" Then
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                    Else
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                    End If
                                                End If
                                                iii = 30
                                            End If
                                 Rem dans sequestre/Rception
                                            If NumReception = "" Then
                                                NumReception = Lire_Un_Champ("num_reception", "reception, preparation", "fk_reception = pk_reception and num_preparation = '" & p_NumPrepa & "'")
                                            End If
                                            If FileExists(G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & NumReception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, vbNormal) Then
                                                If L_ShortFile(L_Nb_Fichier_Joint) <> "" Then
                                                    If pUnzipDir <> "no" Then
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, pUnzipDir & L_ShortFile(L_Nb_Fichier_Joint)
                                                    Else
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_dir_reception & "\" & L_ShortFile(L_Nb_Fichier_Joint)
                                                    End If
                                                Else
                                                    If pUnzipDir <> "no" Then
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                    Else
                                                        FileCopy G_dir_sequestre & "\" & Format(Date - iii, "YYYYMMDD") & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                    End If
                                                End If
                                                iii = 30
                                            End If
                                            If iii = 20 Then
                                                Rem Nouveau patch on essaye sans le nom du contact
                                                If InStr(1, L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), vbTextCompare) > 0 Then
                                                    If FileExists(G_dir_reception & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare)) Then
                                                        FileCopy G_dir_reception & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare), DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                        iv = 30
                                                    Else
                                                        If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare)) Then
                                                            FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare), DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                            iv = 30
                                                        Else
                                                            For iv = 0 To 20
                                                                If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iv, "YYYYMMDD") & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare), vbNormal) Then
                                                                    FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iv, "YYYYMMDD") & "\" & Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "_" & Lire_Un_Champ("web_id", "contact", "pk_contact = " & p_Fk_Contact), "", 1, , vbTextCompare), DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                                                    iv = 30
                                                                End If
                                                            Next
                                                        End If
                                                    End If
                                                End If
                                                If iv = 20 Or iii = 20 Then
                                                    Fusion_fichier_Data = L_tab_Champs_Emis(L_Index).t_champ_emis_data & " manquant! (" & L_tab_Champs_Emis(L_Index).t_champ_emis_data & ")"
                                                    GoTo Clean_Exit
                                                End If
                                            End If
                                        Next iii
                                                
                                    End If
                                End If
                            Else
                                If L_ShortFile(L_Nb_Fichier_Joint) <> "" Then
                                    If Not FileExists(DirSequestre & L_ShortFile(L_Nb_Fichier_Joint)) Then
                                        FileCopy DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_ShortFile(L_Nb_Fichier_Joint)
                                    End If
                                End If
                                        
                            End If
Rem Si le fichier joint est de type DOC, l'ouvrir et le convertir en pdf
ConvertTest:
                            If Not L_Service_Transformation_Mail And Not L_Archivage_Only Then
                                Select Case LCase(L_Fichier_Joint(L_Nb_Fichier_Joint).t_type)
                                Case "doc", "rtf", "ind", "docx"
                                
                                    If L_Service_ePoBox And Not L_ePoBox_Hybrid Then
                                        Rem On continue
                                    Else
                                        L_Result = Word_Save_To_PDFCreator(L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir, L_ID_PE & "pj" & L_Nb_Fichier_Joint)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        Else
                                            L_Fichier_Joint(L_Nb_Fichier_Joint).t_type = "pdf"
                                            L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier = L_ID_PE & "pj" & L_Nb_Fichier_Joint & ".pdf"
                                            L_Type_PDF = True
                                            L_Type_Fichier_joint = "pdf"
                                            If pFkFlow > 0 Then
                                                Rem On d�place le PDF dans le r�pertoire de fusion
                                                L_Fso.MoveFile L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & "\" & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier, DirSequestre & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = DirSequestre
                                            End If
                                        End If
                                    End If
                                Case "pdf", "txt"
                                    Rem On continue
                                Case "zip", "xls", "xlsx"
                                    If L_Service_ePoBox And Not L_ePoBox_Hybrid Then
                                        Rem On continue
                                    Else
                                        Fusion_fichier_Data = "Le fichier joint est inexploitable (" & LCase(L_Fichier_Joint(L_Nb_Fichier_Joint).t_type) & ")!!!" & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                        GoTo Clean_Exit
                                    End If
                                Case Else
                                    Fusion_fichier_Data = "Le fichier joint est inexploitable (" & LCase(L_Fichier_Joint(L_Nb_Fichier_Joint).t_type) & ")!!!" & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier
                                    GoTo Clean_Exit
                                End Select
                            End If 'Cas autres que e-mail
                        End If
                        L_Chaine_data = L_Chaine_data & """" & L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_tab_Champs_Emis(L_Index).t_champ_emis_data & """;"
                        If pUnzipDir = "no" Then
                            p_liste_fichiers_joints_production_locale = p_liste_fichiers_joints_production_locale & L_tab_Champs_Emis(L_Index).t_champ_emis_data & "|"
                        End If
                        p_Liste_Fichiers_Joints_Reception = p_Liste_Fichiers_Joints_Reception & L_tab_Champs_Emis(L_Index).t_champ_emis_data & "|"
                        If L_OriginalFile(L_Nb_Fichier_Joint) <> "" Then
                            p_Liste_Fichiers_Joints_Reception = p_Liste_Fichiers_Joints_Reception & L_OriginalFile(L_Nb_Fichier_Joint) & "|"
                        End If
                                
                    Case "Fichier r�pertori�", "Corps mail r�pertori�", "Fichier r�pertori� facultatif"
                        Rem Si la fusion du mod�le donne lieu � une pi�ce jointe,
                        Rem alors la notion de r�pertori�e ne subit pas d'insertion de fichier
                        
                        If L_Email_Piece_Jointe_Fusionnee Then
                            L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                        Else
                                    
                            If L_tab_Champs_Emis(L_Index).t_champ_emis_type = "Fichier r�pertori� facultatif" Then
                                Rem Dans le cas ou le champ est un fichier r�pertori� facultatif,
                                Rem Deux cas,
                                Rem 1. Le champ est renseign�, on traite comme si on est en pr�sence d'un champ r�pertori� classique
                                Rem 2. Le champ est vide, dans ce cas, on ignore
                                If Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data) <> "" Then
                                    Rem
                                    GoTo SuiteFichierR�pertori�Facultatif
                                End If
                            Else
SuiteFichierR�pertori�Facultatif:
                                L_Type_Fichier_joint = Type_Fichier(Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                                L_Chaine_data = L_Chaine_data & """" & Replace(G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_REPERTORIES & "\" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;", "\", "\\")
                                L_Nb_Fichier_Joint = L_Nb_Fichier_Joint + 1
                                ReDim Preserve L_Fichier_Joint(L_Nb_Fichier_Joint)
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_Signet = L_tab_Champs_Emis(L_Index).t_champ_emis_nom
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_REPERTORIES & "\"
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_type = L_Type_Fichier_joint
                                Rem Rajout d'un controle d'existence du fichier r�pertori�!!!
                                If Not FileExists(L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier) Then
                                    Rem v528
                                    Fusion_fichier_Data = "Le fichier r�pertori� " & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier & " est introuvable!"
                                    GoTo Clean_Exit
                                End If
                            End If
                        End If

                    Case "Fichier rapprochement facultatif"

                        Rem Dans le cas ou le champ est un fichier rapprochement facultatif
                        Rem Le champ est toujours renseign�
                        Rem On Cherche le fichier � rapprocher dans le r�pertoire du client / rapprochement
                        Rem Si le fichier n'est pas trouv�, pas grave
                        Rem Sinon, cela devient unfichier joint classique
                        If Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data) = "" Then
                            Rem On passe...
                        Else
                            If UCase(Right(Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data), 4)) <> ".PDF" Then
                                Rem On force le nommage du PDF
                                L_tab_Champs_Emis(L_Index).t_champ_emis_data = Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & ".pdf"
                            End If
                            
                            If Not FileExists(G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_RAPPROCHEMENT_FACULTATIF & "\" & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) Then
                                L_tab_Champs_Emis(L_Index).t_champ_emis_data = ""
                                L_Chaine_data = L_Chaine_data & """"";"
                            Else
                                L_Type_Fichier_joint = Type_Fichier(Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                                L_Chaine_data = L_Chaine_data & """" & Replace(G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_RAPPROCHEMENT_FACULTATIF & "\" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;", "\", "\\")
                                L_Nb_Fichier_Joint = L_Nb_Fichier_Joint + 1
                                ReDim Preserve L_Fichier_Joint(L_Nb_Fichier_Joint)
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_Signet = L_tab_Champs_Emis(L_Index).t_champ_emis_nom
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir = G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_RAPPROCHEMENT_FACULTATIF & "\"
                                L_Fichier_Joint(L_Nb_Fichier_Joint).t_type = L_Type_Fichier_joint
                                pListePdfRapproches = pListePdfRapproches & L_Fichier_Joint(L_Nb_Fichier_Joint).t_dir & L_Fichier_Joint(L_Nb_Fichier_Joint).t_Fichier & "|"
                                Rem ATTENTION, si la prestation est de type WORD sans Fichier joint, il convient donc forcer le mode Page de garde !!!
                                L_PageDeGarde = True

                            End If
                        End If
                                
                                
                    Case "Corps mail fusionn�"
CorpsMailFusionne:
                        Rem On utilise le mod�le HTML ou HTM ou TXT pour g�n�rer le corps du mail
                        L_Result = FusionCorpsMail(L_ID_PE, L_ModelName, DirSequestre, L_tab_Champs_Emis, "Corps")
                        If L_Result <> "Ok" Then
                            Fusion_fichier_Data = "Fichier corps de mail non fusionn� ! (" & L_Result & ")"
                            GoTo Clean_Exit
                        End If
                        L_ModelName = ""
                        
                    Case "Corps mail r�pertori� fusionn�"
                        L_ModelName = G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_REPERTORIES & "\" & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        GoTo CorpsMailFusionne
                        
                    Case "Corps mail joint fusionn�"
                        
                        If FileExists(decoupePDFsPath & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) Then
                            L_ModelName = decoupePDFsPath & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        Else
                            L_ModelName = DirSequestre & "\" & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        GoTo CorpsMailFusionne
                    
                    Case "Sujet mail r�pertori� fusionn�"
                        L_Result = FusionCorpsMail(L_ID_PE, G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_REPERTORIES & "\" & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data), DirSequestre, L_tab_Champs_Emis, "Sujet")
                        If L_Result <> "Ok" Then
                            Fusion_fichier_Data = "Fichier corps de mail non fusionn� ! (" & L_Result & ")"
                            GoTo Clean_Exit
                        End If

                    Case "Fond de page r�pertori�"

                        Rem Si la fusion du mod�le donne lieu � une pi�ce jointe,
                        Rem alors la notion de r�pertori�e ne subit pas d'insertion de fichier
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                        If Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data) = "" Then
                            Fusion_fichier_Data = "Le fond de page n'est pas renseign�!"
                            GoTo Clean_Exit
                        End If
                        L_Fond_de_Page = G_dir_client & "\" & pCustomerNumber & "\" & G_CONST_FICHIERS_REPERTORIES & "\" & Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        If Not FileExists(L_Fond_de_Page) Then
                            Fusion_fichier_Data = "Fichier fond de page r�pertori� introuvable ! (" & L_Fond_de_Page & ")"
                            GoTo Clean_Exit
                        End If

                    Case "Fichier � transmettre"
                        Rem Faire la liste des fichiers � copier sur le WEB
                        L_ePoBox_FileToXfer = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        Rem - Recherche dans les jours pr�c�dents
                        If pUnzipDir <> "no" Then
                            If FileExists(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                If get_file_size_only(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) = 0 Then
                                    GoTo Zip0:
                                End If
                                GoTo FaTSuite
                            Else
Zip0:
                                If FileExists(pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                    FileCopy pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                    Rem Patch bug zip � 0
                                    If get_file_size_only(pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data) <> get_file_size_only(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                        Fusion_fichier_Data = "Probl�me lors de la copie du fichier " & pUnzipDir & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                        GoTo Clean_Exit
                                    End If
                                End If
                            End If
                        End If
                        
                        If Not FileExists(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
Patch0Reception:
                            If FileExists(G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                FileCopy G_dir_reception & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                            Else
                                If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                    FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                Else
                                    For iii = 0 To 20
                                        If FileExists(G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iii, "YYYYMMDD") & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                                            FileCopy G_dir_reception & "\" & pCustomerNumber & "\" & Format(Date - iii, "YYYYMMDD") & "\" & L_tab_Champs_Emis(L_Index).t_champ_emis_data, DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data
                                            iii = 30
                                        End If
                                        If iii = 20 Then
                                            Fusion_fichier_Data = L_tab_Champs_Emis(L_Index).t_champ_emis_data & " manquant! (" & L_tab_Champs_Emis(L_Index).t_champ_emis_data & ")"
                                            GoTo Clean_Exit
                                        End If
                                    Next iii
                                    
                                End If
                            End If
                        Else
                            If get_file_size_only(DirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) = 0 Then
                                GoTo Patch0Reception
                            End If
                        End If
                                
FaTSuite:
                        p_Liste_Fichiers_Joints_Reception = p_Liste_Fichiers_Joints_Reception & L_tab_Champs_Emis(L_Index).t_champ_emis_data & "|"
                        p_liste_fichiers_joints_production_locale = p_liste_fichiers_joints_production_locale & L_tab_Champs_Emis(L_Index).t_champ_emis_data & "|"
                        L_Chaine_data = L_Chaine_data & """"";"

                    Case "N� Enregistrement", "Nombre documents"
                        If Not IsNumeric(L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                            L_Num_Record = 0
                        Else
                            L_Num_Record = CLng(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        L_Chaine_data = L_Chaine_data & """" & L_Num_Record & """;"
                            
                    Case "Nombre exemplaires"
                        If Not IsNumeric(L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                            NbExemplaireSup = 0
                        Else
                            NbExemplaireSup = CLng(L_tab_Champs_Emis(L_Index).t_champ_emis_data) - 1
                        End If
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "Soci�t�"
                        If UpdateAdresse Then
                            DestSociete = Upcase(Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                        Else
                            DestSociete = Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        L_Chaine_data = L_Chaine_data & """" & DestSociete & """;"
                    Case "Civilit�"
                        If UpdateAdresse Then
                            DestCivilite = Upcase(Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                        Else
                            DestCivilite = Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        L_Chaine_data = L_Chaine_data & """" & DestCivilite & """;"
                    Case "Nom"
                        If UpdateAdresse Then
                            DestNom = Upcase(Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                        Else
                            DestNom = Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        L_Chaine_data = L_Chaine_data & """" & DestNom & """;"
                    Case "Pr�nom"
                        If UpdateAdresse Then
                            DestPrenom = Upcase(Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
                        Else
                            DestPrenom = Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        End If
                        L_Chaine_data = L_Chaine_data & """" & DestPrenom & """;"

                    Case "R�f�rence ePoBox"
                        ePoBoxR = ""
                        L_Result = Analyse_ePobox(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data), L_Msg, ePoBoxR, L_ePoBox_Hybrid, p_Societe_Fk)
                        If L_Result <> "Ok" Then
                            'p_Start_Statut = "Erreur"
                        Else
                            Rem Lecture du service du Type pour Controle des habilitations
                            TypeePoBox = ePoBox__GetType(p_Prestation_Model_Pk)
                            If TypeePoBox = "none" Then
                                L_Msg = "Convention"
                            ElseIf TypeePoBox = "ar" Then
                                Rem Nothing to to... Let's continue...
                            ElseIf Left(TypeePoBox, 6) = "Erreur" Then
                                L_Msg = "Probl�me d'identification du type de flux ePoBox"
                            Else
                                Rem Controle des autorisations
                                If ePoBoxR = "" Then
                                    Rem Pas de convention...
                                    
                                Else
                                    L_Result = ePoBox__CheckCompanyPairing(pCustomerNumber, Left(ePoBoxR, 12), TypeePoBox)
                                    If L_Result <> "Ok" Then
                                        If Left(L_Result, 6) = "Erreur" Then
                                            L_Msg = L_Result
                                        Else
                                            L_Msg = "Convention """ & TypeePoBox & """ inexistante entre les clients (""" & pCustomerNumber & """ et """ & ePoBoxR & """)"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If ePoBoxR = "" Then
                            Rem Pas trouv�
                        Else
                            L_tab_Champs_Emis(L_Index).t_champ_emis_data = ePoBoxR
                            GoTo Case_ePoBox
                        End If
                                
                    Case "ePoBox"
Case_ePoBox:
                        Rem analyse des informations
                        L_Result = Analyse_ePoBox_Adresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data, _
                                                          L_ePoBox_Destinataire_PlateformId, _
                                                          L_ePoBox_Destinataire_ClientID, _
                                                          L_ePoBox_Destinataire_Prestation_Model_Nom, _
                                                          L_ePoBox_Destinataire_ContactId)
                        If L_Result <> "Ok" Then
                            L_Chaine_data = Analyse_ePobox(L_tab_Champs_Emis(L_Index).t_champ_emis_data, L_Result, L_Result)
                            Fusion_fichier_Data = L_Result
                            GoTo Clean_Exit
                        End If
                        L_ePoBox_Destinataire_Adresse = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"

                    Case "R�f�rence Client"
                        L_ePoBox_Destinataire_Reference = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                        
                    Case "e-mail Copie (CC)"
                        lMailCC = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                         
                    Case "e-mail Copie Cach�e (CCI)"
                        lMailCCI = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                        
                    Case "e-mail Reply-TO"
                        lMailReplyTO = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    
                    Case "Id Liste r�capitulative"
                        L_ePobox_Id_Liste_Recapitulative = Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & L_ePobox_Id_Liste_Recapitulative & """;"
                        
                    Rem LCI Impression Dynamique Supprim�e
                    'Case "Impression Dynamique", "Impression dynamique"
                    '    Rem Mettre en variable
                    '    L_tmp_Array = Split(L_tab_Champs_Emis(L_Index).t_champ_emis_data, "|", , vbTextCompare)
                    '    ReDim G_Print_Properties(UBound(L_tmp_Array))
                    '    Rem
                    '    For L_tmp_Index = 0 To UBound(L_tmp_Array) - 1
                    '        L_tmp_Array2 = Split(L_tmp_Array(L_tmp_Index), ";", , vbTextCompare)
                    '        G_Print_Properties(L_tmp_Index).t_Page_From = L_tmp_Array2(0)
                    '        G_Print_Properties(L_tmp_Index).t_Page_To = L_tmp_Array2(1)
                    '        G_Print_Properties(L_tmp_Index).t_Page_Type = L_tmp_Array2(2)
                    '        G_Print_Properties(L_tmp_Index).t_Page_Paper = L_tmp_Array2(3)
                    '    Next
                    '    L_Chaine_data = L_Chaine_data & """" & vbNullString & """;"
                        
                    Case "Adresse"
                        L_Pli_Adresse_Rue = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & L_Pli_Adresse_Rue & """;"
                        
                    Case "Code Postal"
                        L_Pli_Adresse_Cp = OptiCp(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                    Case "Ville"
                        L_Pli_Adresse_Ville = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_Pli_Adresse_Ville) & """;"
                    Case "Pays"
                        L_Pli_Adresse_Pays = Upcase(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                        If InStr(1, L_Pli_Adresse_Pays, vbNewLine, vbBinaryCompare) > 0 Then
                            If L_Pli_Adresse_Pays = vbNewLine Then
                                L_Pli_Adresse_Pays = ""
                            Else
                                L_Pli_Adresse_Pays = Replace(L_Pli_Adresse_Pays, vbNewLine, "")
                            End If
                        End If
                        If Trim(UCase(L_Pli_Adresse_Pays)) = "FRANCE" Or Trim(UCase(L_Pli_Adresse_Pays)) = "" Or UCase(L_Pli_Adresse_Pays) = "FR" Then
                            Rem Si c'est France on ne le met pas dans le pav� adresse
                            L_Pli_Adresse_Pays = ""
                            If Premium Then
                                whiteCB = True
                            End If
                        Else
                            If Premium Then
                                whiteCB = False
                            End If
                            
                            If UCase(L_Pli_Adresse_Pays) = "GUADELOUPE" _
                            Or Left(UCase(L_Pli_Adresse_Pays), 6) = "GUYANE" _
                            Or UCase(L_Pli_Adresse_Pays) = "MARTINIQUE" Then
                                If Premium Then
                                    whiteCB = True
                                End If
                            End If
                            If (InStr(1, L_Pli_Adresse_Pays, "polynesi", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "polyn�sie", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "reunion", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "r�union", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "mayotte", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "miquelon", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "guyane", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "r�union", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "caledonie", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "cal�donie", vbTextCompare) _
                            + InStr(1, L_Pli_Adresse_Pays, "futuna", vbTextCompare)) > 0 Then
                                If Premium Then
                                    whiteCB = True
                                End If
                            End If


                        End If
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_Pli_Adresse_Pays) & """;"
                    Case Else
                        L_Chaine_data = L_Chaine_data & """" & Valid_Champ_Fusion(L_tab_Champs_Emis(L_Index).t_champ_emis_data) & """;"
                        
                    End Select
                    
                Next
                        
                L_Chaine_header = L_Chaine_header & "REFERENCE_POSTEASY"
                
                L_Pli_Adresse = L_Pli_Adresse_Rue & vbNewLine & _
                                L_Pli_Adresse_Cp & " " & L_Pli_Adresse_Ville & vbNewLine & _
                                L_Pli_Adresse_Pays

                DestAdresseComplete = DestSociete
                If DestAdresseComplete <> "" Then
                    DestAdresseComplete = DestAdresseComplete & vbNewLine
                End If
                If Trim(DestCivilite) <> "" Then
                    If Trim(DestNom & DestPrenom) <> "" Then
                        DestAdresseComplete = DestAdresseComplete & DestCivilite & " " & Trim(Trim(DestNom) & " " & Trim(DestPrenom)) & vbNewLine
                    Else
                        DestAdresseComplete = DestAdresseComplete & DestCivilite & vbNewLine
                    End If
                Else
                    If Trim(DestNom) <> "" Or Trim(DestPrenom) <> "" Then
                        DestAdresseComplete = DestAdresseComplete & Trim(Trim(DestNom) & " " & Trim(DestPrenom)) & vbNewLine
                    End If
                End If
                DestAdresseComplete = DestAdresseComplete & L_Pli_Adresse
                
                DestAdresseComplete = OptiAdresse(DestAdresseComplete)
                    
                If L_Service_Transformation_Mail Then
                    L_Chaine_data = L_Chaine_data & """" & L_ID_PE & """"
                Else
                    L_Chaine_data = L_Chaine_data & """*" & L_ID_PE & "*"""
                End If
                
                If L_Sql_CreateUserInInsert <> "" Then
                    If L_Pli_Adresse_Rue <> "" Then
                        TabRue = Split(L_Pli_Adresse_Rue, vbNewLine, , vbBinaryCompare)
                        IndRueInsert = 0
                        For IndRue = 0 To UBound(TabRue)
                            If TabRue(IndRue) <> "" Then
                                IndRueInsert = IndRueInsert + 1
                                L_Sql_CreateUserInInsert = L_Sql_CreateUserInInsert & " contact_adresse" & IndRueInsert & ", "
                                L_Sql_CreateUserInSelect = L_Sql_CreateUserInSelect & "'" & Valid_Text(TabRue(IndRue)) & "', "
                            End If
                        Next
                    End If
                    L_Nb_pages_dans_pli = 0
                    L_PageCount = L_Nb_pages_dans_pli
                    L_SheetCount = L_Nb_pages_dans_pli
                    GoTo Archivage_Only_Suite2
                End If
                        
                If L_Archivage_Only Then
                    GoTo Archivage_Only_Suite
                End If
                
                Select Case L_ModelType
                Case "Aucun"
                    If L_Service_Transformation_fax Or L_Service_Transformation_Mail Or L_Service_Transformation_Sms Then
                        GoTo CBDirect_Suite1
                    Else
                        GoTo FusionSuite1
                    End If
                Rem LCI le 14/08/2017 Suppression de Masque PDF Case "CB Direct", "Masque Pdf", "Mod�le Texte", "Mod�le Html", "CB + Adresse"
                Case "CB Direct", "Mod�le Texte", "Mod�le Html", "CB + Adresse"
                    GoTo CBDirect_Suite1
                End Select

            Rem DEBUT DE LA FUSION WORD
                Rem creer fichier
                Rem ouvrir fichier
                L_Num_data_source = FreeFile
                Open L_Chemin_nom_datasource_doc For Output As #L_Num_data_source
                Rem ecrire dans fichier
                Print #L_Num_data_source, L_Chaine_header
                Print #L_Num_data_source, L_Chaine_data
                Rem fermer fichier
                Close #L_Num_data_source
                Rem Rappel: Cas normal de Fusion
                Rem Page de Garde
CBDirect_Suite1:
                If L_PageDeGarde Then
                    
                    Rem v500 le cas page de garde de type calque NE SE retrouve PAS ICI
                    Rem Si <> pdf on fait comme d'habitude
                    Select Case L_Type_Fichier_joint
                    Case "pdf"
                        'si le fichier joint est de type pdf, on effectue la fusion
                        'on n'enregistrement pas le r�sultat en pcl,
                        'on transforme le doc en pdf, puis on concat�ne le pdf d'origine avec ceux joints
                        If L_ModelName <> "" Then
                            G_Wd.Documents(L_ModelName).MailMerge.OpenDataSource _
                                            Name:=L_Chemin_nom_datasource_doc, _
                                            linktosource:=False
                            Rem fusionner la page
                            G_Wd.Documents(L_ModelName).MailMerge.Execute
                            G_Wd.Documents(L_ModelName).Application.Options.UpdateFieldsAtPrint = True
                            G_Wd.Documents(L_ModelName).Application.Options.PrintFieldCodes = False
                        End If
                    
                    Case Else
                        If LCase(L_Type_Fichier_joint) = "xls" Or LCase(L_Type_Fichier_joint) = "xlsx" Then
                            Fusion_fichier_Data = "Type de fichier joint non support�! (" & L_Type_Fichier_joint & ")"
                            GoTo Clean_Exit
                        End If
                        
                        If Trim(L_Type_Fichier_joint) = "" Then
                            Fusion_fichier_Data = "Aucun fichier joint d�fini"
                            GoTo Clean_Exit
                        End If
                    End Select
                            
                    Select Case L_ModelType
                    Rem LCI le 14/08/2017 Suppression de Masque PDF     Case "Aucun", "CB Direct", "Masque Pdf"
                    Case "Aucun", "CB Direct"
                        Fusion_fichier_Data = "Erreur Param�trage de la prestation : mod�le incompatible avec la page de garde!"
                        GoTo Clean_Exit
                    Case "Mod�le Word"
                        Rem Comme avant
                    Case "CB + Adresse"
                        Rem Cr�er la page blanche + ajouter le CB + adresse
                        Rem Copie de la page blanche
                        FileCopy BlankPage, DirSequestre & L_ID_PE & "_blank.pdf"
                        If L_Type_Impression <> "Recto" Then
                            L_New_pdf = L_ID_PE & "pdgR.pdf"
                        Else
                            L_New_pdf = L_ID_PE & "pdg.pdf"
                        End If
                        Rem ajout du CB + Adresse
                        L_Result = AddTmBarcodeAdresse(DirSequestre & L_ID_PE & "_blank.pdf", DirSequestre & L_New_pdf, L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat, DestAdresseComplete, AdresseFondTab, False, whiteCB)
                        If L_Result <> "Ok" Then
                            Fusion_fichier_Data = L_Result
                            GoTo Clean_Exit
                        End If
                        If L_Type_Impression <> "Recto" Then
                            L_Result = Concat_PdfLib(p_Form, DirSequestre & L_ID_PE & "pdgR.pdf" & "|" & BlankPage, DirSequestre & L_ID_PE & "pdg.pdf")
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                        End If
                        L_New_pdf = L_ID_PE & ".pdf"
                        GoTo SuitePdgCbAdresse
                    End Select
                        
                    Rem Page de garde de type Word !!
                    Rem Sauvegarde du document g�n�r�
                    L_New_doc = L_ID_PE & ".doc"
                    Rem Cr�ation Document Word
                    Call Upgrade_Compteurs(L_Compteur, L_Compteur_precedent)

                    Rem Save Document
                    Call Display_Status("Etat d'avancement de la fusion : (Enregistrement du DOC (local))", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                    G_Wd.Documents("Lettres types" & L_Compteur).SaveAs DirSequestre & L_New_doc
                    
                    L_New_pdf = Replace(L_New_doc, ".doc", ".pdf", , , vbTextCompare)
On Error Resume Next
                    Rem Dans page de garde (sauf calque)
                    If Not L_Service_Transformation_Mail Then
                        Select Case L_Type_Fichier_joint
                        Case "pdf"
                        
                            If Left(G_Wd.Documents(L_New_doc).Application.ActivePrinter, Len(G_Driver_PDFCreator)) <> G_Driver_PDFCreator Then
                                Call Display_Status("Etat d'avancement de la fusion : (Affectation du driver)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                G_Wd.Documents(L_New_doc).Application.ActivePrinter = G_Driver_PDFCreator
                            End If
                            Call Display_Status("Etat d'avancement de la fusion : (G�n�ration du PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)

                            L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE & "pdg")
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                            Call WaitingPdfCreator
                            Rem lire le nombre de page du/des pdf li�s
                            Call Display_Status("Etat d'avancement de la fusion : (Fermeture du Doc)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                            Rem New concat�ner le pdf page de garde avec le pdf joint => dans le pdf de garde
                            
                            If get_file_size_only(DirSequestre & L_ID_PE & "pdg.pdf") < 3750 Then
                                Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                GoTo Clean_Exit
                            End If
                                    
SuitePdgCbAdresse:
                            Rem Tourner dans les pi�ces jointes!!!
                            L_list_Pdf_Files = vbNullString
                            For i = 1 To L_Nb_Fichier_Joint
                                Rem Ajouter la gestion du fond de page, sur les pi�ces jointes, sinon cela se fait plus loin!!!
                                If L_Fond_de_Page <> "" And L_FondDePageSurPJOnly Then
                                    If Not L_Fond_de_Page_P2P Then
                                        L_Result = Merge_Pdf_NewAuto(p_Form, L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier, L_Fond_de_Page, False, pCustomerNumber, L_Fichier_Joint(i).t_dir & Replace(L_Fichier_Joint(i).t_Fichier, ".pdf", "_pj" & i & ".pdf", , , vbTextCompare))
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                        L_Fichier_Joint(i).t_Fichier = Replace(L_Fichier_Joint(i).t_Fichier, ".pdf", "_pj" & i & ".pdf", , , vbTextCompare)
                                    Else
                                        L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = "Cas PJ P2P non test�"
                                            GoTo Clean_Exit
                                        End If
                                    End If
                                    
                                End If
                                L_list_Pdf_Files = L_list_Pdf_Files & "|" & L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier & "|"
                                If NbExemplaireSup > 0 Then
                                    Rem si recto/verso
                                    If L_Type_Impression <> "Recto" Then
                                        Rem Lire le nombre de page
                                        L_Result = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_ID_PE & "pdg.pdf")
                                        If Not IsNumeric(L_Result) Then
                                            Fusion_fichier_Data = "Probl�me lecture Pages n01"
                                            GoTo Clean_Exit
                                        End If
                                        NbPagesIntermediaire = CLng(L_Result)
                                        L_Result = Read_Pdf_Num_Pages(p_Form, L_Fichier_Joint(i).t_dir, L_Fichier_Joint(i).t_Fichier)
                                        If Not IsNumeric(L_Result) Then
                                            Fusion_fichier_Data = "Probl�me lecture Pages n02"
                                            GoTo Clean_Exit
                                        End If
                                        NbPagesIntermediaire = NbPagesIntermediaire + CLng(L_Result)
                                        If NbPagesIntermediaire Mod 2 = 1 Then
                                            Rem Si on est en pr�sence d'un nombre impair de page pour la racine, on ajoute une page blanche
                                            SupBlankPage = BlankPage & "|"
                                        End If
                                    End If
                                    For IndiceExemplaires = 1 To NbExemplaireSup
                                        L_list_Pdf_Files = L_list_Pdf_Files & SupBlankPage & L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier & "|"
                                    Next
                                End If
                            Next
                                    
                            Call Display_Status("Etat d'avancement de la fusion : (Concat�nation PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            
                            If DatamatriX Then
                                L_New_pdf = L_ID_PE & "_woDatamatrix.pdf"
                                Rem test OK
                            End If
                            L_Result = Concat_PdfLib(p_Form, DirSequestre & L_ID_PE & "pdg.pdf" & L_list_Pdf_Files, DirSequestre & L_New_pdf)
                            If L_Result <> "Ok" Then
                                L_Result = Concat_Pdf(p_Form, DirSequestre & L_ID_PE & "pdg.pdf" & L_list_Pdf_Files, DirSequestre & L_New_pdf)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = L_Result & " (" & L_list_Pdf_Files & ")"
                                    GoTo Clean_Exit
                                End If
                            End If
                            
                            Rem Dans page de garde (sauf calque)
                            
                            Rem Si la fusion du mod�le n�cessite une "fusion de fond de page"
                            If L_Fond_de_Page <> vbNullString And Not L_FondDePageSurPJOnly Then
                                Call Display_Status("Etat d'avancement de la fusion : (Fusion fond de page)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                If Not L_Fond_de_Page_P2P Then
                                    L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                Else
                                    L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                End If
                            End If
                            
                            Rem lire le nombre de page du nouveau pli
                            L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                            If L_Nb_pages_dans_pli > 4900 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(1)"
                                GoTo Clean_Exit
                            End If
                            L_PageCount = L_Nb_pages_dans_pli
                            Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                            L_SheetCount = L_Nb_pages_dans_pli
                            If DatamatriX Then
                                L_Result = Insert_Datamatrix6(MW6DataMatrixFusion, 0, 0, DirSequestre & "\" & L_New_pdf, DirSequestre & "\" & Replace(L_New_pdf, "_woDatamatrix", "", 1, , vbTextCompare), "", L_SheetCount, "", DtmxVide, False, False, True)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Probl�me lors de l'ajout du Datamatrix!!! (001)"
                                    GoTo Clean_Exit
                                End If
                                L_New_pdf = Replace(L_New_pdf, "_woDatamatrix", "", , , vbTextCompare)
                            End If
                            Rem rechercher si certains des �l�ments suivants peuvent �tre r�cup�r�s
                            'L_Document_Size = get_file_size_only(DirSequestre & pCustomerNumber & "\" & L_New_pdf)
                            L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                        
                        Case Else   'Fichier joint non PDF
                            
                            Rem Cas Page de garde avec pi�ce jointe non pdf
                            Rem Dans page de garde (sauf calque)
                            If InStr(1, G_Wd.Documents(L_New_doc).Application.ActivePrinter, G_Driver_PDFCreator, vbTextCompare) = 0 Then
                                Call Display_Status("Etat d'avancement de la fusion : (Affectation du driver)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                G_Wd.Documents(L_New_doc).Application.ActivePrinter = G_Driver_PDFCreator
                            End If
                            Call Display_Status("Etat d'avancement de la fusion : (G�n�ration du PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            G_Wd.Documents(L_New_doc).SaveAs DirSequestre & L_New_doc
                            L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = "Impossible d'initialiser le Driver PDFCreator."
                                GoTo Clean_Exit
                            End If
                            G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                            Call WaitingPdfCreator
                            
                            Rem FERMER le DOC WORD!!!!
                            Rem 'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                            L_Result = get_file_size_only(DirSequestre & L_ID_PE & ".pdf")
                            If L_Result = "0" Then
                                Fusion_fichier_Data = "PDF non g�n�r�!!!"
                                GoTo Clean_Exit
                            End If
                            If L_Result < 5000 Then
                                Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                GoTo Clean_Exit
                            End If

                            L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                            If L_Nb_pages_dans_pli > 4900 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(2)"
                                GoTo Clean_Exit
                            End If
                            
                            L_PageCount = L_Nb_pages_dans_pli
                            Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                            L_SheetCount = L_Nb_pages_dans_pli
                            If DatamatriX Then
                                Rem Unique cas ou on recopie le PDF
                                FileCopy DirSequestre & L_New_pdf, DirSequestre & Replace(L_New_pdf, ".pdf", "_woDatamatrix.pdf", 1, , vbTextCompare)
                                Kill DirSequestre & L_New_pdf
                                L_Result = Insert_Datamatrix6(MW6DataMatrixFusion, 0, 0, DirSequestre & Replace(L_New_pdf, ".pdf", "_woDatamatrix.pdf", 1, , vbTextCompare), DirSequestre & "\" & L_New_pdf, "", L_SheetCount, "", DtmxVide, False, False, True)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Probl�me lors de l'ajout du Datamatrix!!! (002)"
                                    GoTo Clean_Exit
                                End If
                            End If

                            Rem rechercher si certains des �l�ments suivants peuvent �tre r�cup�r�s
                            L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                        
                        End Select
                    End If  'Si ce n'est pas du mail...
                            

                    If (L_Service_Transformation_Mail Or L_Service_Transformation_fax) Then
                        Rem Dans page de garde (sauf calque)
                        If L_Service_Transformation_Mail Then
                            L_Result = Word_Save_To_Html(L_New_doc, DirSequestre, L_ID_PE)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            G_Wd.Documents(Replace(L_New_doc, ".doc", "_Mail_Body.html", , , vbTextCompare)).Close savechanges:=wdDoNotSaveChanges
                            L_Result = Replace_Image_Number_By_Image_List(DirSequestre, Replace(L_New_doc, ".doc", "_Mail_Body.html", , , vbTextCompare), L_Url_Dir)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            If L_Email_ReplyTo <> "" And lMailReplyTO = "" Then
                                lMailReplyTO = L_Email_ReplyTo
                            End If
                            L_Result = Mail_Creation(L_ID_PE, _
                                                     L_Email_to, _
                                                     L_Email_From, _
                                                     L_Email_xFer, _
                                                     L_Fk_destinataire, _
                                                     pCustomerNumber, _
                                                     L_Email_Sql_Pli_Update, _
                                                     L_chemin, _
                                                     L_Email_Piece_Jointe_Fusionnee, _
                                                     L_SujetMailFusionne, _
                                                     L_List_Fichiers_Mail_To_Move, _
                                                     p_liste_fichiers_joints_production_locale, _
                                                     L_Liste_Pieces_Jointes, _
                                                     L_Document_Size, False, False, "", _
                                                     DirSequestre, pUnzipDir, _
                                                     lMailCC, lMailCCI, lMailReplyTO, _
                                                     decoupePDFsPath)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            G_Wd.Documents(Replace(L_New_doc, ".doc", "_Mail_Body.html", , , vbTextCompare)).Close savechanges:=wdDoNotSaveChanges
                        End If
                        Rem Dans page de garde (sauf calque)
                    Else
                        Rem D�j� fait pour le pdf
                        Select Case L_Type_Fichier_joint
                        Case "pdf"
                            Rem Nothing
                        Case Else
                            G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                        End Select
                    End If
                    
                Else
                            
Rem ******************************************************************************************************
Rem FUSION NORMALE
Rem                         Pas de page de garde
Rem
Rem ************************************************8*****************************************************
                    Rem v500 le cas page de garde de type calque se retrouve ICI
                    If Left(L_Type_Fichier_joint, 3) <> "htm" Then
                        Rem Sans Fichier joint
                        Call Display_Status("Etat d'avancement de la fusion : (Fusion)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                        If L_Service_Transformation_Sms Then
                            GoTo Suite_Mail_Simple
                        End If
                        If L_ModelName = vbNullString And L_Service_Transformation_Mail Then
                            GoTo Suite_Mail_Simple
                        End If
                        If DatamatriX Then
                            L_New_pdf = L_ID_PE & "_woDatamatrix.pdf"
                        Else
                            L_New_pdf = L_ID_PE & ".pdf"
                        End If
                        
                        Rem G�n�ration du PDF principal
                        If L_NoWord And L_Modele_Merge Then 'rem dans ce cas, il y a au moins un PDF joint !!!!
                            Rem On ajoute le code � barre directement sur la premi�re page du premier PDF joint
                            Select Case L_ModelType
                            Case "Aucun"
                                L_Result = "Ok"
                            
                            Rem LCI le 14/08/2017 - Suppression
                            'Case "Masque Pdf"
                            '
                            '    L_Result = createPdfCBPage(DirSequestre & L_ID_PE & "_CB.PDF", L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat)
                            '    If L_Result <> "Ok" Then
                            '        MsgBox "Fusion_Fichier_Data_Err1"
                            '        End
                            '    End If
                            '    FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_ID_PE & "_DOC.pdf"
                            '
                            '    L_Result = Pdf_MergeRev(p_Form, DirSequestre & L_ID_PE & "_CB.PDF", L_ModelName, DirSequestre & L_ID_PE & "_RESULT.PDF")
                            '    If L_Result <> "Ok" Then
                            '        Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err2)"
                            '        GoTo Clean_Exit
                            '    End If
                            '
                            '    PatchPonceletNbPages = Read_Pdf_Num_Pages(p_Form, "", DirSequestre & L_ID_PE & "_DOC.pdf")
                            '    If PatchPonceletNbPages = 1 Then
                            '        L_Result = Pdf_MergeRev(p_Form, DirSequestre & L_ID_PE & "_RESULT.PDF", DirSequestre & L_ID_PE & "_DOC.pdf", DirSequestre & L_ID_PE & "_RESULT2.PDF")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err3)"
                            '            GoTo Clean_Exit
                            '        End If
                            '    Else
                            '        Rem Extraire la page 1 pour y ajouter le CB
                            '        L_Result = Pdf_ExtractPages(1, 1, DirSequestre & L_ID_PE & "_DOC.pdf", DirSequestre & L_ID_PE & "_DOC1.pdf")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err4)"
                            '            GoTo Clean_Exit
                            '        End If
                            '       Rem Extraire les pages 2 et +
                            '        L_Result = Pdf_ExtractPages(2, PatchPonceletNbPages, DirSequestre & L_ID_PE & "_DOC.pdf", DirSequestre & L_ID_PE & "_DOCS.pdf")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err5)"
                            '            GoTo Clean_Exit
                            '        End If
                            '        Rem Fusion Invers�e de la page1 avec le masque et le CB
                            '        L_Result = Pdf_MergeRev(p_Form, DirSequestre & L_ID_PE & "_RESULT.PDF", DirSequestre & L_ID_PE & "_DOC1.pdf", DirSequestre & L_ID_PE & "_RESULT_PAGE1.PDF")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err6)"
                            '            GoTo Clean_Exit
                            '        End If
                            '
                            '        Rem Fusion Invers�e des pages suivantes avec le masque sans le CB
                            '        L_Result = Pdf_MergeRev(p_Form, L_ModelName, DirSequestre & L_ID_PE & "_DOCS.pdf", DirSequestre & L_ID_PE & "_RESULT_PAGES.PDF")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err7)"
                            '            GoTo Clean_Exit
                            '        End If
                            '
                            '        Rem Concat�nation de la premi�re page et des suivantes
                            '        L_Result = Concat_PdfLib(p_Form, DirSequestre & L_ID_PE & "_RESULT_PAGE1.PDF|" & DirSequestre & L_ID_PE & "_RESULT_PAGES.PDF", DirSequestre & L_ID_PE & "_RESULT2.PDF")
                            '        If L_Result <> "Ok" Then
                            '            Fusion_fichier_Data = L_Result & " (Fusion_Fichier_Data_Err8)"
                            '            GoTo Clean_Exit
                            '        End If
                            '    End If
                            '    FileCopy DirSequestre & L_ID_PE & "_RESULT2.PDF", DirSequestre & L_New_pdf

                            Case "CB + Adresse"
                                L_Result = AddTmBarcodeAdresse(L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf, L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat, DestAdresseComplete, AdresseFondTab, False, whiteCB)

                            Case Else
                                    If Premium Then
                                        If whiteCB Then
                                            Rem Inutile de faire le traitement pour rien !!!
                                            L_Result = "Ok"
                                        Else
                                            L_Result = AddTmBarcode(L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf, L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat)
                                        End If
                                    Else
                                        If decoupePDFsPath <> "" Then
                                            L_Result = AddTmBarcode(decoupePDFsPath & "\" & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf, L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat)
                                        Else
                                            L_Result = AddTmBarcode(L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf, L_ID_PE, L_bcTab(0), L_bcTab(1), L_bcTab(2), L_bcCoverFormat)
                                        End If
                                    End If
'                                        End If
                            End Select
                                
                            If L_Result <> "Ok" Then
                                'Impossible d'ajouter le code � barre dans ce pdf !!!!
                                Rem Test de conversion
                                Fusion_fichier_Data = "Impossible d'ajouter le code � barre dans le pdf !!! (" & L_Result & ") (" & L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier & ")"

                                'If L_Fichier_Joint(1).t_dir = pUnzipDir Then
                                    'Kill L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier
                                'End If

                                'If InStr(1, L_Fichier_Joint(1).t_dir, p_NumPrepa, vbTextCompare) > 0 Then
                                    'Kill L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier
                                'End If
                                GoTo Clean_Exit
                            End If
                        Else 'Comme avant, on fusionne le WORD
                            If L_ModelType = "Mod�le Word" Then
                            Rem Cas ou le document fusionn�
                                G_Wd.Documents(L_ModelName).MailMerge.OpenDataSource _
                                                  Name:=L_Chemin_nom_datasource_doc, _
                                                  linktosource:=False
                                Rem fusionner la page
                                G_Wd.Documents(L_ModelName).Application.Options.UpdateFieldsAtPrint = True
                                G_Wd.Documents(L_ModelName).Application.Options.PrintFieldCodes = False
                                G_Wd.Documents(L_ModelName).MailMerge.Execute
                                Rem Sauvegarde du document g�n�r�
                                L_New_doc = L_ID_PE & ".doc"
                                Rem Cr�ation Document Word
                                Call Upgrade_Compteurs(L_Compteur, L_Compteur_precedent)
                                G_Wd.Documents("Lettres types" & L_Compteur).SaveAs DirSequestre & L_New_doc, , , , False, , , False, True
                            End If
                        End If
                        
                        For i = 1 To L_Nb_Fichier_Joint
                            Call Display_Status("Etat d'avancement de la fusion : (Merge Fichier(s) joint(s))", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            
                            If L_Service_Transformation_Mail Then
                                Select Case L_ModelType
                                Case "Mod�le Word"
                                    If L_Type_Fichier_joint = "pdf" And L_ModelName <> "" Then
                                        GoTo Suite_Mail_Simple
                                    End If
                                Case "Mod�le Texte", "Mod�le Html"
                                    GoTo Suite_Mail_CorpsDYNamique
                                End Select
                                If L_CorpsMailFusionne Then
                                    GoTo Suite_Mail_CorpsDYNamique
                                End If
                            End If
                            
                            If L_Fichier_Joint(i).t_type = "pdf" Then
                                Rem Il y a au moins un PDF joint
                                If L_Fichier_Joint(i - 1).t_type <> "pdf" And Not L_NoWord Then
                                    G_Wd.Documents(L_New_doc).SaveAs DirSequestre & L_New_doc
                                    If InStr(1, G_Wd.Documents(L_New_doc).Application.ActivePrinter, G_Driver_PDFCreator, vbTextCompare) = 0 Then
                                        G_Wd.Documents(L_New_doc).Application.ActivePrinter = G_Driver_PDFCreator
                                    End If
                                    Rem lire le nombre de page du/des pdf li�s
                                    L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = "Impossible d'initialiser PdfCreator " & L_Result
                                        GoTo Clean_Exit
                                    End If
                                    G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                                    Call WaitingPdfCreator
                                    G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                                    L_Pdf_Generated = True
                                    'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                                    If get_file_size_only(DirSequestre & L_ID_PE & ".pdf") < 3500 Then
                                        Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                        GoTo Clean_Exit
                                    End If
                                
                                End If

                                If L_Modele_Merge And i = 1 Then
                                    
                                    Select Case L_ModelType
                                    Case "CB Direct", "CB + Adresse"
                                    Rem LCI le 14/08/2017 'Case "CB Direct", "Masque Pdf", "CB + Adresse"
                                        Rem Dans ce cas, le premier PDF est d�j� cr��, passer au second!!!
                                        Rem on continue donc en bypassant le cas suivant
                                        L_Pdf_Generated = True

                                        If L_Nb_Fichier_Joint = 1 Then
                                            If L_Fond_de_Page <> vbNullString Then
                                                If Not L_Fond_de_Page_P2P Then
                                                    If Read_Pdf_Num_Pages(p_Form, "", L_Fond_de_Page) = 1 Then
                                                        L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                                    Else
                                                        L_Result = Merge_Pdf_NewAuto(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber, DirSequestre & Replace(L_New_pdf, ".pdf", "_pj.pdf", , , vbTextCompare))
                                                        If L_Result = "Ok" Then
                                                            FileCopy DirSequestre & L_New_pdf, DirSequestre & Replace(L_New_pdf, ".pdf", "_bu.pdf", , , vbTextCompare)
                                                            FileCopy DirSequestre & Replace(L_New_pdf, ".pdf", "_pj.pdf", , , vbTextCompare), DirSequestre & L_New_pdf
                                                            Kill DirSequestre & Replace(L_New_pdf, ".pdf", "_pj.pdf", , , vbTextCompare)
                                                            Kill DirSequestre & Replace(L_New_pdf, ".pdf", "_bu.pdf", , , vbTextCompare)
                                                        End If
                                                    End If
                                                    If L_Result <> "Ok" Then
                                                        Fusion_fichier_Data = L_Result
                                                        GoTo Clean_Exit
                                                    End If
                                                Else
                                                    L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                                    L_Result = Merge_Pdf_NewAuto(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                                    If L_Result <> "Ok" Then
                                                        Fusion_fichier_Data = L_Result
                                                        GoTo Clean_Exit
                                                    End If
                                                End If
                                            End If
                                        End If
                                            
                                    Case "Aucun"
                                        FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf
                                        L_Pdf_Generated = True
                                    
                                    Case Else
                                        If Not L_Fond_de_Page_P2P Then
                                            If L_ShortFile(i) <> "" Then

                                                If FileExists(DirSequestre & L_ShortFile(i)) Then
                                                    L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, DirSequestre & L_ShortFile(i), True, pCustomerNumber)
                                                Else
                                                    If Len(L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier) > 253 Then
                                                        If FileExists(G_dir_prereception & "\" & Format(Date, "YYYYMMDD") & "\" & L_Fichier_Joint(i).t_Fichier) Then
                                                            L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, G_dir_prereception & "\" & Format(Date, "YYYYMMDD") & "\" & L_Fichier_Joint(i).t_Fichier, True, pCustomerNumber)
                                                        Else
                                                            L_Result = "Fichier trop long ou introuvable!! (" & L_Fichier_Joint(i).t_Fichier & ")"
                                                        End If
                                                    Else
                                                        L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier, True, pCustomerNumber)
                                                    End If
                                                End If
                                            Else
                                                If Len(L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier) > 253 Then
                                                    If FileExists(G_dir_prereception & "\" & Format(Date, "YYYYMMDD") & "\" & L_Fichier_Joint(i).t_Fichier) Then
                                                        L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, G_dir_prereception & "\" & Format(Date, "YYYYMMDD") & "\" & L_Fichier_Joint(i).t_Fichier, True, pCustomerNumber)
                                                    Else
                                                        L_Result = "Fichier trop long ou introuvable!! (" & L_Fichier_Joint(i).t_Fichier & ")"
                                                    End If
                                                Else
                                                    Rem D�j� contr�l� plus haut !!!!
                                                    L_New_pdf = Replace(L_New_doc, ".doc", IIf(DatamatriX, "_woDatamatrix", "") & ".pdf", , , vbTextCompare)
                                                    If Not FileExists(DirSequestre & L_New_pdf) Then
                                                        If FileExists(Replace(DirSequestre & L_New_doc, ".doc", ".pdf", , , vbTextCompare)) Then
                                                            FileCopy Replace(DirSequestre & L_New_doc, ".doc", ".pdf", , , vbTextCompare), DirSequestre & L_New_pdf
                                                            Kill Replace(DirSequestre & L_New_doc, ".doc", ".pdf", , , vbTextCompare)
                                                        End If
                                                    End If
                                                    L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier, True, pCustomerNumber)
                                                End If
                                            End If
                                            If L_Result <> "Ok" Then
                                                Fusion_fichier_Data = L_Result
                                                GoTo Clean_Exit
                                            End If
                                        Else
                                            Rem Plus utilis� au 30/09/2010 => a tester � nouveau si besoin
                                            L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier, DirSequestre & L_New_pdf)
                                            If L_Result <> "Ok" Then
                                                Fusion_fichier_Data = L_Result
                                                GoTo Clean_Exit
                                            End If
                                        End If
                                        If L_Fond_de_Page <> vbNullString Then
                                            If Not L_Fond_de_Page_P2P Then
                                                L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                                If L_Result <> "Ok" Then
                                                    Fusion_fichier_Data = L_Result
                                                    GoTo Clean_Exit
                                                End If
                                            Else
                                                L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                                If L_Result <> "Ok" Then
                                                    Fusion_fichier_Data = L_Result
                                                    GoTo Clean_Exit
                                                End If
                                            End If
                                        End If
                                      
                                    End Select
                                            
                                Else
                                        
                                    L_list_Pdf_Files = vbNullString
                                    L_list_Pdf_Files = L_list_Pdf_Files & "|" & L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier & "|"
                                    FileCopy DirSequestre & L_New_pdf, DirSequestre & Replace(L_New_pdf, ".pdf", "_main" & i & ".pdf", , , vbTextCompare)
                                    If NbExemplaireSup > 0 Then
                                        Rem On ajoute autant de fois le(s) document(s) qu'il faut d'exemplaire(s)
                                        For ii = 1 To NbExemplaireSup
                                            L_list_Pdf_Files = L_list_Pdf_Files & "|" & L_list_Pdf_Files
                                        Next
                                        L_list_Pdf_Files = Replace(L_list_Pdf_Files, "|||", "|", , , vbTextCompare)
                                    End If
                                            
                                    L_Result = Concat_PdfLib(p_Form, DirSequestre & Replace(L_New_pdf, ".pdf", "_main" & i & ".pdf", , , vbTextCompare) & L_list_Pdf_Files, DirSequestre & L_New_pdf)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                    
                                End If
                                
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = L_Result
                                    GoTo Clean_Exit
                                End If
                                
                            Else
                                        
                                   
                                Rem Ici, le document joint n'est pas en PDF !!!!!
                                Rem A mon avis ne fonction plus
                                Rem envoi d'un mail pour confirmation
                                G_Wd.Documents(L_New_doc).Application.Selection.GoTo What:=wdGoToPage, Which:=wdGoToLast
                            '' Inserer un signet de travail � la hauteur
                            ''  du signet Fichier joint
                                
                                G_Wd.Documents(L_New_doc).Bookmarks.Add Range:=G_Wd.Documents(L_New_doc).Application.Selection.Range, Name:="deb"
                                G_Wd.Documents(L_New_doc).Bookmarks.DefaultSorting = wdSortByName
                                G_Wd.Documents(L_New_doc).Bookmarks.ShowHidden = False
                                ' Insere le fichier joint

                                G_Wd.Documents(L_New_doc).Application.ChangeFileOpenDirectory L_Fichier_Joint(i).t_dir
                                G_Wd.Documents(L_New_doc).Application.Selection.InsertFile FileName:=L_Fichier_Joint(i).t_Fichier, Range:="", _
                                    ConfirmConversions:=False, Link:=False, Attachment:=False
                                If InStr(1, UCase(L_Fichier_Joint(i).t_Fichier), ".TXT") > 0 Then
                                            
                                    ''==============================================
                                    '' Si l'insert est un fichier TXT
                                    ''==============================================
                                    '' Inserer un signet en fin d'insert
                                    G_Wd.Documents(L_New_doc).Bookmarks.Add Range:=G_Wd.Documents(L_New_doc).Application.Selection.Range, Name:="fin"
                                    G_Wd.Documents(L_New_doc).Bookmarks.DefaultSorting = wdSortByName
                                    G_Wd.Documents(L_New_doc).Bookmarks.ShowHidden = False
                                    ' Revenir en d�but d'insert
                                    G_Wd.Documents(L_New_doc).Application.Selection.GoTo What:=wdGoToBookmark, Name:="deb"
                                    'Marquer la zone entre les signets (Texte inser�)
                                    G_Wd.Documents(L_New_doc).Application.Selection.MoveEnd (wdStory)
                                    G_Wd.Documents(L_New_doc).Application.Selection.font.Name = "Courier New"
                                    G_Wd.Documents(L_New_doc).Application.Selection.font.Name = "Courier New"
                                    G_Wd.Documents(L_New_doc).Application.Selection.font.Size = 9
                                    G_Wd.Documents(L_New_doc).Application.Selection.font.Bold = wdToggle
                                    G_Wd.Documents(L_New_doc).Application.Selection.font.Bold = wdToggle
                                End If
                                
                                Rem Ne pas supprimer les fichiers r�pertori�s!!!
                                G_Wd.Documents(L_New_doc).Save
                                    
                            End If

                        Next i
                                
                    End If  'Rem Rappel Cas Normal
                            'Rem Sans page de garde
                            'Rem ET Sans type PCL
                            'rem ET <> htm
                                    
Suite_Mail_Simple:
                    Rem SUPPRIME LE 2017-03-14 ************************************************************************
                    'If L_Nb_champs_lies > 0 Then
                    '    Rem Rappel, par d�finition, pas de champs li�s pour les CB Direct!!!
                    '    Call Display_Status("Etat d'avancement de la fusion : (Champs li�s)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                    '    L_Result = Fusion_Champs_Lies(L_New_doc, _
                    '                                  L_Table_Lies_Tmp, _
                    '                                  L_Nb_champs_lies, _
                    '                                  L_champ_lie_jointure_data, _
                    '                                  L_connexion, _
                    '                                  p_Form)
                    '    If L_Result <> "Ok" Then
                    '        Fusion_fichier_Data = "Erreur dans les champs li�s (erreur:" & L_Result & ")"
                    '        GoTo Clean_Exit
                    '    End If
                    'End If
                    Rem ************************************************************************
                    Rem SUPPRIME LE 2017-03-14 ************************************************************************
                            
                    If Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms Then
                        L_New_pdf = Replace(L_New_doc, ".doc", IIf(DatamatriX, "_woDatamatrix", "") & ".pdf", , , vbTextCompare)
                        If L_New_doc = "" Then
                            L_New_pdf = L_ID_PE & IIf(DatamatriX, "_woDatamatrix", "") & ".pdf"
                        End If
                        If L_Type_PDF Then  'dans Pas de page de garde
                            If Not L_Pdf_Generated Then

                                Call Display_Status("Etat d'avancement de la fusion : (Enregistrement du Doc)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                If Not FileExists(DirSequestre & L_New_doc) Then
                                    Fusion_fichier_Data = "Document fusionn� introuvable : " & DirSequestre & L_New_doc
                                    GoTo Clean_Exit
                                End If
                                    
                                If InStr(1, G_Wd.Documents(L_New_doc).Application.ActivePrinter, G_Driver_PDFCreator, vbTextCompare) = 0 Then
                                    Call Display_Status("Etat d'avancement de la fusion : (Affectation du driver)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                    G_Wd.Documents(L_New_doc).Application.ActivePrinter = G_Driver_PDFCreator
                                End If
                                Call Display_Status("Etat d'avancement de la fusion : (G�n�ration du PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE & IIf(DatamatriX, "_woDatamatrix", ""))
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Impossible d'initialiser PdfCreator " & L_Result
                                    GoTo Clean_Exit
                                End If
                                    
                                G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                                Call WaitingPdfCreator
                                    
                                G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                                'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                                If get_file_size_only(DirSequestre & L_ID_PE & IIf(DatamatriX, "_woDatamatrix", "") & ".pdf") < 5000 Then
                                    Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                    GoTo Clean_Exit
                                End If
                            
                            
                            End If
                                    
                            Rem Si la fusion du mod�le n�cessite une "fusion de fond de page"
                            If Not L_Pdf_Generated Then
                                If L_Fond_de_Page <> vbNullString Then
                                    If Not L_Fond_de_Page_P2P Then
                                        L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                    Else
                                        L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                    End If
                                End If
                            End If
                            L_Pdf_Generated = True

                            L_New_pdf = L_ID_PE & IIf(DatamatriX, "_woDatamatrix", "") & ".pdf"
                            
                            Rem lire le nombre de page du nouveau pli
                            L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                            If L_Nb_pages_dans_pli = -1 Then
                                Rem Probl�me dans lecture du PDF avec ActivePDF (limitation outil)
                                L_Nb_pages_dans_pli = GetPageCount(DirSequestre & "\" & L_New_pdf)
                            End If
                            If L_Nb_pages_dans_pli > 4900 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(3)"
                                GoTo Clean_Exit
                            End If
                            If L_Nb_pages_dans_pli < 1 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(<1)"
                                GoTo Clean_Exit
                            End If
                            L_PageCount = L_Nb_pages_dans_pli
                            Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                            L_SheetCount = L_Nb_pages_dans_pli
                            Rem Cas du Datamatrix sur un document WORD sans PJ
                            If DatamatriX Then
                                L_Result = Insert_Datamatrix6(MW6DataMatrixFusion, 0, 0, DirSequestre & "\" & L_New_pdf, DirSequestre & "\" & Replace(L_New_pdf, "_woDatamatrix", "", , , vbTextCompare), "", L_SheetCount, "", DtmxVide, False, False, True)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Probl�me lors de l'ajout du Datamatrix!!! (003)"
                                    GoTo Clean_Exit
                                End If
                                L_New_pdf = Replace(L_New_pdf, "_woDatamatrix", "", , , vbTextCompare)
                            End If

                            
                            L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)

                        Else ' = If not L_Type_PDF then    (dans Pas de page de garde)
                                
                                    
                            Rem lci : ici, si If not L_Type_PDF then Je suis dans le cas d'un mod�le WORD, d'autres cas?
                            Rem ============================================
                            Rem Cr�ation des PDF
                            Rem ============================================
                            Call Display_Status("Etat d'avancement de la fusion : (G�n�ration du PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE & IIf(DatamatriX, "_woDatamatrix", ""))
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = "Impossible d'initialiser PdfCreator " & L_Result
                                GoTo Clean_Exit
                            End If
                            G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                            Call WaitingPdfCreator
                            
                            Call Display_Status("Etat d'avancement de la fusion : (Fermeture du document)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                            
                            'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                            If get_file_size_only(DirSequestre & L_ID_PE & IIf(DatamatriX, "_woDatamatrix", "") & ".pdf") < 5000 Then
                                Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                GoTo Clean_Exit
                            End If
                            
                            Call Display_Status("Etat d'avancement de la fusion : (Lecture du nombre de pages)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                            Rem lire le nombre de page du nouveau pli
                            L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                            If L_Nb_pages_dans_pli > 4900 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(4)"
                                GoTo Clean_Exit
                            End If
                            
                            L_PageCount = L_Nb_pages_dans_pli
                            Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                            L_SheetCount = L_Nb_pages_dans_pli
                            If DatamatriX Then
                                L_Result = Insert_Datamatrix6(MW6DataMatrixFusion, 0, 0, DirSequestre & "\" & L_New_pdf, DirSequestre & "\" & Replace(L_New_pdf, "_woDatamatrix", "", 1, , vbTextCompare), "", L_SheetCount, "", DtmxVide, False, False, True)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Probl�me lors de l'ajout du Datamatrix!!! (004)"
                                    GoTo Clean_Exit
                                End If
                                L_New_pdf = Replace(L_New_pdf, "_woDatamatrix", "", , , vbTextCompare)
                            End If
                            
                            Rem rechercher si certains des �l�ments suivants peuvent �tre r�cup�r�s
                            L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                            
                        End If
                                
                        If L_Service_Transformation_fax Then
                            If Not L_Type_PDF And LCase(L_Type_Fichier_joint) <> "pdf" And Trim(L_Type_Fichier_joint) <> "" Then
                                'Peux t-on envore avoir ce cas???
                                L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Impossible d'initialiser PdfCreator " & L_Result
                                    GoTo Clean_Exit
                                End If
                                G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                                Call WaitingPdfCreator
                                
                                ' Sauve le .doc sur le r�pertoire de production
                                Call Display_Status("Etat d'avancement de la fusion : (Fermeture du document)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges
                                
                                'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                                If get_file_size_only(DirSequestre & L_ID_PE & ".pdf") < 5000 Then
                                    Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                    GoTo Clean_Exit
                                End If
                                L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                                If L_Nb_pages_dans_pli < 0 Then
                                    Fusion_fichier_Data = "Nombre de pages du PDF illisibles!!!"
                                    GoTo Clean_Exit
                                End If
                                
                                L_PageCount = L_Nb_pages_dans_pli
                                Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                                L_SheetCount = L_Nb_pages_dans_pli
                                Rem Pas de datamatrix pour le fax !!!!
                                Rem rechercher si certains des �l�ments suivants peuvent �tre r�cup�r�s
                                L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                            End If
                        End If
                    Else
                        Rem eMail ou Fax (ou Sms)
                        If L_Service_Transformation_Mail Then
                            Rem S�lection du format d'enregistrement
                            If L_Email_Piece_Jointe_Fusionnee Then
                                
                                L_New_pdf = Replace(L_New_doc, ".doc", ".pdf", , , vbTextCompare)
                                
                                If InStr(1, G_Wd.Documents(L_New_doc).Application.ActivePrinter, G_Driver_PDFCreator, vbTextCompare) = 0 Then
                                    G_Wd.Documents(L_New_doc).Application.ActivePrinter = G_Driver_PDFCreator
                                End If
                                L_Result = InitPdfCreatorPrint(DirSequestre, L_ID_PE)
                                If L_Result <> "Ok" Then
                                    Fusion_fichier_Data = "Impossible d'initialiser PdfCreator " & L_Result
                                    GoTo Clean_Exit
                                End If
                                
                                G_Wd.Documents(L_New_doc).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
                                Call WaitingPdfCreator
                                    
                                Call Display_Status("Etat d'avancement de la fusion : (Fermeture du document)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                G_Wd.Documents(L_New_doc).Close savechanges:=wdDoNotSaveChanges

                                'Controle de la taille du PDF g�n�r� => Si < 8 Ko => erreur!!
                                If get_file_size_only(DirSequestre & L_ID_PE & ".pdf") < 5000 Then
                                    Fusion_fichier_Data = "PDF g�n�r� blanc!!!"
                                    GoTo Clean_Exit
                                End If

                                Rem Si la fusion du mod�le n�cessite une "fusion de fond de page"
                                If L_Fond_de_Page <> vbNullString Then
                                    If Not L_Fond_de_Page_P2P Then
                                        L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                    Else
                                        L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                    End If
                                End If
                                
                                Rem lire le nombre de page du nouveau pli
                                Rem Nous sommes dans un mail!!!
                                L_Nb_pages_dans_pli = 0
                                L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                                    
                            
                            ElseIf (L_ModelType = "Mod�le Texte" Or L_ModelType = "Mod�le Html") Then
Suite_Mail_CorpsDYNamique:
                            Else 'Mail Classique
                                    
                                If L_ModelName <> vbNullString Then
                                    If Left(L_Type_Fichier_joint, 3) <> "htm" Then
                                        If FileExists(DirSequestre & L_New_doc) Then
                                            L_Result = Word_Save_To_Html(L_New_doc, DirSequestre, L_ID_PE)
                                        Else
                                            If FileExists(DirSequestre & L_New_doc) Then
                                                L_Result = Word_Save_To_Html(L_New_doc, DirSequestre, L_ID_PE)
                                            End If
                                        End If
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                        G_Wd.Documents(Replace(L_New_doc, ".doc", "_Mail_Body.html", , , vbTextCompare)).Close savechanges:=wdDoNotSaveChanges
                                        If FileExists(DirSequestre & L_New_doc) Then
                                            L_Result = Replace_Image_Number_By_Image_List(DirSequestre, _
                                                                                      Replace(L_New_doc, ".doc", "_Mail_Body.html"), _
                                                                                      L_Url_Dir)
                                        End If
                                        Err.Clear
                                        L_Result = "Ok"
                                        If L_Result <> "Ok" Then
                                            Fusion_fichier_Data = L_Result
                                            GoTo Clean_Exit
                                        End If
                                    Else
                                        FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, G_dir_production & "\" & pCustomerNumber & "\" & L_ID_PE & "_Mail_Body.html"
                                    End If
                                Else
                                    If L_Nb_Fichier_Joint > 0 Then
                                        'Rem Ca classique LCI 22/05/2017
                                        'If L_Nb_Fichier_Joint = 1 Then
                                            If LCase(L_Fichier_Joint(1).t_type) = "pdf" Then
                                                FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_ID_PE & ".pdf"
                                            Else
                                                If FileExists(L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier) Then
                                                    FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_ID_PE & "_Mail_Body.html"
                                                ElseIf FileExists(L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier) Then
                                                    FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, G_dir_production & "\" & pCustomerNumber & "\" & L_ID_PE & "_Mail_Body.html"
                                                Else
                                                    Fusion_fichier_Data = "001 - Anomalie Copie!"
                                                    GoTo Clean_Exit
                                                End If
                                            End If
                                        'Else
                                        '    L_list_Pdf_Files = vbNullString
                                        '    For i = 1 To L_Nb_Fichier_Joint
                                        '        If LCase(Right(L_Fichier_Joint(i).t_Fichier, 4)) = ".pdf" Then
                                        '            L_list_Pdf_Files = L_list_Pdf_Files & L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier & "|"
                                        '        End If
                                        '    Next
                                        '    'stop
                                        '    If L_list_Pdf_Files <> "" Then
                                        '        If InStr(1, L_list_Pdf_Files, "|", vbTextCompare) > 0 Then
                                        '            L_New_pdf = L_ID_PE & ".pdf"
                                        '            L_Result = Concat_PdfLib(p_Form, L_list_Pdf_Files, DirSequestre & L_New_pdf)
                                        '            If L_Result <> "Ok" Then
                                        '                Fusion_fichier_Data = L_Result
                                        '                GoTo Clean_Exit
                                        '            End If
                                        '        Else
                                        '            FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_ID_PE & ".pdf"
                                        '        End If
                                        '    End If
                                        'End If
                                    Else
                                        If UBound(L_Fichier_Joint) > 0 Then
                                            FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, G_dir_production & "\" & pCustomerNumber & "\" & L_ID_PE & "_Mail_Body.html"
                                        End If
                                    End If
                                End If
                            End If

                            If L_Email_ReplyTo <> "" And lMailReplyTO = "" Then
                                lMailReplyTO = L_Email_ReplyTo
                            End If

                            L_Result = Mail_Creation(L_ID_PE, _
                                                     L_Email_to, _
                                                     L_Email_From, _
                                                     L_Email_xFer, _
                                                     L_Fk_destinataire, _
                                                     pCustomerNumber, _
                                                     L_Email_Sql_Pli_Update, _
                                                     L_chemin, _
                                                     L_Email_Piece_Jointe_Fusionnee, _
                                                     L_SujetMailFusionne, _
                                                     L_List_Fichiers_Mail_To_Move, _
                                                     p_liste_fichiers_joints_production_locale, _
                                                     L_Liste_Pieces_Jointes, _
                                                     L_Document_Size, False, False, "", _
                                                     DirSequestre, pUnzipDir, _
                                                     lMailCC, lMailCCI, lMailReplyTO, _
                                                     decoupePDFsPath)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                        End If
                                
                                
Rem ====================================================================================================================
Rem DEBUT SMS =====================================================================================================
                        If L_Service_Transformation_Sms Then
                            L_Email_From = L_eSmsAddressFrom
                            If Not G_Test Then
                                'Graphnet
                                L_Email_to = "sms=" & L_FaxSmsNumber & "#CR=" & L_ID_PE & L_eSmsAddressTo
                            Else
                                L_Email_to = "exploitation@axessy.fr"
                            End If
                            'If Sigma Then => LCI Supprim� le 14/08/2017
                            '    Rem on applique le bon encodage
                            '    L_Sms_Message = UTF8_Decode(L_Sms_Message)
                            'End If
                            L_Email_xFer = L_FaxSmsNumber
                            If L_Email_ReplyTo <> "" And lMailReplyTO = "" Then
                                lMailReplyTO = L_Email_ReplyTo
                            End If
                            
                            L_Result = Mail_Creation(L_ID_PE, _
                                                     L_Email_to, _
                                                     L_Email_From, _
                                                     L_Email_xFer, _
                                                     L_Fk_destinataire, _
                                                     pCustomerNumber, _
                                                     L_Email_Sql_Pli_Update, _
                                                     L_chemin, _
                                                     L_Email_Piece_Jointe_Fusionnee, _
                                                     L_SujetMailFusionne, _
                                                     L_List_Fichiers_Mail_To_Move, _
                                                     p_liste_fichiers_joints_production_locale, _
                                                     L_Liste_Pieces_Jointes, _
                                                     L_Document_Size, False, True, L_Sms_Message, _
                                                     DirSequestre, pUnzipDir, _
                                                     lMailCC, lMailCCI, lMailReplyTO, _
                                                     decoupePDFsPath)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            Rem Sms Sms Sms Sms
                        End If
Rem FIN SMS =======================================================================================================
Rem ====================================================================================================================
                            
                    End If
                    
                    Rem On ne ferme plus le mod�le
                    On Error Resume Next
                    
                    Rem ***************************************************************************
                    Rem FUSION NORMALE
                    Rem ***************************************************************************
                    If L_Service_Transformation_Mail _
                    Or _
                    L_Service_Transformation_Sms Then
                        
                        If L_Service_Transformation_Mail Then
                            L_Nb_pages_dans_pli = 1
                        End If
                        
                        If L_Service_Transformation_Sms Then
                            L_Nb_pages_dans_pli = ReadNbSms(L_Sms_Message)
                            L_PageCount = L_Nb_pages_dans_pli
                            L_SheetCount = L_Nb_pages_dans_pli
                            Rem Pas de Datamatrix pour le Sms!!!!
                        End If
                    
                    End If
                            
                    Rem LCI R�pertoire local de Production supprim�
                    'If L_Service_Transformation_Mail Then
                    '    If FileExists(G_dir_local_production & "\" & pCustomerNumber & "\" & L_New_doc) Then
                    '        If FileExists(G_dir_production & "\" & pCustomerNumber & "\" & L_New_doc) Then
                    '            Kill G_dir_local_production & "\" & pCustomerNumber & "\" & L_New_doc
                    '        End If
                    '    End If
                    'End If

Rem ====================================================================================================================
Rem DEBUT ARCHIVAGE ONLY ==========================================================================================
Archivage_Only_Suite:
FusionSuite1:
                    If L_Archivage_Only Or (L_ModelType = "Aucun" And Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms) Or L_Service_ePoBox And Not L_ePoBox_Hybrid Then
                        Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Archivage des documents", p_Form)
                                
                        L_Type_PDF = True
                        L_list_Pdf_Files = vbNullString
                        For i = 1 To L_Nb_Fichier_Joint
                            L_list_Pdf_Files = L_list_Pdf_Files & L_Fichier_Joint(i).t_dir & L_Fichier_Joint(i).t_Fichier & "|"
                            If UCase(L_Fichier_Joint(i).t_type) <> "PDF" Then
                                L_Type_PDF = False
                            End If
                        Next
                        If L_Type_PDF Then
                            L_New_pdf = L_ID_PE & ".pdf"
                        Else
                            If InStr(1, L_Fichier_Joint(1).t_Fichier, ".", vbTextCompare) > 0 Then
                                L_New_pdf = L_ID_PE & Mid(L_Fichier_Joint(1).t_Fichier, InStrRev(L_Fichier_Joint(1).t_Fichier, ".", , vbTextCompare))
                            Else
                                L_New_pdf = L_ID_PE
                            End If
                        End If
                                
                        If L_Type_PDF Then
                            If L_Nb_Fichier_Joint > 1 Then
                                Rem Tourner dans les pi�ces jointes!!!
                                If L_list_Pdf_Files <> "" Then
                                    Call Display_Status("Etat d'avancement de la fusion : (Concat�nation PDF)", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
                                    L_Result = Concat_PdfLib(p_Form, L_list_Pdf_Files, DirSequestre & L_New_pdf)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                End If
                            Else
                                Rem Si flow, on ne recopie pas !!!
                                Rem comme avant
                                If Not FileExists(DirSequestre & L_New_pdf) Then
                                    FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf
                                End If
                            End If
                                    
                            If L_Fond_de_Page <> vbNullString Then
                                If Not L_Fond_de_Page_P2P Then
                                    L_Result = Merge_Pdf(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, False, pCustomerNumber)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                Else
                                    L_Result = Merge_Pdf_P2p_New(p_Form, DirSequestre & L_New_pdf, L_Fond_de_Page, DirSequestre & L_New_pdf)
                                    If L_Result <> "Ok" Then
                                        Fusion_fichier_Data = L_Result
                                        GoTo Clean_Exit
                                    End If
                                End If
                            End If

                            L_Nb_pages_dans_pli = Read_Pdf_Num_Pages(p_Form, DirSequestre, L_New_pdf)
                            If L_Nb_pages_dans_pli = -1 Then
                                Rem Probl�me dans lecture du PDF
                                L_Nb_pages_dans_pli = GetPageCount(DirSequestre & "\" & L_New_pdf)
                            End If
                            If L_Nb_pages_dans_pli > 4900 Then
                                Fusion_fichier_Data = "Probl�me lors de la lecture des pages!!!(5)"
                                GoTo Clean_Exit
                            End If
                            L_PageCount = L_Nb_pages_dans_pli
                            Call Update_Nb_Pages(L_Nb_pages_dans_pli, L_Type_Impression)
                            L_SheetCount = L_Nb_pages_dans_pli
                            Rem Pas de Datamatrix pour l'Archivage !!!!
                        Else 'Pas PDF
                            If L_Nb_Fichier_Joint <> 1 Then
                                L_Result = "Un d�pot de type archivage (seulement) ne peut contenir plusieurs documents joints autrement qu'au format PDF!"
                                Fusion_fichier_Data = L_Result
                                GoTo Clean_Exit
                            End If
                            L_Nb_pages_dans_pli = 0
                            FileCopy L_Fichier_Joint(1).t_dir & L_Fichier_Joint(1).t_Fichier, DirSequestre & L_New_pdf
                        End If
                        
                        Rem rechercher si certains des �l�ments suivants peuvent �tre r�cup�r�s
                        L_Document_Size = get_file_size_only(DirSequestre & L_New_pdf)
                    End If
Rem FIN ARCHIVAGE ONLY ==========================================================================================
Rem ====================================================================================================================
                End If
                        
                
Archivage_Only_Suite2:
                Call Display_Status("Mises � jour de la base de donn�es...", "", p_Form)

                Rem Le destinataire vient d'�tre cr��!!!
                        
                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Insertion du pli", p_Form)
                Call CheckApplication("Fusion processing " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter)
                L_SQL = " INSERT INTO PLI (FK_PLI_STATUT_TELEGRAMME, "
                L_SQL = L_SQL & " OF_AUTOMATIQUE, BAT, "
                L_SQL = L_SQL & " FK_PREPARATION, FK_PRESTATION_MODEL, "
                L_SQL = L_SQL & " ID_PE, "
                L_SQL = L_SQL & " NOM_FICHIER_PDF, NOM_FICHIER_EMISSION, "
                Rem DOCUMENT ENVOI XXX
                L_SQL = L_SQL & " document_envoi_nom, "
                L_SQL = L_SQL & " document_envoi_date_creation, "
                L_SQL = L_SQL & " document_envoi_type, "
                L_SQL = L_SQL & " document_envoi_taille, "
                L_SQL = L_SQL & " pli_emission_size, "
                L_SQL = L_SQL & " fk_contact, "
                L_SQL = L_SQL & " prix_pli, "
                L_SQL = L_SQL & " prix_affranchissement, "
                L_SQL = L_SQL & " poids_pli, "
                L_SQL = L_SQL & " pli_code_pays_iso3A, "
                L_SQL = L_SQL & " pli_service_postal, "
                L_SQL = L_SQL & " pli_zone_postale, "
                L_SQL = L_SQL & " FK_PLI_STATUT, FK_DESTINATAIRE, FK_SOCIETE, "
                L_SQL = L_SQL & " REGROUPEMENT_PLI, "
                If L_Sql_InsertPli = "" Then
                    L_SQL = L_SQL & " DATE_EMISSION, "
                Else
                    L_SQL = L_SQL & L_Sql_InsertPli
                End If
                L_SQL = L_SQL & " NUM_RECORD_EMISSION, NOMBRE_PAGES, PAGE_COUNT, SHEET_COUNT, MAJ_DATE, MAJ_USERID "
                L_SQL = L_SQL & ") "
                L_SQL = L_SQL & " VALUES (" & IIf(L_TG, L_Statut_Telegramme_a_traiter, 0) & ", "
                L_SQL = L_SQL & IIf(L_OF_Auto, 1, 0) & ", 0, "
                L_SQL = L_SQL & p_Preparation_Fk & ", " & p_Prestation_Model_Pk & ", "
                L_SQL = L_SQL & " '" & L_ID_PE & "', "
                L_SQL = L_SQL & " '" & L_ID_PE & ".pdf', '" & L_ID_PE & ".doc', "
                Rem DOCUMENT ENVOI XXX
                If L_Archivage_Only Or (L_ModelType = "Aucun" And Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms) Or L_ModelType = "CB Direct" Or L_ModelType = "CB + Adresse" Or (L_Service_ePoBox And Not L_ePoBox_Hybrid) Then
                Rem LCI le 14/08/2017 suppression du cas Masque PDF 'If L_Archivage_Only Or (L_ModelType = "Aucun" And Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms) Or L_ModelType = "CB Direct" Or L_ModelType = "CB + Adresse" Or L_ModelType = "Masque Pdf" Or (L_Service_ePoBox And Not L_ePoBox_Hybrid) Then
                    L_SQL = L_SQL & " '" & L_New_pdf & "', "
                    L_SQL = L_SQL & " now(), "
                    If L_Type_PDF Then
                        L_SQL = L_SQL & " 'pdf', "
                    Else
                        If InStr(1, L_New_pdf, ".", vbTextCompare) > 0 Then
                            Rem
                            L_SQL = L_SQL & " '" & Mid(L_New_pdf, InStr(1, L_New_pdf, ".", vbTextCompare) + 1) & "', "
                        Else
                            L_SQL = L_SQL & " '', "
                        End If
                    End If
                    L_SQL = L_SQL & " " & L_Document_Size & ", "
                ElseIf L_Service_Transformation_Mail Then
                    L_SQL = L_SQL & " '" & L_ID_PE & "_Mail_Body.html', "
                    L_SQL = L_SQL & " now(), "
                    L_SQL = L_SQL & " 'html', "
                    L_Document_Size = get_file_size_only(DirSequestre & L_ID_PE & "_mail.eml")
                    L_SQL = L_SQL & " " & L_Document_Size & ", "
                ElseIf L_Service_Transformation_Sms Then

                    L_SQL = L_SQL & " '" & L_ID_PE & ".txt', "
                    L_SQL = L_SQL & " now(), "
                    L_SQL = L_SQL & " 'txt', "
                    If FileExists(DirSequestre & L_ID_PE & ".txt") Then
                        L_Document_Size = get_file_size_only(DirSequestre & L_ID_PE & ".txt")
                    Else
                        L_Document_Size = get_file_size_only(G_dir_production & "\" & pCustomerNumber & "\" & L_ID_PE & ".txt")
                    End If
                    L_SQL = L_SQL & " " & L_Document_Size & ", "
                Else
                    L_SQL = L_SQL & " '" & L_ID_PE & ".pdf', "
                    L_SQL = L_SQL & " now(), "
                    L_SQL = L_SQL & " 'pdf', "
                    L_SQL = L_SQL & " " & L_Document_Size & ", "
                End If
                L_SQL = L_SQL & " " & L_Document_Size & ", "
                L_SQL = L_SQL & p_Fk_Contact & ", "
                
                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Insertion du pli - Calcul Affranchissement (Service WEB)", p_Form)
                Call CheckApplication("Fusion processing " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter & " (price)")
                L_Mnt_Total = PrixPliGlobal(L_ID_PE, L_SheetCount, _
                                            p_Prestation_Model_Pk, _
                                            L_Poids_Pli, _
                                            L_Mnt_Service, _
                                            L_Mnt_Affranchissement, _
                                            L_Pli_Code_pays_Iso3A, _
                                            L_Pli_Service_Postal, _
                                            L_Pli_Zone_Postale, _
                                            L_Pli_Adresse, _
                                            IIf(L_Service_Transformation_fax = True, FaxNumberWebS(L_FaxSmsNumber), ""), _
                                            p_Form, _
                                            L_PageCount, _
                                            (L_Type_Impression = "Recto/Verso"), _
                                            NewURLforAffranchissement)
                                            
                L_SQL = L_SQL & Replace(L_Mnt_Service, ",", ".") & ", "
                L_SQL = L_SQL & Replace(L_Mnt_Affranchissement, ",", ".") & ", "
                L_SQL = L_SQL & Replace(L_Poids_Pli, ",", ".") & ", "
                L_SQL = L_SQL & " '" & L_Pli_Code_pays_Iso3A & "', "
                L_SQL = L_SQL & " '" & L_Pli_Service_Postal & "', "
                L_SQL = L_SQL & " '" & L_Pli_Zone_Postale & "', "
                
                L_SQL = L_SQL & L_Pli_statut_pk & ", " & L_Fk_destinataire & ", " & p_Societe_Fk & ", "
                L_SQL = L_SQL & "0, "
                
                If L_Sql_InsertPli = "" Then
                    L_SQL = L_SQL & " now() , "
                Else
                    L_SQL = L_SQL & L_Sql_ValuesPli
                End If
                L_SQL = L_SQL & L_Num_Record & ", " & L_Nb_pages_dans_pli & ", " & L_PageCount & ", " & L_SheetCount & ", now() , '" & G_User_Id & "') "
                
                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Insertion du pli - Calcul Affranchissement termin�(Service WEB)", p_Form)
                If Run_Execute_Sql(L_SQL) = -1 Then
                    Fusion_fichier_Data = "Impossible de cr�er le pli " & L_ID_PE
                    GoTo Clean_Exit
                End If
                Rem Insert du pli contenu ? Toute la page
                L_Fk_Pli = Lire_Un_Champ("PK_PLI", "PLI", "ID_PE = '" & L_ID_PE & "'")
                    
                If L_Validation_Contenu Then
                    Rem controle de l'existence d'un statut dynamique
                    Rem Si au moins un champ emis de type statut dynamique, pas de mise � jour du statut
                    If Not L_Presence_Statut_Dynamique Then
                        L_SQL = "update pli set fk_pli_statut=" & L_Statut_A_Valider_Par_Le_Client & " where id_pe='" & L_ID_PE & " '"
                        If Run_Execute_Sql(L_SQL) = -1 Then
                            Fusion_fichier_Data = "Impossible de mettre le statut du pli � jour"
                            GoTo Clean_Exit
                        End If
                    End If
                Else
                    If L_Service_robot_t2c Then
                        t2c_Status_pk_pli_list_go = t2c_Status_pk_pli_list_go & "," & L_Fk_Pli
                    End If
                End If
                  
                If L_Service_Transformation_Mail Then
                    L_Email_Sql_Pli_Update = "UPDATE PLI SET " & L_Email_Sql_Pli_Update & " WHERE id_pe='" & L_ID_PE & "'"
                    Call Run_Execute_Sql(L_Email_Sql_Pli_Update)
                    If L_Liste_Pieces_Jointes <> vbNullString Then
                    Rem Ajout des pi�ces jointes dans la table PLI_EMAIL_PIECES_JOINTES
                        Call Mail_Insert_Pi�ces_Jointes(L_Liste_Pieces_Jointes, L_ID_PE)
                    End If
                End If
                
                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "R�f�rencement des documents", p_Form)
                
                If L_Service_Creation_Des_Coffres Then
                    GoTo Direct_Service_Creation_Des_Coffres
                End If

                If L_Archivage_Only Or (L_ModelType = "Aucun" And Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms) Or L_ModelType = "CB Direct" Or L_ModelType = "CB + Adresse" Then
                Rem LCI le 14/08/2017 Suppression du cas Masque PDF      'If L_Archivage_Only Or (L_ModelType = "Aucun" And Not L_Service_Transformation_Mail And Not L_Service_Transformation_Sms) Or L_ModelType = "CB Direct" Or L_ModelType = "CB + Adresse" Or L_ModelType = "Masque Pdf" Then
                    Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Archivage des documents", p_Form)
                    'LCI SUPPRESSION SEQUESTRE LE 14/08/2017                     L_Result = Pdf_Dir2(DirSequestre, L_New_pdf, Now, L_chemin, M_Make_Sequestre, True, pCustomerNumber)
                    L_Result = Pdf_Dir2(DirSequestre, L_New_pdf, Now, L_chemin, False, True, pCustomerNumber)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                    
                    'If M_Make_Xades Then          'LCI SUPPRESSION XADES LE 14/08/2017
                    '    If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "XADES", M_Make_Worm) Then
                    '        Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                    '        GoTo Clean_Exit
                    '    End If
                    'End If
                    If M_Make_PdfS Then
                        'LCI SUPPRESSION WORM LE 14/08/2017
                        'If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "PDF", M_Make_Worm) Then
                        If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "PDF", False) Then
                            Fusion_fichier_Data = "Impossible de copier le document � signer!"
                            GoTo Clean_Exit
                        End If
                    End If
                    If L_Service_robot_t2c And t2c_WebService_Actif Then
                        If t2c_NomDocument = "" Then
                            t2c_NomDocument = L_New_pdf
                        Else
                            Rem Si le type est diff�rent de celui qui est g�n�r�, on donne le m�me
                            If FileType(t2c_NomDocument) <> FileType(L_New_pdf) Then
                                t2c_NomDocument = t2c_NomDocument & "." & FileType(L_New_pdf)
                            End If
                        End If
                        If M_Make_PdfS Then
                            WaitSignedPdf = 0
                            L_New_pdf = Replace(L_New_pdf, ".pdf", ".tm1.pdf", , , vbTextCompare)
                        End If
NewTryInjector:
                        L_Result = t2c_Injector(CLng(L_Fk_Pli), 0, CLng(p_Societe_Fk), L_chemin & "\" & L_New_pdf, _
                                              t2c_Profil, t2c_NomDocument, t2c_DateDocument, t2c_UserID, t2c_IndexDocument, t2c_Classement, "wait")
                        If L_Result <> "Ok" Then
                            If M_Make_PdfS Then
                                If WaitSignedPdf < 500 Then
                                    DoEvents
                                    Sleep 500
                                    DoEvents
                                    Sleep 500
                                    DoEvents
                                    Sleep 500
                                    DoEvents
                                    Sleep 500
                                    DoEvents
                                    WaitSignedPdf = WaitSignedPdf + 1
                                    Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Attente du PDF sign�... Tentative n�" & WaitSignedPdf, p_Form)
                                    GoTo NewTryInjector
                                End If
                            End If
                            Fusion_fichier_Data = L_Result
                            GoTo Clean_Exit
                        End If
                        If M_Make_PdfS Then
                            L_New_pdf = Replace(L_New_pdf, ".tm1.pdf", ".pdf", , , vbTextCompare)
                        End If
                    End If
                
                ElseIf L_Service_Transformation_Sms Then
                    'LCI SUPPRESSION SEQUESTRE LE 14/08/2017                     L_Result = Pdf_Dir3(DirSequestre, L_ID_PE & "_mail.eml", Now, L_chemin, M_Make_Sequestre)
                    L_Result = Pdf_Dir3(DirSequestre, L_ID_PE & "_mail.eml", Now, L_chemin, False)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                    'If M_Make_Xades Then   'LCI SUPPRESSION XADES LE 14/08/2017
                    '    If Not SignaturePli(L_chemin, L_ID_PE & "_mail.eml", L_ID_PE, "XADES", M_Make_Worm) Then
                    '        Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                    '        GoTo Clean_Exit
                    '    End If
                    'End If
                    If M_Make_PdfS Then
                        'LCI SUPPRESSION WORM LE 14/08/2017
                        'If Not SignaturePli(L_chemin, L_ID_PE & "_mail.eml", L_ID_PE, "PDF", M_Make_Worm) Then
                        If Not SignaturePli(L_chemin, L_ID_PE & "_mail.eml", L_ID_PE, "PDF", False) Then
                            Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                            GoTo Clean_Exit
                        End If
                    End If
                
                ElseIf L_Service_Transformation_Mail Then
                    'Call Display_Status("D�placement des fichiers...", "", p_Form)
                    Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Signature du fichier", p_Form)
                    'LCI SUPPRESSION SEQUESTRE LE 14/08/2017                     L_Result = Html_Dir(DirSequestre, L_ID_PE & "_mail_body.html", L_chemin, M_Make_Sequestre, L_List_Fichiers_Mail_To_Move, DirSequestre, pUnzipDir)
                    L_Result = Html_Dir(DirSequestre, L_ID_PE & "_mail_body.html", L_chemin, False, L_List_Fichiers_Mail_To_Move, DirSequestre, pUnzipDir)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                    'If M_Make_Xades Then       'LCI SUPPRESSION WORM LE 14/08/2017
                    '    If Not SignaturePli(L_chemin, L_ID_PE & "_mail.eml", L_ID_PE, "XADES", M_Make_Worm) Then
                    '        Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                    '        GoTo Clean_Exit
                    '    End If
                    'End If
                    If M_Make_PdfS Then
                        If FileExists(L_chemin & "\" & L_ID_PE & "_pj.pdf") Then
                            FileCopy L_chemin & "\" & L_ID_PE & "_pj.pdf", L_chemin & "\" & L_ID_PE & ".pdf"
                            'LCI SUPPRESSION WORM LE 14/08/2017                If Not SignaturePli(L_chemin, L_ID_PE & "_pj.pdf", L_ID_PE, "PDF", M_Make_Worm) Then
                            If Not SignaturePli(L_chemin, L_ID_PE & "_pj.pdf", L_ID_PE, "PDF", False) Then
                                Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                                GoTo Clean_Exit
                            End If
                            L_Fk_Pli = Lire_Un_Champ("PK_PLI", "PLI", "ID_PE = '" & L_ID_PE & "'")
Waiting4PdfS:
                            L_Result = Lire_Un_Champ("status", "pli_signature", "fk_pli = " & L_Fk_Pli)
                            If L_Result < 4 Then
                                Rem On attend
                                Sleep (500)
                                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Signature du fichier / Attente du fichier sign�...", p_Form)
                                Sleep (200)
                                GoTo Waiting4PdfS
                            ElseIf L_Result = 4 Then
                                Rem une erreur
                                Fusion_fichier_Data = "Impossible de signer le PDF! Fichier : " & L_Fichier_Joint(1).t_Fichier
                                GoTo Clean_Exit
                            ElseIf L_Result = 5 Then
                                Rem Ok
                                Rem On substitue les PJ
                                FileCopy L_chemin & "\" & L_ID_PE & ".tm1.pdf", L_chemin & "\" & L_ID_PE & "_pj.pdf"
                                Rem On �change la PJ dans le mail
                            End If
                            L_Result = Mail_Creation_ReplaceAttachedFile(L_chemin & "\" & L_ID_PE & "_mail.eml", L_chemin & "\" & L_ID_PE & "_mail.eml", L_chemin & "\" & L_ID_PE & "_pj.pdf")
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = "Impossible de copier le document � archiver! " & L_Result
                                GoTo Clean_Exit
                            End If
                            Rem On modifie la taille du document
                            Rem On copie dans le s�questre
                            FileCopy L_chemin & "\" & L_ID_PE & "_mail.eml", DirSequestre & L_ID_PE & "_mail.eml"
                            L_Document_Size = get_file_size_only(DirSequestre & L_ID_PE & "_mail.eml")
                            L_Email_Sql_Pli_Update = "UPDATE PLI SET document_envoi_taille = " & L_Document_Size & ", pli_emission_size = " & L_Document_Size & "  WHERE id_pe='" & L_ID_PE & "'"
                            Call Run_Execute_Sql(L_Email_Sql_Pli_Update)
                        Else
Rem ADDED BY LCI le 28/11/2017 pour le CAS MULTI PJ
                            If FileExists(L_chemin & "\" & L_ID_PE & "_pj_1.pdf") Then
                                FileCopy L_chemin & "\" & L_ID_PE & "_pj_1.pdf", L_chemin & "\" & L_ID_PE & ".pdf"
                            End If
                            If Not SignaturePli(L_chemin, L_ID_PE & "_pj_1.pdf", L_ID_PE, "PDF", False) Then
                                Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                                GoTo Clean_Exit
                            End If
                            L_Fk_Pli = Lire_Un_Champ("PK_PLI", "PLI", "ID_PE = '" & L_ID_PE & "'")
Waiting4PdfSPjs:
                            L_Result = Lire_Un_Champ("status", "pli_signature", "fk_pli = " & L_Fk_Pli)
                            If L_Result < 4 Then
                                Rem On attend
                                Sleep (500)
                                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "Signature du fichier / Attente du fichier sign�...", p_Form)
                                Sleep (200)
                                GoTo Waiting4PdfSPjs
                            ElseIf L_Result = 4 Then
                                Rem une erreur
                                Fusion_fichier_Data = "Impossible de signer le PDF! Fichier : " & L_Fichier_Joint(1).t_Fichier
                                GoTo Clean_Exit
                            ElseIf L_Result = 5 Then
                                Rem Ok
                                Rem On substitue les PJ
                                FileCopy L_chemin & "\" & L_ID_PE & ".tm1.pdf", L_chemin & "\" & L_ID_PE & "_pj_1.pdf"
                                Rem On �change la PJ dans le mail
                            End If
                            L_Result = Mail_Creation_ReplaceAttachedFile(L_chemin & "\" & L_ID_PE & "_mail.eml", L_chemin & "\" & L_ID_PE & "_mail.eml", L_chemin & "\" & L_ID_PE & "_pj_1.pdf", True)
                            If L_Result <> "Ok" Then
                                Fusion_fichier_Data = "Impossible de copier le document � archiver! " & L_Result
                                GoTo Clean_Exit
                            End If
                            Rem On modifie la taille du document
                            Rem On copie dans le s�questre
                            FileCopy L_chemin & "\" & L_ID_PE & "_mail.eml", DirSequestre & L_ID_PE & "_mail.eml"
                            L_Document_Size = get_file_size_only(DirSequestre & L_ID_PE & "_mail.eml")
                            L_Email_Sql_Pli_Update = "UPDATE PLI SET document_envoi_taille = " & L_Document_Size & ", pli_emission_size = " & L_Document_Size & "  WHERE id_pe='" & L_ID_PE & "'"
                            Call Run_Execute_Sql(L_Email_Sql_Pli_Update)
Rem FIN AJOUT BY LCI le 28/11/2017 pour le CAS MULTI PJ
                        End If
                    End If
                    
                    If L_Service_robot_t2c And t2c_WebService_Actif Then
                        L_New_pdf = L_ID_PE & ".pdf"
                        If t2c_NomDocument = "" Then
                            t2c_NomDocument = L_New_pdf
                        Else
                            Rem Si le type est diff�rent de celui qui est g�n�r�, on donne le m�me
                            If FileType(t2c_NomDocument) <> FileType(L_New_pdf) Then
                                t2c_NomDocument = t2c_NomDocument & "." & FileType(L_New_pdf)
                            End If
                        End If
                        If M_Make_PdfS Then
                            L_New_pdf = Replace(L_New_pdf, ".pdf", ".tm1.pdf", , , vbTextCompare)
                        End If
                        L_Result = t2c_Injector(CLng(L_Fk_Pli), 0, CLng(p_Societe_Fk), L_chemin & "\" & L_New_pdf, _
                                              t2c_Profil, t2c_NomDocument, t2c_DateDocument, t2c_UserID, t2c_IndexDocument, t2c_Classement, "wait")
                        If L_Result <> "Ok" Then
                            Fusion_fichier_Data = L_Result
                            GoTo Clean_Exit
                        End If
                    End If

                Else
                    'LCI SUPPRESSION SEQUESTRE LE 14/08/2017                     L_Result = Pdf_Dir2(DirSequestre, L_New_pdf, Now, L_chemin, M_Make_Sequestre, False, pCustomerNumber)
                    L_Result = Pdf_Dir2(DirSequestre, L_New_pdf, Now, L_chemin, False, False, pCustomerNumber)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                        
                    'If M_Make_Xades Then                    'LCI SUPPRESSION WORM LE 14/08/2017
                    '    If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "XADES", M_Make_Worm) Then
                    '        Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                    '        GoTo Clean_Exit
                    '    End If
                    'End If
                    If M_Make_PdfS Then
                        'LCI SUPPRESSION WORM LE 14/08/2017                     If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "PDF", M_Make_Worm) Then
                        If Not SignaturePli(L_chemin, L_New_pdf, L_ID_PE, "PDF", False) Then
                            Fusion_fichier_Data = "Impossible de copier le document � archiver!"
                            GoTo Clean_Exit
                        End If
                    End If
                    
                End If
                    
                Rem Fax => Ajouter la demande d'envoi !!! Attention � la validation !!!
                If L_Service_Transformation_fax Then
                    ' P_PkPli               : pk du pli � faxer
                    ' p_FaxNumber           : Num�ro de fax du destinataire => � mettre a format international
                    ' p_DisplayName         : Nom affich� sur le fax du destinataire
                    ' p_NotificationEmail   : si renseign�, envoi les notifications � cette adresse mail
                    ' p_NotificationFilter  : si renseign� et adresse mail pr�c�dente �galement, permet de filtrer les mails retourn�s ('ALL', 'Failures Only','None')
                    ' p_HeaderRecipient     : si renseign�, apparait en haut de la page fax�e
                    ' p_HeaderSender        : ?
                    ' p_Resolution          : qualit� du document fax� ('Normal', 'Fine')
                    ' p_AutoValidation      : si (autovalidation = 1 => envoi direct) sinon (autovalidation = 0, fax en attente d'�tre envoy�)
                    
                    If L_Nb_pages_dans_pli < 0 Then
                        Fusion_fichier_Data = "Nombre de pages du PDF illisibles!!!"
                        GoTo Clean_Exit
                    End If

                    L_Result = Fax__Create(L_Fk_Pli, FaxNumberWebS(L_FaxSmsNumber), False)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                End If
                    
                If L_Service_Creation_Des_Coffres Then
Direct_Service_Creation_Des_Coffres:
                    Rem Insert de l'enregistrement
                    L_Sql_CreateUserInInsert = L_Sql_CreateUserInInsert & " maj_date, maj_userid) "
                    L_Sql_CreateUserInSelect = L_Sql_CreateUserInSelect & " now(), '" & G_User_Id & "' "
                    L_Sql_CreateUserInSelect = L_Sql_CreateUserInSelect & " from destinataire where pk_destinataire = " & L_Fk_destinataire
                    L_Result = Run_Execute_Sql(L_Sql_CreateUserInInsert & L_Sql_CreateUserInSelect)
                    If L_Result < 0 Then
                        Fusion_fichier_Data = L_Result
                        GoTo Clean_Exit
                    End If
                End If
                L_List_Pk_Pli = L_List_Pk_Pli & L_Fk_Pli & ","

                Rem 1
                Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "D�tail du pli...", p_Form)
                
                Rem LCI le 14/08/2017 suppression de l'impression dynamique
                'If UBound(G_Print_Properties) > 0 Then
                '    L_Result = Load_Print_Properties(CLng(L_Fk_Pli))
                '    If L_Result <> "Ok" Then
                '        Fusion_fichier_Data = L_Result
                '        GoTo Clean_Exit
                '    End If
                '    Call Add_Pli_Fourniture(CLng(L_Fk_Pli), p_Prestation_Model_Pk, "Print Dynamic", CLng(L_Nb_pages_dans_pli))
                'Else
                    Call Add_Pli_Fourniture(CLng(L_Fk_Pli), p_Prestation_Model_Pk, "Normal", CLng(L_Nb_pages_dans_pli))
                'End If
                    
                
                L_SQL = " INSERT INTO PLI_PRIX (fk_pli, service_price_ht, maj_date, maj_userid) "
                L_SQL = L_SQL & " values (" & L_Fk_Pli & ", " & Replace(L_Mnt_Service, ",", ".") & ", now(), '" & G_User_Id & "')"
                Call Run_Execute_Sql(L_SQL)
                

                If L_Service_ePoBox Then
                    Rem ajouter un enregistrement epobox emetteur associ� au pli
                    Rem L_ePoBox_Destinataire_Prestation_Model_Nom
                    Rem L_ePoBox_Destinataire_Plateformeid
                    Rem L_ePoBox_Destinataire_clientid
                    Call Display_Status("Etat d'avancement de la fusion : (Fusion) " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, "ePobox...", p_Form)
                    Call CheckApplication("Fusion processing " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter & " (ePobox)")
                    If M_Make_PdfS Then
                        L_New_pdf = Replace(L_New_pdf, ".pdf", ".tm1.pdf", , , vbTextCompare)
                    End If
                    L_Pk_Pli_EpoBox_Emetteur = Add_Pli_ePoBox_Emetteur(CLng(L_Fk_Pli), _
                                                                       L_ePoBox_Destinataire_PlateformId, _
                                                                       L_ePoBox_Destinataire_ClientID, _
                                                                       L_ePoBox_Prestation_Model_Nom, _
                                                                       L_ePoBox_Destinataire_ContactId, _
                                                                       L_ePoBox_Destinataire_Reference, _
                                                                       L_ePoBox_Destinataire_Adresse, _
                                                                       L_New_pdf, _
                                                                       IIf(L_ePoBox_FileToXfer = vbNullString, "", L_ID_PE & "." & Type_Fichier(L_ePoBox_FileToXfer)), _
                                                                       L_ePobox_Id_Liste_Recapitulative)
                                                                       
                    If L_Pk_Pli_EpoBox_Emetteur = 0 Then
                        Fusion_fichier_Data = "Impossible d'enregistrer les informations relatives � l'ePoBox!"
                        GoTo Clean_Exit
                    End If
                    
                    Rem Mettre les informations de d�tail
                    For i = 0 To L_ePoBox_Nb_Fields - 1
                        Rem Toujours mettre le nom du fichier � envoyer
                        If L_tab_ePoBox_Data(i).t_Field_Type = "File" Then
                            Rem Il ne peut y en avoir qu'un!!
                            L_tab_ePoBox_Data(i).t_Field_Value = L_ID_PE & "." & Type_Fichier(L_ePoBox_FileToXfer)
                            If get_file_size_only(DirSequestre & L_ePoBox_FileToXfer) = 0 Then
                                Fusion_fichier_Data = "Le fichier " & L_ePoBox_FileToXfer & " p�se 0 ko!"
                                GoTo Clean_Exit
                            End If
                            FileCopy DirSequestre & L_ePoBox_FileToXfer, L_chemin & "\" & L_tab_ePoBox_Data(i).t_Field_Value
                            If get_file_size_only(DirSequestre & L_ePoBox_FileToXfer) <> get_file_size_only(L_chemin & "\" & L_tab_ePoBox_Data(i).t_Field_Value) Then
                                Fusion_fichier_Data = "Le fichier " & L_ePoBox_FileToXfer & " et sa copie ont une taille diff�rente!"
                                GoTo Clean_Exit
                            End If
                        End If
                        L_Result = Add_Pli_ePoBox_Emetteur_Detail(L_Pk_Pli_EpoBox_Emetteur, _
                                                                  CLng(L_Fk_Pli), _
                                                                  L_tab_ePoBox_Data(i).t_Field_Fk, _
                                                                  L_tab_ePoBox_Data(i).t_Field_Value)
                        If L_Result <> "Ok" Then
                            Fusion_fichier_Data = "Impossible d'enregistrer les informations d�taill�es relatives � l'ePoBox!"
                            GoTo Clean_Exit
                        End If
                        Rem If type = file then le d�placer
                    Next i
                Else
                    L_Pk_Pli_EpoBox_Emetteur = 0
                End If
                    
                If L_Archivage_Only And Not L_Validation_Contenu Then
                    Rem Mettre � jour le statut et le statut depart � 1
                    Call Run_Update_Sql(G_Adoconnection, "PLI", " STATUT_DEPART = 1, DATE_EXPEDITION = now()", " PK_PLI = " & L_Fk_Pli)
                End If
                
                Rem **********************************************************
                Rem A faire si conservation des donn�es coch�
                Rem **********************************************************
                'If L_Service_Conservation_Data Then
                '    'LCI SUPPRESSION CONSERVATION LE 14/08/2017
                '    Rem Remarque tres importante
                '    Rem plac�e ici cette requ�te est execut�e � chaque fois !!!
                '    Rem La sortir de la boucle !!!
                '    L_SQL = " SELECT CHAMP_EMIS_NOM,CHAMP_EMIS_A_CONSERVER_ORDRE "
                '    L_SQL = L_SQL & " FROM CHAMP_EMIS"
                '    L_SQL = L_SQL & " WHERE CHAMP_EMIS_A_CONSERVER = 1"
                '    L_SQL = L_SQL & " AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk
                '    L_SQL = L_SQL & " ORDER BY CHAMP_EMIS_A_CONSERVER_ORDRE "
                '    Set L_Recordset_Tmp = New ADODB.Recordset
                '    L_Recordset_Tmp.Open L_SQL, L_connexion
                '    If L_Recordset_Tmp.EOF Then
                '        L_Recordset_Tmp.Close
                '        Set L_Recordset_Tmp = Nothing
                '        Fusion_fichier_Data = "Aucun Champ �mis � conserver!!"
                '        GoTo Clean_Exit
                '    Else
                '        L_Recordset_Tmp.MoveFirst
                '    End If
                '    L_SQL = vbNullString
                '
                '    Rem V�rification que le H est cr��, le cr�e sinon
                '    If Not HEADER_FLAG Then
                '        HEADER_FLAG = True
                '        Rem R�initialise le tableau
                '        ReDim Nom_Champ(0)
                '        L_Index_Champ = 1
                '        While Not L_Recordset_Tmp.EOF
                '            Rem Alimente le tableau
                '            ReDim Preserve Nom_Champ(UBound(Nom_Champ) + 1)
                '            Nom_Champ(L_Index_Champ) = Valid_Text(L_Recordset_Tmp("CHAMP_EMIS_NOM"))
                '            L_Index_Champ = L_Index_Champ + 1
                '            L_Recordset_Tmp.MoveNext
                '        Wend
                '        NB_Champ = L_Index_Champ - 1
                '    End If
                '    Rem Lecture des donn�es
                '    L_Index_Champ = 1
                '
                '    L_Sql_Insert = " FK_PLI, RECORD, CHAMPS , MAJ_DATE"
                '    L_Sql_value = L_Fk_Pli & ", " & ENREG & ", "
                '
                '    l = vbNullString
                '    For L_Index_Champ = 1 To NB_Champ
                '        Rem Cas sp�ciaux - Le premier (pas dans le tableau!!!)
                '        If UCase$(Nom_Champ(L_Index_Champ)) = "REFERENCE_POSTEASY" Then
                '            l = l & Nom_Champ(L_Index_Champ) & "=" & L_ID_PE & "|"
                '        Else
                '            For L_Index2 = 0 To L_Nb_champs_emis_nom - 1
                '                If UCase$(Trim$(L_tab_Champs_Emis(L_Index2).t_champ_emis_nom)) = UCase$(Nom_Champ(L_Index_Champ)) Then
                '                    l = l & Nom_Champ(L_Index_Champ) & "=" & Valid_Text(Trim$(L_tab_Champs_Emis(L_Index2).t_champ_emis_data)) & "|"
                '                    Exit For
                '                End If
                '            Next L_Index2
                '        End If
                '    Next L_Index_Champ
                '    l = Left$(l, Len(l) - 1)
                '    L_Sql_value = L_Sql_value & "'" & l & "', now()"
                '    L_SQL = "INSERT INTO PLI_CONSERVATION (" & L_Sql_Insert & ")"
                '    L_SQL = L_SQL & " VALUES (" & L_Sql_value & ")"
                '    Call Run_Execute_Sql(L_SQL)
                '
                '    ENREG = ENREG + 1
                '    L_Recordset_Tmp.Close
                '    Set L_Recordset_Tmp = Nothing
                'End If
                
                If L_Nb_champs_lies > 0 And L_ModelType <> "Mod�le Word" Then
                    L_SQL = " FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk
                    L_SQL = L_SQL & " AND FIELD_TYPE <> '" & G_CONST_Jointure_Champ_Lie & "' "
                    ListeChamps = Lire_Une_Liste("champ_lie", "field_name", "PIPE", L_SQL, "")
                    TabChamps = Split(ListeChamps, "|")
                    For IndexCL = 0 To UBound(TabChamps)
                        If TabChamps(IndexCL) <> "" Then
                            Rem Lire la pk du champ li�
                            PkCL = Lire_Un_Champ("pk_champ_lie", "champ_lie", " field_name = '" & TabChamps(IndexCL) & "' AND FK_PRESTATION_MODEL = " & p_Prestation_Model_Pk)
                            If PkCL = "" Then
                                Fusion_fichier_Data = "Impossible de lire l'information sur le champ li� " & TabChamps(IndexCL) & " !"
                                GoTo Clean_Exit
                            End If
                            Rem Insert dans la table destinataire_index
                            L_SQL = " INSERT INTO DESTINATAIRE_INDEX(fk_destinataire, fk_champ_lie, valeur, maj_userid, maj_date) "
                            L_SQL = L_SQL & " SELECT " & L_Fk_destinataire & ", " & PkCL & ", " & TabChamps(IndexCL) & ", '" & G_User_Id & "', now() "
                            L_SQL = L_SQL & " FROM " & L_Table_Lies_Tmp
                            L_SQL = L_SQL & " WHERE CHAMP_LIE_JOINTURE = '" & Valid_Text(L_champ_lie_jointure_data) & "'"
                            L_SQL = L_SQL & " ORDER BY PK_CHAMPS_LIES_TMP"
                            L_Result = Run_Execute_Sql(L_SQL)
                            If L_Result = -1 Then
                                Fusion_fichier_Data = "Probl�me lors de l'insertion des index li�s (" & TabChamps(IndexCL) & ") !"
                                GoTo Clean_Exit
                            End If
                        End If
                    Next
                End If
                
                L_Nb_lignes_traitees = L_Nb_lignes_traitees + 1
                If Err.Number > 0 Then
                    Fusion_fichier_Data = Err.Description
                    GoTo Clean_Exit
                End If
                
            Else
                L_Sequence = L_Sequence + 1
            End If
        
        Rem Si la balise du fichier ne correspond pas � la balise attendue
        Else
            Rem Anomalie dans la structure du fichier
            If Existe_Balise_Ligne(UTF8_Decode(L_ligne), "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">") Then
                L_ligne = UTF8_Decode(L_ligne)
                GoTo DecodeUTF8
            End If
            If L_Sequence > 1 Then
                If Err.Number <> 0 Then
                    Fusion_fichier_Data = "Erreur " & Err.Number & " - " & Err.Description
                Else
                    Fusion_fichier_Data = "Format de donn�es non respect� (s�quence)!"
                End If
                GoTo Clean_Exit
            Rem ELSE
            Rem l'analyse du fichier n'a pas encore d�mar�e
            Rem ou est entre deux blocs!
            End If
            If Err.Number <> 0 Then
                Fusion_fichier_Data = "Erreur " & Err.Number & " - " & Err.Description
                GoTo Clean_Exit
            End If
        End If
        Call Display_Status("Etat d'avancement de la fusion:", L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter, p_Form)
        Call CheckApplication("Fusion processing " & L_Nb_lignes_traitees & " / " & p_Nb_lignes_a_traiter)
    Wend
    Rem Fermeture du fichier
    Close #L_Fnum
    
    L_List_Pk_Pli = Left(L_List_Pk_Pli, Len(L_List_Pk_Pli) - 1)

    If Not L_Lignes_lisibles Then
        Fusion_fichier_Data = "Fichier Data illisible (Anomalie Balises)"
        GoTo Clean_Exit
    End If
    
Rem *******************************************************************************************
Rem EPOBOX - ENVOI DIRECT SI PLATEFORME DESTINATAIRE + EMETTEUR
Rem *******************************************************************************************
    If L_Service_ePoBox And Not L_ePoBox_Hybrid Then
        Rem Ajouter un OF pour Ce client et Cette Prestation mod�le, pour ce jour s'il n'y en existe pas d�j� un
        Rem Le mettre en statut g�n�r� ET r�cup�rer sa FK et mettre � jour la quantit�
        Rem Ne s�lectionner � envoyer, que les plis de cette fusion!!!!
            Rem Attention, s�lection du statut "A affecter seulement!" Fusion_Robot.P_Fk_Pli_Statut_AAFFC
            Rem Pk de la fusion = p_preparation_fk
                L_SQL = " Select pk_pli, ID_PE, CHEMIN_PLI, FAX_ENVOYE "
        L_SQL = L_SQL & " from pli "
        L_SQL = L_SQL & " where fk_preparation = " & p_Preparation_Fk
        L_SQL = L_SQL & " and fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_AAFFC
        L_SQL = L_SQL & " and pk_pli in (" & L_List_Pk_Pli & ") "
        L_SQL = L_SQL & " and fk_of = 0"
        Call Init_Adodc(Fusion_Robot.Adodc_Tool)
        Fusion_Robot.Adodc_Tool.RecordSource = L_SQL
        Fusion_Robot.Adodc_Tool.Refresh
        If Not Fusion_Robot.Adodc_Tool.Recordset.EOF Then
            L_Pk_Of = Add_OF(CLng(p_Societe_Fk), p_Prestation_Model_Pk, p_Nb_lignes_a_traiter, "EPOBOX")

            Rem Faire la mise � jour des plis pour l'OF
                    L_SQL = " update pli set fk_of = " & L_Pk_Of
            L_SQL = L_SQL & " where fk_preparation = " & p_Preparation_Fk
            L_SQL = L_SQL & " and fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_AAFFC
            L_SQL = L_SQL & " and pk_pli in (" & L_List_Pk_Pli & ") "
            L_SQL = L_SQL & " and fk_of = 0"
            Call Run_Execute_Sql(L_SQL)

            Call ePobox_Send_Order(L_List_Pk_Pli)
            Call Display_Status("Envoi des flux ePoBox en cours.", p_Nb_lignes_a_traiter & " plis / " & p_Nb_lignes_a_traiter, p_Form)
            Rem Mettre � jour le nombre de plis li�s � l'of s�lectionn�
            Call Update_Pli_OF(L_Pk_Of)
        End If
    End If
Rem *******************************************************************************************
Rem EPOBOX - ENVOI DIRECT SI PLATEFORME DESTINATAIRE + EMETTEUR
Rem *******************************************************************************************


Rem *******************************************************************************************
Rem EPOBOX - ENVOI DIRECT SI FAX
Rem *******************************************************************************************
    If L_Service_Transformation_fax Then
        Rem Ajouter un OF pour Ce client et Cette Prestation mod�le, pour ce jour s'il n'y en existe pas d�j� un
        Rem Le mettre en statut g�n�r� ET r�cup�rer sa FK et mettre � jour la quantit�
        Rem Ne s�lectionner � envoyer, que les plis de cette fusion!!!!
        Rem Attention, s�lection du statut "A affecter seulement!" Fusion_Robot.P_Fk_Pli_Statut_AAFFC
        Rem Pk de la fusion = p_preparation_fk
        Rem Faire la mise � jour des plis pour l'OF
                L_SQL = " update pli inner join sent_fax on pli.pk_pli = sent_fax.document_id "
        L_SQL = L_SQL & " set pli.fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_ENPRD & ", sent_fax.status = 1 "
        L_SQL = L_SQL & " where pli.fk_preparation = " & p_Preparation_Fk
        Rem La mise � jour sur le statut du pli "A affecter" permet de ne traiter que ceux qui sont valid�s!!!
        L_SQL = L_SQL & " and pli.fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_AAFFC
        L_SQL = L_SQL & " and pli.pk_pli in (" & L_List_Pk_Pli & ") "
        Call Run_Execute_Sql(L_SQL)
    End If
    
Rem *******************************************************************************************
Rem FAX - ENVOI DIRECT
Rem *******************************************************************************************
Rem *******************************************************************************************
Rem SMS - ENVOI DIRECT SI SERVICE
Rem *******************************************************************************************
    If L_Service_Transformation_Sms And L_Envoi_Automatique Then
        Rem Ajouter un OF pour Ce client et Cette Prestation mod�le, pour ce jour s'il n'y en existe pas d�j� un
        Rem Le mettre en statut g�n�r� ET r�cup�rer sa FK et mettre � jour la quantit�
        Rem Ne s�lectionner � envoyer, que les plis de cette fusion!!!!
        L_Pk_Of = Add_OF(CLng(p_Societe_Fk), p_Prestation_Model_Pk, p_Nb_lignes_a_traiter, "SMS")

        Rem Faire la mise � jour des plis pour l'OF
                L_SQL = " update pli set fk_of = " & L_Pk_Of & ", "
        L_SQL = L_SQL & " fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_ENPRD
        L_SQL = L_SQL & " where fk_preparation = " & p_Preparation_Fk
        L_SQL = L_SQL & " and fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_AAFFC
        L_SQL = L_SQL & " and pk_pli in (" & L_List_Pk_Pli & ") "
        L_SQL = L_SQL & " and fk_of = 0"
        Call Run_Execute_Sql(L_SQL)

        Call Init_Adodc(Fusion_Robot.Adodc_Tool)
        Fusion_Robot.Adodc_Tool.RecordSource = SQL("TRAITEMENT_FLOW", CStr(L_Pk_Of), "SMS", " fk_preparation = " & p_Preparation_Fk & " AND pk_pli in (" & L_List_Pk_Pli & ") ", "", "")
        Fusion_Robot.Adodc_Tool.Refresh
        
        If Not Fusion_Robot.Adodc_Tool.Recordset.EOF Then
            Rem Tourner dans le recordset
            L_Result = Mail_Smtp_Send_New(Fusion_Robot.Adodc_Tool, _
                                    p_Prestation_Model_Pk, _
                                    L_Result, _
                                    p_Nb_lignes_a_traiter, _
                                    p_Form, _
                                    False, _
                                    "sms")
            Select Case L_Result
            Case "Ok"
                Call Display_Status("Envoi des Sms termin�.", p_Nb_lignes_a_traiter & " plis / " & p_Nb_lignes_a_traiter, p_Form)
            Case Else
                Rem Si erreur lors de l'envoi, les plis sont en statut en production
                Rem Mettre L'of en mode visible pour la production suffit � reprendre le flux
                L_SQL = " update of set statut_of = 'Erreur' where pk_of = " & L_Pk_Of
                Call Run_Execute_Sql(L_SQL)
            End Select
        End If
        Rem Mettre � jour le nombre de plis li�s � l'of s�lectionn�
        Call Update_Pli_OF(L_Pk_Of)
    End If
Rem *******************************************************************************************
Rem SMS - ENVOI DIRECT
Rem *******************************************************************************************


Rem *******************************************************************************************
Rem EMAIL - ENVOI DIRECT SI SERVICE
Rem *******************************************************************************************
    If L_Service_Transformation_Mail And L_Envoi_Automatique Then
        
        Rem Ajouter un OF pour Ce client et Cette Prestation mod�le, pour ce jour s'il n'y en existe pas d�j� un
        Rem Le mettre en statut g�n�r� ET r�cup�rer sa FK et mettre � jour la quantit�
        Rem Ne s�lectionner � envoyer, que les plis de cette fusion!!!!
        L_Pk_Of = Add_OF(CLng(p_Societe_Fk), p_Prestation_Model_Pk, p_Nb_lignes_a_traiter, "EMAIL")

        Rem Faire la mise � jour des plis pour l'OF
                L_SQL = " update pli set fk_of = " & L_Pk_Of & ", "
        L_SQL = L_SQL & " fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_ENPRD
        L_SQL = L_SQL & " where fk_preparation = " & p_Preparation_Fk
        L_SQL = L_SQL & " and fk_pli_statut = " & Fusion_Robot.P_Fk_Pli_Statut_AAFFC
        L_SQL = L_SQL & " and pk_pli in (" & L_List_Pk_Pli & ") "
        L_SQL = L_SQL & " and fk_of = 0"
        Call Run_Execute_Sql(L_SQL)

        Call Init_Adodc(Fusion_Robot.Adodc_Tool)
        Fusion_Robot.Adodc_Tool.RecordSource = SQL("TRAITEMENT_FLOW", CStr(L_Pk_Of), "EMAIL", " fk_preparation = " & p_Preparation_Fk & " AND pk_pli in (" & L_List_Pk_Pli & ") ", "", "")
        Fusion_Robot.Adodc_Tool.Refresh
        
        If Not Fusion_Robot.Adodc_Tool.Recordset.EOF Then
            Rem Tourner dans le recordset
            L_Result = Mail_Smtp_Send_New(Fusion_Robot.Adodc_Tool, _
                                    p_Prestation_Model_Pk, _
                                    L_Result, _
                                    p_Nb_lignes_a_traiter, _
                                    p_Form, _
                                    False, _
                                    "")
            Select Case L_Result
            Case "Ok"
                Call Display_Status("Envoi des e-mails termin�.", p_Nb_lignes_a_traiter & " plis / " & p_Nb_lignes_a_traiter, p_Form)
            Case Else
                Rem Si erreur lors de l'envoi, les plis sont en statut en production
                Rem Mettre L'of en mode visible pour la production suffit � reprendre le flux
                L_SQL = " update of set statut_of = 'Erreur' where pk_of = " & L_Pk_Of
                Call Run_Execute_Sql(L_SQL)
            End Select
        End If
        Rem Mettre � jour le nombre de plis li�s � l'of s�lectionn�
        Call Update_Pli_OF(L_Pk_Of)
    End If
Rem *******************************************************************************************
Rem EMAIL - ENVOI DIRECT
Rem *******************************************************************************************


Rem *******************************************************************************************
Rem SAE - ENVOI DIRECT dans le Coffre
Rem *******************************************************************************************
    If t2c_Status_pk_pli_list_go <> "" Then
        t2c_Status_pk_pli_list_go = Mid(t2c_Status_pk_pli_list_go, 2)
        L_SQL = " Update robot_t2c_injector set status = 'start' where status = 'wait' and fk_pli in (" & t2c_Status_pk_pli_list_go & ") "
        Call Run_Execute_Sql(L_SQL)
    End If

Rem *******************************************************************************************
Rem SAE - ENVOI DIRECT dans le Coffre
Rem *******************************************************************************************

Rem Etiquette sp�cial export group� (en attendant mieux)
    Fusion_fichier_Data = "Ok"
    
    Rem Fermeture du mod�le,
    Rem Dans le cas sans regroupement page !!!
    Rem ET Si fusion
    If L_ModelType = "Mod�le Word" Then
        If L_ModelName <> "" Then
            G_Wd.Documents(L_ModelName).Close savechanges:=wdDoNotSaveChanges
        End If
        Rem Fermeture du document et de Word
        If G_wd_opened Then
            Fusion_fichier_Data = Word_Close
        End If
    End If
    
Clean_Exit:

    Rem Fermeture du Fichier
    If L_Fnum > 0 Then
        Close #L_Fnum
    End If

    Rem Fermeture des connexions
    L_connexion.Close
    Set L_connexion = Nothing

Clean_Exit_Word:
    Rem Fermeture de Word
    If G_wd_opened Then
        Call Word_Close
    End If

    Rem Suppression de la table de champ li�s - A CONSERVER pour g�rer le cas NON WORD
    If L_Table_Lies_Tmp <> "" Then
        L_SQL = "DROP TABLE " & L_Table_Lies_Tmp
        p_Liste_Tables_Temporaires_a_Supprimer = vbNullString
        Call Run_Execute_Sql(L_SQL)
    End If
    
End Function

Public Sub Nettoyage_Fusion(p_what As String, _
                            ByVal p_preparation_pk As String, _
                            ByVal p_Customer_Number As String, _
                            ByRef p_Form As Form)

Dim L_liste_document_page   As String
Dim L_liste_document_pli    As String
Dim L_liste_pk_destinataire As String
Dim L_liste_pli_pk          As String
Dim L_Liste_pli_epobox_pk   As String

Dim L_SQL                   As String
Dim L_liste_document_WEB    As String

Dim L_Liste_document_chemin As String
Dim L_Prestation_Model_Pk   As Long
Dim L_Fk_Contact            As String
Dim L_Montant               As Double
Dim L_Montant2              As String


    L_Prestation_Model_Pk = Lire_Un_Champ("fk_prestation_model", "preparation", "pk_preparation = " & p_preparation_pk)
    
    Rem Crit�re de s�lection des plis = fk_preparation
    Call Display_Status(p_what & " de la fusion (analyse des plis d�j� g�n�r�s)...", "", p_Form)
    L_liste_document_page = Lire_Une_Liste("REGROUPEMENT_PAGE, PLI", "MODEL_WORD_FUSION_PAGE", "INFORMATION", "PK_PLI = FK_PLI AND FK_PREPARATION = " & p_preparation_pk)
    L_liste_document_pli = Lire_Une_Liste("PLI", "NOM_FICHIER_EMISSION", "INFORMATION", "FK_PREPARATION = " & p_preparation_pk)
    If L_liste_document_page <> "" Then
        Rem La liste se presente ainsi: nom1, NOM2, NOM3
        Rem La transformer en nom1.doc|nom2.doc|nom3.doc|
        L_liste_document_page = Replace(L_liste_document_page, ",", ".doc|", , , vbTextCompare) & ".doc|"
        L_liste_document_page = Replace(L_liste_document_page, " ", "")
        Call Display_Status(p_what & " de la fusion (suppression des documents d�j� g�n�r�s)...", "", p_Form)
        Call Supprimer_Les_Fichiers(G_dir_production & "\" & p_Customer_Number, L_liste_document_page, 0, 0, 0)
    End If
    If L_liste_document_pli <> "" Then
        Rem La liste se presente ainsi: nom1.doc, NOM2.doc, NOM3.doc
        Rem La transformer en nom1.doc|nom2.doc|
        L_liste_document_pli = Replace(L_liste_document_pli, ",", "|") & "|"
        L_liste_document_pli = Replace(L_liste_document_pli, " ", "")
        Call Display_Status(p_what & " de la fusion (suppression des documents de production d�j� g�n�r�s)...", "", p_Form)
        Call Supprimer_Les_Fichiers(G_dir_production & "\" & p_Customer_Number, L_liste_document_pli, 0, 0, 0)
    End If
    L_liste_document_WEB = Lire_Une_Liste("PLI", "concat(CHEMIN_PLI, '|',ID_PE)", "INFORMATION", "FK_PREPARATION = " & p_preparation_pk)
    If L_liste_document_WEB <> vbNullString Then
        L_liste_document_WEB = Replace(L_liste_document_WEB, "|", "\")
        L_liste_document_WEB = Replace(L_liste_document_WEB, ",  ", "*|")
        Call Display_Status(p_what & " de la fusion (suppression des documents WEB d�j� g�n�r�s)...", "", p_Form)
        Call Supprimer_Les_Fichiers(G_dir_Web_PDF, L_liste_document_WEB, 0, 0, 0)
    End If
Rem v43xx
    Call Display_Status(p_what & " de la fusion (mise � jour du syst�me)...", "", p_Form)
    Rem Liste des fk_destinataire de la table pli, des plis qui vont �tre supprim�s
    L_liste_pk_destinataire = Lire_Une_Liste("PLI", "FK_DESTINATAIRE", "LONG", "FK_PREPARATION = " & p_preparation_pk)
    Rem Suppression de tous les plis contenus dont les plis vont �tres supprim�s
    L_liste_pli_pk = Lire_Une_Liste("PLI", "PK_PLI", "LONG", "FK_PREPARATION = " & p_preparation_pk)
    If L_liste_pli_pk <> "" Then
        Rem Suppression du d�tail du pli � supprimer
        L_SQL = "DELETE FROM PLI_DETAIL WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM PLI_PRINT WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM PLI_PRIX WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM PLI_SIGNATURE WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM REGROUPEMENT_PAGE WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM sent_fax WHERE document_id IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM pli__fourniture WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM robot_t2c_injector WHERE FK_PLI IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        
    End If
    Rem Suppression de tous les plis g�n�r�s....
    L_SQL = "DELETE FROM PLI WHERE FK_PREPARATION =  " & p_preparation_pk
    Call Run_Execute_Sql(L_SQL)
    If L_liste_pk_destinataire <> "" Then
        Rem Liste de tous les destinataires des plis supprim�s, qui n'ont plus aucun pli
        L_liste_pk_destinataire = Lire_Une_Liste("DESTINATAIRE LEFT OUTER JOIN PLI ON PK_DESTINATAIRE = FK_DESTINATAIRE", "PK_DESTINATAIRE", "LONG", "PK_DESTINATAIRE IN (" & L_liste_pk_destinataire & ") AND PK_PLI IS NULL")
        Rem Supression de tous les destinataires de la liste qui ne sont plus rattach�s � aucun pli!!
        L_SQL = "DELETE FROM DESTINATAIRE WHERE PK_DESTINATAIRE IN (" & L_liste_pk_destinataire & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM creation_utilisateur_in WHERE FK_DESTINATAIRE IN (" & L_liste_pk_destinataire & ")"
        Call Run_Execute_Sql(L_SQL)
        
    End If
    If L_liste_pli_pk <> vbNullString Then
        L_Liste_pli_epobox_pk = Lire_Une_Liste("pli_epobox_emetteur", "pk_pli_epobox_emetteur", "LONG", "FK_PLI IN (" & L_liste_pli_pk & ")")
        If L_Liste_pli_epobox_pk <> vbNullString Then
            Rem Suppression du d�tail du pli � supprimer
            L_SQL = "DELETE FROM pli_epobox_emetteur_reponse WHERE fk_pli_epobox_emetteur  IN (" & L_Liste_pli_epobox_pk & ")"
            Call Run_Execute_Sql(L_SQL)
            Rem new
            L_SQL = "DELETE FROM pli_epobox_emetteur_detail WHERE fk_pli_epobox_emetteur IN (" & L_Liste_pli_epobox_pk & ")"
            Call Run_Execute_Sql(L_SQL)
            L_SQL = "DELETE FROM pli_epobox_emetteur WHERE pk_pli_epobox_emetteur IN (" & L_Liste_pli_epobox_pk & ")"
            Call Run_Execute_Sql(L_SQL)
        End If
    End If
    
Rem Cas epobosx IN
    L_liste_pli_pk = Lire_Une_Liste("pli_epobox_destinataire ", "pk_pli_epobox_destinataire", "INFORMATION", "FK_PREPARATION = " & p_preparation_pk)
    If L_liste_pli_pk <> vbNullString Then
        Rem Suppression du d�tail du pli � supprimer
        L_SQL = "DELETE FROM pli_epobox_destinataire_reponse WHERE fk_pli_epobox_destinataire  IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        Rem new
        L_SQL = "DELETE FROM pli_epobox_destinataire_detail WHERE fk_pli_epobox_destinataire IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        L_SQL = "DELETE FROM pli_epobox_destinataire WHERE pk_pli_epobox_destinataire IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
    End If

End Sub

Public Sub Nettoyage_Fusion_ePoBox(p_what As String, _
                                   ByVal p_preparation_pk As String, _
                                   ByVal p_Customer_Number As String, _
                                   ByRef p_Form As Form)

Dim L_liste_document_page   As String
Dim L_liste_document_pli    As String
Dim L_liste_pk_destinataire As String
Dim L_liste_pli_pk          As String
Dim L_SQL                   As String

    Rem Crit�re de s�lection des plis = fk_preparation
    Call Display_Status(p_what & " de la fusion (analyse des plis d�j� g�n�r�s)...", "", p_Form)
    Call Display_Status(p_what & " de la fusion (mise � jour du syst�me)...", "", p_Form)
    Rem Suppression de tous les plis contenus dont les plis vont �tres supprim�s
    L_liste_pli_pk = Lire_Une_Liste("PLI_EPOBOX_DESTINATAIRE", "PK_PLI_EPOBOX_DESTINATAIRE", "LONG", "FK_PREPARATION = " & p_preparation_pk)
    If L_liste_pli_pk <> "" Then
        Rem  ' d�gager + tard
        L_SQL = "DELETE FROM PLI_EPOBOX_DESTINATAIRE_DETAIL WHERE FK_PLI_EPOBOX_DESTINATAIRE IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        Rem Suppression du d�tail du pli � supprimer
        L_SQL = "DELETE FROM PLI_EPOBOX_DESTINATAIRE_REPONSE WHERE FK_PLI_EPOBOX_DESTINATAIRE IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
        Rem
        L_SQL = "DELETE FROM PLI_EPOBOX_DESTINATAIRE WHERE PK_PLI_EPOBOX_DESTINATAIRE IN (" & L_liste_pli_pk & ")"
        Call Run_Execute_Sql(L_SQL)
    End If

End Sub

Public Function Add_Destinataire(ByVal p_Nb_champs_emis_nom As Integer, _
                                 ByVal p_Prestation_Model_Pk As Long, _
                                 p_Pli_Statut_pk As Long, _
                                 p_Sql_Mail As String, _
                                 p_Champ_Lie_Jointure_Data As String, _
                                 ByVal p_Societe_Fk As Long, _
                                 ByVal p_Id_Pe As String, _
                                 ByRef p_Email_From As String, _
                                 ByRef p_Email_Xfer As String, _
                                 ByRef p_Fax_Sms_Number As String, _
                                 ByRef p_Sms_Message As String, _
                                 ByVal p_Premium As Boolean, _
                                 pDirSequestre As String, _
                                 pDirUnzip As String, _
                                 pCCFEIns As String, _
                                 pCCFESel As String, _
                                 ByRef pInsertPli As String, _
                                 ByRef pValuesPli As String) _
                                 As String


    Dim L_Index             As Integer
    Dim L_ePoBox_Index      As Integer
    Dim L_SQL               As String
    Dim L_Sql_Values_Mail   As String
    Dim L_Sql_Insert_mail   As String
    Dim L_Sql_Insert        As String
    Dim L_Sql_value         As String
    Dim Tmp                 As String
    
    
    On Error GoTo 0
    On Error GoTo Err_Add_Destinataire
    
    p_Champ_Lie_Jointure_Data = vbNullString
    
    L_Sql_Insert = vbNullString
    L_Sql_value = vbNullString
    L_Index = 0


    L_ePoBox_Index = 0
    p_Fax_Sms_Number = vbNullString
    p_Sms_Message = vbNullString


    L_Sql_Insert_mail = "INSERT INTO DESTINATAIRE_DETAIL("
    L_Sql_Insert_mail = L_Sql_Insert_mail & "fk_destinataire, type, data,maj_date,maj_userid)"
    L_Sql_Values_Mail = vbNullString
    
    Rem Lecture des donn�es
    For L_Index = 0 To p_Nb_champs_emis_nom - 1
        Select Case Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_type)
        
        Case "Civilit�"
            L_Sql_Insert = L_Sql_Insert & "DEST_CIVILITE, "
        
        Case "Nom"
            L_Sql_Insert = L_Sql_Insert & "DEST_NOM, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_nom, "
                pCCFESel = pCCFESel & " DEST_NOM, "
            End If
        
        Case "Pr�nom"
            L_Sql_Insert = L_Sql_Insert & "DEST_PRENOM, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_prenom, "
                pCCFESel = pCCFESel & " DEST_PRENOM, "
            End If
        
        Case "Adresse"
            L_Sql_Insert = L_Sql_Insert & "DEST_ADRESSE, "
        
        Case "Code Postal"
            L_Sql_Insert = L_Sql_Insert & "DEST_CP, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_cp, "
                pCCFESel = pCCFESel & " DEST_CP, "
            End If
            
        Case "Ville"
            L_Sql_Insert = L_Sql_Insert & "DEST_VILLE, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_ville, "
                pCCFESel = pCCFESel & " DEST_VILLE, "
            End If
            
        Case "Pays"
            L_Sql_Insert = L_Sql_Insert & "DEST_PAYS_NOM, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_pays, "
                If Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data) = "" Then
                    pCCFESel = pCCFESel & "'FRANCE', "
                Else
                    pCCFESel = pCCFESel & " DEST_PAYS_NOM, "
                End If
            End If
            
        Case "T�l�phone Domicile"
            L_Sql_Insert = L_Sql_Insert & "DEST_TEL_PERSO, "
            
        Case "T�l�phone Bureau"
            L_Sql_Insert = L_Sql_Insert & "DEST_TEL_BUREAU, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_tel, "
                pCCFESel = pCCFESel & " DEST_TEL_BUREAU, "
            End If
            
        Case "T�l�phone Portable"
            L_Sql_Insert = L_Sql_Insert & "DEST_TEL_GSM, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_gsm, "
                pCCFESel = pCCFESel & " DEST_TEL_GSM, "
            End If
            
        Case "Fax", "Sms"
            L_Sql_Insert = L_Sql_Insert & "DEST_FAX, "
            p_Fax_Sms_Number = Trim(Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, " ", "", 1, , vbTextCompare))
            
        Case "e-mail"
            L_Sql_Insert = L_Sql_Insert & "DEST_EMAIL, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " contact_email, "
                pCCFESel = pCCFESel & " DEST_EMAIL, "
            End If

        Case "Soci�t�"
            L_Sql_Insert = L_Sql_Insert & "DEST_SOCIETE, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " societe_nom, "
                pCCFESel = pCCFESel & " DEST_SOCIETE, "
            End If
        
        Case "R�f�rence Client"
            L_Sql_Insert = L_Sql_Insert & "DEST_REF_CLIENT, "
            
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " reference_unique, "
                pCCFESel = pCCFESel & " DEST_REF_CLIENT, "
            End If
            
        Case "R�f�rence Client2"
            L_Sql_Insert = L_Sql_Insert & "DEST_REF_CLIENT2, "
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " destinataire_unique, "
                pCCFESel = pCCFESel & " DEST_REF_CLIENT2, "
            End If
        
        Case "Date client"
            L_Sql_Insert = L_Sql_Insert & "DEST_DATE, "
            
        Case "Date exp�dition"
            pInsertPli = pInsertPli & "DATE_EXPEDITION, DATE_EMISSION, STATUT_DEPART, "
        
        Case "Num�ro de recommand�"
            pInsertPli = pInsertPli & "NUM_RA, "
        
        Case "Montant client"
            L_Sql_Insert = L_Sql_Insert & "DEST_AMOUNT, "
        
        Case "R�f�rence Comptable"
            L_Sql_Insert = L_Sql_Insert & "DEST_REF_COMPTABLE, " 'dest_ref_comptable
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " societe_siren, "
                pCCFESel = pCCFESel & " DEST_REF_COMPTABLE, "
            End If
            
        Case "Sc�nario"
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " fk_creation_utilisateur_type, "
                pCCFESel = pCCFESel & " " & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & ", "
            End If
            
        Case "R�f�rencement"
            If pCCFEIns <> "" Then
                pCCFEIns = pCCFEIns & " fk_societe_referencement, "
                pCCFESel = pCCFESel & " " & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & ", "
            End If
        
        Case "Date d�bouclement"
            L_Sql_Insert = L_Sql_Insert & "DEST_DATE_DEBOUCLEMENT, "
        
        Case "Dispatch_AR"
            If Len(Trim(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) <= 25 Then
                L_Sql_Insert = L_Sql_Insert & "DEST_DISPATCH_AR, "
            Else
                L_Sql_Insert = L_Sql_Insert & "DEST_DISPATCH_AR, R1, "
            End If
    
        Case "ePoBox", "R�f�rence ePoBox"
            L_Sql_Insert = L_Sql_Insert & "DEST_EPOBOX, "
        
        Case "Id Liste r�capitulative"
            Rem Nothing to do
        
        Case "Adresse mail de transfert"
            p_Email_Xfer = Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
            
        Case "Adresse mail de l'exp�diteur"
            p_Email_From = Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))

        Case "Statut dynamique"
            Rem Lire la fk du statut relatif � la valeur contenue
            p_Pli_Statut_pk = Fk_Pli_Statut_Dynamique(p_Prestation_Model_Pk, _
                                                    L_tab_Champs_Emis(L_Index).t_champ_emis_data)
        Rem remarque tres importante
        Rem Pour l'instant Une seule pi�ce jointe (ou r�pertori�e)
        Rem La notion de regroupement (pour les mails) n'a de sens que si
        Rem les pi�ces sont r�pertori�es ou jointes et unique
        Case "Sujet mail"               'sm
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'sm', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Sujet mail joint"         'smj
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'smj', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Sujet mail r�pertori�"    'smr
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'smr', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Sujet mail r�pertori� fusionn�"    'smf
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'smf', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Corps mail"               'cm
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'cm', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Corps mail joint", "Corps mail joint fusionn�"         'cmj
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'cmj', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Corps mail r�pertori�"    'cmr
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'cmr', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Corps mail fusionn�", "Corps mail r�pertori� fusionn�"    'cmf
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'cmf', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Pi�ce mail jointe"        'pmj
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'pmj', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Pi�ce mail r�pertori�e"   'pmr
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'pmr', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
        Case "Pi�ce mail fusionn�e"   'pmf
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'pmf', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail

        Case "Message Sms"              'sms
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'sms', '" & Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
            p_Sms_Message = Replace(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data), Chr(164), "�", , , vbBinaryCompare)

        
        Case "Message Sms Joint"              'sms
        
            
            If FileExists(pDirSequestre & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                p_Sms_Message = ReadTextFile(pDirSequestre, L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            ElseIf FileExists(pDirUnzip & L_tab_Champs_Emis(L_Index).t_champ_emis_data) Then
                p_Sms_Message = ReadTextFile(pDirUnzip, L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            Else
                p_Sms_Message = ""
            End If
            p_Sms_Message = Replace(p_Sms_Message, vbCr, "", , , vbBinaryCompare)
            p_Sms_Message = Replace(p_Sms_Message, vbLf, "", , , vbBinaryCompare)
            p_Sms_Message = Trim(p_Sms_Message)
            
            L_Sql_Values_Mail = " VALUES (fk_destinataire,'sms', '" & Valid_Text(p_Sms_Message) & "' , now(), '" & G_User_Id & "'" & ")|"
            p_Sql_Mail = p_Sql_Mail & L_Sql_Insert_mail & L_Sql_Values_Mail
            
            p_Sms_Message = Replace(p_Sms_Message, Chr(164), "�", , , vbBinaryCompare)

        Case "", _
             "Fichier joint", _
             "Fichier r�pertori�", _
             "Fichier r�pertori� facultatif", _
             "Fichier rapprochement facultatif", _
             "Fichier(s) multiple(s)", _
             "Fichier � transmettre", _
             "Fond de page r�pertori�", _
             "Nombre exemplaires", _
             "Nombre documents", _
             "t2c_Profil", _
             "t2c_Classement", _
             "t2c_IndexDocument", _
             "t2c_DateDocument", _
             "t2c_NomDocument", _
             "t2c_UserID"
             
            Rem "Impression Dynamique", "Impression dynamique", _ suppression le 14/08/2017 LCI
            
        End Select
        
        Select Case Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_type)
        Rem
        Case ""

        Case "t2c_Profil", _
             "t2c_Classement", _
             "t2c_IndexDocument", _
             "t2c_DateDocument", _
             "t2c_NomDocument", _
             "t2c_UserID"
            Rem Nothing

        Case "e-mail Copie (CC)"
        Case "e-mail Copie Cach�e (CCI)"
        Case "e-mail Reply-TO"
        
        Case "Fichier joint"
        Case "Fichier(s) multiple(s)"
        Case "Fichier � transmettre" 'R�serv� � ePoBox
        Case "Fond de page r�pertori�" '
        Case "Fichier r�pertori�"

        Case "Fichier r�pertori� facultatif"
        Case "Fichier rapprochement facultatif"
        Case "Nombre exemplaires", "Nombre documents"
        Case "Statut dynamique"
        'Case "Impression Dynamique", "Impression dynamique" 'LCI suppression le 14/08/2017
        Case "PM Origine"
        Case "Sujet mail"               'sm
        Case "Sujet mail joint"         'sml
        Case "Sujet mail r�pertori�"    'smr
        Case "Sujet mail r�pertori� fusionn�"                           'smf
        Case "Corps mail"                                               'cm
        Case "Corps mail joint", "Corps mail joint fusionn�"            'cmj
        Case "Corps mail r�pertori�"    'cmr
        Case "Corps mail fusionn�", "Corps mail r�pertori� fusionn�"    'cmf
        Case "Pi�ce mail jointe"        'pmj
        Case "Pi�ce mail r�pertori�e"   'pmr
        Case "Pi�ce mail fusionn�e"     'pmf
        Case "Adresse mail de l'exp�diteur"

        Case "Message Sms"              'sms
        Case "Message Sms Joint"        'sms
        
        Case "Date d�bouclement"    'JJ/MM/AAAA => Mysql YYYY-MM-DD
            If IsDate(Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) Then
                L_Sql_value = L_Sql_value & "'" & Format(CDate(Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))), G_FORMAT_DATE) & "', "
            Else
                L_Sql_value = L_Sql_value & "'" & Format(Date, G_FORMAT_DATE) & "', "
            End If
        
        Case "Date client"
            L_tab_Champs_Emis(L_Index).t_champ_emis_data = Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            If IsDate(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) Then
                L_Sql_value = L_Sql_value & "'" & Format(L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_FORMAT_DATE) & "', "
            Else
                L_Sql_value = L_Sql_value & " null, "
            End If
            
        Case "Date exp�dition"
            L_tab_Champs_Emis(L_Index).t_champ_emis_data = Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            If IsDate(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) Then
                pValuesPli = pValuesPli & "'" & Format(L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_FORMAT_SHORTDATE) & "', " '" & Format(L_tab_Champs_Emis(L_Index).t_champ_emis_data, G_FORMAT_SHORTDATE) & " 08:00:00', 1, "
            End If
            
        Case "Num�ro de recommand�"
            L_tab_Champs_Emis(L_Index).t_champ_emis_data = Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            If IsDate(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)) Then
                pValuesPli = pValuesPli & "'" & L_tab_Champs_Emis(L_Index).t_champ_emis_data & "', "
            End If
            
        Case "Montant client"
            L_tab_Champs_Emis(L_Index).t_champ_emis_data = Trim(Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, ".", ","))
            If IsNumeric(Valid_Text(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) Then
                L_Sql_value = L_Sql_value & Replace(Replace(L_tab_Champs_Emis(L_Index).t_champ_emis_data, ",", ".") & ", ", " ", "")
            Else
                L_Sql_value = L_Sql_value & " null, "
            End If
        
        Case "Jointure Champ Li�"
            p_Champ_Lie_Jointure_Data = Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            
        Case "Id Liste r�capitulative"
            Rem Nothing to do
            
        Case "Dispatch_AR"
            Tmp = Trim(Trim(L_tab_Champs_Emis(L_Index).t_champ_emis_data))
            If Len(Tmp) <= 25 Then
                L_Sql_value = L_Sql_value & "'" & Valid_Text(Tmp) & "', "
            Else
                L_Sql_value = L_Sql_value & "'" & Valid_Text(Left(Tmp, 25)) & "', '" & Valid_Text(Mid(Tmp, 26, 50)) & "', "
            End If
            
        Case "Soci�t�", "Civilit�", "Nom", "Pr�nom"
            If p_Premium Then
                L_tab_Champs_Emis(L_Index).t_champ_emis_data = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            End If
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) & "', "
            
        Case "Adresse"
            If p_Premium Then
                L_tab_Champs_Emis(L_Index).t_champ_emis_data = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            End If
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(CleanString(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data), vbTab, " "))) & "', "
        
        Case "Code Postal"
            L_tab_Champs_Emis(L_Index).t_champ_emis_data = OptiCp(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) & "', "
        
        Case "Ville"
            If p_Premium Then
                L_tab_Champs_Emis(L_Index).t_champ_emis_data = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            End If
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) & "', "
            
        Case "Pays"
            If p_Premium Then
                L_tab_Champs_Emis(L_Index).t_champ_emis_data = OptiAdresse(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
            End If
            
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) & "', "
            
        Case "Sc�nario", "R�f�rencement"
            'Nothing
        Case Else
        
            L_Sql_value = L_Sql_value & "'" & Valid_Text(Trim(Valid_Champ_Xml(L_tab_Champs_Emis(L_Index).t_champ_emis_data))) & "', "

        End Select
        
        If L_tab_Champs_Emis(L_Index).t_champ_emis_epobox = True Then
            L_tab_ePoBox_Data(L_ePoBox_Index).t_Field_Fk = L_tab_Champs_Emis(L_Index).t_champ_emis_pk
            L_tab_ePoBox_Data(L_ePoBox_Index).t_Field_Value = L_tab_Champs_Emis(L_Index).t_champ_emis_data
            Select Case L_tab_Champs_Emis(L_Index).t_champ_emis_type
            Case "Fichier � transmettre"
                Rem Type File
                L_tab_ePoBox_Data(L_ePoBox_Index).t_Field_Type = "File"
            Case Else
                Rem type data
                L_tab_ePoBox_Data(L_ePoBox_Index).t_Field_Type = "Data"
            End Select
            L_ePoBox_Index = L_ePoBox_Index + 1
        End If

    Next
    
    Rem ******************************************************************************************************************************************************
    Rem
    Rem Insertion du destinataire
    Rem
    Rem ******************************************************************************************************************************************************
    If InStr(1, L_Sql_Insert, "dest_cp", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_cp, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    If InStr(1, L_Sql_Insert, "dest_ville", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_ville, "
        L_Sql_value = L_Sql_value & "'', "
    End If

    If InStr(1, L_Sql_Insert, "dest_adresse", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_adresse, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    If InStr(1, L_Sql_Insert, "dest_pays_nom", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_pays_nom, "
        L_Sql_value = L_Sql_value & "'', "
    End If

    If InStr(1, L_Sql_Insert, "dest_email", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_email, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    If InStr(1, L_Sql_Insert, "dest_ref_client", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_ref_client, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    If InStr(1, L_Sql_Insert, "dest_date_debouclement", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_date_debouclement, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    If InStr(1, L_Sql_Insert, "dest_nom", vbTextCompare) = 0 Then
        L_Sql_Insert = L_Sql_Insert & " dest_nom, "
        L_Sql_value = L_Sql_value & "'', "
    End If
    
    Rem Suite de la liste des champs
    L_Sql_Insert = L_Sql_Insert & " FK_PRESTATION_MODEL, FK_SOCIETE, "
    L_Sql_Insert = L_Sql_Insert & " MAJ_DATE, MAJ_USERID, DEST_ID_PE"
    

    Rem Suite de la liste des valeurs
    L_Sql_value = L_Sql_value & p_Prestation_Model_Pk & ", " & p_Societe_Fk & ", "
    L_Sql_value = L_Sql_value & "now(), '" & G_User_Id & "', '" & p_Id_Pe & "'"

    Rem Ordre SQL
    L_SQL = "INSERT INTO DESTINATAIRE (" & L_Sql_Insert & ")"
    L_SQL = L_SQL & " VALUES (" & L_Sql_value & ")"
    If Run_Execute_Sql(L_SQL) = 0 Then
        Add_Destinataire = "Ok"
    Else
         Add_Destinataire = L_SQL
    End If
    
Exit Function

Err_Add_Destinataire:
    Add_Destinataire = Err.Number & " - " & Err.Description
    

End Function

Public Sub Add_Destinataire_detail(ByVal p_Fk_destinataire As Long, _
                                   ByVal p_Sql_Mail As String)

Dim L_pos               As String
Dim L_sql_to_execute    As String
Dim L_sql_mail          As String

L_sql_mail = Replace(p_Sql_Mail, "VALUES (fk_destinataire,", "VALUES (" & p_Fk_destinataire & ",")
L_pos = InStr(1, L_sql_mail, "|")
While L_pos > 0
    L_sql_to_execute = Left(L_sql_mail, L_pos - 1)
    L_sql_mail = Right(L_sql_mail, Len(L_sql_mail) - L_pos)
    Call Run_Execute_Sql(L_sql_to_execute)
    L_pos = InStr(1, L_sql_mail, "|")
Wend

End Sub

Public Function MergeB(f1 As String, f2 As String, f3 As String) As Long

    Dim Nf1, Nf3    As Long
    Dim L_ligne()         As Byte
    Dim RCLF(2)            As Byte
    Dim Reste As Long
    
RCLF(1) = 10
RCLF(2) = 11
    Nf3 = FreeFile
    Reste = FileLen(f1)
    Open f1 For Binary Access Write As Nf3
        Put Nf3, Reste + 1, 13
        Put Nf3, Reste + 2, 10
        'Close (Nf1)
        Nf1 = FreeFile
        Open f2 For Binary Access Read As Nf1
        'While Loc(Nf1) < LOF(Nf1)
            L_ligne = InputB(LOF(Nf1), Nf1)
            Put Nf3, Reste + 3, L_ligne
        'Wend
        Close (Nf1)
        Close (Nf3)
    
End Function

Public Function Fusion_fichier_Data_ePoBox(p_Dir_From As String, _
                                           p_Data_File_Name As String, _
                                           p_Prestation_Model_Pk As Long, _
                                           p_Cust_subfolder As String, _
                                           p_Societe_Fk As String, _
                                           p_Preparation_Fk As Long, _
                                           p_Liste_Fichiers_Joints_Reception As String, _
                                           p_liste_fichiers_joints_production_locale As String, _
                                           ByRef p_Form As Form, _
                                           ByVal p_Nombre_Data As Long, _
                                           ByVal pUnzipDir As String, _
                                           ByVal pFkFlow As Long) _
                                           As String


    Rem Declaration des variables locales
        Rem Fichier data en lecture
        Dim L_Fnum                                                  As Long
        Dim L_Fnom                                                  As String
        Dim L_ligne                                                 As String
        Dim L_Nb_champs_emis_detail                                 As Long
Rem v478
        Dim L_Referencement                                         As Boolean
Rem v478
        Dim L_Valeur_balise_ligne                                   As String
    Rem Chaine SQL
        Dim L_SQL                                                   As String
    Rem Chaine des champs pour Word
        Dim L_Word_fields                                           As String
    
    Rem Divers
        Dim L_Nb_champs_emis_nom                                    As Long
        Dim L_Index                                                 As Long
        Dim L_Chaine_header                                         As String
        Dim L_Chaine_data                                           As String
        Dim L_Bool                                                  As Boolean
        
Rem Suivi des plis
        Dim L_Nb_lignes_a_traiter                                   As Long
        Dim L_Nb_lignes_traitees                                    As Long
        Rem Num�ro d'enregistrement dans le fichier d'origine
        'Dim L_Num_Record                                            As Long
        Dim i                                                       As Long
        Dim L_Sequence                                              As Long
        Dim L_Balise                                                As String
        Dim L_chemin                                                As String
        
Rem // Variable multi-emploi
        Dim L_Result                                                As String
        
        Dim L_Notification                                          As Boolean
        Dim L_Root                                                  As String

        Dim Continuer                                               As Boolean
        

Rem // Initialisation

Rem Initialisation des champs li�s
    L_chemin = vbNullString
    
Rem suivi
    p_Liste_Fichiers_Joints_Reception = vbNullString
    L_Nb_lignes_a_traiter = p_Nombre_Data
    L_Nb_lignes_traitees = 0

Rem Liste des fichiers de donn�es
    
    p_Liste_Fichiers_Joints_Reception = vbNullString
    p_liste_fichiers_joints_production_locale = vbNullString
    
    Fusion_fichier_Data_ePoBox = "Erreur"
    If pFkFlow = 0 Then
        L_Fnom = G_dir_production & "\" & p_Cust_subfolder & "\" & p_Data_File_Name
    Else
        L_Fnom = Replace(pUnzipDir, "incoming", "fusion", , , vbTextCompare) & p_Data_File_Name
    End If

Rem Insert du destinataire ePoBox
    Dim L_Insert_Dest                               As String
    
    Dim L_Fk_destinataire                       As String
    Dim L_Emetteur_Nom                          As String
    Dim L_Emetteur_Reference                    As String
    Dim L_Emetteur_ePoBox_Site_Id               As String
    Dim L_Emetteur_ePoBox_Adresse               As String
    Dim L_Emetteur_Login                        As String
    Dim L_Emetteur_Validation_Date              As String
    Dim L_Emetteur_Tid_Magicaxess               As String
    Dim L_Destinataire_ePoBox_Site_Id           As String
    Dim L_Destinataire_ePoBox_Adresse           As String
    Dim L_Destinataire_Reference                As String
    Dim L_Destinataire_Reference2               As String
    Dim L_Destinataire_Reference_Comptable      As String
    Dim L_Destinataire_Date                     As String
    Dim L_Destinataire_Amount                   As String
    Dim L_ID_PE                                 As String
    Dim L_Chemin_Pli                            As String
    Dim L_Pdf_Document_Nom                      As String
    Dim L_Xml_Document_Nom                      As String
    Dim L_eMail                                 As String
    
    Dim L_DateS                                 As String
    Dim L_DateD                                 As Date
    
Rem Programm� � l'arrache, � refaire!!!
    Dim L_ePobox_Id_Liste_Recapitulative        As String


Rem
    Dim t2c_Profil                                              As String
    Dim t2c_IndexDocument                                       As String 'Index
    Dim t2c_NomDocument                                         As String 'Nom du doc dans le SAE
    Dim t2c_Classement                                          As String 'Classement
    Dim t2c_DateDocument                                        As String 'Date du doc
    Dim L_Service_robot_t2c                                     As Boolean
    Dim t2c_WebService_Actif                                    As Boolean
    Dim t2c_Status_pk_pli_list_go                               As String
    Dim t2c_UserID                                              As String

    Rem **************************************************************************************
    Rem Champs Emis
    Rem **************************************************************************************
    Call Init_Tab_Champs_Emis_Detail(L_tab_Champs_Emis, _
                                     L_tab_Champs_Emis_detail, _
                                     p_Prestation_Model_Pk, _
                                     L_Nb_champs_emis_detail, _
                                     L_Nb_champs_emis_nom, _
                                     L_Word_fields)
    

    L_Notification = Lire_Un_Champ("epobox_notification", "prestation_model", "pk_prestation_model = " & p_Prestation_Model_Pk)
Rem v596
    t2c_WebService_Actif = True
    t2c_Status_pk_pli_list_go = "" 'On laisse toujours � Wait, sauf � la fin de la fusion
    L_Service_robot_t2c = Existence_Service_Pe(p_Prestation_Model_Pk, G_CONST_SERVICE_T2C)
    'If L_Service_robot_t2c Then
    '    L_Result = Lire_Un_Champ("valeur", "sys_parameters", "code = 'TMP_FUS_T2C'")
    '    If L_Result = "MAIL" Then 'Seul cas ou l'on peut d�sactiver le WEB SERVICE
    '        t2c_WebService_Actif = False
    '    End If
    'End If
Rem v595
    
Rem v596


Rem **************************************************************************************
Rem Parcours du fichier de donn�es
Rem **************************************************************************************
    Call Display_Status("Lecture des donn�es...", "", p_Form)
    Rem Parcourir le fichier et lire le tableau
    L_Sequence = 0
Rem Securidad: L_Line_read
    Dim L_Line_read As Long
    L_Line_read = 0

    L_Fnum = FreeFile
    Open L_Fnom For Input As #L_Fnum
    While Not EOF(L_Fnum)
        Line Input #L_Fnum, L_ligne
        L_Line_read = L_Line_read + 1
        
        Rem Si la balise du fichier correspond � la balise attendue
        If Existe_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">") Then
            Rem Le num�ro de s�quence est valide!
            L_Bool = False
            L_Valeur_balise_ligne = Valeur_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">", L_Bool)
            While L_Bool
                Line Input #L_Fnum, L_ligne
                L_Line_read = L_Line_read + 1
                L_Valeur_balise_ligne = L_Valeur_balise_ligne & vbNewLine & Valeur_Balise_Ligne(L_ligne, "<" & L_tab_Champs_Emis(L_Sequence).t_champ_emis_nom & ">", L_Bool)
            Wend
Rem // hll - trim$(L_valeur_balise_ligne) ?
            L_tab_Champs_Emis(L_Sequence).t_champ_emis_data = L_Valeur_balise_ligne
            Rem *****************************
            Rem Incrementation de la s�quence
            Rem *****************************
            Rem En fin de s�quence
            Rem     et on fusionne
            Rem     et on red�marre
            
            If L_Sequence = L_Nb_champs_emis_nom - 1 Then
                Rem A ce stade, tous les champs d'une ligne sont lus
                Rem initialisation des variables
                L_Sequence = 0
                L_chemin = vbNullString
                Rem Et dans le tableau Champ_emis!!
                L_ID_PE = vbNullString
                                
                Rem Cas "Normal" de Fusion
                L_Index = 0
                Rem Lecture des champs et donn�es
                L_Chemin_Pli = vbNullString
                L_Chaine_data = vbNullString
Rem v458
                'L_Num_Record = 0
Rem v458

Rem v596
                t2c_Profil = ""
                t2c_IndexDocument = ""
                t2c_NomDocument = ""
                t2c_Classement = ""
                t2c_DateDocument = ""
                t2c_UserID = ""
Rem v596

                L_Chaine_header = vbNullString
                L_Chaine_header = L_Chaine_header & " INSERT INTO pli_epobox_destinataire_detail ("
                L_Chaine_header = L_Chaine_header & " fk_pli_epobox_destinataire, fk_champ_emis, valeur, maj_date, maj_userid)"
                
                L_eMail = vbNullString
                L_ePobox_Id_Liste_Recapitulative = vbNullString
                
                For L_Index = 0 To L_Nb_champs_emis_nom - 1
                
                    L_Chaine_data = L_Chaine_data & " VALUES(fk_pli_epobox_destinataire, " & L_tab_Champs_Emis(L_Index).t_champ_emis_pk & ", '" & Replace(Valid_Text(L_tab_Champs_Emis(L_Index).t_champ_emis_data), "|", ";", , , vbTextCompare) & "', now(), '" & G_User_Id & "')|"

                    If L_tab_Champs_Emis(L_Index).t_champ_emis_type = "e-mail" Then
Rem // v400 LCHLL
Rem v400 e-mail: politique diff�rente, ne pas lire cette valeur l�, mais lire la valeur dans societe/email_notification???
Rem L_Notification
                        L_eMail = L_tab_Champs_Emis(L_Index).t_champ_emis_data
Rem // v400 LCHLL
                    End If
                    If L_tab_Champs_Emis(L_Index).t_champ_emis_type = "Id Liste r�capitulative" Then
                        L_ePobox_Id_Liste_Recapitulative = L_tab_Champs_Emis(L_Index).t_champ_emis_data
                    End If
                    Rem v596
'stop
                    Select Case L_tab_Champs_Emis(L_Index).t_champ_emis_type
                    Case "t2c_Profil" 'Obligatoire si service t2c activ�
                        t2c_Profil = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    Case "t2c_Classement" 'Par d�faut, aucun
                        t2c_Classement = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    Case "t2c_IndexDocument" 'PAr d�faut Aucun
                        t2c_IndexDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    Case "t2c_DateDocument" 'Par D�faut Aucune
                        t2c_DateDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    Case "t2c_NomDocument" 'Par d�faut identique � celui g�n�r�
                        t2c_NomDocument = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    Case "t2c_UserID"      'A quoi cela sert=il?
                        t2c_UserID = Trim$(L_tab_Champs_Emis(L_Index).t_champ_emis_data)
                    End Select
                    Rem v596
                Next
Rem
                
                Line Input #L_Fnum, L_ligne
                
                If InStr(UCase(L_ligne), G_Balise_Data_Epobox_In) = 0 Then
                    Fusion_fichier_Data_ePoBox = "Les champs re�us d'une prestation mod�le ePoBox entrante doivent exactement correspondre au(x) champs �mis ET export�s (au sens epobox) de la prestation mod�le Sortante!"
                    Exit Function
                End If
                
                While InStr(UCase(L_ligne), G_Balise_Data_Epobox_In) > 0
                    Rem La balise <line_in> est pr�sente, tourner jusqu'a </line_in>
                    While InStr(UCase(L_ligne), G_Balise_Data_Epobox_Out) = 0
                        
                        Line Input #L_Fnum, L_ligne
                        Rem
                        L_Balise = Trim$(L_ligne)
                        If InStr(1, L_Balise, ">", vbTextCompare) > 0 Then
                            L_Balise = Mid(L_Balise, 1, InStr(1, L_Balise, ">", vbTextCompare))
                        End If
                        Continuer = (InStr(1, L_ligne, Replace(L_Balise, "<", "</", 1, 1, vbTextCompare)) = 0)
                        
                        Select Case LCase(L_Balise)
                        Case "<emetteur_nom>"
                            L_Emetteur_Nom = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<emetteur_reference>"
                            L_Emetteur_Reference = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<emetteur_epobox_site_id>"
                            L_Emetteur_ePoBox_Site_Id = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<emetteur_epobox_adresse>"
                            L_Emetteur_ePoBox_Adresse = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<emetteur_login>"
                            L_Emetteur_Login = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
Rem V495 - Fusion - Ajout
                        Case "<emetteur_validation_date>"
                            'If L_ligne <> "<emetteur_validation_date></emetteur_validation_date>" Then
                                'stop
                            'End If
                            L_Emetteur_Validation_Date = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<emetteur_tid_magicaxess>"
                            L_Emetteur_Tid_Magicaxess = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
Rem V495 - Fusion - Ajout (fin)
                            
                        Case "<destinataire_epobox_site_id>"
                            L_Destinataire_ePoBox_Site_Id = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<destinataire_epobox_adresse>"
                            L_Destinataire_ePoBox_Adresse = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<destinataire_reference>"
                            L_Destinataire_Reference = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
Rem V495 - Fusion - Ajout
                        Case "<destinataire_reference2>"
                            L_Destinataire_Reference2 = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<destinataire_reference_comptable>"
                            L_Destinataire_Reference_Comptable = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                            If Continuer Then
                                Line Input #L_Fnum, L_ligne
                                L_Destinataire_Reference_Comptable = L_Destinataire_Reference_Comptable & Valeur_Balise_Ligne(L_ligne, L_Balise, True)
                            End If
                            
                        Case "<destinataire_date>"
                            'If L_ligne <> "<destinataire_date></destinataire_date>" Then
                            '    stop
                            'End If
                            L_Destinataire_Date = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<destinataire_amount>"
                            'If L_ligne <> "<destinataire_amount>0</destinataire_amount>" Then
                            '    stop
                            'End If
                            L_Destinataire_Amount = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
Rem V495 - Fusion - Ajout (Fin)
                        Case "<id_pe>"
                            L_ID_PE = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                        Case "<pdf_document_nom>"
                            L_Pdf_Document_Nom = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
                            
                            If Len(L_Pdf_Document_Nom) > 40 Then
                                If L_Root = "" Then
                                    L_Root = Left(L_Pdf_Document_Nom, 29)
                                Else
                                    If L_Root <> Left(L_Pdf_Document_Nom, 29) Then
                                        L_Root = Left(L_Pdf_Document_Nom, 29)
                                    End If
                                End If
                            Else
                                L_Root = ""
                            End If
                            
                            Rem Controle de la pr�sence du pdf, sinon recherche dans les r�pertoires de rangement
                            If pFkFlow = 0 Then
                                If pUnzipDir <> "" Then
                                    If Not FileExists(G_dir_reception & "\" & L_Pdf_Document_Nom) Then
                                        If FileExists(pUnzipDir & L_Pdf_Document_Nom) Then
                                            If Copy_File(pUnzipDir, G_dir_reception, L_Pdf_Document_Nom) = True Then
                                                GoTo PdfGo
                                            End If
                                        End If
                                    End If
                                End If
                                If Not FileExists(G_dir_reception & "\" & L_Pdf_Document_Nom) Then
                                    For i = 0 To 20
                                        If FileExists(G_dir_reception & "\" & p_Cust_subfolder & "\" & Format(Date - i, "YYYYMMDD") & "\" & L_Pdf_Document_Nom) Then
                                            If Copy_File(G_dir_reception & "\" & p_Cust_subfolder & "\" & Format(Date - i, "YYYYMMDD"), G_dir_reception, L_Pdf_Document_Nom) = True Then
                                                i = 30
                                            Else
                                                Fusion_fichier_Data_ePoBox = "Le fichier " & L_Pdf_Document_Nom & " introuvable !"
                                                Exit Function
                                            End If
                                        End If
                                        If i = 20 Then
                                            Fusion_fichier_Data_ePoBox = "Le fichier " & L_Pdf_Document_Nom & " introuvable !"
                                            Exit Function
                                        End If
                                    Next i
                                End If
                            End If 'Si le flox est renseign�, nous savons d�j� que les fichiers sont dans incoming !!!!
PdfGo:
                            Call Pdf_ePoBox_Dir(L_Pdf_Document_Nom, L_Chemin_Pli, p_Cust_subfolder, True, True, L_Root, pFkFlow, pUnzipDir)
                            If L_Chemin_Pli = vbNullString Then
                                Fusion_fichier_Data_ePoBox = "Impossible de copier le fichier: " & L_Pdf_Document_Nom
                                Exit Function
                            Else
                                p_Liste_Fichiers_Joints_Reception = p_Liste_Fichiers_Joints_Reception & L_Pdf_Document_Nom & "|"
                            End If
                        Case "<xml_document_nom>"
                            L_Xml_Document_Nom = Valeur_Balise_Ligne(L_ligne, L_Balise, False)
Xml_Document_Nom:
                            If L_Xml_Document_Nom <> vbNullString Then
                                Rem Controle de la pr�sence du pdf, sinon recherche dans les r�pertoires de rangement
                                If Len(L_Xml_Document_Nom) > 40 Then
                                    If L_Root = "" Then
                                        L_Root = Left(L_Xml_Document_Nom, 29)
                                    Else
                                        If L_Root <> Left(L_Xml_Document_Nom, 29) Then
                                            L_Root = Left(L_Xml_Document_Nom, 29)
                                        End If
                                    End If
                                Else
                                    L_Root = ""
                                End If
                                If Len(L_Xml_Document_Nom) = 29 Then
                                    Rem Controle que ce n'est pas une anomalie de d�p�t
                                    If L_Root <> "" Then
                                        If L_Root = L_Xml_Document_Nom Then
                                            L_Xml_Document_Nom = ""
                                            GoTo Xml_Document_Nom
                                        End If
                                    End If
                                    If L_Pdf_Document_Nom <> "" Then
                                        If Left(L_Pdf_Document_Nom, 29) = L_Xml_Document_Nom Then
                                            L_Xml_Document_Nom = ""
                                            GoTo Xml_Document_Nom
                                        End If
                                    End If
                                End If
                                
                                If pFkFlow = 0 Then
                                    If pUnzipDir <> "" Then
                                        If Not FileExists(G_dir_reception & "\" & L_Xml_Document_Nom) Then
                                            If FileExists(pUnzipDir & L_Xml_Document_Nom) Then
                                                If Copy_File(pUnzipDir, G_dir_reception, L_Xml_Document_Nom) = True Then
                                                    GoTo XmlGo
                                                End If
                                            End If
                                        End If
                                    End If
                                    If Not FileExists(G_dir_reception & "\" & L_Xml_Document_Nom) Then
                                        For i = 0 To 20
                                            If FileExists(G_dir_reception & "\" & p_Cust_subfolder & "\" & Format(Date - i, "YYYYMMDD") & "\" & L_Xml_Document_Nom) Then
                                                If Copy_File(G_dir_reception & "\" & p_Cust_subfolder & "\" & Format(Date - i, "YYYYMMDD"), G_dir_reception, L_Xml_Document_Nom) = True Then
                                                    i = 30
                                                Else
                                                    Fusion_fichier_Data_ePoBox = "Le fichier " & L_Xml_Document_Nom & " introuvable !"
                                                    Exit Function
                                                End If
                                            End If
                                            If i = 20 Then
                                                Fusion_fichier_Data_ePoBox = "Le fichier " & L_Xml_Document_Nom & " introuvable !"
                                                Exit Function
                                            End If
                                        Next i
                                    End If
                                End If
XmlGo:
                                Call Pdf_ePoBox_Dir(L_Xml_Document_Nom, L_Chemin_Pli, p_Cust_subfolder, True, True, L_Root, pFkFlow, pUnzipDir)
                                If L_Chemin_Pli = vbNullString Then
                                    Fusion_fichier_Data_ePoBox = "Impossible de copier le fichier: " & L_Pdf_Document_Nom
                                    Exit Function
                                Else
                                    p_Liste_Fichiers_Joints_Reception = p_Liste_Fichiers_Joints_Reception & L_Xml_Document_Nom & "|"
                                End If
                            End If
                        Case "</epobox_in>"
                            Rem Nothing to do
                        Case Else
                            Fusion_fichier_Data_ePoBox = "Champ ePoBox r�serv� inconnu! (" & L_Balise & ")"
                            Exit Function
                        End Select
                        L_Line_read = L_Line_read + 1
                    Wend
                    Rem Si je sors de la boucle, charger la nouvelle ligne � analyser
                    Line Input #L_Fnum, L_ligne
                    L_Line_read = L_Line_read + 1
                Wend
                Rem Here, all data are read, we can add ePoBox Destinataire
                
                                L_Insert_Dest = " INSERT INTO PLI_EPOBOX_DESTINATAIRE ("
                L_Insert_Dest = L_Insert_Dest & " emetteur_nom, emetteur_reference, "
                L_Insert_Dest = L_Insert_Dest & " emetteur_epobox_site_id, emetteur_epobox_adresse, "
                L_Insert_Dest = L_Insert_Dest & " emetteur_login, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_epobox_site_id, destinataire_epobox_adresse, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_reference, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_reference2, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_reference_comptable, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_date, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_amount, "
                L_Insert_Dest = L_Insert_Dest & " emetteur_validation_date, "
                L_Insert_Dest = L_Insert_Dest & " emetteur_tid_magicaxess, "
                L_Insert_Dest = L_Insert_Dest & " destinataire_prestation_model_nom, "
                L_Insert_Dest = L_Insert_Dest & " id_pe, fk_preparation, "
                L_Insert_Dest = L_Insert_Dest & " document_visualisable, document_transmis, "
                L_Insert_Dest = L_Insert_Dest & " date_last_download, date_last_response, "
                L_Insert_Dest = L_Insert_Dest & " document_chemin, date_reception, "
                L_Insert_Dest = L_Insert_Dest & " fk_pli_epobox_statut, fk_prestation_model, "

                If L_ePobox_Id_Liste_Recapitulative <> "" Then
                    L_Insert_Dest = L_Insert_Dest & " fk_recap_list_completed, "
                End If
                L_Insert_Dest = L_Insert_Dest & " date_last_event, maj_userid, maj_date)"
                
                L_SQL = vbNullString
                L_SQL = L_SQL & " VALUES ("
                L_SQL = L_SQL & "'" & Valid_Text(L_Emetteur_Nom) & "', '" & Valid_Text(L_Emetteur_Reference) & "', "
                L_SQL = L_SQL & "'" & Valid_Text(L_Emetteur_ePoBox_Site_Id) & "', '" & Valid_Text(L_Emetteur_ePoBox_Adresse) & "', "
                L_SQL = L_SQL & "'" & Valid_Text(L_Emetteur_Login) & "', "
                L_SQL = L_SQL & "'" & Valid_Text(L_Destinataire_ePoBox_Site_Id) & "', '" & Valid_Text(L_Destinataire_ePoBox_Adresse) & "', "
                L_SQL = L_SQL & "'" & Valid_Text(L_Destinataire_Reference) & "', "

                L_SQL = L_SQL & "'" & Valid_Text(L_Destinataire_Reference2) & "', "
                L_SQL = L_SQL & "'" & Valid_Text(L_Destinataire_Reference_Comptable) & "', "
                If Trim(L_Destinataire_Date) = "01/01/0001 00:00:00" Then
                    L_Destinataire_Date = ""
                End If
                If Trim(L_Destinataire_Date) <> "" Then
                    L_SQL = L_SQL & "'" & Valid_Text(Format(L_Destinataire_Date, "YYYY-MM-DD HH:mm:SS")) & "', "
                Else
                    L_SQL = L_SQL & "'', "
                End If
                L_SQL = L_SQL & "'" & Valid_Text(Replace(L_Destinataire_Amount, ",", ".")) & "', "
                If Trim(L_Emetteur_Validation_Date) = "0001-01-01T00:00:00" Then
                    L_Emetteur_Validation_Date = ""
                End If
                If L_Emetteur_Validation_Date = "" Then
                    L_SQL = L_SQL & "'', "
                Else
                    L_SQL = L_SQL & "'" & Valid_Text(Format(L_Emetteur_Validation_Date, "YYYY-MM-DD HH:mm:SS")) & "', "
                End If
                L_SQL = L_SQL & "'" & Valid_Text(L_Emetteur_Tid_Magicaxess) & "', "
                
Rem v495 - Fusion - Ajout (fin)
                L_SQL = L_SQL & "'?', "
                L_SQL = L_SQL & "'" & Valid_Text(L_ID_PE) & "', " & p_Preparation_Fk & ", "
                L_SQL = L_SQL & "'" & Valid_Text(Replace(L_Pdf_Document_Nom, L_Root, "")) & "', '" & Valid_Text(Replace(L_Xml_Document_Nom, L_Root, "")) & "', "
                L_SQL = L_SQL & " NULL, NULL, "
                L_SQL = L_SQL & "'" & Valid_Text(L_Chemin_Pli) & "', now(), "
                L_SQL = L_SQL & Fk_Pli_Statut_ePoBox("OP") & ", " & p_Prestation_Model_Pk & ", "
                If L_ePobox_Id_Liste_Recapitulative <> "" Then
                    L_SQL = L_SQL & L_ePobox_Id_Liste_Recapitulative & ", "
                End If
                L_SQL = L_SQL & "now(), '" & Valid_Text(G_User_Id) & "', now()"
                L_SQL = L_SQL & ")"
                
                'Call Executer_un_ordre_SQL(L_Insert_Dest & L_SQL, True)
                If Run_Execute_Sql(L_Insert_Dest & L_SQL) = -1 Then
                    If Err.Number = 1062 Then
                        Err.Clear
                        If Run_Execute_Sql("UPDATE PLI_EPOBOX_DESTINATAIRE SET FK_PREPARATION = " & p_Preparation_Fk & ", maj_userid = '" & Valid_Text(G_User_Id) & "', maj_date = now() WHERE ID_PE = '" & L_ID_PE & "'") < 0 Then
                            Fusion_fichier_Data_ePoBox = "Erreur d'execution de la requ�te suivante : UPDATE PLI_EPOBOX_DESTINATAIRE SET FK_PREPARATION = " & p_Preparation_Fk & ", maj_userid = '" & Valid_Text(G_User_Id) & "', maj_date = now() WHERE ID_PE = '" & L_ID_PE & "'"
                            Close #L_Fnum
                            Exit Function
                        End If
                    Else
                        Fusion_fichier_Data_ePoBox = "Erreur d'execution de la requ�te suivante : " & L_Insert_Dest & L_SQL
                        Close #L_Fnum
                        Exit Function
                    End If
                End If
                Rem Le record �tait d�j� l�
                
                Rem Lecture de la PK du destinataire
                L_Fk_destinataire = Lire_Un_Champ("PK_PLI_EPOBOX_DESTINATAIRE", "PLI_EPOBOX_DESTINATAIRE", "ID_PE = '" & L_ID_PE & "'")
                
                L_Chaine_data = Replace(L_Chaine_data, "fk_pli_epobox_destinataire", L_Fk_destinataire, , , vbTextCompare)
                
                Dim L_array() As String
                L_array = Split(L_Chaine_data, "|", , vbTextCompare)
                For i = 0 To UBound(L_array)
                    If L_array(i) <> "" Then
                        If Run_Execute_Sql(L_Chaine_header & L_array(i)) = -1 Then
                            If Err.Number = 1062 Then
                                Err.Clear
                            Else
                                Fusion_fichier_Data_ePoBox = "Erreur d'execution de la requ�te suivante : " & L_Chaine_header & L_array(i)
                                Close #L_Fnum
                                Exit Function
                            End If
                        End If
                    End If
                Next
                
                If L_Service_robot_t2c Then
                    If t2c_NomDocument = "" Then
                        t2c_NomDocument = Valid_Text(Replace(L_Pdf_Document_Nom, L_Root, ""))
                    Else
                        Rem Si le type est diff�rent de celui qui est g�n�r�, on donne le m�me
                        If FileType(t2c_NomDocument) <> FileType(L_Pdf_Document_Nom) Then
                            t2c_NomDocument = t2c_NomDocument & "." & FileType(L_Pdf_Document_Nom)
                        End If
                    End If
                    L_Result = t2c_Injector(0, CLng(L_Fk_destinataire), CLng(p_Societe_Fk), G_dir_Web_PDF & "\" & L_Chemin_Pli & "\" & Replace(L_Pdf_Document_Nom, L_Root, ""), _
                                          t2c_Profil, t2c_NomDocument, t2c_DateDocument, t2c_UserID, t2c_IndexDocument, t2c_Classement, "wait")
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data_ePoBox = L_Result
                        Close #L_Fnum
                        Exit Function
                    End If
                    t2c_Status_pk_pli_list_go = t2c_Status_pk_pli_list_go & "," & L_Fk_destinataire
                End If
                
                L_Nb_lignes_traitees = L_Nb_lignes_traitees + 1
                If Err.Number > 0 Then
                    Fusion_fichier_Data_ePoBox = Err.Description
                    Close #L_Fnum
                    Exit Function
                End If
                
                If Existence_Service_Pe(CLng(p_Prestation_Model_Pk), G_CONST_SERVICE_WORM) Then
                    
                    L_DateS = Left(L_Chemin_Pli, 10)
                    L_DateD = Right(L_DateS, 2) & "/" & Mid(L_DateS, 6, 2) & "/" & Left(L_DateS, 4)
                    L_Result = worm__create(CLng(p_Societe_Fk), "IN", L_DateD, L_Chemin_Pli, Replace(L_Pdf_Document_Nom, L_Root, ""))
                    Err.Clear
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data_ePoBox = "Erreur Worm entrant : " & L_Result
                        Exit Function
                    End If
                End If
                
                Rem envoi d'email de notification
                Select Case p_Cust_subfolder
                Case "033411000517"
                    Call Mail_Notification_Send(G_Sae_Demo_Mail_To, "robot-fusion@axessy.fr", p_Cust_subfolder, "[AMSYNDIC-" & p_Cust_subfolder & "]", "[AMSYNDIC-" & p_Cust_subfolder & "]", L_ID_PE, G_dir_reception & "\" & L_Pdf_Document_Nom, "exploitation@axessy.fr")
                Case "033411005732"
                    Call Mail_Notification_Send(G_Sae_Demo_Mail_To, "robot-fusion@axessy.fr", p_Cust_subfolder, "[AMSYNDIC-" & p_Cust_subfolder & "]", "[AMSYNDIC-" & p_Cust_subfolder & "]", L_ID_PE, G_dir_reception & "\" & L_Pdf_Document_Nom, "exploitation@axessy.fr")
                End Select
                
                If L_Notification Then
                    Rem � finir/tester
                    If CLng(Lire_Un_Champ("num_referencement", "societe", "num_societe = '" & p_Cust_subfolder & "'")) = 235 Then
                        L_Referencement = True
                    Else
                        L_Referencement = False
                    End If
                    L_eMail = Lire_Un_Champ("epobox_email_notification", "societe", "num_societe = '" & p_Cust_subfolder & "'")
                    L_Result = Mail_Creation_Notification(L_ID_PE, _
                                                          L_eMail, _
                                                          "", _
                                                          p_Cust_subfolder, L_Chemin_Pli, L_Referencement, p_Prestation_Model_Pk)
                    If L_Result <> "Ok" Then
                        Fusion_fichier_Data_ePoBox = L_Result
                        Exit Function
                    End If
                    
                End If
                
                
                Rem A finir
                L_eMail = vbNullString
                
            Else
                L_Sequence = L_Sequence + 1
            End If
        
        Rem Si la balise du fichier ne correspond pas � la balise attendue
        Else
            Rem Anomalie dans la structure du fichier
            If L_Sequence > 1 Then
                If Err.Number <> 0 Then
                    Fusion_fichier_Data_ePoBox = "Erreur " & Err.Number & " - " & Err.Description
                Else
                    Fusion_fichier_Data_ePoBox = "Format de donn�es non respect� (s�quence)!"
                End If
                Close #L_Fnum
                Exit Function
            Rem ELSE
            Rem l'analyse du fichier n'a pas encore d�mar�e
            Rem ou est entre deux blocs!
            End If
            If Err.Number <> 0 Then
                If Err.Number = -2147217900 Then
                    Err.Clear
                Else
                    Fusion_fichier_Data_ePoBox = "Erreur " & Err.Number & " - " & Err.Description
                    Close #L_Fnum
                    Exit Function
                End If
            End If
        End If
        Call Display_Status("Etat d'avancement de la fusion:", L_Nb_lignes_traitees & " / " & L_Nb_lignes_a_traiter, p_Form)
    Wend
    Rem Fermeture du fichier
    Close #L_Fnum
    
    
    Rem v596
    Rem *******************************************************************************************
    Rem SAE - ENVOI DIRECT dans le Coffre
    Rem *******************************************************************************************
    If t2c_Status_pk_pli_list_go <> "" Then
        t2c_Status_pk_pli_list_go = Mid(t2c_Status_pk_pli_list_go, 2)
        L_SQL = " Update robot_t2c_injector set status = 'start' where status = 'wait' and fk_pli_epobox_destinataire in (" & t2c_Status_pk_pli_list_go & ") "
        Call Run_Execute_Sql(L_SQL)
    End If
    Rem *******************************************************************************************
    Rem SAE - ENVOI DIRECT dans le Coffre
    Rem *******************************************************************************************
    Rem v596 - fin
    
    Fusion_fichier_Data_ePoBox = "Ok"
    
Exit Function

sql_Error:
    G_msg_error_p1 = Err.Number
    G_msg_error_p2 = Err.Description
    G_msg_error_p3 = "Regles_gestion"
    G_msg_error_p4 = "Fusion_fichier_Data_ePoBox"
    G_msg_error_p5 = vbNullString
    G_msg_error_p6 = vbNullString
    error_manager.Show vbModal
    Close #L_Fnum
    End
End Function

Private Function Load_Print_Properties(p_fk_pli As Long) As String

    Dim L_Index     As Long
    Dim L_Insert    As String
    Dim L_Values    As String
    
    L_Insert = "fk_pli, page_from, page_to, page_type, paper, maj_date, maj_userid"
    
    If UBound(G_Print_Properties) = 0 Then
        Load_Print_Properties = "Aucune information d'impression n'est renseign�e ()"
        Exit Function
    End If
    
    For L_Index = 0 To UBound(G_Print_Properties) - 1
        L_Values = p_fk_pli & "," & G_Print_Properties(L_Index).t_Page_From & ", " & G_Print_Properties(L_Index).t_Page_To & ", '" & G_Print_Properties(L_Index).t_Page_Type & "', '" & G_Print_Properties(L_Index).t_Page_Paper & "', now(), '" & G_User_Id & "'"
        If Run_Insert_Sql(G_Adoconnection, "PLI_PRINT", L_Insert, L_Values) <> 0 Then
            Load_Print_Properties = "Probl�me d'enregistrement des param�tres d'impression"
            Exit Function
        End If
    Next
    
    Load_Print_Properties = "Ok"
    
End Function

'Private Function Load_Invoice_Data(p_ligne As String, ByRef p_Tab_Epobox_Facture As T_ePoBox_Facture)

'
    

'End Function

'Private Function Init_T_ePoBox_Facture(p_Tab_Epobox_Facture() As T_ePoBox_Facture) As Boolean

'    On Error GoTo Err_Init_T_ePoBox_Facture
'    ReDim p_Tab_Epobox_Facture(0)
'    Dim L_Tab_List()    As String
'    Dim i               As Integer
    

'    L_Tab_List = Split(G_Liste_Balise_Epobox_facture, ";")
'    ReDim p_Tab_Epobox_Facture(0)
'    For i = 0 To UBound(L_Tab_List)
'        ReDim Preserve p_Tab_Epobox_Facture(UBound(p_Tab_Epobox_Facture) + 1)
'        p_Tab_Epobox_Facture(UBound(p_Tab_Epobox_Facture)).field_name = L_Tab_List(i)
'        p_Tab_Epobox_Facture(UBound(p_Tab_Epobox_Facture)).field_value = ""
'    Next
    
'    Init_T_ePoBox_Facture = True
'    Exit Function
    
'Err_Init_T_ePoBox_Facture:
'    Init_T_ePoBox_Facture = False
    
'End Function

'Private Function Write_T_ePoBox_Facture(p_Tab_Epobox_Facture() As T_ePoBox_Facture, ByVal p_Line As String) As Boolean

'    On Error GoTo Err_Write_T_ePoBox_Facture
'    Dim L_Balise    As String
'    Dim L_valeur    As String
'    Dim L_array()   As String
'    Dim L_Writed    As Boolean
'    Dim i           As Integer
    
'    Rem Lire les valeurs
'    p_Line = Trim(p_Line)                       '      <ll>blabla<\ll>     '    => '<ll>blabla<\ll>'
'    p_Line = Mid(p_Line, 2, Len(p_Line) - 2)    '<ll>blabla<\ll>'               => 'll>blabla<\ll'
'    p_Line = Replace(p_Line, "<\", "|")         'll>blabla<\ll                  => 'll>blabla|ll'
'    p_Line = Replace(p_Line, ">", "|")          ''ll>blabla|ll'                 => 'll|blabla|ll'
    
'    L_array = Split(p_Line, "|")
'    L_Balise = LCase(L_array(0))
'    L_valeur = LCase(L_array(1))
'    L_Writed = False
    
'    For i = 0 To UBound(p_Tab_Epobox_Facture)
'        If p_Tab_Epobox_Facture(UBound(p_Tab_Epobox_Facture)).field_name = L_Balise Then
'            p_Tab_Epobox_Facture(UBound(p_Tab_Epobox_Facture)).field_value = L_valeur
'            L_Writed = True
'        End If
'    Next
        
'    Write_T_ePoBox_Facture = L_Writed
'    Exit Function
    
'Err_Write_T_ePoBox_Facture:
'    Write_T_ePoBox_Facture = False
    
'End Function

'Private Function Read_T_ePoBox_Facture(p_Tab_Epobox_Facture() As T_ePoBox_Facture, P_OutPut As String, Optional p_Separator As String) As String

'    On Error GoTo Err_Read_T_ePoBox_Facture
    
'    Dim L_Output        As String
'    Dim L_Header        As String
'    Dim L_Values        As String
'    Dim L_Separateur    As String
'    Dim i               As Integer
    
'    P_OutPut = LCase(P_OutPut)
'    Select Case P_OutPut
'    Case "csv"
'        L_Separateur = ";"
'    Case "sql"
'        Rem Ok
'        L_Separateur = ","
'    Case Else
        Rem
'        Read_T_ePoBox_Facture = "Erreur: Format de sortie non support�"
'        Exit Function
'    End Select
    
'    L_Output = vbNullString
    
'    For i = 0 To UBound(p_Tab_Epobox_Facture)
'        L_Header = L_Header & p_Tab_Epobox_Facture(i).field_name & L_Separateur
'        L_Values = L_Values & p_Tab_Epobox_Facture(i).field_value & L_Separateur
'    Next
    
'    If p_Separator <> vbNullString Then
'        p_Separator = "|"
'    End If
    
'    Read_T_ePoBox_Facture = Left(L_Header, Len(L_Header) - 1) & p_Separator & Left(L_Values, Len(L_Values) - 1)
    
'    Exit Function
    
'Err_Read_T_ePoBox_Facture:
'    Read_T_ePoBox_Facture = "Erreur:" & Err.Number & " - " & Err.Description
    
'End Function


Public Sub Update_Nb_Pages(ByRef L_Nb_pages_dans_pli As Long, ByVal L_Type_Impression As String)

    If L_Type_Impression = "Recto/Verso" Then
        If (L_Nb_pages_dans_pli / 2 - CLng(L_Nb_pages_dans_pli / 2)) <> 0 Then
            Rem Page Impair
            L_Nb_pages_dans_pli = L_Nb_pages_dans_pli + 1
        End If
        L_Nb_pages_dans_pli = L_Nb_pages_dans_pli / 2
    End If
    
End Sub

Public Function ReadNbSms(ByVal pMsg As String) As Long

    If Len(pMsg) Mod 160 = 0 Then
        ReadNbSms = Len(pMsg) / 160
    Else
        ReadNbSms = Int(Len(pMsg) / 160) + 1
    End If
    
End Function

Public Function Word2Pdf(p_FullWord As String, Optional p_dir_to As String) As String

    Dim L_FullWord() As String
    Dim DefaultPrinter
    Dim i As Integer
    
    Word2Pdf = "Ko"
    On Error GoTo Err_Word2Pdf
    'Screen.MousePointer = vbHourglass
    'DefaultPrinter = PDFCreator1.cDefaultPrinter ' Sauvegarde de l'imprimante par d�faut
    'PDFCreator1.cVisible = False
    L_FullWord = Split(p_FullWord, "|")
    ' Transformer le liste des fichiers DOC en PDF
    'For i = 0 To List1.ListCount - 1
    For i = 0 To UBound(L_FullWord) - 1
        StartTime = Now
        'ImprimePDF (List1.List(i))
        Call ImprimePDF(L_FullWord(i), p_dir_to)
    Next i
    'PDFCreator1.cVisible = True
    'PDFCreator1.cDefaultPrinter = DefaultPrinter ' R�affectation de l'imprimante par d�faut
    'Screen.MousePointer = vbNormal
    
    Word2Pdf = "Ok"
    Exit Function
    
Err_Word2Pdf:
    Word2Pdf = Err.Number & " - " & Err.Description
    
End Function

Private Sub ImprimePDF(ByVal fichier As String, Optional p_DirTo As String)
' Transformer en fichier DOC en PDF
    
    Dim sRep, sNomFic As String
    Dim i As Integer
     
    'S�parer le chemin du nom de fichier
    If InStr(fichier, "\") Then
        For i = Len(fichier) To 0 Step -1 ' Je commence la recherche du "\" par la fin
            If Mid(fichier, i, 1) = "\" Then ' D�s que je le trouve, je sors du boucle For
                Exit For
            End If
        Next i
        sRep = Left(fichier, i)
        sNomFic = Mid(fichier, i + 1)
        sNomFic = Left(sNomFic, Len(sNomFic) - 4) ' Nom du fichier sans son extension
    Else
        sRep = ""
        sNomFic = Left(fichier, Len(fichier) - 4)
    End If
    If p_DirTo <> "" Then
        sRep = p_DirTo
    End If
  
    'AddStatus "D�but de la cr�ation du fichier PDF ..."
    With opt
        .AutosaveDirectory = sRep ' Chemin du fichier qui est le m�me que celui d'origine
        .AutosaveFilename = sNomFic ' Nom du fichier qui est le m�me sans extension
        .UseAutosave = 1
        .UseAutosaveDirectory = 1
        .AutosaveFormat = 0 ' PDF
        .PaperSize = "a4"
        .PDFGeneralResolution = 120
    End With
    With PDFCreator1
        Set .cOptions = opt
        .cClearCache
        .cDefaultPrinter = "PDFCreator"
        .cPrintFile (fichier)
        .cPrinterStop = False
    End With
    
    ' Pour g�rer l'�tat de PDFCreator, tant que l'impression n'est pas fini,
    ' on ne passe pas � la suivante
    Do While PDFCreator1.cPrinterStop = False
        DoEvents
    Loop
End Sub

Public Function InitPdfCreatorPrint(pDir As String, pIdPe As String) As String

    InitPdfCreatorPrint = "Ko"
    On Error GoTo errInitPdfCreatorPrint
    
    'AddStatus "D�but de la cr�ation du fichier PDF ..."
    With opt
        .AutosaveDirectory = pDir ' Chemin du fichier qui est le m�me que celui d'origine
        .AutosaveFilename = pIdPe ' Nom du fichier qui est le m�me sans extension
        .UseAutosave = 1
        .UseAutosaveDirectory = 1
        .AutosaveFormat = 0 ' PDF
        .PaperSize = "a4"
        .PDFGeneralResolution = 120
    End With
    With PDFCreator1
        Set .cOptions = opt
        .cClearCache
        .cDefaultPrinter = "PDFCreator"
        '.cPrintFile (fichier)
        .cPrinterStop = False
    End With
    
    InitPdfCreatorPrint = "Ok"

    Exit Function
    
errInitPdfCreatorPrint:
    InitPdfCreatorPrint = Err.Number & " - " & Err.Description
    
End Function

Private Function WaitingPdfCreator()
    ' Pour g�rer l'�tat de PDFCreator, tant que l'impression n'est pas fini,
    ' on ne passe pas � la suivante
    Do While PDFCreator1.cPrinterStop = False
        DoEvents
    Loop

End Function

Public Function Word_Save_To_PDFCreator(p_Document_Name As String, _
                                        p_Document_Dir As String, _
                                        p_Document_PdfName) _
                                        As String

Dim L_Word_Closed   As Boolean
Dim L_Result        As String

On Error GoTo W_Error:
    
    L_Word_Closed = G_wd_opened
    If Not G_wd_opened Then
        Set G_Wd = CreateObject("Word.Application")
        G_wd_opened = True
        G_Wd.Visible = False
    End If
    G_Wd.Documents.Open p_Document_Dir & "\" & p_Document_Name, , , False
Ignore_5127:
    If Left(G_Wd.Documents(p_Document_Name).Application.ActivePrinter, Len(G_Driver_PDFCreator)) <> G_Driver_PDFCreator Then
        G_Wd.Documents(p_Document_Name).Application.ActivePrinter = G_Driver_PDFCreator
    End If
    L_Result = InitPdfCreatorPrint(p_Document_Dir, CStr(p_Document_PdfName))
    If L_Result <> "Ok" Then
        Word_Save_To_PDFCreator = L_Result
        Exit Function
    End If
    G_Wd.Documents(p_Document_Name).PrintOut Background:=False, Range:=wdPrintAllDocument, PrintToFile:=False
    Call WaitingPdfCreator
    G_Wd.Documents(p_Document_Name).Close savechanges:=wdDoNotSaveChanges
    
    On Error Resume Next
    Word_Save_To_PDFCreator = "Ok"
    
    If Not L_Word_Closed Then
        Call Word_Close
    End If
    Exit Function

W_Error:

    If Err.Number = 5121 Then
        Word_Save_To_PDFCreator = "Impossible d'ouvrir le document Word (err=5121, " & Err.Description & ")"
    ElseIf Err.Number = 5127 Then
        Err.Clear
        GoTo Ignore_5127
    Else
        Word_Save_To_PDFCreator = Err.Number & " - " & Err.Description & vbNewLine & "Source: Word_Save_To_PDFCreator"
    End If
    
    Exit Function
    
W_Error2:
    Word_Save_To_PDFCreator = "Impossible de convertir le document"

End Function


'Public Function createPdfSpecialPoweo(ByRef pForm As Form, pNumClient As String, pDir As String) As String
'
'    Dim p               As PDFlib_com.PDF
'    Dim r               As Long
'    Dim y               As Long
'    Dim OutPutPDFInt    As Long
'    Dim PdfOut          As String
'    Dim PdfGenerated    As String
'    Dim ProPart         As String
'    Dim Masque          As String
'    Dim ListContrat As String
'    Dim Result          As String
    
'    If Trim(pNumClient) = "" Then
'        createPdfSpecialPoweo = ""
'        Exit Function
'    End If
    
'    ListContrat = Lire_Une_Liste("_poweo_contrats", "num_contrat", "information", "num_client = '" & pNumClient & "' and operation = 'A traiter'")
    
'    If Trim(ListContrat) = "" Then
'        createPdfSpecialPoweo = ""
'        Exit Function
'    End If
    
'    ProPart = UCase(Lire_Un_Champ("type_client", "_poweo_contrats", "num_client = '" & pNumClient & "'"))
'    Select Case ProPart
'    Case "PRO", "PART"
'        Masque = G_dir_client & "\033411000334\" & G_CONST_FICHIERS_REPERTORIES & "\MIGRATION - LOT2 - " & ProPart & ".pdf"
'        If Not FileExists(Masque) Then
'            createPdfSpecialPoweo = "ko"
'            Exit Function
'        End If
'    Case Else
'        createPdfSpecialPoweo = ""
'        Exit Function
'    End Select
    
'    PdfGenerated = pDir & "LOT2_" & pNumClient & "_" & ProPart & ".pdf"
'    PdfOut = Replace(PdfGenerated, ".pdf", "_FUSION.pdf", , , vbTextCompare)
'    If FileExists(PdfOut) Then
'        Kill PdfOut
'    End If
'    If FileExists(PdfGenerated) Then
'        Kill PdfGenerated
'    End If
    
    
'    Set p = New PDFlib_com.PDF
'    OutPutPDFInt = p.begin_document(PdfGenerated, "")
'    p.set_info "Creator", "Axessy"
'    Call CreatePdfClientContrat(p, pNumClient, ListContrat)
'    p.end_document ""
'    Set p = Nothing

    
'    FileCopy PdfGenerated, PdfOut
'    Result = Merge_Pdf(pForm, PdfOut, Masque, True)
'    If Result = "Ok" Then
'        createPdfSpecialPoweo = PdfOut
'        Exit Function
'    Else
'        createPdfSpecialPoweo = "ko"
'    End If
    
'Exit Function

'Err_createPdfSpecialPoweo:
'    createPdfSpecialPoweo = "ko"

'End Function
