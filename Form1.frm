VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Datos Despacho"
      Height          =   3255
      Left            =   7560
      TabIndex        =   24
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Gabinete"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Numero Ferrovial"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Factura"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Guia Despacho"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Ingreso"
      Height          =   3255
      Left            =   3720
      TabIndex        =   15
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "OC Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "OC Proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Proyecto"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10560
      Top             =   1800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   13
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Ingreso"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Numero de Serie"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Lectura"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modo"
      Height          =   1215
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Consulta"
         Height          =   615
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Despacho"
         Height          =   615
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ingreso"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public i As Integer
Public counter As Integer

Public modo As Integer

Public obj As Excel.Application
Public Hoja As Excel.Worksheet
'Dim rec1 As Recordset
'Dim con As Connection


Public mark2 As Integer

Private Sub Command1_Click()

modo = 1

Set obj = GetObject(, "Excel.Application")

' Lectura de Datos desde el excel
obj.ActiveSheet.Select

Command1.BackColor = &HC0C0C0

Text1.Enabled = True
Text1.BackColor = vbWhite

Text1.SetFocus

End Sub

Private Sub Command2_Click()

modo = 1

End Sub

Private Sub Command3_Click()

modo = 3

End Sub

Private Sub Form_Load()

i = 2

counter = 1

End Sub

Private Sub Text1_Change()

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""



mark2 = InStr(Text1.Text, "91")

Text2.Text = Replace((Left(Text1.Text, (mark2 - 1))), "93", "")

Text3.Text = Replace(Text1.Text, ("93" & Text2.Text & "91"), "")

Text3.Text = Replace(Left(Text3.Text, (Len(Text3.Text) - 1)), "92", "")

Text9.Text = Now()

While (obj.Worksheets("Codigos").Cells(i, 2).Value) <> "Final"

If (obj.Worksheets("Codigos").Cells(i, 2).Value) = Text2.Text Then

Text4.Text = obj.Worksheets("Codigos").Cells(i, 1).Value

End If

i = i + 1

Wend


While (obj.Worksheets("BD").Cells(counter, 1).Value) <> ""

counter = counter + 1

Wend

obj.Worksheets("BD").Cells(counter, 1).Value = Text2.Text
obj.Worksheets("BD").Cells(counter, 2).Value = Text3.Text
obj.Worksheets("BD").Cells(counter, 3).Value = Text4.Text
obj.Worksheets("BD").Cells(counter, 4).Value = Text5.Text
obj.Worksheets("BD").Cells(counter, 5).Value = Text6.Text
obj.Worksheets("BD").Cells(counter, 6).Value = Text7.Text
obj.Worksheets("BD").Cells(counter, 7).Value = Text8.Text

obj.Worksheets("BD").Cells(counter, 8).Value = "N/A"
obj.Worksheets("BD").Cells(counter, 9).Value = "N/A"
obj.Worksheets("BD").Cells(counter, 10).Value = "N/A"
obj.Worksheets("BD").Cells(counter, 11).Value = "N/A"

obj.Worksheets("BD").Cells(counter, 12).Value = Text9.Text

obj.Worksheets("BD").Cells(counter, 13).Value = "N/A"

Text1.Text = ""

counter = 1
i = 1

Timer1.Enabled = False

End Sub
