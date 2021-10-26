VERSION 5.00
Begin VB.Form frmTelaPrincipal 
   Caption         =   "Reading UTF-8 files in Visual Basic 6"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResultado 
      Height          =   8655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmTelaPrincipal.frx":0000
      Top             =   840
      Width           =   10935
   End
   Begin VB.TextBox txtArquivo 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin VB.CommandButton btnAbrir 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTelaPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function ReadUTF8File(sFile) As String
   Const ForReading = 1
   Dim sPrefix
 
   With CreateObject("Scripting.FileSystemObject")
     sPrefix = .OpenTextFile(sFile, ForReading, False, False).Read(3)
   End With
   If Left(sPrefix, 3) <> Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
     With CreateObject("Scripting.FileSystemObject")
       pvReadFile = .OpenTextFile(sFile, ForReading, False, Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE)).ReadAll()
       ReadUTF8File = pvReadFile
     End With
   Else
     With CreateObject("ADODB.Stream")
       .Open
       If Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE) Then
         .Charset = "Unicode"
       ElseIf Left(sPrefix, 3) = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
         .Charset = "UTF-8"
       Else
         .Charset = "_autodetect"
       End If
       .LoadFromFile sFile
       pvReadFile = .ReadText
       ReadUTF8File = pvReadFile
     End With
   End If
End Function


Private Sub btnAbrir_Click()
    
    txtResultado = ReadUTF8File(txtArquivo)
    
End Sub
