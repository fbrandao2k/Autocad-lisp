Attribute VB_Name = "ModuloImportarGraficos"
Option Explicit

'routine to open a dialog box asking to select the desirable file
Function GOFN(Optional sPath As String, Optional sTitle As String, _
    Optional sFilter As String, Optional bMulti As Boolean) As Variant

  Dim vTemp As Variant
  Dim sHomeDir As String

  ' save current directory
  On Error Resume Next
  sHomeDir = CurDir
  If Len(sPath) > 1 Then
    ' change to new directory if it is specified
    ChDrive sPath
    ChDir sPath
  End If
  On Error GoTo 0
  vTemp = Application.GetOpenFilename( _
      FileFilter:=sFilter, Title:=sTitle, MultiSelect:=bMulti)
  If TypeName(vTemp) = "Boolean" Then
    ' canceled by user
    ' check for Boolean in calling procedure
  End If
  GOFN = vTemp

  ' restore current directory
  On Error Resume Next
  ChDrive sHomeDir
  ChDir sHomeDir
  On Error GoTo 0
End Function

'ORIGEM = name of the file that contain the excel chart
Public Sub ImportarGraficos()
    
    Dim mensagem, FilePath, ORIGEM, Z, chart
    'Dim resposta, DESTINO, Z, tipoFonte, linhaDestino, rng As Variant
    
    mensagem = True
    
    'DESTINO = ThisWorkbook.Name
    FilePath = GOFN(ThisWorkbook.Path, "Selecione el archivo deseado: ", "Excel Files (*.xls;*.xlsx),*.xls;*.xlsx", True)
    'FilePath = Application.GetOpenFilename(  (“Excel Files (*.TXT; *.CSV), *.txt; *.csv”)
    
    If TypeName(FilePath) = "Boolean" Then
        MsgBox "Archivo no Encontrado.", vbCritical, "Error en la importación"
        mensagem = False
        GoTo SAIDA
    End If
    
    'it is possible to select more than one file
    DoEvents
    For Z = 1 To 1 'UBound(FilePath)
        Application.Workbooks.Open (FilePath(Z))
        ORIGEM = ActiveWorkbook.Name
        'ThisWorkbook.Activate
        

        Dim wb As Workbook

        Set wb = Workbooks(Dir(FilePath(Z)))

        With wb
            'MsgBox .FullName
            '.Select
            .Activate
            '~~> Do what you want
        End With
        
        Main2

    Next Z

    mensagem = False
    GoTo SAIDA

'
SAIDA:
    On Error Resume Next
    If mensagem Then MsgBox "Problema con la importación.", vbCritical, "Error en la importación"
'    Range("b6").Select
    'Workbooks(ORIGEM).Close SaveChanges:=False
    On Error GoTo 0


End Sub
