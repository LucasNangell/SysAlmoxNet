Attribute VB_Name = "Módulo 99a - pbSubsSystemDev"
Option Compare Database
Option Explicit


'no form [ frmDev_90dDevPickColor ] há uma rotina ColorPiker


Public Sub CheckFormPosition()
    
    Dim fForm As Form, fForm2 As Form

    'Retorna a diferença em pix entre dois formulários abertos
    ' como usar:
    ' . abrir dois forms
    ' . no painel [ Verificação imediata ] do VB digitar CheckFormPosition + Enter
    For Each fForm In Forms
        
        If gBbDebugOn Then Debug.Print fForm.Name & vbCrLf & "-----------------------------------------------" & vbCrLf & " Diferenças de posição dos formulários abertos: "
        
        For Each fForm2 In Forms
            If gBbDebugOn Then Debug.Print fForm.Name & " > "; fForm2.Name & ": " & vbCrLf & " - Left: " & fForm.WindowLeft - fForm2.WindowLeft & vbCrLf & " - Top: " & fForm.WindowTop - fForm2.WindowTop
        
        Next fForm2
        If gBbDebugOn Then Debug.Print "-----------------------------------------------" & vbCrLf & "-----------------------------------------------" & vbCrLf & vbCrLf
    
    Next fForm

End Sub


