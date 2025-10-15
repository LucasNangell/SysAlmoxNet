Attribute VB_Name = "Módulo 99a - pbSubsSystemDev"
Option Compare Database
Option Explicit


'no form [ frmDev_90dDevPickColor ] há uma rotina ColorPiker


Public Sub CheckFormPosition()
    
    Dim fForM As Form, fForm2 As Form

    'Retorna a diferença em pix entre dois formulários abertos
    ' como usar:
    ' . abrir dois forms
    ' . no painel [ Verificação imediata ] do VB digitar CheckFormPosition + Enter
    For Each fForM In Forms
        
        If gBbDebugOn Then Debug.Print fForM.Name & vbCrLf & "-----------------------------------------------" & vbCrLf & " Diferenças de posição dos formulários abertos: "
        
        For Each fForm2 In Forms
            If gBbDebugOn Then Debug.Print fForM.Name & " > "; fForm2.Name & ": " & vbCrLf & " - Left: " & fForM.WindowLeft - fForm2.WindowLeft & vbCrLf & " - Top: " & fForM.WindowTop - fForm2.WindowTop
        
        Next fForm2
        If gBbDebugOn Then Debug.Print "-----------------------------------------------" & vbCrLf & "-----------------------------------------------" & vbCrLf & vbCrLf
    
    Next fForM

End Sub


