Attribute VB_Name = "M�dulo 99a - pbSubsSystemDev"
Option Compare Database
Option Explicit


'no form [ frmDev_90dDevPickColor ] h� uma rotina ColorPiker


Public Sub CheckFormPosition()
    
    Dim fForM As Form, fForm2 As Form

    'Retorna a diferen�a em pix entre dois formul�rios abertos
    ' como usar:
    ' . abrir dois forms
    ' . no painel [ Verifica��o imediata ] do VB digitar CheckFormPosition + Enter
    For Each fForM In Forms
        
        If gBbDebugOn Then Debug.Print fForM.Name & vbCrLf & "-----------------------------------------------" & vbCrLf & " Diferen�as de posi��o dos formul�rios abertos: "
        
        For Each fForm2 In Forms
            If gBbDebugOn Then Debug.Print fForM.Name & " > "; fForm2.Name & ": " & vbCrLf & " - Left: " & fForM.WindowLeft - fForm2.WindowLeft & vbCrLf & " - Top: " & fForM.WindowTop - fForm2.WindowTop
        
        Next fForm2
        If gBbDebugOn Then Debug.Print "-----------------------------------------------" & vbCrLf & "-----------------------------------------------" & vbCrLf & vbCrLf
    
    Next fForM

End Sub


