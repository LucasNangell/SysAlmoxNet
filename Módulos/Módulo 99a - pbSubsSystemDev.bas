Attribute VB_Name = "M�dulo 99a - pbSubsSystemDev"
Option Compare Database
Option Explicit


'no form [ frmDev_90dDevPickColor ] h� uma rotina ColorPiker


Public Sub CheckFormPosition()
    
    Dim fForm As Form, fForm2 As Form

    'Retorna a diferen�a em pix entre dois formul�rios abertos
    ' como usar:
    ' . abrir dois forms
    ' . no painel [ Verifica��o imediata ] do VB digitar CheckFormPosition + Enter
    For Each fForm In Forms
        
        If gBbDebugOn Then Debug.Print fForm.Name & vbCrLf & "-----------------------------------------------" & vbCrLf & " Diferen�as de posi��o dos formul�rios abertos: "
        
        For Each fForm2 In Forms
            If gBbDebugOn Then Debug.Print fForm.Name & " > "; fForm2.Name & ": " & vbCrLf & " - Left: " & fForm.WindowLeft - fForm2.WindowLeft & vbCrLf & " - Top: " & fForm.WindowTop - fForm2.WindowTop
        
        Next fForm2
        If gBbDebugOn Then Debug.Print "-----------------------------------------------" & vbCrLf & "-----------------------------------------------" & vbCrLf & vbCrLf
    
    Next fForm

End Sub


