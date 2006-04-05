Attribute VB_Name = "mod_DirectX"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

Private error As Boolean
Public X As String, Y As String, dxVer As String
Public directx As New DirectX7
Public DirectDraw As DirectDraw7
Public DirectSound As DirectSound

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub VersionDirectX()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 14/03/2006
'10/03/06: Maraxus - Si Shell funciona espero hasta que la ejecución
'de dxdiag termine para evitar errores (se hace asincrónicamente y puede no terminar antes del open)
'*************************************************
On Error GoTo ErrHandler

    Dim handle As Integer
    
    Call Shell("dxdiag C:\DXTest.txt")
    
    Do Until FileExist("C:\DXTest.txt", vbArchive)
        DoEvents
    Loop
    
    handle = FreeFile
    
    Open "C:\DXTest.txt" For Input As handle
    
    Do While Not EOF(handle)
        Line Input #handle, X
        If Left$(LTrim$(X), 15) = "DirectX Version" Then
            Y = LTrim$(X)
            dxVer = Mid$(Y, 17, Len(Y) - 16)
            frmAOSetup.lDirectX.Caption = dxVer
            Exit Do
        End If
    Loop
    
    Close handle
    Kill "C:\DXTest.txt"
Exit Sub

ErrHandler:
    If Err.Number = 70 Then
        'Permission denied. Es posible el archivo exista pero como todavía
        'está abierto por dxdiag no podemos abrirlo. Dormimos 5 ms e intentamos de nuevo
        Sleep 5
        Resume
    End If
    frmAOSetup.Caption = "Error"
End Sub

Public Sub ProbarDirectX()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 10/03/06
'10/03/06: ^[GS]^ - Adapte el codigo al nuevo formulario principal
'*************************************************
    error = False
    
    'Create DX object
    If IniciarDXobject(directx) Then
        frmAOSetup.lblDX.Visible = True
        frmAOSetup.lblDX.Caption = "OK"
        frmAOSetup.lblDX.ForeColor = &H8000&
        frmAOSetup.lblDX.Font.Bold = True
    Else
        frmAOSetup.lblDX.Visible = True
        frmAOSetup.lblDX.Caption = "OK"
        frmAOSetup.lblDX.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDX.Font.Bold = True
        error = True
    End If
    
    frmAOSetup.Text1.BackColor = frmAOSetup.lblDX.ForeColor
    DoEvents
    
    'Create DirectSound
    If IniciarDirectSound() Then
        frmAOSetup.lblDS.Visible = True
        frmAOSetup.lblDS.Caption = "OK"
        frmAOSetup.lblDS.ForeColor = &H8000&
        frmAOSetup.lblDS.Font.Bold = True
    Else
        frmAOSetup.lblDS.Visible = True
        frmAOSetup.lblDS.Caption = "OK"
        frmAOSetup.lblDS.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDS.Font.Bold = True
        error = True
    End If
    frmAOSetup.Text3.BackColor = frmAOSetup.lblDS.ForeColor
    DoEvents
    
    'Create DirectDraw
    If IniciarDDobject(DirectDraw) Then
        frmAOSetup.lblDD.Visible = True
        frmAOSetup.lblDD.Caption = "OK"
        frmAOSetup.lblDD.ForeColor = &H8000&
        frmAOSetup.lblDD.Font.Bold = True
    Else
        frmAOSetup.lblDD.Visible = True
        frmAOSetup.lblDD.Caption = "ERROR"
        frmAOSetup.lblDD.ForeColor = RGB(255, 0, 0)
        frmAOSetup.lblDD.Font.Bold = True
        error = True
    End If
    frmAOSetup.Text2.BackColor = frmAOSetup.lblDD.ForeColor
    DoEvents
    
    If error Then
        MsgBox "Necesita reinstalar DirectX", vbCritical, "Argentum Online Setup"
    End If
End Sub

Private Function IniciarDirectSound() As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
On Error Resume Next
    Set DirectSound = directx.DirectSoundCreate("")
    
    If Err Then
        IniciarDirectSound = False
        Exit Function
    End If
    
    IniciarDirectSound = True
End Function

Private Function IniciarDXobject(ByRef dx As DirectX7) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
On Error Resume Next

    Set dx = New DirectX7
    
    If Err Then
        Err.Clear
        IniciarDXobject = False
        Exit Function
    End If
    
    IniciarDXobject = True
End Function

Private Function IniciarDDobject(ByRef DD As DirectDraw7) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'*************************************************
On Error Resume Next
    Set DD = directx.DirectDrawCreate("")
    
    If Err Then
        Err.Clear
        IniciarDDobject = False
        Exit Function
    End If
    
    IniciarDDobject = True
End Function
