Attribute VB_Name = "mod_MD5"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
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
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


' MD5.bas - wrapper for RSA MD5 DLL
'   derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm
' Functions:
'   MD5String (some string) -> MD5 digest of the given string as 32 bytes string
'   MD5File (some filename) -> MD5 digest of the file's content as a 32 bytes string
'      returns a null terminated "FILE NOT FOUND" if unable to open the
'      given filename for input
' Bugs, complaints, etc:
'   Francisco Carlos Piragibe de Almeida
'   piragibe@esquadro.com.br
' History
'       Apr, 17 1999 - fixed the null byte problem
' Contains public domain RSA C-code for MD5 digest (see MD5-original.txt file)
' The aamd532.dll DLL MUST be somewhere in your search path
'   for this to work

Option Explicit

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(ByVal p As String) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space$(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(ByVal f As String) As String
'*************************************************
'Author: Unkwown
'Last modified: ?/?/?
'*************************************************
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space$(32)
    MDFile f, r
    MD5File = r
End Function
