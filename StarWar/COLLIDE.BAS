Attribute VB_Name = "Module2"
' Function when collide will happen
Function COLLIDE(XPOSITION1 As Integer, YPOSITION1 As Integer, XPOSITION2 As Integer, YPOSITION2 As Integer, BOXSIZE As Integer) As Integer
     If XPOSITION1 > XPOSITION2 - (BOXSIZE / 2) Then
          If XPOSITION1 < XPOSITION2 + (BOXSIZE / 2) Then
               If YPOSITION1 > YPOSITION2 - (BOXSIZE / 2) Then
                    If YPOSITION1 < YPOSITION2 + (BOXSIZE / 2) Then
                         COLLIDE = 1
                    End If
               End If
          End If
     End If
End Function

