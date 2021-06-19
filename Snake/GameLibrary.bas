Attribute VB_Name = "GameLibrary"
'This game library contains a number of functions which can
'be useful in calculating things like collisions, distances,
'and directions from one object to another, which are useful
'for programming the artifical intelligence.

Public Function CollisionDetection(Object1 As Object, Object2 As Object) As Boolean
'This function finds whether two objects have
'collided or not. Returns true or false
'
If Object1.Left < Object2.Left + Object2.Width And Object1.Left + Object1.Width > Object2.Left And Object1.Top < Object2.Top + Object2.Height And Object1.Top + Object1.Height > Object2.Top Then
    CollisionDetection = True
Else
    CollisionDetection = False
End If
End Function

Public Function CollisionDetection2(Object1 As Object, XPos2 As Integer, YPos2 As Integer, Width2 As Integer, Height2 As Integer) As Boolean
'This function is similar to the last but allows
'you to find the collision between an object and
'a specific point or section.
'
'It asks you to specify the x and y position and the
'width and height of the point or section. If you are
'finding a collision with a point, input zero as the
'height and width.
'
'Returns true or false
'
If Object1.Left < XPos2 + Width2 And Object1.Left + Object1.Width > XPos2 And Object1.Top < YPos2 + Height2 And Object1.Top + Object1.Height > YPos2 Then
    CollisionDetection2 = True
Else
    CollisionDetection2 = False
End If
End Function

Public Function CollisionDetection3(XPos1 As Integer, YPos1 As Integer, Width1 As Integer, Height1 As Integer, XPos2 As Integer, YPos2 As Integer, Width2 As Integer, Height2 As Integer) As Boolean
'This function is similar to the last but allows
'you to find the collision between two different
'points or sections.
'
'It asks you to specify the x and y position and the
'width and height of each point or section. If you are
'finding a collision with points, input zero as the
'height and width.
'
'Returns true or false
'
If XPos1 < XPos2 + Width2 And XPos1 + Width1 > XPos2 And YPos1 < YPos2 + Height2 And YPos1 + Height1 > YPos2 Then
    CollisionDetection3 = True
Else
    CollisionDetection3 = False
End If
End Function

Public Function ObjectCenterWidth(Object1 As Object) As Integer
'This function helps finds the centre of an object
'It improves the readability of the code
'
ObjectCenterWidth = Object1.Width / 2
End Function

Public Function ObjectCenterHeight(Object1 As Object) As Integer
'Similarly, this function also helps in finding the centre
'It makes code easier to read
'
ObjectCenterHeight = Object1.Height / 2
End Function

Public Function ConvertRadianstoDegrees(Angle As Integer) As Integer
'This function converts angles in radians to degrees.
'This is especially useful as Visual Basic returns
'the results of trig functions in radians, and you
'may want the result in degrees.
'
'Here 's how you it to find the tan of 45 degrees :
'
'   x = tan(ConvertRadianstoDegrees(45))
'
ConvertRadianstoDegrees = Angle * (4 * Atn(1)) / 180
End Function

Public Function Direction(FromObject As Object, ToObject As Object) As Integer
'This function returns a number between 1 and 16
'depending on the direction from the first object
'to the second object.
'
'Let x represent the first object. If the second object
'was at the position of where the number seven was, the
'function would return the direction of seven
'
'                   1
'               16      2
'           15              3
'       14                      4
'   13              x               5
'       12                      6
'           11              7
'               10      8
'                   9
'
'
'It's a actually calculates the direction more like this :
'
'
'           15      16      1       2       3
'
'
'           14                              4
'
'
'           13              x               5
'
'
'           12                              6
'
'
'           11      10      9       8       7
'
Dim a As Integer
Dim b As Integer

If ToObject.Top + (ToObject.Height / 2) = FromObject.Top + (FromObject.Height / 2) Then
    If ToObject.Left + (ToObject.Width / 2) = FromObject.Left + (FromObject.Width / 2) Then
        Direction = 0
    ElseIf ToObject.Left + (ToObject.Width / 2) > FromObject.Left + (FromObject.Width / 2) Then
        Direction = 5
    Else
        Direction = 13
    End If
ElseIf ToObject.Top + (ToObject.Height / 2) < FromObject.Top + (FromObject.Height / 2) Then
    If ToObject.Left + (ToObject.Width / 2) = FromObject.Left + (FromObject.Width / 2) Then
        Direction = 1
    ElseIf ToObject.Left + (ToObject.Width / 2) < FromObject.Left + (FromObject.Width / 2) Then
        a = Abs(FromObject.Left + (FromObject.Width / 2) - (ToObject.Left + (ToObject.Width / 2)))
        b = Abs(FromObject.Top + (FromObject.Height / 2) - (ToObject.Top + (ToObject.Height / 2)))
        If a = b Then
            Direction = 15
        ElseIf a < b Then
            Direction = 16
        Else
            Direction = 14
        End If
    Else
        a = Abs(ToObject.Left + (ToObject.Width / 2) - (FromObject.Left + (FromObject.Width / 2)))
        b = Abs(FromObject.Top + (FromObject.Width / 2) - (ToObject.Top + (ToObject.Height / 2)))
        If a = b Then
            Direction = 3
        ElseIf a < b Then
            Direction = 2
        Else
            Direction = 4
        End If
    End If
Else
    If ToObject.Left = FromObject.Left Then
        Direction = 9
    ElseIf ToObject.Left < FromObject.Left Then
        a = Abs(FromObject.Left + (FromObject.Width / 2) - (ToObject.Left + (ToObject.Width / 2)))
        b = Abs(ToObject.Top + (ToObject.Height / 2) - (FromObject.Top + (FromObject.Height / 2)))
        If a = b Then
            Direction = 11
        ElseIf a < b Then
            Direction = 10
        Else
            Direction = 12
        End If
    Else
        a = Abs(ToObject.Left + (ToObject.Width / 2) - (FromObject.Left + (FromObject.Width / 2)))
        b = Abs(ToObject.Top + (ToObject.Width / 2) - (FromObject.Top + (FromObject.Height / 2)))
        If a = b Then
            Direction = 7
        ElseIf a < b Then
            Direction = 8
        Else
            Direction = 6
        End If
    End If
End If
End Function

Public Function DistanceBetween(FromObject As Object, ToObject As Object, FromCorner As Integer, ToCorner As Integer) As Single
'This function finds the distance between two objects
'using pythogras theorm. Consider objects x and y :
'
'
'                               Y
'                               |
'                               |   6
'                               |
'       X-----------------------
'               8
'
'           d^2 = 6^2 + 8^2
'             d = 10
'
'It also takes into account the corners of the objects.
'1=Top Left   2=Top Right   3=Bottom Right   4=Bottom Left
'
'Anything thing else (eg 0) will use the centre of the object
'
Dim Direct As Integer
Dim a As Integer
Dim b As Integer
Dim Corner1Left As Integer
Dim Corner1Top As Integer
Dim Corner2Left As Integer
Dim Corner2Top As Integer

Select Case FromCorner
    Case Is = 1
        Corner1Left = FromObject.Left
        Corner1Top = FromObject.Top
    Case Is = 2
        Corner1Left = FromObject.Left + FromObject.Width
        Corner1Top = FromObject.Top
    Case Is = 3
        Corner1Left = FromObject.Left + FromObject.Width
        Corner1Top = FromObject.Top + FromObject.Height
    Case Is = 4
        Corner1Left = FromObject.Left
        Corner1Top = FromObject.Top + FromObject.Height
    Case Else
        Corner1Left = FromObject.Left + FromObject.Width / 2
        Corner1Top = FromObject.Top + FromObject.Height / 2
End Select
Select Case ToCorner
    Case Is = 1
        Corner2Left = ToObject.Left
        Corner2Top = ToObject.Top
    Case Is = 2
        Corner2Left = ToObject.Left + ToObject.Width
        Corner2Top = ToObject.Top
    Case Is = 3
        Corner2Left = ToObject.Left + ToObject.Width
        Corner2Top = ToObject.Top + ToObject.Height
    Case Is = 4
        Corner2Left = ToObject.Left
        Corner2Top = ToObject.Top + ToObject.Height
    Case Else
        Corner2Left = ToObject.Left + ToObject.Width / 2
        Corner2Top = ToObject.Top + ToObject.Height / 2
End Select

a = Abs(Corner1Left - Corner2Left)
b = Abs(Corner1Top - Corner2Top)
DistanceBetween = Sqr(a ^ 2 + b ^ 2)
End Function

Public Function RelativeDistanceBetween(FromObject As Object, ToObject As Object, FromCorner As Integer, ToCorner As Integer) As Long
'In some instances you will only want to compare the
'results of distances with different objects and you
'may not want the exact value of the distance.
'
'In this function, the square root is excluded as it
'is not neccessary when comparing distances. This
'exclusion will save valuable processing time which
'can be useful in getting the game to run better.
'
'
'                               Y
'                               |
'                               |   6
'                               |
'       X-----------------------
'               8
'
'             d = 6^2 + 8^2
'             d = 100
'
'Again the corners are of the objects are taken into account.
'1=Top Left   2=Top Right   3=Bottom Right   4=Bottom Left
'
'Anything thing else (eg 0) will use the centre of the object
'
Dim Direct As Integer
Dim a As Integer
Dim b As Integer
Dim Corner1Left As Integer
Dim Corner1Top As Integer
Dim Corner2Left As Integer
Dim Corner2Top As Integer

Select Case FromCorner
    Case Is = 1
        Corner1Left = FromObject.Left
        Corner1Top = FromObject.Top
    Case Is = 2
        Corner1Left = FromObject.Left + FromObject.Width
        Corner1Top = FromObject.Top
    Case Is = 3
        Corner1Left = FromObject.Left + FromObject.Width
        Corner1Top = FromObject.Top + FromObject.Height
    Case Is = 4
        Corner1Left = FromObject.Left
        Corner1Top = FromObject.Top + FromObject.Height
    Case Else
        Corner1Left = FromObject.Left + FromObject.Width / 2
        Corner1Top = FromObject.Top + FromObject.Height / 2
End Select
Select Case ToCorner
    Case Is = 1
        Corner2Left = ToObject.Left
        Corner2Top = ToObject.Top
    Case Is = 2
        Corner2Left = ToObject.Left + ToObject.Width
        Corner2Top = ToObject.Top
    Case Is = 3
        Corner2Left = ToObject.Left + ToObject.Width
        Corner2Top = ToObject.Top + ToObject.Height
    Case Is = 4
        Corner2Left = ToObject.Left
        Corner2Top = ToObject.Top + ToObject.Height
    Case Else
        Corner2Left = ToObject.Left + ToObject.Width / 2
        Corner2Top = ToObject.Top + ToObject.Height / 2
End Select

a = Abs(Corner1Left - Corner2Left)
b = Abs(Corner1Top - Corner2Top)
RelativeDistanceBetween = a ^ 2 + b ^ 2
End Function

