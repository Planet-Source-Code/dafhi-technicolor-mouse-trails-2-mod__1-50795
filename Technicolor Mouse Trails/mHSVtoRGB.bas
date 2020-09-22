Attribute VB_Name = "mHSVtoRGB"
Option Explicit

Public Function HSV(Optional ByVal Hue As Single = -1, Optional ByVal Saturation As Single = 1, Optional ByVal Lightness As Single = 1) As Long

    
    ' ==============================================================================================
    ' This code comes from one of my other PlanetSourceCode submissions.
    ' ==============================================================================================
    ' Given a Hue, Saturation and Lightness, return the Red-Green_Blue equivalent as a Long data type.
    ' This funtion is intended to replace VB's RGB function.
    '
    ' Ranges:
    '   Hue -1 (no hue)
    '       or
    '   Hue 0 to 360
    '
    '   Saturation 0 to 1
    '   Lightness 0 to 1
    '
    ' ie. Bright-RED = (Hue=0, Saturation=1, Lightness=1)
    '
    ' Example:
    '   Picture1.ForeColor = HSV(0,1,1)
    '
    ' ==============================================================================================
    
    Dim Red As Single
    Dim Green As Single
    Dim Blue As Single
    
    Dim I As Single
    Dim f As Single
    Dim p As Single
    Dim q As Single
    Dim t As Single
    
    If Saturation = 0 Then  '   The colour is on the black-and-white center line.
        If Hue = -1 Then    '   Achromatic color: There is no hue.
            Red = Lightness
            Green = Lightness
            Blue = Lightness
        Else
            ' *** Make sure you've turned on 'Break on unhandled Errors' ***
            Err.Raise vbObjectError + 1000, "HSV_to_RGB", "A Hue was given with no Saturation. This is invalid."
        End If
    Else
        Hue = (Hue Mod 360) / 60
        I = Int(Hue)    ' Return largest integer
        f = Hue - I     ' f is the fractional part of Hue
        p = Lightness * (1 - Saturation)
        q = Lightness * (1 - (Saturation * f))
        t = Lightness * (1 - (Saturation * (1 - f)))
        Select Case I
            Case 0
                Red = Lightness
                Green = t
                Blue = p
            Case 1
                Red = q
                Green = Lightness
                Blue = p
            Case 2
                Red = p
                Green = Lightness
                Blue = t
            Case 3
                Red = p
                Green = q
                Blue = Lightness
            Case 4
                Red = t
                Green = p
                Blue = Lightness
            Case 5
                Red = Lightness
                Green = p
                Blue = q
        End Select
    End If
    
    HSV = RGB(255 * Red, 255 * Green, 255 * Blue)
        
End Function
