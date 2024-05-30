Attribute VB_Name = "modPublicVars"
'A couple constants to simplify tracking various variables
Public Const DRAWMETHOD_LUMINANCE As Long = 3
Public Const DRAWMETHOD_BARS As Long = 0

'Used to track the last options we used on the histogram form
Public lastHistSource As Long, lastHistMethod As Long


'We'll use this routine only to draw the gradient below the histogram window
'(like Photoshop does).  This code is old, but it works ;)
Public Sub DrawGradient(ByRef DstObject As PictureBox, ByVal Color1 As Long, ByVal Color2 As Long)
    'RGB() variables for each color
    Dim R As Long, G As Long, B As Long
    Dim R2 As Long, G2 As Long, B2 As Long
    
    'Extract the r,g,b values from the colors passed by the user
    R = Color1 Mod 256
    G = (Color1 \ 256) And 255
    B = (Color1 \ 65536) And 255
    R2 = Color2 Mod 256
    G2 = (Color2 \ 256) And 255
    B2 = (Color2 \ 65536) And 255
    
    'Calculation variables for the gradiency
    Dim VR As Single, VG As Single, VB As Single
    
    'Size of the picture box we'll be drawing to
    Dim iWidth As Long, iHeight As Long
    iWidth = DstObject.ScaleWidth
    iHeight = DstObject.ScaleHeight
    
    'Here, create a calculation variable for determining the step between
    'each level of the gradient
    VR = Abs(R - R2) / iWidth
    VG = Abs(G - G2) / iWidth
    VB = Abs(B - B2) / iWidth
    'If the second value is lower then the first value, make the step negative
    If R2 < R Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < B Then VB = -VB
    'Last, run a loop through the width of the picture box, incrementing the color as
    'we go (thus creating a gradient effect)
    Dim x As Long
    For x = 0 To iWidth
        R2 = R + VR * x
        G2 = G + VG * x
        B2 = B + VB * x
        DstObject.Line (x, 0)-(x, iHeight), RGB(R2, G2, B2)
    Next x
End Sub
