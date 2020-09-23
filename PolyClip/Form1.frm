VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'**************************
   'Sutherland-Hodgman Polygon Clipping Algorithm v1.0
   'By : Hamed Sheidaian
   'December 2006
   'Tehran , Iran
'**************************


Private Type point
  x As Integer
  y As Integer
End Type
Private Type edge
  A As point
  B As point
End Type
Dim ed(1 To 4) As edge
Dim poly(1 To 4) As point
Dim Res(1 To 8) As point

Private Sub Form_Load()
  poly(1).x = 150
  poly(1).y = 20
  poly(2).x = 280
  poly(2).y = 150
  poly(3).x = 180
  poly(3).y = 260
  poly(4).x = 50
  poly(4).y = 250
End Sub

Private Sub Form_Paint()
   ed(1).A.x = 100
   ed(1).A.y = 100
   ed(1).B.x = 300
   ed(1).B.y = 100
   ed(2).A.x = 300
   ed(2).A.y = 100
   ed(2).B.x = 300
   ed(2).B.y = 200
   ed(3).A.x = 300
   ed(3).A.y = 200
   ed(3).B.x = 100
   ed(3).B.y = 200
   ed(4).A.x = 100
   ed(4).A.y = 200
   ed(4).B.x = 100
   ed(4).B.y = 100

   Call drawPoly(poly, 4)
   Dim e As edge
   Dim k As Integer
   Dim n As Integer
   
   n = 4
   
   For i = 1 To n
     Res(i) = poly(i)
   Next
   
   For k = 1 To 4
       Line (ed(k).A.x, ed(k).A.y)-(ed(k).B.x, ed(k).B.y), vbRed
       Counter = n
       n = 0
       e.A.x = Res(Counter).x
       e.A.y = Res(Counter).y
       Dim Temp(1 To 8) As point
       For j = 1 To Counter
             e.B.x = Res(j).x
             e.B.y = Res(j).y
             If Inside(e.B, k) = True Then
                If Inside(e.A, k) = True Then
                   n = n + 1
                   Temp(n) = e.B
                Else
                    If (ed(k).A.y = ed(k).B.y) Then
                      iy = ed(k).A.y
                      ix = e.A.x + (ed(k).A.y - e.A.y) * (e.B.x - e.A.x) / (e.B.y - e.A.y)
                    Else
                      ix = ed(k).A.x
                      iy = e.A.y + (ed(k).A.x - e.A.x) * (e.B.y - e.A.y) / (e.B.x - e.A.x)
                    End If
                    n = n + 1
                    Temp(n).x = ix
                    Temp(n).y = iy
                    n = n + 1
                    Temp(n) = e.B
                End If
            Else
                If Inside(e.A, k) = True Then
                    If (ed(k).A.y = ed(k).B.y) Then
                      iy = ed(k).A.y
                      ix = e.A.x + (ed(k).A.y - e.A.y) * (e.B.x - e.A.x) / (e.B.y - e.A.y)
                    Else
                      ix = ed(k).A.x
                      iy = e.A.y + (ed(k).A.x - e.A.x) * (e.B.y - e.A.y) / (e.B.x - e.A.x)
                    End If
                    n = n + 1
                    Temp(n).x = ix
                    Temp(n).y = iy
                End If
            End If
            e.A = e.B
       Next
       For i = 1 To n
          Res(i) = Temp(i)
       Next
   Next
     
   Form1.DrawWidth = 3
   Call drawPoly(Res, n)
End Sub
Private Function Inside(tv As point, c As Integer) As Boolean
  If (ed(c).B.x > ed(c).A.x) Then
    If (tv.y >= ed(c).A.y) Then
       Inside = True
       Exit Function
    End If
  End If
  If (ed(c).B.x < ed(c).A.x) Then
    If (tv.y <= ed(c).A.y) Then
       Inside = True
       Exit Function
    End If
  End If
  If (ed(c).B.y > ed(c).A.y) Then
    If (tv.x <= ed(c).B.x) Then
       Inside = True
       Exit Function
    End If
  End If
  If (ed(c).B.y < ed(c).A.y) Then
    If (tv.x >= ed(c).B.x) Then
       Inside = True
       Exit Function
    End If
  End If
  Inside = False
End Function


Private Sub drawPoly(p() As point, n As Integer)
  For i = 1 To n - 1
    Line (p(i).x, p(i).y)-(p(i + 1).x, p(i + 1).y), vbBlue
  Next
  Line (p(n).x, p(n).y)-(p(1).x, p(1).y), vbBlue
End Sub
