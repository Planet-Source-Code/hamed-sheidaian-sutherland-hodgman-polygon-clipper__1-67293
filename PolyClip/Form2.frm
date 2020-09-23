VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Vertex
   x As Single
   y As Single
End Type
Const Max = 5
Dim edge(1 To 2) As Vertex
Dim VertexArray(1 To Max) As Vertex
