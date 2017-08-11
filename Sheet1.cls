VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub CommandButton1_Click()
   Dim sname As String
   Dim gname As String
   Dim mstyle As Integer
   Dim size As Integer
   Dim mcolor As Long
   Dim cnt As Integer
   Dim mysheet As String

   sname = Range("sname")
   gname = Range("gname")
   mstyle = getMarkID(Range("mstyle"))
   msize = Range("msize")
   mcolor = getMarkColor(Range("mcolor"))
   cnt = 100 'ç≈ëÂÉfÅ[É^êî
   

   

   mysheet = ActiveSheet.Name

   Sheets(sname).Select
   ActiveSheet.ChartObjects(gname).Activate
   
   For i = 1 To cnt
      On Error GoTo BREAK
   
      ActiveChart.SeriesCollection(i).Select
      Selection.Format.Line.Visible = msoFalse
      Selection.MarkerStyle = mstyle
      Selection.MarkerSize = msize
      Selection.Format.Fill.ForeColor.RGB = mcolor
   Next i
BREAK:

   ActiveChart.PlotArea.Select
   ActiveChart.ChartArea.Select
   ActiveChart.Deselect

   Sheets(mysheet).Select
End Sub


Function getMarkID(smark)
   Select Case smark
      Case "Å°": getMarkID = 1
      Case "Åü": getMarkID = 2
      Case "Å£": getMarkID = 3
      Case "Å|": getMarkID = 7
      Case "Åú": getMarkID = 8
      Case Else: getMarkID = 8
   End Select
End Function


Function getMarkColor(cmark)

   Select Case cmark
      Case "Red": getMarkColor = RGB(255, 0, 0)
      Case "Blue": getMarkColor = RGB(0, 0, 255)
      Case "Green": getMarkColor = RGB(0, 190, 0)
      Case "Pink": getMarkColor = RGB(0, 255, 255)
      Case "Yellow": getMarkColor = RGB(255, 255, 0)
      Case "Black": getMarkColor = RGB(0, 0, 0)
      Case Else: getMarkColor = RGB(255, 0, 0)
   End Select
End Function
