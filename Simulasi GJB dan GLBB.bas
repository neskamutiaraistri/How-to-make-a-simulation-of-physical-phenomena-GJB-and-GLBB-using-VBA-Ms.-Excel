Attribute VB_Name = "Module1"
Sub ButtonGJB_Click()
 Range("B9").Value = 0 '0
 While Range("B9").Value < 10
 Range("B9") = Range("B9").Value + 0.05
 
 DoEvents
Wend
 
End Sub
Sub GLBB_Click()
 Range("I9").Value = 0 '0
 While Range("I9").Value < 10
 Range("I9") = Range("I9").Value + 0.05
 
 DoEvents
Wend
End Sub
