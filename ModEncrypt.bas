Attribute VB_Name = "ModEncrypt"
Dim MCount(100000) As Integer
Dim out As String

Sub reverse(txt As String, txtoutput As RichTextBox, pbar As ProgressBar)
d = ""
pbar.Max = Len(txt)
For x = Len(txt) To 1 Step -1
  s = Mid(txt, x, 1)
  pbar.Value = -(x)
    d = d & s
Next
pbar.Value = 0
txtoutput.Text = d
End Sub


Sub Encrypt(txt As String, txtoutput As RichTextBox, Strgth As Long, pbar As ProgressBar)
Dim gs As String
gs = ""
pbar.Max = Len(txt)
For x = 1 To Len(txt)
  f = Mid(txt, x, 1)
  s = Asc(f)
  s = s * 2
  dt = Chr(s)
  gs = gs & dt
For md = 1 To Strgth
MCount(md) = Int(255 * Rnd) + 1
ld = Chr(MCount(md))
gs = gs & ld
Next
  pbar.Value = x
Next
pbar.Value = 0
txtoutput.Text = gs
End Sub

Sub Decrypt(txt As String, txtoutput As RichTextBox, Strgth As Long, pbar As ProgressBar)
gs = ""
pbar.Max = Len(txt)
For x = 1 To Len(txt) Step Strgth + 1
 l = Mid(txt, x, 1)
 s = Asc(l)
 s = s / 2
 f = Chr(s)
 gs = gs & f
 DoEvents
 pbar.Value = x
Next
pbar.Value = 0
txtoutput.Text = gs
End Sub
