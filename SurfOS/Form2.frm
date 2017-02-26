VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vance"
   ClientHeight    =   495
   ClientLeft      =   5340
   ClientTop       =   9585
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Ask me anything"
      Top             =   120
      Width           =   7575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   7575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function chk()

If Text3.Text = "tell value of pi" Then
Text2.Text = "According to me, the value of pie, is 3.141592"
Call speak
End If

If Text3.Text = "tell your name" Then
Text2.Text = "My name is Vance"
Call speak
End If

If Text3.Text = "quit" Then
Text2.Text = "Goodbye"
Call speak
End
End If

If Text3.Text = "what is your name" Then
Text2.Text = "My name is Vance"
Call speak
End If

If Text3.Text = "what is your job" Then
Text2.Text = "I am here to interact with you, when you get bored!"
Call speak
End If

If Text3.Text = "how can you help me" Then
Text2.Text = "I'll help you to pass your time"
Call speak
End If

If Text3.Text = "help" Then
Text2.Text = "What should I do!"
Call speak
End If

If Text3.Text = "who is your creator" Then
Text2.Text = "My creator is Naveen. Visit surf O S.webs.com"
Call speak
End If

If Text3.Text = "tell time" Then
Text2.Text = Time
Call speak
End If

If Text3.Text = "laugh" Then
Text2.Text = "ha ha ha"
Call speak
End If

If Text3.Text = "tell date" Then
Text2.Text = Date
Call speak
End If

If Text3.Text = "open facebook" Then
Text2.Text = "Sure! Give me a second. But don't sit the whole day!"
Form1.b.Navigate2 ("www.facebook.com")
Call speak
End If

If Text3.Text = "thank you" Then
Text2.Text = "You're welcome"
Call speak
End If

If Text3.Text = "thanks" Then
Text2.Text = "Welcome!"
Call speak
End If

If Text3.Text = "thank you so much" Then
Text2.Text = "You're welcome"
Call speak
End If

If Text3.Text = "did you know i will get my result tomorrow" Then
Text2.Text = "Don't worry, everything will be alright! Cheer up!"
Call speak
End If

If Text3.Text = "ill get my result" Then
Text2.Text = "Don't worry, everything will be alright! Cheer up!"
Call speak
End If

If Text3.Text = "ill get my result tomorow" Then
Text2.Text = "Don't worry, everything will be alright! Cheer up!"
Call speak
End If

If Text3.Text = "i have a test tomorrow and i have studied" Then
Text2.Text = "good! All the best for tomorrow"
Call speak
End If

If Text3.Text = "i have a test tomorrow and i have not studied" Then
Text2.Text = "The please go and study. I'll be back after an hour."
Call speak
End
End If

If Text3.Text = "help me in an equation" Then
Text2.Text = "I'll hook you"
Call speak
Form1.b.Navigate2 ("www.wolframalpha.com")
End If

If Text3.Text = "tell india's capital" Then
Text2.Text = "New Delhi"
Call speak
End If

If Text3.Text = "who is anna hazare" Then
Text2.Text = "Anna Hazare is an activist fighting against corruption in India"
Call speak
End If

If Text3.Text = "who is rajeswari" Then
Text2.Text = "Rajeswari is Naveen's mother"
Call speak
End If

If Text3.Text = "who am i" Then
Text2.Text = "You are a, user of Surf O S,,.... Greetings!"
Call speak
End If

If Text3.Text = "how are you" Then
Text2.Text = "I'm fine! How are you?"
Call speak
End If

If Text3.Text = "how are you today" Then
Text2.Text = "I'm fine. What about you?"
Call speak
End If

If Text3.Text = "i'm fine" Then
Text2.Text = "Good!"
Call speak
End If

If Text3.Text = "im fine" Then
Text2.Text = "Great!"
Call speak
End If

If Text3.Text = "close internet" Then
Text2.Text = "OK!"
Form1.b.Navigate2 ("about:blank")
Call speak

End If







End Function

Private Function speak()
Dim sapi
Set sapi = CreateObject("SAPI.SpVoice")
sapi.speak Text2.Text
End Function

Private Sub Form_Load()
Text2.Text = "Ask me anything"
Call speak
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call chk
End If
End Sub
