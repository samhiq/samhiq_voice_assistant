Set voice = CreateObject("SAPI.SpVoice")
voice.Rate = 1
voice.Volume = 90
Say = InputBox("Ethical Hacker Samhiq please give me Order iam your assistant", "Md Sameer iqbal(Samhiq) Voice Robot", "Hii Ethical Hacker")
If (Len(Say) > 0) Then
	voice.Speak Say
End If 


