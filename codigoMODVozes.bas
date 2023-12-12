Attribute VB_Name = "MODVozes"
'@Folder("SGES2020")
Option Explicit

Public Sub TalkToMe2()
    Application.Speech.Speak "Excel is talking to me", speakasync:=True
    MsgBox "test"
End Sub
Public Sub AvailableVoices()
    Dim i     As Long
    Dim Voc   As SpeechLib.SpVoice
    Set Voc = New SpVoice
    Debug.Print Voc.GetVoices.Count & " available voices:"
    For i = 0 To Voc.GetVoices.Count - 1
        Set Voc.Voice = Voc.GetVoices.Item(i)
        Debug.Print " " & i & " - " & Voc.Voice.GetDescription
        Voc.Speak "test audio"
    Next i
End Sub


Public Sub ChangeVoiceDemo()
    SuperTalk "Excel is talking to me.", "BOY", 2, 100
    SuperTalk "Excel is talking to me.", "GIRL", 2, 100
    SuperTalk "Excel is talking to me.", "BOY", -10, 30
    SuperTalk "Excel is talking to me.", "GIRL", 10, 70
End Sub

Private Sub SuperTalk(Words As String, Person As String, Rate As Long, Volume As Long)
    Dim Voc   As SpeechLib.SpVoice
    Set Voc = New SpVoice

    With Voc
        If UCase$(Person) = "BOY" Then
            Set .Voice = .GetVoices.Item(2)
        ElseIf UCase$(Person) = "GIRL" Then
            Set .Voice = .GetVoices.Item(0)
        End If
        .Rate = Rate
        .Volume = Volume
        .Speak Words
    End With
End Sub
