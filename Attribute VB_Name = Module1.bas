Attribute VB_Name = "Module1"

Option Explicit

Public Sub TestGitHub()
    ' Procedura testowa weryfikująca działanie kodu pobranego z repozytorium
    Dim userLogin As String
    Dim messageText As String
    
    userLogin = Environ$("USERNAME")
    messageText = "Repozytorium działa poprawnie!" & vbCrLf & _
                  "Uruchomiono przez użytkownika: " & userLogin & vbCrLf & _
                  "Data i czas: " & Now()
    
    MsgBox messageText, vbInformation, "GitHub VBA Test"
End Sub

Public Function DodajLiczby(ByVal a As Double, ByVal b As Double) As Double
    ' Prosta funkcja użytkownika (UDF) do testów jednostkowych
    DodajLiczby = a + b
End Function
public
