Attribute VB_Name = "modUpdate"
Public CurrentVersion As String
Public Path As String
Public Section(1 To 10) As String
Type FileData
    Name As String
    Version As String
End Type
Public Files(0 To 50) As FileData




Sub Main()
Path = "http://nerfmud.r0x0rsoft.com/upd/"

    frmUpdate.Show
    DoEvents
    If Dir("version.dat") = "" Then
        CurrentVersion = ""
    Else
        Open "version.dat" For Input As #1
        If Not EOF(1) Then
            Line Input #1, CurrentVersion
        End If
        Close #1
    End If

    If CurrentVersion <> "" Then
        frmUpdate.curver = "Version " & CurrentVersion
        frmUpdate.list.AddItem "Current version is " & CurrentVersion
    Else
        frmUpdate.curver = "No version info"
    End If
End Sub

Public Function DownloadFile(srcFileName As String, targetFileName As String)
  Dim B() As Byte
  Dim FID As Byte
  B() = frmUpdate.Inet.OpenURL(Path & srcFileName, icByteArray)
  FID = FreeFile
  Open targetFileName For Binary Access Write As #FID
    Put #FID, , B()
  Close #FID
  DoEvents
End Function


Sub DoFiles()
Dim fSt As String
    Dim i As Long
    i = 0
    If Dir("files.dat") <> "" Then
        Open "files.dat" For Input As #22
            Do Until EOF(22)
                Line Input #22, fSt
                GetSections fSt, ","
                If Section(1) <> "" Then
                    Files(i).Name = Section(1)
                    Files(i).Version = Section(2)
                    
                End If
                i = i + 1
            Loop
        Close #22
    End If
End Sub


Sub GetSections(St, Deliminator As String)
    Dim a As Integer, B As Integer, C As Integer
    B = 1
    Erase Section
    For a = 1 To 10
TryAgain:
        C = InStr(B, St, Deliminator)
        If C - B = 0 Then B = B + 1: GoTo TryAgain
        If C <> 0 Then
                Section(a) = Mid$(St, B, C - B)
        Else
                Section(a) = Mid$(St, B, Len(St) - B + 1)
                Exit For
        End If
        B = C + 1
    Next a
End Sub


Sub CheckFile(Name As String, Version As String)
    Dim a As Long
    For a = 0 To 50
        If Files(a).Name = Name Then
            If Files(a).Version < Version Then
                frmUpdate.list.AddItem "Updating " & Name & " to version " & Version
                DownloadFile Name, Name
                Files(a).Version = Version
            End If
            GoTo done
        End If
    Next a
    frmUpdate.list.AddItem "Downloading " & Name
    DownloadFile Name, Name
    For a = 0 To 50
        If Files(a).Name = "" Then
            Files(a).Name = Name
            Files(a).Version = Version
            Exit For
        End If
    Next a
done:
End Sub

Sub CreateFileList()
Dim a As Long
Open "files.dat" For Output As #23
For a = 0 To 50
    If Files(a).Name <> "" Then
        Print #23, Files(a).Name & "," & Files(a).Version
    End If
Next a
Close #23
End Sub
