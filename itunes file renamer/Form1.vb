Imports System.Collections.ObjectModel

Public Class Form1
    Const Movie As Integer = 1
    Const TVShow As Integer = 3

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim iTunesApp As New iTunesLib.iTunesApp
        Dim mainLibrary = iTunesApp.LibraryPlaylist
        Dim mainLibrarySource = iTunesApp.LibrarySource
        Dim tracks = mainLibrary.Tracks
        Dim numTracks = tracks.Count
        Dim numPlaylistsCreated = 0
        'Dim protokoll As String = ""

        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = tracks.Count
        Me.ProgressBar1.Step = 1

        For Each currTrack As Object In tracks
            Me.txtFiles.BackColor = Color.White
            Me.txtFiles.Text = currTrack.Name
            Me.ProgressBar1.PerformStep()
            Application.DoEvents()
            If currTrack.Kind = iTunesLib.ITTrackKind.ITTrackKindFile Then

                If currTrack.VideoKind = 3 Then
                    If Not IsNothing(currTrack.Location) Then
                        'TV Show

                        Me.txtFiles.BackColor = Color.Yellow
                        Me.txtFiles.Text = currTrack.Location
                        Application.DoEvents()

                        Dim DateiPfad As String = currTrack.Location.ToString
                        Dim pfad As String = System.IO.Path.GetDirectoryName(DateiPfad)
                        Dim Dateiname As String = System.IO.Path.GetFileNameWithoutExtension(DateiPfad)
                        Dim Dateiname_new As String = Dateiname
                        Dim extension As String = System.IO.Path.GetExtension(DateiPfad)

                        If Not Dateiname.StartsWith("S") And IsNumeric(Mid(Dateiname, 1, 1)) And InStr(Mid(Dateiname, 1, 4), "-") > 0 Then

                            Dim tempFilename() As String = Dateiname.Split("-")
                            tempFilename(0) = "S" & Format(Val(tempFilename(0)), "00")
                            tempFilename(1) = "E" & tempFilename(1)

                            Dateiname_new = ""
                            For i As Integer = 0 To tempFilename.Length - 1
                                Dateiname_new &= tempFilename(i)
                                If i > 0 And i < tempFilename.Length - 1 Then
                                    Dateiname_new &= "-"
                                End If
                            Next
                        End If
                        If My.Computer.FileSystem.FileExists(pfad & "\" & Dateiname & extension) Then
                            If Dateiname <> Dateiname_new Then
                                Me.txtFiles.BackColor = Color.LightGreen
                                Application.DoEvents()
                                My.Computer.FileSystem.RenameFile(pfad & "\" & Dateiname & extension, Dateiname_new & extension)
                                currTrack.Location = pfad & "\" & Dateiname_new & extension
                            End If

                        Else
                            'Stop
                        End If
                    Else
                        Dim files As ReadOnlyCollection(Of String)
                        'If InStr(currTrack.Show, "familie", CompareMethod.Text) Then Stop
                        If My.Computer.FileSystem.DirectoryExists("Z:\iTunes Media\TV Shows\" & Mid(currTrack.Show, 1, 40) & "\Season " & currTrack.SeasonNumber & "\") Then
                            files = My.Computer.FileSystem.GetFiles("Z:\iTunes Media\TV Shows\" & Mid(currTrack.Show, 1, 40) & "\Season " & currTrack.SeasonNumber & "\", FileIO.SearchOption.SearchAllSubDirectories, "S" & Format(currTrack.SeasonNumber, "00") & "E" & Format(currTrack.EpisodeNumber, "00") & "*.*")
                            If files.Count = 0 Then
                                files = My.Computer.FileSystem.GetFiles("Z:\iTunes Media\TV Shows\" & Mid(currTrack.Show, 1, 40) & "\Season " & currTrack.SeasonNumber & "\", FileIO.SearchOption.SearchAllSubDirectories, Format(currTrack.SeasonNumber, "0") & "-" & Format(currTrack.EpisodeNumber, "00") & "*.*")
                            End If

                            If files.Count = 1 Then
                                'Stop
                                currTrack.Location = files(0)
                            Else
                                'Stop
                            End If
                        End If
                    End If
                Else
                        'protokoll &= currTrack.Name & vbCrLf
                    End If
                Else
                    'Stop
                End If
        Next
        '        If protokoll.Trim.Length > 0 Then
        ' MsgBox(protokoll)
        ' Else
        'MsgBox("fertig")
        ' End If
        End
    End Sub
End Class