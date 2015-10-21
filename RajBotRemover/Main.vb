Imports Microsoft.Win32
Imports System
Imports System.IO
Imports System.Text
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Security.Permissions
Imports System.Runtime.InteropServices



Public Class Main
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)> _
    Shared Function MoveFile(ByVal src As String, ByVal dst As String) As Boolean
    End Function
    <DllImport("shlwapi.dll", EntryPoint:="PathFileExistsW",  SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Shared Function PathFileExists(<MarshalAs(UnmanagedType.LPTStr)> ByVal pszPath As String) As <MarshalAs(UnmanagedType.Bool)> Boolean


    End Function
    Private Sub CheckInfected()
        Dim regk As RegistryKey
        Dim loadval As String

        regk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows", False)
        loadval = regk.GetValue("Load")

        If (loadval IsNot Nothing AndAlso loadval.EndsWith("msnuv.exe")) Or My.Computer.FileSystem.FileExists("C:\\ProgramData\\msnuv.exe") Then
            Label1.Text = "Your Machine is Infected"
            Label1.ForeColor = Color.Red
            btRemovePC.Enabled = True
            btCleanDrives.Enabled = False
        Else
            Label1.ForeColor = Color.DarkGreen
            Label1.Text = "Your Machine is clean"
            btRemovePC.Enabled = False
            btCleanDrives.Enabled = True
        End If
    End Sub

    Private Sub CleanPC()
        'Termina il processo msiexec.exe che tiene aperto msnuv.exe
        txtlog.Clear()
        If My.Computer.FileSystem.FileExists("C:\\ProgramData\\msnuv.exe") Then
            Dim processi = Process.GetProcessesByName("msiexec")
            For i As Integer = 0 To processi.Count - 1
                processi(i).Kill()
                processi(i).WaitForExit()
                txtlog.AppendText("Killing process " & processi(i).Id & vbNewLine)
            Next
            'Togli l'attributo nascosto e di sistema a C:\ProgramData\msnuv.exe
            Dim attrs As FileAttributes

            attrs = File.GetAttributes("C:\\ProgramData\\msnuv.exe")
            attrs = attrs And Not (FileAttributes.System Or FileAttributes.Hidden)
            txtlog.AppendText("Removing System&Hidden attributes from C:\\ProgramData\\msnuv.exe" & vbNewLine)
            File.SetAttributes("C:\\ProgramData\\msnuv.exe", attrs)
            'Cancella msnuv.exe
            txtlog.AppendText("Deleting C:\\ProgramData\\msnuv.exe" & vbNewLine)
            File.Delete("C:\\ProgramData\\msnuv.exe")
        End If
        'Ripristina i permessi di modifica all'utente corrente sulla chiave Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows
        Dim user As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim regk As RegistryKey



        'Dim f As New RegistryPermission(RegistryPermissionAccess.AllAccess, "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows")

        txtlog.AppendText("Restoring permissions on registry key HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows" & vbNewLine)
        Dim err As String
        err = ""
        RegistryOwnership.SeizeRegistryKey(BaseKeySelector.CurrentUser, "Software\Microsoft\Windows NT\CurrentVersion\Windows", err)
        regk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows", RegistryKeyPermissionCheck.ReadWriteSubTree)


        'Elimina il valore stringa Load
        txtlog.AppendText("Removing registry value Load from Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows" & vbNewLine)
        'regk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows", True)

        regk.DeleteValue("Load")
        txtlog.AppendText("Done")

        CheckInfected()

    End Sub

    Private Sub CleanUSB()
        txtlog.Clear()
        For Each d As System.IO.DriveInfo In My.Computer.FileSystem.Drives
            If d.DriveType = DriveType.Removable Then
                txtlog.AppendText("Checking " & d.Name & vbNewLine)
                If PathFileExists(d.Name & Chr(&HA0)) Then
                    txtlog.AppendText(d.Name & " is infected, cleaning" & vbNewLine)
                    For Each f In My.Computer.FileSystem.GetFiles(d.Name)
                        ' txtlog.AppendText(f & vbNewLine)
                        If f.EndsWith(".lnk") Then
                            txtlog.AppendText("Delete " & f & vbNewLine)
                            My.Computer.FileSystem.DeleteFile(f)
                        End If
                    Next
                    

                    MoveFile(d.Name & Chr(&HA0), d.Name & "_____TMP")

                    For Each f In My.Computer.FileSystem.GetFiles(d.Name & "_____TMP")
                        If My.Computer.FileSystem.GetFileInfo(f).Length > 10 * 1024 * 1024 Then
                            Dim reader As New BinaryReader(File.Open(f, FileMode.Open))
                            Dim magic As Integer = reader.ReadInt32()
                            reader.Close()

                            If magic = 9460301 Then
                                My.Computer.FileSystem.DeleteFile(f)
                            End If
                        End If
                    Next

                    'My.Computer.FileSystem.RenameDirectory()
                    If My.Computer.FileSystem.FileExists(d.Name & "_____TMP" & "\desktop.ini") Then
                        My.Computer.FileSystem.DeleteFile(d.Name & "_____TMP" & "\desktop.ini")
                    End If
                    If My.Computer.FileSystem.FileExists(d.Name & "_____TMP" & "\IndexerVolumeGuid") Then
                        My.Computer.FileSystem.DeleteFile(d.Name & "_____TMP" & "\IndexerVolumeGuid")
                    End If



                    For Each fX In My.Computer.FileSystem.GetFiles(d.Name & "_____TMP")
                        Dim f As String = System.IO.Path.GetFileName(fX)
                        My.Computer.FileSystem.MoveFile(d.Name & "_____TMP" & "\" & f, d.Name & f)
                    Next
                    For Each fX In My.Computer.FileSystem.GetDirectories(d.Name & "_____TMP")
                        Dim f As String = System.IO.Path.GetFileName(fX)
                        My.Computer.FileSystem.MoveDirectory(d.Name & "_____TMP" & "\" & f, d.Name & f)
                    Next
                    My.Computer.FileSystem.DeleteDirectory(d.Name & "_____TMP", FileIO.DeleteDirectoryOption.ThrowIfDirectoryNonEmpty)
                    txtlog.AppendText(d.Name & " cleaned")
                Else
                    txtlog.AppendText(d.Name & " is clean")
                End If
            End If
        Next

        'Elimina da dentro la cartella speciale i file desktop.ini, il file puntato dal lnk e IndexerVolumeGuid
        'Elimina il file lnk
        'Sposta i file dalla cartella \xA0 sulla root della pennetta
        'Elimina la cartella \xA0

    End Sub
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.Text = "Please wait..."
        CheckInfected()


    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub btRemovePC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRemovePC.Click
        CleanPC()
    End Sub

    Private Sub btCleanDrives_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCleanDrives.Click
        CleanUSB()
    End Sub

    Private Sub txtlog_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtlog.TextChanged

    End Sub
End Class
