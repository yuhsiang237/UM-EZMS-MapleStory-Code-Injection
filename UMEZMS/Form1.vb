Imports System.IO
Public Class Form1
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Declare Function CEInitialize Lib "ceautoassembler" (ByVal ngPassedPID As Integer, ByVal Phandle As Integer) As Integer
    Public Declare Function CEAutoAsm Lib "ceautoassembler" (ByVal Script As String, ByVal AllocID As Boolean, ByVal Alloc As Integer) As Boolean
    Dim KEYDOWNVALUE As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '遊戲鎖定區
        OpenProcessByWindow("MapleStory", "StartUpDlgClass")
        If pid <> 0 Then
            CEInitialize(pid, hprocess)
            MsgBox("遊戲鎖定成功囉！", 64, "UM.MSG")
            ToolStripStatusLabel1.Text = "Success get!"
        Else
            MsgBox("鎖定遊戲失敗！", 16, "UM.MSG")
            ToolStripStatusLabel1.Text = "ERROR!"
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)


    End Sub

    Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.Click

    End Sub
    Dim filenum As Integer
    Dim filename As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''====清空buffer====
        'filenum = FreeFile()
        ''====寫入TEXT====
        'System.IO.File.CreateText(".\Script\" + TextBox3.Text + ".txt")
        'FileOpen(FreeFile(), ".\Script\" + TextBox3.Text + ".txt", OpenMode.Output)
        'Write(filenum, TextBox1.Text) '寫入存放位置
        'FileClose(filenum)
        ''=====
        'CheckedListBox1.Items.Add(TextBox3.Text)

        LogWrite(TextBox3.Text, TextBox2.Text)
        TextBox2.Clear()
        TextBox3.Clear()



    End Sub
    Private Function LogRead(ByVal fileName As String)
        Dim line As String

        ' Create new StreamReader instance with Using block.
        Using reader As System.IO.StreamReader = New System.IO.StreamReader(".\Script\" + fileName)
            ' Read one line from file
            line = reader.ReadToEnd
        End Using

        ' Write the line we read from "file.txt"
        '  MessageBox.Show(line)
        Return line
    End Function

    Private Sub LogWrite(ByVal fileName As String, ByVal script As String)
        Dim FilePath As String = ".\Script\" + fileName + ".txt"
        '移除以前的檔案
        Try
            My.Computer.FileSystem.DeleteFile(FilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
        Catch ex As Exception

        End Try
        '寫入
        Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(FilePath, True)
        sw.WriteLine(script)
        sw.Flush()
        sw.Close()

        CheckedListBox1.Items.Add(TextBox3.Text + ".txt")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            Dim indexName As String = CheckedListBox1.SelectedItem.ToString
            Dim FileExists As Boolean

            Dim filePath = ".\Script\" + indexName

            FileExists = My.Computer.FileSystem.FileExists(filePath)

            If FileExists = False Then

                '檔案不存在
                ToolStripStatusLabel1.Text = "Remove Error!"
                Return

            Else

                '檔案 存在則刪除檔案 

                My.Computer.FileSystem.DeleteFile(filePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

            End If
            CheckedListBox1.Items.RemoveAt(CheckedListBox1.SelectedIndex)
            ToolStripStatusLabel1.Text = "Remove complete"

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Clear()
    End Sub
    Private Sub testEven(ByVal sender As Object, ByVal e As System.EventArgs)



    End Sub
    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged




    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        '=====判斷索引名稱=====
        Dim indexSelect As Integer = CheckedListBox1.SelectedIndex
        Dim indexName As String = CheckedListBox1.SelectedItem.ToString
        '=====暫存取得的Script=
        Dim scriptTmp As String = LogRead(indexName)
        '=====判斷是否選取=====
        If (CheckedListBox1.GetItemChecked(indexSelect) = False) Then
            'MessageBox.Show("IN" + scriptTmp)
            CEAutoAsm(scriptTmp, True, 0)
        Else
            'MessageBox.Show("out" + scriptTmp)
            CEAutoAsm(scriptTmp, False, 0)
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dir As String = ".\Script\"
        Dim fn As String                '取得檔案完整路徑檔名
        Dim nn As String                '取得檔名
        'Dim SourceFile, DestinationFile As String


        CheckedListBox1.Items.Clear()          '清除ListBox
        For Each fn In My.Computer.FileSystem.GetFiles(dir, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
            nn = My.Computer.FileSystem.GetName(fn)         '取得檔案名稱
            CheckedListBox1.Items.Add(nn)                          '列入檔名
            ''-------------------------------------------
            'SourceFile = fn                                 '來源完整路徑
            'DestinationFile = "c:目的" & nn               '指定路徑檔名
            'FileCopy(SourceFile, DestinationFile)           'Copy source to target
        Next

        '====名言
        Dim i As New Random()

        Dim value As Integer = i.Next(1, 4)

        If (value = 1) Then
            Label13.Text = "問題發生的當下，沒辦法解決" + vbCrLf + "是因為你的思想尚未突破，需" + vbCrLf + "要一點耐心，等你思想更成熟" + vbCrLf + "時就會迎刃而解。"

        ElseIf (value = 2) Then
            Label13.Text = "不認真看待小事的人，大事上" + vbCrLf + "也無法被信任"

        ElseIf (value = 3) Then
            Label13.Text = "邏輯會把你從 A 帶到 B，但" + vbCrLf + "想像力會帶你到任何角落。"

        ElseIf (value = 4) Then
            Label13.Text = "如果你不能簡單說明，代表你" + vbCrLf + "還未完全瞭解它。"


        End If



    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://umworksite.blogspot.tw/")
    End Sub

    Dim timerTime As Integer = 60
    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If (CheckBox1.Checked = True) Then
            timerTime = TextBox4.Text
            Label9.Text = TextBox4.Text
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        timerTime -= 1
        Label9.Text = timerTime
        If (timerTime = 0) Then
            Shell("ShutDown -t 1")
            Timer1.Enabled = False
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If (CheckBox2.Checked = True) Then
            CK1.Interval = TextBox5.Text
            CK1.Enabled = True
        Else
            CK1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If (CheckBox3.Checked = True) Then
            CK2.Interval = TextBox6.Text
            CK2.Enabled = True
        Else
            CK2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If (CheckBox4.Checked = True) Then
            CK3.Interval = TextBox7.Text
            CK3.Enabled = True
        Else
            CK3.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If (CheckBox5.Checked = True) Then
            CK4.Interval = TextBox8.Text
            CK4.Enabled = True
        Else
            CK4.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If (CheckBox6.Checked = True) Then
            CK5.Interval = TextBox9.Text
            CK5.Enabled = True
        Else
            CK5.Enabled = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If (CheckBox7.Checked = True) Then
            CK6.Interval = TextBox10.Text
            CK6.Enabled = True
        Else
            CK6.Enabled = False
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If (CheckBox8.Checked = True) Then
            CK7.Interval = TextBox11.Text
            CK7.Enabled = True
        Else
            CK7.Enabled = False
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If (CheckBox9.Checked = True) Then
            CK8.Interval = TextBox12.Text
            CK8.Enabled = True
        Else
            CK8.Enabled = False
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If (CheckBox10.Checked = True) Then
            CK9.Interval = TextBox13.Text
            CK9.Enabled = True
        Else
            CK9.Enabled = False
        End If
    End Sub

    Private Sub CK1_Tick(sender As Object, e As EventArgs) Handles CK1.Tick

        Dim word As String = ComboBox1.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

        PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
        ' PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
        '   PostMessage(KEYDOWNVALUE, WM_KEYUP, KEY, MakeKeyLparam(KEY, WM_KEYUP))

    End Sub

    Private Sub CK2_Tick(sender As Object, e As EventArgs) Handles CK2.Tick
        Dim word As String = ComboBox2.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If
 PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK3_Tick(sender As Object, e As EventArgs) Handles CK3.Tick
        Dim word As String = ComboBox3.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

         PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK4_Tick(sender As Object, e As EventArgs) Handles CK4.Tick
        Dim word As String = ComboBox4.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

        PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK5_Tick(sender As Object, e As EventArgs) Handles CK5.Tick
        Dim word As String = ComboBox5.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

   PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK6_Tick(sender As Object, e As EventArgs) Handles CK6.Tick
        Dim word As String = ComboBox6.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

         PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK7_Tick(sender As Object, e As EventArgs) Handles CK7.Tick
        Dim word As String = ComboBox7.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If
 PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK8_Tick(sender As Object, e As EventArgs) Handles CK8.Tick
        Dim word As String = ComboBox8.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If

       PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub CK9_Tick(sender As Object, e As EventArgs) Handles CK9.Tick
        Dim word As String = ComboBox9.Text
        Dim KEY As Integer
        If (word = "1") Then
            KEY = VK_1
        ElseIf (word = "2") Then
            KEY = VK_2
        ElseIf (word = "3") Then
            KEY = VK_3
        ElseIf (word = "4") Then
            KEY = VK_4
        ElseIf (word = "5") Then
            KEY = VK_5

        ElseIf (word = "6") Then
            KEY = VK_6
        ElseIf (word = "7") Then
            KEY = VK_7
        ElseIf (word = "8") Then
            KEY = VK_8
        ElseIf (word = "9") Then
            KEY = VK_9
        ElseIf (word = "0") Then
            KEY = VK_0
        ElseIf (word = "A") Then
            KEY = VK_A
        ElseIf (word = "B") Then
            KEY = VK_B

        ElseIf (word = "C") Then
            KEY = VK_C

        ElseIf (word = "D") Then
            KEY = VK_D

        ElseIf (word = "E") Then
            KEY = VK_E

        ElseIf (word = "F") Then
            KEY = VK_F

        ElseIf (word = "G") Then
            KEY = VK_G

        ElseIf (word = "H") Then
            KEY = VK_H

        ElseIf (word = "I") Then
            KEY = VK_I

        ElseIf (word = "J") Then
            KEY = VK_J
        ElseIf (word = "K") Then
            KEY = VK_K
        ElseIf (word = "L") Then
            KEY = VK_L
        ElseIf (word = "M") Then
            KEY = VK_M
        ElseIf (word = "N") Then
            KEY = VK_N
        ElseIf (word = "O") Then
            KEY = VK_O
        ElseIf (word = "P") Then
            KEY = VK_P
        ElseIf (word = "Q") Then
            KEY = VK_Q
        ElseIf (word = "R") Then
            KEY = VK_R
        ElseIf (word = "S") Then
            KEY = VK_S
        ElseIf (word = "T") Then
            KEY = VK_T
        ElseIf (word = "U") Then
            KEY = VK_U
        ElseIf (word = "V") Then
            KEY = VK_V
        ElseIf (word = "W") Then
            KEY = VK_W
        ElseIf (word = "X") Then
            KEY = VK_X
        ElseIf (word = "Y") Then
            KEY = VK_Y
        ElseIf (word = "Z") Then
            KEY = VK_Z
        End If
 PostMessage(KEYDOWNVALUE, WM_KEYDOWN, KEY, MakeKeyLparam(KEY, WM_KEYDOWN))
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        KEYDOWNVALUE = FindWindow(vbNullString, TextBox14.Text)

        Label12.Text = TextBox14.Text + "已鎖定!"



    End Sub
End Class
