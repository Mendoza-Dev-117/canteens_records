Module Program
    Sub Main()
        On Error GoTo ERR_LABEL

        Dim lRet As String
        Dim strSQL As String
        Dim Term_clock As String
        Dim date_rec As String
        Dim time_rec As String
        Dim Company As String
        Dim payroll_num As String
        Dim I_O As String
        Dim Group As String
        Dim Department As String
        Dim Facility As String
        Dim Position As String
        Dim Bus_Unit As String

        Dim CNT1 As Integer
        Dim CNT2 As Integer
        Dim CNT3 As Integer
        Dim CNT4 As Integer
        Dim CNT5 As Integer
        Dim CNT6 As Integer
        Dim CNT7 As Integer
        Dim CNT8 As Integer
        Dim CNT9 As Integer
        Dim CNT10 As Integer

        Dim text_Detail As String
        Dim textDetail(3000) As String
        Dim TextPath As String
        Dim TextOutPath As String
        Dim CopyPath As String
        Dim TextLine As String
        Dim fileNo As Integer
        Dim ifileNo As Integer
        Dim I As Integer
        Dim J As Integer
        Dim ITEM_NAME As String
        Dim wk_copypath As String

        TextPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HP.txt")
        TextOutPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reloj.dat")

        If File.Exists(TextPath) Then
            Using fileReader As New StreamReader(TextPath)
                Using fileWriter As New StreamWriter(TextOutPath)
                    I = 1

                    Do While Not fileReader.EndOfStream
                        TextLine = fileReader.ReadLine()

                        CNT1 = InStr(TextLine, ",")
                        If Len(TextLine) > 0 Then
                            Term_clock = fStrField(TextLine, 1, CNT1 - 1)
                            CNT2 = InStr(CNT1 + 1, TextLine, ",")
                            date_rec = fStrField(TextLine, CNT1 + 1, CNT2 - CNT1 - 1)
                            CNT3 = InStr(CNT2 + 1, TextLine, ",")
                            time_rec = fStrField(TextLine, CNT2 + 1, CNT3 - CNT2 - 1)
                            CNT4 = InStr(CNT3 + 1, TextLine, ",")
                            Company = fStrField(TextLine, CNT3 + 1, CNT4 - CNT3 - 1)
                            CNT5 = InStr(CNT4 + 1, TextLine, ",")
                            payroll_num = fStrField(TextLine, CNT4 + 1, CNT5 - CNT4 - 1)
                            CNT6 = InStr(CNT5 + 1, TextLine, ",")
                            I_O = fStrField(TextLine, CNT5 + 1, CNT6 - CNT5 - 1)
                            CNT7 = InStr(CNT6 + 1, TextLine, ",")
                            Group = fStrField(TextLine, CNT6 + 1, CNT7 - CNT6 - 1)
                            CNT8 = InStr(CNT7 + 1, TextLine, ",")
                            Department = fStrField(TextLine, CNT7 + 1, CNT8 - CNT7 - 1)
                            CNT9 = InStr(CNT8 + 1, TextLine, ",")
                            Bus_Unit = fStrField(TextLine, CNT9 + 1, Len(TextLine))

                            If I_O = "#" Then I_O = 1 Else I_O = 0
                            If Len(Term_clock) = 1 Then
                                fileWriter.WriteLine("000" & Term_clock & "000000@0000" & payroll_num & "A" & Mid(date_rec, 5) & Left(time_rec, 4) & I_O)
                            Else
                                fileWriter.WriteLine("00" & Term_clock & "000000@0000" & payroll_num & "A" & Mid(date_rec, 5) & Left(time_rec, 4) & I_O)
                            End If
                        End If
                    Loop
                End Using fileNo
                
            End Using ifileNo
            

            CopyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Backup", "Bck_" & date_rec & time_rec & ".txt")
            File.Copy(TextPath, CopyPath)
            File.Delete(TextPath)
 '       End If
'
 '       Exit Sub

ERR_LABEL:
        Dim ERRMSG As String
        ERRMSG = Err.Description
        On Error Resume Next
        Err_Call(ERRMSG)
        Exit Sub
    End Sub

    Private Function fStrField(ByVal Mydat As String, Sn As Long, Kn As Long) As String
        Return Mid(Mydat, Sn, Kn)
    End Function

    Private Sub Err_Call(strERR As String)
        On Error Resume Next
        MsgBox(strERR)
    End Sub
End Module