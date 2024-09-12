Imports System.Data.SqlClient
Imports System.IO

Module Module1

    Sub Main()
        On Error GoTo ERR_LABEL

        ' Variables para la conexión y consulta SQL
        Dim connectionString As String = "Server=your_server_name;Database=your_database_name;User Id=your_username;Password=your_password;"
        Dim query As String = "SELECT Term_clock, date_rec, time_rec, Company, payroll_num, I_O, Group, Department, Bus_Unit FROM YourTableName"

        ' Variables de datos
        Dim Term_clock As String
        Dim date_rec As String
        Dim time_rec As String
        Dim Company As String
        Dim payroll_num As String
        Dim I_O As String
        Dim Group As String
        Dim Department As String
        Dim Bus_Unit As String

        Dim TextOutPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reloj.dat")

        ' Conexión a la base de datos
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            ' Ejecutar la consulta y leer los resultados
            Using command As New SqlCommand(query, connection)
                Using reader As SqlDataReader = command.ExecuteReader()
                    Using fileWriter As New StreamWriter(TextOutPath)
                        While reader.Read()

                            ' Consulta SQL con INNER JOIN 
                            Dim query As String = "SELECT punch_time, employee_id, terminal_id" &
                                "FROM att_punches ap" &
                                "INNER JOIN hr_employee he ON ap.employee_id = he.id" &
                                "INNER JOIN att_terminal at2 ON ap.terminal_id = at2.id" &
                                "WHERE punch_time >= datetime('now', '-1 day', 'start of day', '+7 hours')" &
                                "AND punch_time < datetime('now', 'start of day', '+7 hours');"
                            ' Procesar I/O
                            If I_O = "#" Then I_O = 1 Else I_O = 0

                            ' Escribir en el archivo de salida
                            If Len(Term_clock) = 1 Then
                                fileWriter.WriteLine("000" & Term_clock & "000000@1000" & payroll_num & "A" & Mid(date_rec, 5) & Left(time_rec, 4) & I_O)
                            Else
                                fileWriter.WriteLine("00" & Term_clock & "000000@1000" & payroll_num & "A" & Mid(date_rec, 5) & Left(time_rec, 4) & I_O)
                            End If
                        End While
                    End Using
                End Using
            End Using
        End Using

        Exit Sub

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
