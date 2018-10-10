Imports System.Data.SqlClient

Module Promise
    Public connstr = "data source=DESKTOP-JBRFH9E;initial catalog=HEADLETTERS_ENGINE;integrated security=True;"
    'Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS_ENGINE;integrated security=True;multipleactiveresultsets=True;"
    'Public connstr = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    'Public connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;multipleactiveresultsets=True;User Id=sa;password=pSI)TA1t0K[);"
    Sub Main()
        Dim s As New Stopwatch
        s.Start()
        Dim HCUSP_CALCList As New List(Of HCUSP_CALC)
        Dim HCUSP_SIGN_SUBList As New List(Of HCUSP_SIGN_SUB)
        Dim C1List As New List(Of C1)
        Dim Y1List As New List(Of Y1)
        Select_From_HCUSP("vcdubai@gmail.com", "3", HCUSP_CALCList)
        Dim HCUSP_CALC_TypeNot2nor7 As IList(Of HCUSP_CALC) = (From HCUSP_CALC In HCUSP_CALCList Where HCUSP_CALC.CUSPTYPE <> 2 And HCUSP_CALC.CUSPTYPE <> 7 Select HCUSP_CALC).ToList()
        Insert_Into_HCUSP_SIGN_SUB("vcdubai@gmail.com", "3", HCUSP_CALC_TypeNot2nor7)
        Dim HCUSP_CALC_T1C1 As IList(Of HCUSP_CALC) = (From HCUSP_CALC In HCUSP_CALCList Where HCUSP_CALC.CUSPTYPE = 1 And HCUSP_CALC.CUSPCAT = "1" Select HCUSP_CALC).ToList()
        Dim HCUSP_CALC_T3C1 As IList(Of HCUSP_CALC) = (From HCUSP_CALC In HCUSP_CALCList Where HCUSP_CALC.CUSPTYPE = 3 And HCUSP_CALC.CUSPCAT = "1" Select HCUSP_CALC).ToList()
        Dim HCUSP_CALC_T3C3 As IList(Of HCUSP_CALC) = (From HCUSP_CALC In HCUSP_CALCList Where HCUSP_CALC.CUSPTYPE = 3 And HCUSP_CALC.CUSPCAT = "3" Select HCUSP_CALC).ToList()
        Select_From_HCUSP_SIGN_SUB("vcdubai@gmail.com", "3", HCUSP_SIGN_SUBList)
        Insert_Into_C1("vcdubai@gmail.com", "3", HCUSP_CALC_T1C1, HCUSP_CALC_T3C1, HCUSP_CALC_T3C3, HCUSP_SIGN_SUBList)
        Select_From_C1("vcdubai@gmail.com", "3", C1List)
        DoSomething("vcdubai@gmail.com", "3", C1List, HCUSP_CALCList)
        Select_From_Y1("vcdubai@gmail.com", "3", Y1List)
        DoSomethingWith_Y1("vcdubai@gmail.com", "3", Y1List)
        Promise1.Main()
        s.Stop()
        Console.Write(s.Elapsed.ToString())
        Console.ReadKey()
    End Sub
    Sub Select_From_HCUSP(ByVal UID As String, ByVal HID As String, ByRef HCUSP_CALCList As List(Of HCUSP_CALC))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_CALC WHERE CUSPUSERID = '" + UID + "' AND CUSPHID = '" + HID + "';"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim hCusp_Calc = New HCUSP_CALC With {
                    .CUSPUSERID = reader("CUSPUSERID").ToString().Trim,
                    .CUSPHID = reader("CUSPHID").ToString().Trim,
                    .CUSPID = reader("CUSPID").ToString().Trim,
                    .CUSPPLANET = reader("CUSPPLANET").ToString().Trim,
                    .CUSPTYPE = reader("CUSPTYPE"),
                    .CUSPCAT = reader("CUSPCAT").ToString().ToUpper().Trim
                }
                HCUSP_CALCList.Add(hCusp_Calc)
            End While
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_HCUSP_SIGN_SUB(ByVal UID As String, ByVal HID As String, ByRef HCUSP_SIGN_SUBList As List(Of HCUSP_SIGN_SUB))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT DISTINCT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_SIGN_SUB;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim hCusp_Sign_Sub = New HCUSP_SIGN_SUB With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .CUSP = reader("CUSP").ToString().Trim,
                    .PLANET = reader("PLANET").ToString().Trim
                }
                HCUSP_SIGN_SUBList.Add(hCusp_Sign_Sub)
            End While
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_C1(ByVal UID As String, ByVal HID As String, ByRef C1List As List(Of C1))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT DISTINCT * FROM HEADLETTERS_ENGINE.DBO.C1;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim c1 = New C1 With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .CUSP = reader("CUSP").ToString().Trim,
                    .P = reader("P").ToString().Trim,
                    .TYPE = reader("TYPE").ToString().Trim,
                    .CAT = reader("CAT").ToString().Trim
                }
                C1List.Add(c1)
            End While
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_Y1(ByVal UID As String, ByVal HID As String, ByRef Y1List As List(Of Y1))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT DISTINCT * FROM HEADLETTERS_ENGINE.DBO.Y1;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim y1 = New Y1 With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .FKEY = reader("FKEY").ToString().Trim,
                    .TKEY = reader("TKEY").ToString().Trim,
                    .PLANET = reader("PLANET").ToString().Trim
                }
                Y1List.Add(y1)
            End While
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Insert_Into_HCUSP_SIGN_SUB(ByVal UID As String, ByVal HID As String, ByRef HCUSP_CALCList As List(Of HCUSP_CALC))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con

            cmd.CommandText = $"DELETE FROM HEADLETTERS_ENGINE.DBO.HCUSP_SIGN_SUB WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            cmd.ExecuteNonQuery()
            For i As Integer = 0 To HCUSP_CALCList.Count - 1
                cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_SIGN_SUB VALUES ('" + HCUSP_CALCList.Item(i).CUSPUSERID + "','" + HCUSP_CALCList.Item(i).CUSPHID + "','" + HCUSP_CALCList.Item(i).CUSPID + "','" + HCUSP_CALCList.Item(i).CUSPPLANET + "');
                                INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_SIGN_SUB VALUES ('" + HCUSP_CALCList.Item(i).CUSPUSERID + "','" + HCUSP_CALCList.Item(i).CUSPHID + "','99','" + HCUSP_CALCList.Item(i).CUSPPLANET + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Insert_Into_C1(ByVal UID As String, ByVal HID As String, ByRef HCUSP_CALC_T1C1 As List(Of HCUSP_CALC), ByRef HCUSP_CALC_T3C1 As List(Of HCUSP_CALC), ByRef HCUSP_CALC_T3C3 As List(Of HCUSP_CALC), ByVal HCUSP_SIGN_SUBList As List(Of HCUSP_SIGN_SUB))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con

            cmd.CommandText = $"DELETE FROM HEADLETTERS_ENGINE.DBO.C1 WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            cmd.ExecuteNonQuery()

            For i As Integer = 0 To HCUSP_CALC_T1C1.Count - 1
                cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.C1 VALUES ('" + HCUSP_CALC_T1C1.Item(i).CUSPUSERID + "','" + HCUSP_CALC_T1C1.Item(i).CUSPHID + "','" + HCUSP_CALC_T1C1.Item(i).CUSPID + "','" + HCUSP_CALC_T1C1.Item(i).CUSPPLANET + "','" + HCUSP_CALC_T1C1.Item(i).CUSPTYPE.ToString() + "','" + HCUSP_CALC_T1C1.Item(i).CUSPCAT + "');"
                cmd.ExecuteNonQuery()
            Next
            For i As Integer = 0 To HCUSP_CALC_T3C1.Count - 1
                cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.C1 VALUES ('" + HCUSP_CALC_T3C1.Item(i).CUSPUSERID + "','" + HCUSP_CALC_T3C1.Item(i).CUSPHID + "','" + HCUSP_CALC_T3C1.Item(i).CUSPID + "','" + HCUSP_CALC_T3C1.Item(i).CUSPPLANET + "','" + HCUSP_CALC_T3C1.Item(i).CUSPTYPE.ToString() + "','" + HCUSP_CALC_T3C1.Item(i).CUSPCAT + "');"
                cmd.ExecuteNonQuery()
            Next

            Dim HCUSP_CALC_T1C1_CUSP99 = (From HCUSP_CALC In HCUSP_CALC_T3C3 From HCUSP_SIGN_SUB In HCUSP_SIGN_SUBList
                                          Where HCUSP_SIGN_SUB.CUSP = "99" And HCUSP_SIGN_SUB.UID = HCUSP_CALC.CUSPUSERID And HCUSP_SIGN_SUB.HID = HCUSP_CALC.CUSPHID And HCUSP_SIGN_SUB.PLANET = HCUSP_CALC.CUSPPLANET
                                          Select New With {HCUSP_CALC.CUSPUSERID, HCUSP_CALC.CUSPHID, HCUSP_CALC.CUSPPLANET, HCUSP_CALC.CUSPID, HCUSP_CALC.CUSPTYPE, HCUSP_CALC.CUSPCAT}).ToList()
            For i As Integer = 0 To HCUSP_CALC_T1C1_CUSP99.Count - 1
                cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.C1 VALUES ('" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPUSERID + "','" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPHID + "','" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPID + "','" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPPLANET + "','" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPTYPE.ToString() + "','" + HCUSP_CALC_T1C1_CUSP99.Item(i).CUSPCAT + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub DoSomething(ByVal UID As String, ByVal HID As String, ByVal C1List As List(Of C1), ByRef HCUSP_CALCList As List(Of HCUSP_CALC))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con

            cmd.CommandText = $"DELETE FROM HEADLETTERS_ENGINE.DBO.Y1 WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            cmd.ExecuteNonQuery()

            For x As Integer = 0 To C1List.Count - 1
                Dim s_key = C1List.Item(x).UID + C1List.Item(x).HID + C1List.Item(x).CUSP + C1List.Item(x).P + C1List.Item(x).TYPE + C1List.Item(x).CAT
                For i As Integer = 1 To 12
                    Dim s_key1 = C1List.Item(x).UID + C1List.Item(x).HID + i.ToString("D2") + C1List.Item(x).P + "11"
                    If s_key <> s_key1 Then
                        Dim FoundCount = (From HCUSP_CALC In HCUSP_CALCList
                                          Where HCUSP_CALC.CUSPUSERID = C1List.Item(x).UID And HCUSP_CALC.CUSPHID = C1List.Item(x).HID And HCUSP_CALC.CUSPID = i.ToString("D2") And HCUSP_CALC.CUSPPLANET = C1List.Item(x).P And HCUSP_CALC.CUSPTYPE = "1" And HCUSP_CALC.CUSPCAT = "1" Select HCUSP_CALC).ToList()
                        If FoundCount.Count > 0 Then
                            cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.Y1 VALUES ('" + C1List.Item(x).UID + "','" + C1List.Item(x).HID + "','" + C1List.Item(x).CUSP + "','" + i.ToString("D2") + "','" + C1List.Item(x).P + "');"
                            cmd.ExecuteNonQuery()
                        End If
                    End If
                    s_key1 = C1List.Item(x).UID + C1List.Item(x).HID + i.ToString("D2") + C1List.Item(x).P + "31"
                    If s_key <> s_key1 Then
                        Dim FoundCount = (From HCUSP_CALC In HCUSP_CALCList
                                          Where HCUSP_CALC.CUSPUSERID = C1List.Item(x).UID And HCUSP_CALC.CUSPHID = C1List.Item(x).HID And HCUSP_CALC.CUSPID = i.ToString("D2") And HCUSP_CALC.CUSPPLANET = C1List.Item(x).P And HCUSP_CALC.CUSPTYPE = "3" And HCUSP_CALC.CUSPCAT = "1" Select HCUSP_CALC).ToList()
                        If FoundCount.Count > 0 Then
                            cmd.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.Y1 VALUES ('" + C1List.Item(x).UID + "','" + C1List.Item(x).HID + "','" + C1List.Item(x).CUSP + "','" + i.ToString("D2") + "','" + C1List.Item(x).P + "');"
                            cmd.ExecuteNonQuery()
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Console.WriteLine(ex)
        Finally
            con.Close()
        End Try
    End Sub
    Sub Select_From_HCUSP_PROMISE_LINK(ByVal UID As String, ByVal HID As String, ByRef HCUSP_PROMISE_LINKList As List(Of HCUSP_PROMISE_LINK))
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "SELECT * FROM HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK;"
            reader = cmd.ExecuteReader()
            While reader.Read()
                Dim HCusp_Promise_Link = New HCUSP_PROMISE_LINK With {
                    .UID = reader("UID").ToString().Trim,
                    .HID = reader("HID").ToString().Trim,
                    .CUSPKEY = reader("CUSPKEY").ToString().Trim,
                    .TCOUNT = reader("TCOUNT"),
                    .SKEY = reader("SKEY").ToString().Trim
                }
                HCUSP_PROMISE_LINKList.Add(HCusp_Promise_Link)
            End While
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Sub DoSomethingWith_Y1(ByVal UID As String, ByVal HID As String, ByVal Y1List As List(Of Y1))
        Try
            Dim c0n As New SqlConnection
            Dim c0d As New SqlCommand
            c0n.ConnectionString = connstr
            c0n.Open()
            c0d.Connection = c0n
            c0d.CommandText = "DELETE FROM HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK WHERE UID = '" + UID + "' AND HID = '" + HID + "';"
            c0d.ExecuteNonQuery()
            c0n.Close()
            For i As Integer = 0 To Y1List.Count - 1
                Dim k1 = Y1List.Item(i).UID + Y1List.Item(i).HID + Y1List.Item(i).FKEY + Y1List.Item(i).TKEY
                Dim c1n As New SqlConnection
                Dim c1d As New SqlCommand
                Dim rdr As SqlDataReader
                c1n.ConnectionString = connstr
                c1n.Open()
                c1d.Connection = c1n
                c1d.CommandText = $"SELECT COUNT(*) AS COUNTS FROM HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK WHERE UID = '" + Y1List.Item(i).UID + "'
                                        AND HID = '" + Y1List.Item(i).HID + "' AND SKEY = '" + k1 + "';"
                rdr = c1d.ExecuteReader()
                Dim FoundCount As Integer
                While rdr.Read()
                    FoundCount = Convert.ToInt32(rdr("counts"))
                End While
                c1n.Close()
                If FoundCount > 0 Then
                    Dim TCountReader As SqlDataReader
                    Dim TCOUNT = 0
                    Dim c2n As New SqlConnection
                    Dim c2d As New SqlCommand
                    c2n.ConnectionString = connstr
                    c2n.Open()
                    c2d.Connection = c2n
                    c2d.CommandText = $"SELECT TCOUNT FROM HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK WHERE UID = '" + Y1List.Item(i).UID + "'
                                        AND HID = '" + Y1List.Item(i).HID + "' AND SKEY = '" + k1 + "';"
                    TCountReader = c2d.ExecuteReader()
                    While TCountReader.Read()
                        TCOUNT = TCountReader("TCOUNT")
                    End While
                    c2n.Close()
                    Dim c3n As New SqlConnection
                    Dim c3d As New SqlCommand
                    c3n.ConnectionString = connstr
                    c3n.Open()
                    c3d.Connection = c3n
                    c3d.CommandText = $"UPDATE HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK SET TCOUNT = " + (TCOUNT + 1).ToString() + " WHERE UID = '" + Y1List.Item(i).UID + "' AND HID = '" + Y1List.Item(i).HID + "' AND SKEY = '" + k1 + "';"
                    c3d.ExecuteNonQuery()
                    c3n.Close()
                Else
                    Dim c4n As New SqlConnection
                    Dim c4d As New SqlCommand
                    c4n.ConnectionString = connstr
                    c4n.Open()
                    c4d.Connection = c4n
                    c4d.CommandText = $"INSERT INTO HEADLETTERS_ENGINE.DBO.HCUSP_PROMISE_LINK VALUES ('" + Y1List.Item(i).UID + "','" + Y1List.Item(i).HID + "','" + Y1List.Item(i).FKEY + "" + Y1List.Item(i).TKEY + "'," + 1.ToString() + ",'" + k1 + "');"
                    c4d.ExecuteNonQuery()
                    c4n.Close()
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("Error Occured : " + ex.StackTrace)
        End Try
    End Sub
    Class HCUSP_CALC
        Property CUSPUSERID As String
        Property CUSPHID As String
        Property CUSPID As String
        Property CUSPPLANET As String
        Property CUSPTYPE As Integer
        Property CUSPCAT As String
    End Class
    Class HCUSP_SIGN_SUB
        Property UID As String
        Property HID As String
        Property CUSP As String
        Property PLANET As String
    End Class
    Class C1
        Property UID As String
        Property HID As String
        Property CUSP As String
        Property P As String
        Property TYPE As String
        Property CAT As String
    End Class
    Class Y1
        Property UID As String
        Property HID As String
        Property FKEY As String
        Property TKEY As String
        Property PLANET As String
    End Class
End Module
