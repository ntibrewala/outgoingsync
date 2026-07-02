Imports System
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Sap.Data.Hana
Imports Bhavya.SAP.Core

Module OutgoingPaymentSync

    Dim sapSchema As String = "BHAVYA_POLYFILMS_01"
    Dim DBSBankGLAccount As String = "230002"

    ' Service Layer Config
    Dim slUrl As String = "https://10.10.0.113:50000/b1s/v1"
    Dim slUser As String = "manager"
    Dim slPass As String = "bppl@123"

    Private Sub BypassSSL()
        ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors) True
        ServicePointManager.Expect100Continue = False
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
    End Sub

    Private Function SlLogin() As CookieContainer
        Dim cookieContainer As New CookieContainer()
        Dim handler As New HttpClientHandler()
        handler.CookieContainer = cookieContainer
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        
        Using client As New HttpClient(handler)
            Dim loginUrl As String = $"{slUrl}/Login"
            
            Dim payloadObj As New JObject()
            payloadObj("CompanyDB") = sapSchema
            payloadObj("UserName") = slUser
            payloadObj("Password") = slPass
            Dim payload As String = JsonConvert.SerializeObject(payloadObj)
            
            Dim request As New HttpRequestMessage(HttpMethod.Post, loginUrl)
            request.Content = New StringContent(payload, Encoding.UTF8, "application/json")
            request.Headers.ExpectContinue = False
            
            Dim response = client.SendAsync(request).Result
            If response.IsSuccessStatusCode Then
                Return cookieContainer
            Else
                Throw New Exception("Service Layer Login failed: " & response.Content.ReadAsStringAsync().Result)
            End If
        End Using
    End Function

    Private Sub SlLogout(cookieContainer As CookieContainer)
        Dim handler As New HttpClientHandler()
        handler.CookieContainer = cookieContainer
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim request As New HttpRequestMessage(HttpMethod.Post, $"{slUrl}/Logout")
            request.Headers.ExpectContinue = False
            Dim dummy = client.SendAsync(request).Result
        End Using
    End Sub

    '============================================================
    Sub Main()

        Dim runId As String = Guid.NewGuid().ToString()

        Logger.StartSession()
        Logger.Log("========== [START] RunId=" & runId & " ==========")

        Try
            '=========================
            ' HANA CONNECTION
            '=========================
            Using conn = HanaConnectionManager.GetConnection(sapSchema, HanaConnectionManager.HanaConnectionMode.ReadWriteMode)
                conn.Open()
                Logger.Log("[" & runId & "] HANA Connected")

                '-------------------------
                ' Update Bank Status
                '-------------------------
                If HasStatusUpdates(conn) Then
                    Logger.Log("[" & runId & "] Updating bank payment statuses")

                    Using cmd As New HanaCommand(
                        "CALL ""DBS_BANK"".""BHV_UPDATE_PAYMENT_STATUS""()", conn)
                        cmd.ExecuteNonQuery()
                    End Using

                    Logger.Log("[" & runId & "] Bank status updated")
                End If

                '-------------------------
                ' Check completed payments
                '-------------------------
                If Not HasCompletedPayments(conn) Then
                    Logger.Log("[" & runId & "] No completed payments. Exit.")
                    Exit Sub
                End If

                '=========================
                ' SAP CONNECT (Service Layer)
                '=========================
                BypassSSL()
                Dim slCookies As CookieContainer = SlLogin()
                Logger.Log("[" & runId & "] SAP Service Layer Connected")

                ProcessCompletedPayments(conn, slCookies, runId)
                
                SlLogout(slCookies)

            End Using

        Catch ex As Exception
            Logger.Log("[" & runId & "] FATAL: " & ex.ToString())

        Finally
            Logger.Log("========== [END] RunId=" & runId & " ==========")
            Logger.EndSession()
        End Try

    End Sub

    '============================================================
    Function HasStatusUpdates(conn As HanaConnection) As Boolean
        Using cmd As New HanaCommand(
            "SELECT 1 FROM ""DBS_BANK"".""PENDING_PAYMENTS"" WHERE ""BankStatus""='PENDING' LIMIT 1", conn)
            Return cmd.ExecuteScalar() IsNot Nothing
        End Using
    End Function

    Function HasCompletedPayments(conn As HanaConnection) As Boolean
        Using cmd As New HanaCommand(
            "SELECT 1 FROM ""DBS_BANK"".""PENDING_PAYMENTS"" " &
            "WHERE ""BankStatus""='COMPLETED' AND IFNULL(""Processed"",'N')<>'Y' LIMIT 1", conn)
            Return cmd.ExecuteScalar() IsNot Nothing
        End Using
    End Function

    '============================================================
    Sub ProcessCompletedPayments(conn As HanaConnection, slCookies As CookieContainer, runId As String)

        Logger.Log("[" & runId & "] Fetching completed payments")

        Using cmd As New HanaCommand("CALL ""DBS_BANK"".""BHV_GET_COMPLETED_PAYMENTS""()", conn)

            Using reader = cmd.ExecuteReader()

                While reader.Read()

                    Dim id As String = reader("Id").ToString()
                    Dim vendor As String = reader("VendorCode").ToString()
                    Dim amount As Double = Double.Parse(
                        reader("PaymentAmount").ToString().Replace(",", ""),
                        System.Globalization.CultureInfo.InvariantCulture)

                    Dim txnDate As Date = GetApprovedDate(conn, id)

                    Logger.Log($"[{runId}] Processing ID={id} | Date={txnDate:yyyy-MM-dd}")

                    Try
                        Dim paymentDocEntry As Integer = CreateOutgoingPayment(vendor, amount, txnDate, conn, slCookies)

                        UpdatePaymentProcessed(id, paymentDocEntry, conn)

                        Logger.Log($"[{runId}] SUCCESS | ID={id} | DocEntry={paymentDocEntry}")

                    Catch ex As Exception
                        UpdatePaymentError(id, ex.Message, conn)
                        Logger.Log($"[{runId}] ERROR | ID={id} | {ex.Message}")
                    End Try

                End While

            End Using
        End Using

    End Sub

    '============================================================
    Function GetApprovedDate(conn As HanaConnection, id As String) As Date

        Using cmd As New HanaCommand(
        "SELECT ""ApprovedAt"" FROM ""DBS_BANK"".""PENDING_PAYMENTS"" WHERE ""Id""=?", conn)

            cmd.Parameters.AddWithValue("p1", id)
            Dim result = cmd.ExecuteScalar()

            If result IsNot Nothing Then
                Return Convert.ToDateTime(result).Date
            End If

        End Using

        Return DateTime.Today

    End Function

    '============================================================
    Function CreateOutgoingPayment(vendor As String, amount As Double, txnDate As Date, conn As HanaConnection, slCookies As CookieContainer) As Integer

        ' Determine CardType from HANA directly instead of DI API
        Dim cardType As String = ""
        Using cmd As New HanaCommand($"SELECT ""CardType"" FROM ""{sapSchema}"".""OCRD"" WHERE ""CardCode""=?", conn)
            cmd.Parameters.AddWithValue("p_vendor", vendor)
            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    cardType = reader("CardType").ToString()
                End If
            End Using
        End Using

        Dim payloadObj As New JObject()
        payloadObj("TransferAccount") = DBSBankGLAccount
        payloadObj("TransferSum") = amount
        payloadObj("TransferDate") = txnDate.ToString("yyyy-MM-dd")
        payloadObj("DocDate") = txnDate.ToString("yyyy-MM-dd")
        payloadObj("Remarks") = "DBS Payment - " & vendor

        If cardType = "S" OrElse cardType = "C" Then
            payloadObj("CardCode") = vendor
            payloadObj("DocType") = If(cardType = "S", "rSupplier", "rCustomer")
        Else
            ' It's a G/L Account! We completely bypass Journal Entries using rAccount doc type
            payloadObj("DocType") = "rAccount"
            
            Dim controlAcct As String = "210001"
            Using cmd As New HanaCommand($"SELECT ""DebPayAcct"" FROM ""{sapSchema}"".""OCRD"" WHERE ""CardCode""=?", conn)
                cmd.Parameters.AddWithValue("p_vendor", vendor)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then controlAcct = reader("DebPayAcct").ToString()
                End Using
            End Using

            Dim accountArray As New JArray()
            Dim acctRow As New JObject()
            acctRow("AccountCode") = controlAcct
            acctRow("SumPaid") = amount
            acctRow("Decription") = "DBS Bank Payment - " & vendor
            accountArray.Add(acctRow)
            
            payloadObj("PaymentAccounts") = accountArray
        End If

        Dim payloadStr As String = JsonConvert.SerializeObject(payloadObj)
        
        Dim handler As New HttpClientHandler()
        handler.CookieContainer = slCookies
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim url As String = $"{slUrl}/VendorPayments"
            Dim request As New HttpRequestMessage(HttpMethod.Post, url)
            request.Content = New StringContent(payloadStr, Encoding.UTF8, "application/json")
            request.Headers.ExpectContinue = False
            
            Dim response = client.SendAsync(request).Result
            If response.IsSuccessStatusCode OrElse response.StatusCode = HttpStatusCode.Created Then
                Dim jsonResp As JObject = JObject.Parse(response.Content.ReadAsStringAsync().Result)
                Return Convert.ToInt32(jsonResp("DocEntry"))
            Else
                Throw New Exception("Service Layer Error: " & response.Content.ReadAsStringAsync().Result)
            End If
        End Using

    End Function

    '============================================================
    Sub UpdatePaymentProcessed(id As String, docEntry As Integer, conn As HanaConnection)

        Using cmd As New HanaCommand(
            "CALL ""DBS_BANK"".""BHV_MARK_PAYMENT_PROCESSED""(?,?)", conn)
            cmd.Parameters.AddWithValue("p_id", id)
            cmd.Parameters.AddWithValue("p_docnum", docEntry)
            cmd.ExecuteNonQuery()
        End Using

        ' Clear error
        Using cmd2 As New HanaCommand(
            "UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""ErrorDescription""='' WHERE ""Id""=?", conn)
            cmd2.Parameters.AddWithValue("p_id", id)
            cmd2.ExecuteNonQuery()
        End Using

    End Sub

    Sub UpdatePaymentError(id As String, errorMsg As String, conn As HanaConnection)

        Dim safeError = If(errorMsg.Length > 500, errorMsg.Substring(0, 500), errorMsg)

        Using cmd As New HanaCommand(
            "UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""ErrorDescription""=? WHERE ""Id""=?", conn)
            cmd.Parameters.AddWithValue("p_error", safeError)
            cmd.Parameters.AddWithValue("p_id", id)
            cmd.ExecuteNonQuery()
        End Using

    End Sub
    
    Public Sub RunOutgoing()
        Main()
    End Sub
End Module
