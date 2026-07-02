Imports System
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports Sap.Data.Hana
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Bhavya.SAP.Core

Module manualrecocom

    Dim sapSchema As String = "BHAVYA_POLYFILMS_01"
    
    ' HANA DB Config
    Dim hanaHost As String = "10.10.0.113"
    Dim hanaPort As String = "30015"
    Dim hanaUser As String = "SYSTEM"
    Dim hanaPass As String = "B1sap#2025"

    ' Service Layer Config
    Dim slUrl As String = "https://10.10.0.113:50000/b1s/v1"
    Dim slUser As String = "manager"
    Dim slPass As String = "bppl@123"

    ' Disable SSL verification for Service Layer
    Private Sub BypassSSL()
        ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors) True
        ServicePointManager.Expect100Continue = False
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
    End Sub

    Sub Main()
        BypassSSL()
        Dim runId As String = Guid.NewGuid().ToString()
        Logger.Log("========== [ManualReco START] RunId=" & runId & " ==========")

        Dim connString As String = $"Server={hanaHost}:{hanaPort};UserID={hanaUser};Password={hanaPass}"
        Using conn As New HanaConnection(connString)
            Try
                conn.Open()
                Logger.Log("[" & runId & "] Connected to HANA")

                ' Fetch pending payments queue from Bank DB
                Dim sql As String = $"
                    SELECT P.""Id"", P.""VendorCode"", P.""DocNum"" AS ""PaymentDocEntry"", 
                           TO_NVARCHAR(P.""InvoicesJson"") AS ""InvoicesJson"", 
                           IFNULL(P.""Reconciled"",'N') AS ""ReconState"" 
                    FROM ""DBS_BANK"".""PENDING_PAYMENTS"" P 
                    WHERE P.""Processed"" = 'Y' 
                    AND IFNULL(P.""Reconciled"",'N') IN ('N','1','2') 
                    AND P.""InvoicesJson"" IS NOT NULL 
                    AND TO_NVARCHAR(P.""InvoicesJson"") <> '[]'
                "
                Dim cmd As New HanaCommand(sql, conn)
                Dim reader As HanaDataReader = cmd.ExecuteReader()
                
                Dim records As New List(Of Object)()
                While reader.Read()
                    records.Add(New With {
                        .Id = reader.GetString(0),
                        .VendorCode = reader.GetString(1),
                        .DocEntry = reader.GetInt32(2),
                        .InvoicesJson = reader.GetString(3),
                        .State = reader.GetString(4)
                    })
                End While
                reader.Close()

                If records.Count = 0 Then
                    Logger.Log("[" & runId & "] No pending records found. Exiting.")
                    Return
                End If

                Logger.Log("[" & runId & "] Fetching pending payments... Found " & records.Count)

                ' Login to Service Layer
                Dim slCookies As CookieContainer = SlLogin()

                For Each rec In records
                    Logger.Log("[" & runId & "] START Payment | ID=" & rec.Id & " | Vendor=" & rec.VendorCode & " | DocEntry=" & rec.DocEntry & " | State=" & rec.State)
                    
                    Try
                        ReconcilePayment(slCookies, conn, rec.VendorCode, rec.DocEntry, rec.InvoicesJson, runId)
                        UpdateReconciled(conn, rec.Id, "Y")
                        Logger.Log("[" & runId & "] SUCCESS Payment ID=" & rec.Id)
                    Catch ex As Exception
                        Dim nextState As String = "1"
                        If rec.State = "1" Then nextState = "2"
                        If rec.State = "2" Then nextState = "F"
                        
                        UpdateReconciled(conn, rec.Id, nextState)
                        Logger.Log("[" & runId & "] Updated Reconciled state to " & nextState & " for ID=" & rec.Id)
                        Logger.Log("[" & runId & "] ERROR: FAILED Payment ID=" & rec.Id & " | Error=" & ex.Message)
                    End Try
                    
                    Logger.Log("[" & runId & "] END Payment ID=" & rec.Id)
                Next

                ' Logout of Service Layer
                SlLogout(slCookies)

            Catch ex As Exception
                Logger.Log("[" & runId & "] FATAL ERROR: " & ex.Message)
            Finally
                If conn.State = System.Data.ConnectionState.Open Then
                    conn.Close()
                End If
                Logger.Log("========== [ManualReco END] RunId=" & runId & " ==========")
            End Try
        End Using
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

    Private Sub UpdateReconciled(conn As HanaConnection, id As String, state As String)
        Dim sql As String = $"UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""Reconciled""='{state}' WHERE ""Id""='{id}'"
        Using cmd As New HanaCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function GetPaymentJEInfo(conn As HanaConnection, paymentDocEntry As Integer, vendorCode As String) As Tuple(Of Integer, Integer)
        Dim sql As String = $"
            SELECT V.""TransId"", J.""Line_ID"" 
            FROM ""{sapSchema}"".""OVPM"" V 
            INNER JOIN ""{sapSchema}"".""JDT1"" J ON V.""TransId"" = J.""TransId"" 
            WHERE V.""DocEntry"" = {paymentDocEntry} 
            AND J.""Debit"" > 0 AND J.""ShortName"" = '{vendorCode}'
        "
        Using cmd As New HanaCommand(sql, conn)
            Using reader As HanaDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    Return New Tuple(Of Integer, Integer)(reader.GetInt32(0), reader.GetInt32(1))
                End If
            End Using
        End Using
        Throw New Exception($"Payment JE not found | DocEntry={paymentDocEntry}")
    End Function

    Private Function GetInvoiceJEInfo(conn As HanaConnection, invTransId As Integer, vendorCode As String) As Tuple(Of Integer, Double)
        Dim sql As String = $"
            SELECT J.""Line_ID"", J.""Debit"", J.""Credit"" 
            FROM ""{sapSchema}"".""JDT1"" J 
            WHERE J.""TransId"" = {invTransId} 
            AND J.""ShortName"" = '{vendorCode}'
        "
        Using cmd As New HanaCommand(sql, conn)
            Using reader As HanaDataReader = cmd.ExecuteReader()
                If reader.Read() Then
                    Dim debit As Double = reader.GetDouble(1)
                    Dim credit As Double = reader.GetDouble(2)
                    Dim maxAmt As Double = Math.Max(debit, credit)
                    Return New Tuple(Of Integer, Double)(reader.GetInt32(0), maxAmt)
                End If
            End Using
        End Using
        Throw New Exception($"Invoice JE line not found | TransId={invTransId} | BP={vendorCode}")
    End Function

    Private Sub ReconcilePayment(cookies As CookieContainer, conn As HanaConnection, vendor As String, paymentDocEntry As Integer, invoicesJson As String, runId As String)
        
        ' 1. Get Payment JE Line Details
        Dim payInfo = GetPaymentJEInfo(conn, paymentDocEntry, vendor)
        Dim payTransId As Integer = payInfo.Item1
        Dim payLineId As Integer = payInfo.Item2
        Logger.Log("[" & runId & "] Payment JE | TransId=" & payTransId & " | LineId=" & payLineId)

        ' 2. Initialize Service Layer DI-API wrapper payload
        Dim payloadObj As New JObject()
        Dim internalReconObj As New JObject()
        internalReconObj("CardOrAccount") = "coaCard"
        Dim rowsArray As New JArray()
        
        Dim totalAmount As Double = 0.0
        Dim invoices As JArray = JArray.Parse(invoicesJson)

        ' 3. Extract Invoice JE Lines
        For Each inv In invoices
            Dim invTransId As Integer = CInt(inv("InvoiceNumber"))
            Dim amount As Double = CDbl(inv("Amount"))
            
            Dim invInfo = GetInvoiceJEInfo(conn, invTransId, vendor)
            Dim invLineId As Integer = invInfo.Item1
            
            Logger.Log("[" & runId & "] Invoice | TransId=" & invTransId & " | LineId=" & invLineId & " | ReconAmount=" & amount)
            
            Dim invRow As New JObject()
            invRow("Selected") = "tYES"
            invRow("ShortName") = vendor
            invRow("TransId") = invTransId
            invRow("TransRowId") = invLineId
            invRow("ReconcileAmount") = Math.Abs(amount)
            rowsArray.Add(invRow)
            
            totalAmount += Math.Abs(amount)
        Next
        
        Logger.Log("[" & runId & "] Total Invoice Amount = " & totalAmount)
        
        ' 4. Add the Payment Row to balance it exactly
        Dim payRow As New JObject()
        payRow("Selected") = "tYES"
        payRow("ShortName") = vendor
        payRow("TransId") = payTransId
        payRow("TransRowId") = payLineId
        payRow("ReconcileAmount") = totalAmount
        rowsArray.Add(payRow)
        
        ' 5. Wrap everything inside InternalReconciliationOpenTrans
        internalReconObj("InternalReconciliationOpenTransRows") = rowsArray
        payloadObj("InternalReconciliationOpenTrans") = internalReconObj
        Dim payloadStr As String = JsonConvert.SerializeObject(payloadObj)
        
        ' 6. Post to the Hidden Service Layer Action Endpoint
        Logger.Log("[" & runId & "] Posting reconciliation to Service Layer...")
        Dim handler As New HttpClientHandler()
        handler.CookieContainer = cookies
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim url As String = $"{slUrl}/InternalReconciliationsService_Add"
            Dim request As New HttpRequestMessage(HttpMethod.Post, url)
            request.Content = New StringContent(payloadStr, Encoding.UTF8, "application/json")
            request.Headers.ExpectContinue = False
            
            Dim response = client.SendAsync(request).Result
            If response.IsSuccessStatusCode OrElse response.StatusCode = HttpStatusCode.Created OrElse response.StatusCode = HttpStatusCode.NoContent Then
                Logger.Log("[" & runId & "] Reconciliation DONE | PaymentDocEntry=" & paymentDocEntry)
            Else
                Throw New Exception($"Service Layer Error: {response.Content.ReadAsStringAsync().Result}")
            End If
        End Using
    End Sub

    Public Sub RunManualReco()
        Main()
    End Sub

End Module
