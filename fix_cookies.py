import os

repo = "/Users/nikhiltibrewala/.gemini/antigravity/playground/vector-sun/outgoingsync"

for filename in ["OutgoingPaymentSync.vb", "manualrecocom.vb"]:
    filepath = os.path.join(repo, filename)
    with open(filepath, "r") as f:
        content = f.read()

    # 1. Update SlLogin signature and implementation
    old_login = '''    Private Function SlLogin() As String
        Dim handler As New HttpClientHandler()
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
                Dim cookies = response.Headers.GetValues("Set-Cookie")
                Return String.Join(";", cookies)
            Else
                Throw New Exception("Service Layer Login failed: " & response.Content.ReadAsStringAsync().Result)
            End If
        End Using
    End Function'''
    
    new_login = '''    Private Function SlLogin() As CookieContainer
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
    End Function'''
    
    content = content.replace(old_login, new_login)
    
    # 2. Update SlLogout
    old_logout = '''    Private Sub SlLogout(cookies As String)
        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("Cookie", cookies)
            Dim dummy = client.PostAsync($"{slUrl}/Logout", Nothing).Result
        End Using
    End Sub'''
    new_logout = '''    Private Sub SlLogout(cookieContainer As CookieContainer)
        Dim handler As New HttpClientHandler()
        handler.CookieContainer = cookieContainer
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim request As New HttpRequestMessage(HttpMethod.Post, $"{slUrl}/Logout")
            request.Headers.ExpectContinue = False
            Dim dummy = client.SendAsync(request).Result
        End Using
    End Sub'''
    content = content.replace(old_logout, new_logout)
    
    # 3. Update usages of SlLogin
    content = content.replace('Dim slCookies As String = SlLogin()', 'Dim slCookies As CookieContainer = SlLogin()')
    
    # 4. Update CreateOutgoingPayment signature and usage (only in OutgoingPaymentSync.vb)
    if filename == "OutgoingPaymentSync.vb":
        content = content.replace('Sub ProcessCompletedPayments(conn As HanaConnection, slCookies As String, runId As String)', 'Sub ProcessCompletedPayments(conn As HanaConnection, slCookies As CookieContainer, runId As String)')
        content = content.replace('Function CreateOutgoingPayment(vendor As String, amount As Double, txnDate As Date, conn As HanaConnection, slCookies As String) As Integer', 'Function CreateOutgoingPayment(vendor As String, amount As Double, txnDate As Date, conn As HanaConnection, slCookies As CookieContainer) As Integer')
        
        # Replace the HttpClient block in CreateOutgoingPayment
        old_http = '''        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("Cookie", slCookies)
            Dim url As String = $"{slUrl}/VendorPayments"
            Dim content As New StringContent(payloadStr, Encoding.UTF8, "application/json")
            
            Dim response = client.PostAsync(url, content).Result'''
        
        new_http = '''        Dim handler As New HttpClientHandler()
        handler.CookieContainer = slCookies
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim url As String = $"{slUrl}/VendorPayments"
            Dim request As New HttpRequestMessage(HttpMethod.Post, url)
            request.Content = New StringContent(payloadStr, Encoding.UTF8, "application/json")
            request.Headers.ExpectContinue = False
            
            Dim response = client.SendAsync(request).Result'''
            
        content = content.replace(old_http, new_http)
        
    # 5. Update ReconcilePayment signature and usage (only in manualrecocom.vb)
    if filename == "manualrecocom.vb":
        content = content.replace('Private Sub ReconcilePayment(cookies As String, conn As HanaConnection, vendor As String, paymentDocEntry As Integer, invoicesJson As String, runId As String)', 'Private Sub ReconcilePayment(cookies As CookieContainer, conn As HanaConnection, vendor As String, paymentDocEntry As Integer, invoicesJson As String, runId As String)')
        
        # Replace the HttpClient block in ReconcilePayment
        old_http = '''        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("Cookie", cookies)
            Dim url As String = $"{slUrl}/InternalReconciliationsService_Add"
            Dim content As New StringContent(payloadStr, Encoding.UTF8, "application/json")
            
            Dim response = client.PostAsync(url, content).Result'''
            
        new_http = '''        Dim handler As New HttpClientHandler()
        handler.CookieContainer = cookies
        handler.UseCookies = True
        handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, sslPolicyErrors) True
        Using client As New HttpClient(handler)
            Dim url As String = $"{slUrl}/InternalReconciliationsService_Add"
            Dim request As New HttpRequestMessage(HttpMethod.Post, url)
            request.Content = New StringContent(payloadStr, Encoding.UTF8, "application/json")
            request.Headers.ExpectContinue = False
            
            Dim response = client.SendAsync(request).Result'''
            
        content = content.replace(old_http, new_http)

    with open(filepath, "w") as f:
        f.write(content)

print("Updated both files to use CookieContainer!")
