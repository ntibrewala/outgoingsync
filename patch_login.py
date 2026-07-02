import os
import re

repo_dir = "/Users/nikhiltibrewala/.gemini/antigravity/playground/vector-sun/outgoingsync"

files = ["manualrecocom.vb", "OutgoingPaymentSync.vb"]

new_login = """
    Private Function SlLogin() As String
        Using client As New HttpClient()
            Dim loginUrl As String = $"{slUrl}/Login"
            
            Dim payloadObj As New JObject()
            payloadObj("CompanyDB") = sapSchema
            payloadObj("UserName") = slUser
            payloadObj("Password") = slPass
            Dim payload As String = JsonConvert.SerializeObject(payloadObj)
            
            Dim content As New StringContent(payload, Encoding.UTF8, "application/json")
            
            Dim response = client.PostAsync(loginUrl, content).Result
            If response.IsSuccessStatusCode Then
                Dim cookies = response.Headers.GetValues("Set-Cookie")
                Return String.Join(";", cookies)
            Else
                Throw New Exception("Service Layer Login failed: " & response.Content.ReadAsStringAsync().Result)
            End If
        End Using
    End Function
"""

old_login_pattern = r"Private Function SlLogin\(\) As String.*?End Function"

for filename in files:
    filepath = os.path.join(repo_dir, filename)
    with open(filepath, "r") as f:
        content = f.read()
    
    new_content = re.sub(old_login_pattern, new_login.strip(), content, flags=re.DOTALL)
    
    with open(filepath, "w") as f:
        f.write(new_content)

print("Patched SlLogin to use JObject instead of string interpolation.")
