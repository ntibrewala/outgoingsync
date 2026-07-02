import os

repo_dir = "/Users/nikhiltibrewala/.gemini/antigravity/playground/vector-sun/outgoingsync"
files = ["manualrecocom.vb", "OutgoingPaymentSync.vb"]

for filename in files:
    filepath = os.path.join(repo_dir, filename)
    with open(filepath, "r") as f:
        content = f.read()
    
    # Add Expect100Continue = False to BypassSSL
    content = content.replace(
        'ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors) True',
        'ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors) True\n        ServicePointManager.Expect100Continue = False'
    )
    
    with open(filepath, "w") as f:
        f.write(content)

print("Added Expect100Continue = False.")
