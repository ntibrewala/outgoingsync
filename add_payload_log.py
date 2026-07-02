import os
import re

repo_dir = "/Users/nikhiltibrewala/.gemini/antigravity/playground/vector-sun/outgoingsync"
files = ["manualrecocom.vb", "OutgoingPaymentSync.vb"]

for filename in files:
    filepath = os.path.join(repo_dir, filename)
    with open(filepath, "r") as f:
        content = f.read()
    
    # Add Logger.Log or Console.WriteLine
    content = content.replace(
        'Dim payload As String = JsonConvert.SerializeObject(payloadObj)',
        'Dim payload As String = JsonConvert.SerializeObject(payloadObj)\n            Console.WriteLine("SL PAYLOAD: " & payload)'
    )
    
    with open(filepath, "w") as f:
        f.write(content)

print("Added payload logging.")
