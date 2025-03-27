> [!NOTE]
> Microsoft Outlook Client in Legacy Mode


Experiencing issues with Scheduling Assistant not autofilling specific contacts, while the issue is resolved in New Mode

> [!IMPORTANT]
> Here's a **single-line PowerShell command** that ensures:  
✅ **Runs with Administrator Privileges**  
✅ **Sets Execution Policy to Bypass**  
✅ **Executes the Refresh-OutlookCache Script**  


````
powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command \"iwr -useb https://raw.githubusercontent.com/ulyweb/outlook/refs/heads/main/scripts/Refresh-OutlookCache.ps1 | iex\"' -Verb RunAs"
````
