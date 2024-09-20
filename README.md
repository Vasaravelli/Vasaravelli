Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub SearchAndSave()
    ' Excel sheet with search terms in Column A
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust sheet name as per your file

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim searchItem As String
    Dim i As Long
    
    For i = 1 To lastRow
        searchItem = ws.Cells(i, 1).Value
        Call AutomateChrome(searchItem)
    Next i
End Sub

Sub AutomateChrome(ByVal searchItem As String)
    ' Define the path to the Chrome executable file
    Dim chromePath As String
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe" ' Adjust the path to where Chrome is installed on your machine
    
    ' Command to launch Chrome with DevTools Protocol enabled
    Dim chromeCmd As String
    chromeCmd = chromePath & " --remote-debugging-port=9222" ' Ensure no Chrome debugging port option

    ' Open Chrome with DevTools Protocol enabled
    Shell chromeCmd, vbNormalFocus
    Sleep 2000 ' Wait for browser to open

    ' Insert the code to interact with the page, send the search term to the search box using CDP, and click the search button.
    ' Use JSON formatted messages for CDP commands.

    ' Example JSON structure for pasting into the search box:
    ' Use DevTools command Input.insertText to input text into the search box
    ' Then use Page.printToPDF for saving the page.

    ' Simulate Ctrl + P keyboard event and Save to routed folder
    ' Save PDF using Chrome DevTools command

    ' Sample JSON commands for sending text to search box
    Dim searchBoxCmd As String
    searchBoxCmd = "{'method':'Runtime.evaluate','params':{'expression':'document.querySelector(""#searchbox"").value = """ & searchItem & """'}}"
    
    ' Execute this via CDP
    
    ' Trigger search button click
    Dim clickCmd As String
    clickCmd = "{'method':'Runtime.evaluate','params':{'expression':'document.querySelector(""#searchbutton"").click()'}}"
    
    ' Send this JSON command to trigger the click action
    
    ' Wait for the page to load results (Sleep as a simple way, more advanced would be using the CDP to monitor the page status)
    Sleep 3000

    ' Press Ctrl + P and Save
    SendKeys "^p", True ' Ctrl + P to open print dialog
    Sleep 1000
    
    ' Define the folder and file path for saving the PDF
    Dim savePath As String
    savePath = "C:\Path\To\Save\Folder\" & searchItem & ".pdf" ' Adjust your save path
    
    ' Save the page as a PDF using the Page.printToPDF command
    ' Simulate further interaction if needed to complete the save

    ' Close Chrome after saving
    SendKeys "%{F4}", True ' Alt + F4 to close browser
    Sleep 2000 ' Allow time for closing

End Sub
