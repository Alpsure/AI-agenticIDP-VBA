' Constants
Const sTesseract As String = "C:\Program Files\Tesseract-OCR\tesseract.exe"

Sub ProcessDocuments()
    Dim ws As Worksheet
    Dim fileList As Object
    Dim fileName As String
    Dim filePath As String
    Dim rowIndex As Long
    Dim extractedText As String
    Dim category As String
    Dim sFolderInput As String
    Dim sFolderOutput As String
    Dim categorySheet As Worksheet

    ' Initialize
    Set ws = ThisWorkbook.Sheets("Files")
    rowIndex = 2

    Call clean_wb(ThisWorkbook)

    ' Loop through input folder
    sFolderInput = ThisWorkbook.Path & "\INPUT"
    sFolderOutput = ThisWorkbook.Path & "\OUTPUT"
    
    DeleteAllFilesInOutputFolder (sFolderOutput)
    
    Set fileList = CreateObject("Scripting.FileSystemObject").GetFolder(sFolderInput).Files
    For Each file In fileList
        Set categorySheet = Nothing
        fileName = file.Name
        filePath = file.Path
        ws.Cells(rowIndex, 1).Value = fileName

        ' Check file type
        If Not IsKnownFileType(fileName) Then
            ws.Cells(rowIndex, 2).Value = "Unknown file type"
        Else
            ' Extract text using Tesseract
            extractedText = PerformOCR(filePath)
            SaveToFile sFolderOutput & "\" & GetBaseFileName(fileName) & "_RAW.txt", extractedText

            ' Read text and categorize
            extractedText = ReadFile(sFolderOutput & "\" & GetBaseFileName(fileName) & "_RAW.txt")
            category = CategorizeText(SanitizeText(extractedText))

            ' Create category sheet if not exists
            On Error Resume Next
            Set categorySheet = ThisWorkbook.Sheets(category)
            If categorySheet Is Nothing Then
                Set categorySheet = ThisWorkbook.Sheets.Add
                categorySheet.Name = category
            End If
            On Error GoTo 0

            ' Extract data
            Dim jsonData As String
            jsonData = ExtractData(extractedText, category)
            Call WriteToCategorySheet(categorySheet, jsonData)
        End If
        rowIndex = rowIndex + 1
    Next file

    MsgBox "Processing complete!"
End Sub
Sub clean_wb(wb As Workbook)
    
    Dim wsF As Worksheet
    Set wsF = wb.Worksheets("Files")
    
    ' Clear existing data
    wsF.Cells.Clear
    wsF.Cells(1, 1).Value = "File Name"
    wsF.Cells(1, 2).Value = "Status"
    
    For Each ws In wb.Worksheets
        sheetName = ws.Name
        ' Check if the sheet is not "Files" or "Secrets"
        If sheetName <> "Files" And sheetName <> "Secrets" Then
            ' Delete the sheet
            Application.DisplayAlerts = False ' Turn off confirmation alert
            ws.Delete
            Application.DisplayAlerts = True ' Turn it back on
        End If
    Next ws

End Sub
Sub DeleteAllFilesInOutputFolder(sFolderOutput)
    Dim fso As Object
    Dim outputFolder As String
    Dim folder As Object
    Dim file As Object
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fso.FolderExists(sFolderOutput) Then
        MsgBox "OUTPUT folder not found: " & sFolderOutput, vbExclamation
        Exit Sub
    End If
    
    ' Get the folder object
    Set folder = fso.GetFolder(sFolderOutput)
    
    ' Loop through each file in the folder and delete it
    For Each file In folder.Files
        On Error Resume Next ' Handle any potential errors
        file.Delete True ' True for force deletion
        On Error GoTo 0
    Next file

End Sub
Function IsKnownFileType(fileName As String) As Boolean
    Dim ext As String
    ext = LCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
    IsKnownFileType = (ext = "png" Or ext = "jpg" Or ext = "jpeg")
End Function

Function PerformOCR(filePath As String) As String
    Dim shellObj As Object
    Dim command As String
    Dim tempFilePath As String

    tempFilePath = Environ("TEMP") & "\ocr_output.txt"
    command = """" & sTesseract & """ """ & filePath & """ """ & Left(tempFilePath, Len(tempFilePath) - 4) & """"

    Set shellObj = CreateObject("WScript.Shell")
    shellObj.Run command, 0, True
    PerformOCR = ReadFile(tempFilePath)
End Function

Function SaveToFile(filePath As String, content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Function

Function ReadFile(filePath As String) As String
    Dim fileNum As Integer
    Dim content As String
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input(LOF(fileNum), fileNum)
    Close #fileNum
    ReadFile = content
End Function
Function CallAPI(sprompt As String, role_description As String, apiKey As String) As String
    Dim xhr As Object
    Dim jsonBody As String
    Dim jsonPrompt As String
    Dim jsonRole As String
    Dim sanitizedPrompt As String
    Dim sanitizedRole As String
    
    ' Sanitize inputs
    sanitizedPrompt = SanitizeText(sprompt)
    sanitizedRole = SanitizeText(role_description)
    
    ' Construct the JSON payload
    jsonRole = "{""role"":""system"",""content"":""" & sanitizedRole & """}"
    jsonPrompt = "{""role"":""user"",""content"":""" & sanitizedPrompt & """}"
    jsonBody = "{""model"":""gpt-4o-mini"",""messages"":[" & jsonRole & "," & jsonPrompt & "],""max_tokens"":2000,""temperature"":0.7}"
    
    ' Debugging: Print the constructed JSON body
    Debug.Print "JSON Body: " & jsonBody
    
    ' Create the HTTP request object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    
    ' Open a POST request to the API endpoint
    On Error GoTo ErrorHandler
    xhr.Open "POST", "https://api.openai.com/v1/chat/completions", False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Send the request
    xhr.send jsonBody
    
    ' Return the response text
    CallAPI = xhr.ResponseText
    Debug.Print CallAPI
    Exit Function

ErrorHandler:
    ' Return a detailed error message
    CallAPI = "Error: " & Err.Description & " | Response: " & xhr.ResponseText
    Debug.Print "Error: " & Err.Description
    Debug.Print "Response: " & xhr.ResponseText
End Function
Function SanitizeText(inputText As String) As String
    Dim sanitizedText As String
    Dim i As Long
    Dim charCode As Long
    
    ' Escape special JSON characters
    sanitizedText = Replace(inputText, "\", "\\") ' Escape backslashes
    sanitizedText = Replace(sanitizedText, """", "\""") ' Escape double quotes
    sanitizedText = Replace(sanitizedText, Chr(10), "\n") ' Line break (LF) -> \n
    sanitizedText = Replace(sanitizedText, Chr(13), "") ' Remove carriage returns (CR)

    ' Remove any remaining non-printable ASCII characters
    For i = 1 To Len(sanitizedText)
        charCode = Asc(Mid(sanitizedText, i, 1))
        If charCode >= 32 And charCode <= 126 Then
            SanitizeText = SanitizeText & Mid(sanitizedText, i, 1)
        End If
    Next i
End Function
Function CategorizeText(text As String) As String
    Dim sprompt As String
    Dim roleDescription As String
    Dim apiResponse As String
    Dim apiKey As String
    Dim category As String
    
    ' Get the API key from the "Secrets" sheet, cell B2
    apiKey = ThisWorkbook.Sheets("Secrets").Range("B2").Value
    
    ' Define the role and prompt
    roleDescription = "You are a categorization assistant. Based on the input text, identify the most likely category from the predefined list."
    sprompt = "Here is the document text:\n" & text & vbNewLine & "Categories: Invoice, Receipt, Report, Letter, Unknown. Return only the category name."
    
    ' Call the GPT API
    apiResponse = CallAPI(sprompt, roleDescription, apiKey)
    Debug.Print apiResponse
    
    ' Parse the response to extract the category
    category = ParseCategoryFromResponse(apiResponse)
    
    ' Return the category
    CategorizeText = category
End Function
Function ParseCategoryFromResponse(apiResponse As String) As String
    Dim jsonResponse As String
    Dim content As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find the position of the "choices" array
    startPos = InStr(apiResponse, "choices")
    temp_apiresponse = Mid(apiResponse, startPos)
    
    startPos = InStr(temp_apiresponse, "content") + 11
    temp_apiresponse = Mid(temp_apiresponse, startPos)
    
    startPos = InStr(temp_apiresponse, """") - 1
    temp_apiresponse = Mid(temp_apiresponse, 1, startPos)
    
    ParseCategoryFromResponse = temp_apiresponse
    
    ' Output the content
    Debug.Print "Content: " & ParseCategoryFromResponse
End Function
Function ParseDataExtractionFromResponse(apiResponse As String) As String
    Dim jsonResponse As String
    Dim content As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find the position of the "choices" array
    startPos = InStr(apiResponse, "choices")
    temp_apiresponse = Mid(apiResponse, startPos)
    
    startPos = InStr(temp_apiresponse, "content") + 11
    temp_apiresponse = Mid(temp_apiresponse, startPos)
    
    endPos = InStr(temp_apiresponse, "refusal") - 1
    temp_apiresponse = Mid(temp_apiresponse, 1, endPos)
    Debug.Print temp_apiresponse
    
    ParseDataExtractionFromResponse = temp_apiresponse
End Function
Sub ParseAndWriteJsonData(jsonResponse As String, ws As Worksheet)
    Dim keyStart As Long, keyEnd As Long
    Dim valueStart As Long, valueEnd As Long
    Dim valueStartList As Long
    Dim currentKey As String, currentValue As String
    Dim rowNum As Long
    Dim headers As Collection
    Dim i As Long
    Dim max_items As Integer
    max_items = 10
    Dim nr_items As Integer
    nr_items = 0
    
    jsonResponse = Replace(jsonResponse, "`", "")
    jsonResponse = Replace(jsonResponse, "\n", "")
    jsonResponse = Replace(jsonResponse, "\", "")
    jsonResponse = Replace(jsonResponse, "  ", " ")
    
    ' Initialize collection for headers
    Set headers = New Collection
    
    ' Start extracting keys and values
    rowNum = 1 ' Row for the headers
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Loop through the JSON string and find key-value pairs
    Do While InStr(jsonResponse, """") > 0 And Len(Trim(jsonResponse)) > 9 And nr_items < max_items
        Debug.Print jsonResponse
        Debug.Print Len(jsonResponse)
        ' Find the start of the key (name of the field)
        keyStart = InStr(jsonResponse, """") + 1
        keyEnd = InStr(keyStart, jsonResponse, """")
        currentKey = Trim(Mid(jsonResponse, keyStart, keyEnd - keyStart))

        ' Find the start and end of the value for the current key
        valueStart = InStr(keyEnd + 2, jsonResponse, """") + 1
        valueStartList = InStr(keyEnd + 2, jsonResponse, "[")
        
        If valueStartList <= 0 Or valueStartList > valueEnd Then
            valueEnd = InStr(valueStart, jsonResponse, """")
            currentValue = Trim(Mid(jsonResponse, valueStart, valueEnd - valueStart))
        Else
            valueEnd = InStr(valueStartList, jsonResponse, "]")
            If valueEnd <= 0 Then
                currentValue = Trim(Mid(jsonResponse, valueStartList))
            Else
                currentValue = Trim(Mid(jsonResponse, valueStartList, valueEnd))
            End If
        End If
        
        ' Add key to the collection (if it's not already in there)
        On Error Resume Next
        headers.Add currentKey, currentKey
        On Error GoTo 0

        ' Write header in the first row (column names)
        ws.Cells(1, headers.Count).Value = currentKey

        ws.Cells(lastRow, headers.Count).Value = currentValue
        
        ' Move to the next key-value pair
        jsonResponse = Mid(jsonResponse, valueEnd + 2)
        
        ' Increment column count
        rowNum = rowNum + 1
        jsonResponse = Replace(jsonResponse, "  ", " ")
        nr_items = nr_items + 1
    Loop
End Sub

Function ExtractData(text As String, category As String) As String
    Dim schema As String
    Dim sprompt As String
    Dim apiResponse As String
    Dim apiKey As String
    Dim jsonFilePath As String
    
    ' Get the API key from the "Secrets" sheet
    apiKey = ThisWorkbook.Sheets("Secrets").Range("B2").Value
    
    ' Path to the JSON schema file
    jsonFilePath = ThisWorkbook.Path & "\extraction_schema.json"
    
    ' Load the JSON schema
    schema = LoadJSONSchema(jsonFilePath, category)
    If schema = "" Then
        ExtractData = "{}" ' Return an empty JSON object if no schema is found
        Exit Function
    End If
    
    ' Create the prompt
    sprompt = "From this document:" & vbNewLine & text & vbNewLine & _
              "...<end of doc>" & vbNewLine & vbNewLine & _
              "Extract the following elements:" & vbNewLine & schema & vbNewLine & vbNewLine & _
              "Return ONLY the JSON data."
    Debug.Print sprompt
    ' Call the GPT API
    apiResponse = CallAPI(sprompt, "You are a data extraction assistant. Extract data based on the provided schema and return JSON only.", apiKey)
    Debug.Print apiResponse
    ' Return the API response
    ExtractData = apiResponse
End Function
Function LoadJSONSchema(filePath As String, category As String) As String
    Dim jsonText As String
    Dim categoryStart As Long
    Dim categoryEnd As Long
    Dim schema As String
    Dim fileContent As String
    Dim fileNumber As Integer

    ' Read the JSON file into a string
    fileNumber = FreeFile
    On Error Resume Next
    Open filePath For Input As #fileNumber
    fileContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    On Error GoTo 0
    
    ' If the file couldn't be read, return an empty string
    If fileContent = "" Then
        LoadJSONSchema = ""
        Exit Function
    End If

    ' Find the schema for the given category
    categoryStart = InStr(fileContent, """" & category & """")
    If categoryStart = 0 Then
        LoadJSONSchema = ""
        Exit Function
    End If
    categoryStart = InStr(categoryStart, fileContent, "{") ' Find the opening brace
    categoryEnd = InStr(categoryStart, fileContent, "}") ' Find the closing brace
    
    If categoryStart > 0 And categoryEnd > 0 Then
        schema = Mid(fileContent, categoryStart, categoryEnd - categoryStart + 1)
    Else
        schema = ""
    End If

    LoadJSONSchema = schema
End Function

Sub WriteToCategorySheet(ws As Worksheet, jsonData As String)
    Dim rowIndex As Long
    Dim parsed_extraction As String
    
    parsed_extraction = ParseDataExtractionFromResponse(jsonData)
    
    Call ParseAndWriteJsonData(parsed_extraction, ws)
    
End Sub
Function GetBaseFileName(fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        GetBaseFileName = Left(fileName, pos - 1)
    Else
        GetBaseFileName = fileName ' No extension found
    End If
End Function


