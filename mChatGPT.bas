Attribute VB_Name = "mChatGPT"
' ===================================================================
' Project Name: OpenAI ChatGPT in Excel
' Author: Sven Bosau
' Website: https://pythonandvba.com
' YouTube: https://youtube.com/@codingisfun
' Date Created: 2023/10/07
' Last Modified: 2023/10/07
' Version: 1.0
' ===================================================================
'
' Description:
' This VBA module enables users to interact with OpenAI's GPT model
' directly from Excel by sending text prompts via OpenAI's API.
' It allows capturing AI model outputs and display them in Excel.
'
' ===================================================================
'
' Promotional Note:
' For an even more enhanced and user-friendly interaction with OpenAI
' and other advanced features, check out the Excel add-in "MyToolBelt"!
' It offers a superior user interface, easier settings management, and
' additional error handling capabilities.
' Find out more at: https://pythonandvba.com/mytoolbelt
'
' ===================================================================
'
' DISCLAIMER:
' This code is distributed "as is" and the author makes no representations
' or warranties, express or implied, regarding the functionality, operability,
' or use of the code, including, without limitation, any implied warranties of
' merchantability or fitness for a particular purpose. The user of this code
' assumes the entire risk as to its quality and performance. Should any part
' of the code prove defective, the user assumes the entire cost of all necessary
' servicing or repair.
'
' The user must comply with all applicable local laws and regulations in
' using the code, including, without limitation, all intellectual property laws.
'
' Furthermore, by using this code, the user acknowledges and agrees that
' they have read and understand OpenAI's use-case policies and agree to abide by them.
' OpenAI's use-case policy can be found at https://platform.openai.com/docs/use-case-policy
'
' The API key is confidential and should be kept secure. Sharing or exposing
' the API key is strictly prohibited. Use the API key responsibly and ensure
' it is stored, transmitted, and used securely.
'
' ===================================================================


' ===================================================================
'                             OPENAI SETTINGS
' ===================================================================
'
' API Key:
'   Obtain your OpenAI API key from: https://platform.openai.com/account/api-keys
Const OPENAI_API_KEY As String = "sk-XXX"
'
' Model Specification:
'   Specify the OpenAI model to interact with. Default set to GPT-4.
'   Adjust as per your usage and API capabilities.
Const OPENAI_MODEL As String = "gpt-4" ' Alternatively: 'gpt-3.5-turbo
'
' System Content:
'   Set the initial system content to establish the context of the
'   assistant (useful in maintaining consistent interactions).
Const OPENAI_SYSTEM_CONTENT As String = "You are a helpful assistant"
'
' Max Tokens:
'   Define the maximum number of tokens (words/characters) in the
'   response. Adjust according to your requirements.
Const OPENAI_MAX_TOKENS As String = "4096"
'
' Temperature:
'   Determine the randomness of the model’s output. Set it between 0 and 1,
'   where 0 is deterministic and 1 is very random.
Const OPENAI_TEMPERATURE As String = "0.5"
'
' ===================================================================
Option Explicit

' Declare the InternetGetConnectedState function from "wininet.dll" for checking internet connection.
' VBA7 or later (64-bit) uses PtrSafe, while earlier VBA versions (32-bit) don't.
#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#End If

' =====================================================================================
' Procedure: ChatGPT
' Description: Initiates a chat with OpenAI's GPT model, using user's Excel selection
'              as input, and outputs the AI's response to a dedicated worksheet. Handles
'              error notification and user warnings regarding system compatibility,
'              internet connection, and input validation.
' Notes: Requires internet connection and is not compatible with macOS.
' =====================================================================================
Sub ChatGPT()

10        On Error GoTo ErrorHandler
    #If Mac Then
20            MsgBox "This add-in is designed exclusively for Windows. It is not compatible with macOS.", _
                  vbOKOnly, "Windows Compatibility Only"
30            Exit Sub
    #End If
              
          'Microsoft Scripting Runtime needs to be enabled
40        If Not CheckReference Then Exit Sub
              
          'Check if user has internet connection
          Dim HasInternet As Boolean
50        HasInternet = GetInternetConnectedState()
60        If Not HasInternet Then
70            MsgBox "Internet access is required to use the AI Companion. Please ensure you are connected to the internet and try again.", _
                  vbOKOnly Or vbInformation, "No Internet Access"
80            Exit Sub
90        End If

          ' Ensure a range is selected
100       Application.ScreenUpdating = False
110       If TypeName(Selection) <> "Range" Then
120           MsgBox "Please select only cells to proceed.", vbCritical, "Invalid Selection"
130           Exit Sub
140       End If
          
          Dim prompt As String
150       prompt = ""
          
          ' Loop through each cell in the selection
          Dim cell As Range
160       For Each cell In Selection
              ' If the cell is not empty, add its content to the prompt
170           If Trim(cell.Value) <> "" Then
180               prompt = prompt & CleanJSONString(CStr(cell.Value)) & " "
190           End If
200       Next cell
          
          ' If the concatenated prompt is empty, show an error message and exit
210       If Trim(prompt) = "" Then
220           MsgBox "All selected cells are empty. Please enter some text and try again.", vbCritical, "Empty Input"
230           Application.ScreenUpdating = True
240           Exit Sub
250       End If

          ' Show status in status bar
260       Application.StatusBar = "Processing OpenAI request..."

          ' Create XMLHTTP object
          Dim httpRequest As Object
270       Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
280       requestBody = "{" & _
              """model"": """ & OPENAI_MODEL & """," & _
              """messages"": [" & _
              "{""role"":""system"", ""content"":""" & OPENAI_SYSTEM_CONTENT & """}," & _
              "{""role"":""user"", ""content"":""" & prompt & """}" & _
              "]," & _
              """max_tokens"": " & OPENAI_MAX_TOKENS & "," & _
              """temperature"": " & OPENAI_TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
290       With httpRequest
300           .Open "POST", "https://api.openai.com/v1/chat/completions", False
310           .SetRequestHeader "Content-Type", "application/json"
320           .SetRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
330           .send (requestBody)
340       End With
          
          'Check if the request is successful
350       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
360           response = httpRequest.responseText
              
              'Get the completion and clean it up
              Dim completion As String
370           completion = ParseResponse(response)

              'Split the completion into lines
              Dim lines As Variant
380           lines = Split(Replace(Replace(completion, vbCrLf, vbLf), vbCr, vbLf), vbLf)
              
              ' Get the output worksheet, create if it doesn't exist
              Dim outputWs As Worksheet
390           Set outputWs = GetOrCreateSheet(ThisWorkbook, "AI_OUTPUT")
              Dim outputRange As Range
400           Set outputRange = outputWs.Range("A1") 'Start writing from A1 in the output sheet

              ' Write the lines to the output range
              Dim outputWritten As Boolean
              Dim lastRow As Long
410           lastRow = WriteLinesToRange(lines, outputRange)
420           outputWritten = lastRow > 0
              
              ' Show completion message only if the output was written
430           If outputWritten Then
440               MsgBox "AI completion request has been successfully processed. Please check the output.", vbInformation, "AI Request Completed"
                  
                  ' Select the whole output range
450               outputRange.Parent.Activate
460               outputRange.Resize(RowSize:=lastRow).Select
                  
                  'Autofit columns
470               outputRange.Resize(RowSize:=lastRow).EntireColumn.AutoFit
480           End If
          
490       Else
500           MsgBox "The OpenAI request has failed with status code: " & httpRequest.Status & vbCrLf & vbCrLf & "Error message:" & vbCrLf & httpRequest.responseText, vbCritical, "AI Request Failure"
510       End If
          
520       Application.StatusBar = False
530       Application.ScreenUpdating = True
          
540       Exit Sub
          
ErrorHandler:
550       Application.StatusBar = False
560       Application.ScreenUpdating = True
570       MsgBox "An error occurred: " & Err.Description & vbCrLf & _
              "Error number: " & Err.Number & vbCrLf & _
              "Line number: " & Erl, _
              vbCritical, "Error"
End Sub
' =====================================================================================
' Function: CheckReference
' Description: Verifies if the "Microsoft Scripting Runtime" reference is enabled.
' Returns: BOOLEAN - True if reference is enabled, otherwise False.
' Notes: Informs the user via a message box on how to enable the reference if not enabled.
' =====================================================================================
Private Function CheckReference() As Boolean
          Dim ref As Object
          Dim found As Boolean
580       found = False
          
          ' Loop through all references in the VBA project to find "Scripting"
590       For Each ref In ThisWorkbook.VBProject.References
600           If InStr(1, ref.Name, "Scripting") > 0 Then
610               found = True
620               Exit For
630           End If
640       Next ref
          
          ' Notify user if the "Scripting" reference is not found, otherwise confirm it's available
650       If Not found Then
660           MsgBox "Please enable the Microsoft Scripting Runtime reference!" & vbCrLf & _
                  "Go to Tools -> References... -> and check 'Microsoft Scripting Runtime'", _
                  vbCritical, "Reference Error"
670           CheckReference = False
680       Else
690           CheckReference = True
700       End If
End Function
' =====================================================================================
' Function: ParseResponse
' Parameters:
'   - response As String: The JSON response string obtained from the OpenAI API.
' Description: Parses the JSON response from the API, extracting and returning the
'              content message from the first choice. Provides specific error messages
'              if expected keys are not found in the response.
' Returns: STRING - The content message extracted from the API response.
' Notes: Requires a valid JSON response string as input.
' =====================================================================================
Function ParseResponse(ByVal response As String) As String
710       On Error GoTo ErrorHandler

          ' Initialize the JSON converter and parse the response
          Dim json As Object
720       Set json = JsonConverter.ParseJson(response)

          ' Check if "choices" key exists and it has at least one item
730       If Not json.Exists("choices") Then
740           Err.Raise Number:=vbObjectError + 1024, _
                  Description:="JSON response does not contain 'choices' key."
750       ElseIf json("choices").Count = 0 Then
760           Err.Raise Number:=vbObjectError + 1024, _
                  Description:="JSON response contains 'choices' key but it is empty."
770       End If

          ' Check if "message" key exists in the first choice
780       If Not json("choices")(1).Exists("message") Then
790           Err.Raise Number:=vbObjectError + 1024, _
                  Description:="First choice does not contain 'message' key."
800       End If

          ' Check if "content" key exists in the message of the first choice
810       If Not json("choices")(1)("message").Exists("content") Then
820           Err.Raise Number:=vbObjectError + 1024, _
                  Description:="Message does not contain 'content' key."
830       End If

          ' Extract the "content" field from the JSON response
          Dim content As String
840       content = json("choices")(1)("message")("content")

          ' Return the content
850       ParseResponse = content

860       Exit Function

ErrorHandler:
          ' Return the error description if an error occurs
870       ParseResponse = "Error: " & Err.Description
End Function
' =====================================================================================
' Function: CleanJSONString
' Parameters:
'   - inputStr As String: The string to be cleaned.
' Description: Cleans the provided JSON string by removing line breaks and replacing
'              double quotes with single quotes.
' Returns: STRING - The cleaned string.
' Notes: Utilizes On Error Resume Next to handle potential run-time errors.
' =====================================================================================
Private Function CleanJSONString(inputStr As String) As String
880       On Error Resume Next
          ' Remove line breaks
890       CleanJSONString = Replace(inputStr, vbCrLf, "")
900       CleanJSONString = Replace(CleanJSONString, vbCr, "")
910       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
920       CleanJSONString = Replace(CleanJSONString, """", "'")
930       On Error GoTo 0
End Function
' =====================================================================================
' Function: ReplaceBackslash
' Parameters:
'   - text As Variant: The input text that may contain backslash characters.
' Description: Replaces the backslash character only if it is immediately followed by
'              a double quote character.
' Returns: STRING - The modified string with the backslash characters replaced.
' Notes: Utilizes On Error Resume Next to manage potential run-time errors.
' =====================================================================================
Private Function ReplaceBackslash(text As Variant) As String
940       On Error Resume Next
          Dim i As Integer
          Dim newText As String
950       newText = ""
960       For i = 1 To Len(text)
970           If Mid(text, i, 2) = "\" & Chr(34) Then
980               newText = newText & Chr(34)
990               i = i + 1
1000          Else
1010              newText = newText & Mid(text, i, 1)
1020          End If
1030      Next i
1040      ReplaceBackslash = newText
1050      On Error GoTo 0
End Function
' =====================================================================================
' Function: WriteLinesToRange
' Parameters:
'   - lines As Variant: An array of lines to be written to the Excel range.
'   - rng As Range: The Excel range where the lines should be written.
' Description: Writes an array of strings (lines) into a specified Excel range, starting
'              from the first cell and continuing down. Clears previous content and
'              ensures no Excel formulas are accidentally triggered by prepending an
'              apostrophe if a line begins with "=".
' Returns: LONG - The last row index where the data was written.
' Notes: The rng parameter specifies the starting cell for output writing.
' =====================================================================================
Private Function WriteLinesToRange(lines As Variant, rng As Range) As Long
          Dim i As Long
          Dim overwriteWarningShown As Boolean
1060      overwriteWarningShown = False

          ' Clear output
1070      rng.Worksheet.Cells.ClearContents

1080      For i = LBound(lines) To UBound(lines)
              Dim line As String
1090          line = ReplaceBackslash(lines(i))

              ' Add a single quote if the line starts with an "=" sign (Excel formula!)
1100          If Left(line, 1) = "=" Then
1110              line = "'" & line
1120          End If

1130          rng.Cells(i + 1, 1).Value = line
1140      Next i

1150      WriteLinesToRange = i
End Function
' =====================================================================================
' Function: GetInternetConnectedState
' Description: Checks the internet connection status of the user's machine by invoking
'              the InternetGetConnectedState Windows API function.
' Returns: BOOLEAN - True if the internet is connected, otherwise False.
' Notes: Utilizes On Error Resume Next to gracefully manage potential run-time errors.
' =====================================================================================
Private Function GetInternetConnectedState() As Boolean
          'Check if user has internet connection
1160      On Error Resume Next
1170      GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function
' =====================================================================================
' Function: GetOrCreateSheet
' Parameters:
'   - wb As Workbook: The workbook where the sheet is located or will be created.
'   - sheetName As String: The name of the sheet to find or create.
' Description: Finds a worksheet within the specified workbook with the provided name.
'              If not found, it creates a new sheet with that name.
' Returns: Worksheet - The found or newly created worksheet.
' Notes: Ensure sheet names comply with Excel’s naming rules to prevent run-time errors.
' =====================================================================================
Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
          Dim sheet As Worksheet
          
1180      For Each sheet In wb.Sheets
1190          If sheet.Name = sheetName Then
1200              Set GetOrCreateSheet = sheet
1210              Exit Function
1220          End If
1230      Next sheet
          
          ' If the sheet does not exist, create it
1240      Set GetOrCreateSheet = Sheets.Add(After:=Sheets(Sheets.Count))
1250      GetOrCreateSheet.Name = sheetName
End Function
