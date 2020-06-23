Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Imports Google.Apis.Services
Imports System.IO
Imports Google.Apis.Http

Public Class Form1


    Private credentials As IConfigurableHttpClientInitializer

    Public Sub deleteOnSheet()

        Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'Access to the SpreadSheets
        Dim ApplicationName As String = "TestGoogleSheets" 'Only the name -> Does't correspond to anything
        Dim spreadsheetId As String = "1xv0f2W4THCBgE7ahkb3-BjqdfbCLKxsrgwO8enLnSMg" ' Id on the GoogleSheets Link Copy-> Paste
        Dim subSheetID As String = "congress" 'Page id Tab (Downpage on Sheets)

        'Take credentials to access
        Dim cred As GoogleCredential
        Using stream = New FileStream("../../client_secrets.json", FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
            cred = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {
                                        .HttpClientInitializer = cred,
                                        .ApplicationName = ApplicationName}) 'Create Google Sheets API service for connecting to the API.

        'Start the Delete code
        Dim range As String = subSheetID & "!E522:H522"
        Dim requestBody = New ClearValuesRequest()
        Dim deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range)
        Dim deleteReponse = deleteRequest.Execute()
    End Sub

    Public Sub updateSheet()

        Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'Access to the SpreadSheets
        Dim ApplicationName As String = "TestGoogleSheets" 'Only the name -> Does't correspond to anything
        Dim spreadsheetId As String = "1xv0f2W4THCBgE7ahkb3-BjqdfbCLKxsrgwO8enLnSMg" ' Id on the GoogleSheets Link Copy-> Paste
        Dim subSheetID As String = "congress" 'Page id Tab (Downpage on Sheets)



        Dim cred As GoogleCredential
        Using stream = New FileStream("../../client_secrets.json", FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
            cred = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {
                                        .HttpClientInitializer = cred,
                                        .ApplicationName = ApplicationName}) 'Create Google Sheets API service for connecting to the API.


        'Start the Update code
        Dim range As String = subSheetID & "!D541" 'Range to update
        Dim valueRange = New ValueRange()
        Dim oblist = New List(Of Object)() From { 'Data to update on thar range
            "updated"
        }
        valueRange.Values = New List(Of IList(Of Object)) From {
            oblist
        }
        'Request to update
        Dim updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
        Dim appendReponse = updateRequest.Execute()
    End Sub

    Public Sub insertToSheet()
        Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'Access to the SpreadSheets
        Dim ApplicationName As String = "TestGoogleSheets" 'Only the name -> Does't correspond to anything
        Dim spreadsheetId As String = "1xv0f2W4THCBgE7ahkb3-BjqdfbCLKxsrgwO8enLnSMg" ' Id on the GoogleSheets Link Copy-> Paste
        Dim subSheetID As String = "congress" 'Page id Tab (Downpage on Sheets)



        Dim cred As GoogleCredential
        Using stream = New FileStream("../../client_secrets.json", FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
            cred = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {
                                        .HttpClientInitializer = cred,
                                        .ApplicationName = ApplicationName}) 'Create Google Sheets API service for connecting to the API.

        'Start the Insert code
        Dim range As String = subSheetID & "!A:F"  'Cells where will be insert data
        Dim oblist = New List(Of Object)() From {"Data", "to", "write", "via", "vb.net", "on sheets"} 'Data to write
        Dim valueRange As New ValueRange()
        valueRange.MajorDimension = "ROWS" 'or "COLUMNS" depends on we want
        valueRange.Values = New List(Of IList(Of Object))() From {oblist}

        'Write request that require a ValueRange - spreadsheetID and the range
        Dim writeCellRowRequest As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range)
        writeCellRowRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW
        writeCellRowRequest.Execute()

    End Sub

    Public Sub readSheet()

        Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'Access to the SpreadSheets
        Dim ApplicationName As String = "TestGoogleSheets" 'Only the name -> Does't correspond to anything
        Dim spreadsheetId As String = "1xv0f2W4THCBgE7ahkb3-BjqdfbCLKxsrgwO8enLnSMg" ' Id on the GoogleSheets Link Copy-> Paste
        Dim subSheetID As String = "congress" 'Page id Tab (Downpage on Sheets)


        Dim cred As GoogleCredential
        Using stream = New FileStream("../../client_secrets.json", FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
            cred = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
        End Using

        Dim service = New SheetsService(New BaseClientService.Initializer() With {
                                        .HttpClientInitializer = cred,
                                        .ApplicationName = ApplicationName}) 'Create Google Sheets API service for connecting to the API.



        'Start the Read code
        Dim range As String = subSheetID & "!A1:F10" 'Range for read
        Dim request As SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get(spreadsheetId, range)
        Dim response = request.Execute()
        Dim values As IList(Of IList(Of Object)) = response.Values

        'Read values on the range specified
        If values IsNot Nothing AndAlso values.Count > 0 Then
            For Each row In values
                Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5}", row(0), row(1), row(2), row(3), row(4), row(5))
            Next
        Else
            Console.WriteLine("No data found.")
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        insertToSheet()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        readSheet()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        updateSheet()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        deleteOnSheet()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class

