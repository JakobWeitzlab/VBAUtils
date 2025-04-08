VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_CreateMultipleRowsFromOneRow 
   Caption         =   "UserForm1"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2895
   OleObjectBlob   =   "F_CreateMultipleRowsFromOneRow.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "F_CreateMultipleRowsFromOneRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_CreateMultipleRowsFromOneRow_Click()
    CreateMultipleRowsFromOneRow
End Sub
Public Sub CreateMultipleRowsFromOneRow()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Rem Parameter definieren
    Dim PrmAnzahl, PrmRows As Integer
    
    PrmAnzahl = Box_PrmAnzahl.Value
    PrmRows = Box_PrmRows.Value
    
    Dim currentName As String
    Dim currentProgram As String
    Dim currentDefaultSoftware As String
    Dim currentClientSoftware As String
    Dim currentUsage As String
    Dim currentProgramVersion As String
    Dim currentEOL As String
    Dim currentBewertungsDatum As String
    Dim currentSoftwarePublisher As String
    Dim currentLastPatchDate As String
    
    Dim currentRow As Integer
    
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    For X = 1 To PrmAnzahl
        currentName = ws.Cells(X, 1).Value
                
        Rem Multiply Hostnames/Programs
        For i = 1 To PrmRows
            currentProgram = ws.Cells(i, 2).Value
            currentDefaultSoftware = ws.Cells(i, 3).Value
            currentClientSoftware = ws.Cells(i, 4).Value
            currentUsage = ws.Cells(i, 5).Value
            currentProgramVersion = ws.Cells(i, 6).Value
            currentEOL = ws.Cells(i, 7).Value
            currentBewertungsDatum = ws.Cells(i, 8).Value
            currentSoftwarePublisher = ws.Cells(i, 11).Value
            currentLastPatchDate = ws.Cells(i, 11).Value
            
            ws.Cells(currentRow + i, 13).Value = currentName
            ws.Cells(currentRow + i, 14).Value = currentProgram
            ws.Cells(currentRow + i, 15).Value = currentDefaultSoftware
            ws.Cells(currentRow + i, 16).Value = currentClientSoftware
            ws.Cells(currentRow + i, 17).Value = currentUsage
            ws.Cells(currentRow + i, 18).Value = currentProgramVersion
            ws.Cells(currentRow + i, 19).Value = currentEOL
            ws.Cells(currentRow + i, 20).Value = currentBewertungsDatum
            ws.Cells(currentRow + i, 23).Value = currentSoftwarePublisher
            ws.Cells(currentRow + i, 24).Value = currentLastPatchDate
            
            
        Next i
        
        currentRow = currentRow + PrmRows
    Next X
End Sub
