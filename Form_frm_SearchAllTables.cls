VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SearchAllTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub btn_Search_Click()
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : btn_Search_Click
' Author    : AEC - Anders Ebro Christensen / TheSmileyCoder
' Date      : 2013-02-22
' Version   : 1.0
' Purpose   :
' Bugs?     : Email: SmileyCoderTools@gmail.com
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
   On Error GoTo Err_Handler

   If Me.txtSearch & "" = "" Then
      MsgBox "No search word entered"
      Exit Sub
   End If
   
   'Start
      DoCmd.Hourglass True
      Me.txtStatus.Visible = True
   'Clean out old results
      CurrentDb.Execute "DELETE * FROM tbl_SearchResults"
      
   Dim tdf As DAO.TableDef
   Dim oDB As DAO.Database
   Set oDB = CurrentDb
   Dim AutoField As DAO.Field
   Dim bHasAuto As Boolean
   Dim strSQL As String
   Dim rs As DAO.Recordset
   Dim fld As DAO.Field
   Dim rsResults As DAO.Recordset
   Set rsResults = oDB.OpenRecordset("tbl_SearchResults", dbOpenDynaset)
   Dim strCrit As String
   Dim i As Integer
   Me.txtStatus.SetFocus
   Me.txtStatus.SelLength = 0
   For Each tdf In oDB.TableDefs
      'Exclude Microsoft and custom system tables, and the results table
      If Mid(tdf.Name, 2, 3) <> "sys" And tdf.Name <> "tbl_SearchResults" Then
         'Determine if table has a autonumberfield
            bHasAuto = GetAutoNumberField(tdf, AutoField)
         
         'Open recordset
            strSQL = "SELECT * FROM [" & tdf.Name & "]"
            If bHasAuto Then
               strSQL = strSQL & " ORDER BY [" & AutoField.Name & "]"
            End If
            Set rs = oDB.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
            i = 0
         'Search each field
            For Each fld In tdf.Fields
               i = i + 1
               Select Case fld.Type
                  Case dbBigInt, dbByte, dbChar, dbCurrency, dbDate, dbDecimal, dbDouble, dbFloat, dbGUID, dbInteger, dbLong, dbMemo, dbNumeric, dbSingle, dbText, dbTime, dbTimeStamp
                     If Not rs.EOF Then rs.MoveFirst
                     
                     Me.txtStatus = Me.txtStatus & "Searching [" & tdf.Name & "] field " & i & " of " & tdf.Fields.Count & vbNewLine
                     Me.Repaint
                     Me.txtStatus.SelStart = Len(Me.txtStatus)
                     
                     
                     
                     strCrit = "[" & fld.Name & "]" & " LIKE ""*" & Me.txtSearch & "*"""
                     rs.FindFirst strCrit
                     Do While Not rs.NoMatch
                        rsResults.AddNew
                           If bHasAuto Then
                              rsResults!lng_AutoNumber = rs.Fields(AutoField.Name).Value
                           End If
                           rsResults!tx_TableName = tdf.Name
                           rsResults!tx_FieldName = fld.Name
                           rsResults!mem_Value = rs.Fields(fld.Name).Value
                        rsResults.Update
                        rs.FindNext strCrit
                     Loop
            
                  Case Else
                     Me.txtStatus = Me.txtStatus & "Exluding [" & tdf.Name & "][" & fld.Name & "],  field " & i & " of " & tdf.Fields.Count & vbNewLine
                     Me.Repaint
                     Me.txtStatus.SelStart = Len(Me.txtStatus)
               End Select
            Next
         
         
      End If
   Next
   Me.frm_SearchResults.Requery
   Me.frm_SearchResults.Visible = True
   
   Set rs = Nothing
   Set rsResults = Nothing
   Set oDB = Nothing
   

Exit_Sub:
   DoCmd.Hourglass False
   On Error GoTo 0
   Exit Sub

Err_Handler:
   DoCmd.Hourglass False
   Me.frm_SearchResults.Visible = False
   MsgBox Err.Number & " - " & Err.Description & vbNewLine & _
            "In procedure: btn_Search_Click in Form_frm_Search"
   Resume Exit_Sub
   
End Sub




Private Function GetAutoNumberField(oTDF As DAO.TableDef, oField As DAO.Field) As Boolean
   Dim fld As DAO.Field
   For Each fld In oTDF.Fields
        If fld.Attributes And dbAutoIncrField Then
            Set oField = fld
            GetAutoNumberField = True
            Exit Function
        End If
    Next fld
   Set oField = Nothing
   GetAutoNumberField = False
End Function


   
