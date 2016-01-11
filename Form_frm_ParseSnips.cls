VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_ParseSnips"
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
' Modified by: Basil Khan
' Date: June 22, 2015
' Version   : 2.0
'---------------------------------------------------------------------------------------
   On Error GoTo Err_Handler
  

   If Me.txtSearch1 & "" = "" Then
       If Me.txtSearch2 & "" = "" Then
            If Me.txtSearch3 & "" = "" Then
                If Me.txtSearch4 & "" = "" Then
                    If Me.txtSearch5 & "" = "" Then
                        If Me.txtSearch6 & "" = "" Then
                  MsgBox "No search word entered"
                    Exit Sub
                             End If
                         End If
                    End If
                End If
            End If
         End If

   
    
   'Start
      DoCmd.Hourglass True
      Me.txtStatus.Visible = True
      Dim Start, Finish, TotalTime
      Start = Timer
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
   Dim j As Integer
   Dim k As Integer
   Dim seconds As Integer
   Dim minutes As Integer
   Me.txtStatus.SetFocus
   Me.txtStatus.SelLength = 0
   
  
    
         k = 1
         Do While k <= 6 ' Iterations for each Keyword
                          
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
            Me.txtStatus = ""   ' Eliminate overflow errors
         'Search each field
            For Each fld In tdf.Fields
               i = i + 1
                
                If SearchText(k) = "NoWordFound" Then
                     k = k + 1
                     i = 0
                     End If
                     
                If k >= 7 Then
                             Me.frm_SearchResults.Requery
                             Me.frm_SearchResults.Visible = True
                             
                                                 
                               Set rs = Nothing
                               Set rsResults = Nothing
                               Set oDB = Nothing
                               Set tdf = Nothing
                               Set AutoField = Nothing
                               
                               DoCmd.Hourglass False
                               On Error GoTo 0
                               
                               Finish = Timer    ' Set end time.
                                TotalTime = (Finish - Start) / 60  ' Calculate total time.
                                minutes = CInt(TotalTime)
                                If TotalTime > minutes Then
                                    seconds = (TotalTime - minutes) * 60
                                Else
                                    seconds = (TotalTime - (minutes - 1)) * 60
                                End If
                                MsgBox "Search Completed Successfully!" & " Search took " & minutes & " minute(s) and " & seconds & " seconds."
                               Exit Sub
                            
                                            
                             End If
                     
               Select Case fld.Type
                  Case dbBigInt, dbByte, dbChar, dbCurrency, dbDate, dbDecimal, dbDouble, dbFloat, dbGUID, dbInteger, dbLong, dbMemo, dbNumeric, dbSingle, dbText, dbTime, dbTimeStamp
                     If Not rs.EOF Then rs.MoveFirst
                     
                     Me.txtStatus = Me.txtStatus & "Searching [" & tdf.Name & "] field " & i & " of " & tdf.Fields.Count & " for Keyword " & k & vbNewLine
                     Me.Repaint
                     Me.txtStatus.SelStart = Len(Me.txtStatus)
                     If CheckBox(k) = False Then
                     strCrit = "[" & fld.Name & "]" & " LIKE ""*" & SearchText(k) & "*"""
                     Else
                     strCrit = "[" & fld.Name & "]" & " LIKE """ & SearchText(k) & """"
                     End If
                    

                     rs.FindFirst strCrit
                     Do While Not rs.NoMatch
                        rsResults.AddNew
                           If bHasAuto Then
                              rsResults!lng_AutoNumber = rs.Fields(AutoField.Name).Value
                           End If
                           rsResults!tx_TableName = tdf.Name
                           rsResults!tx_FieldName = fld.Name
                           rsResults!mem_Value = rs.Fields(fld.Name).Value
                           
                           
                           'Append Relevant data from row
                           
                           j = 1
                           rsResults!mem_Desc = ""
                           Do While j < CurrentDb.TableDefs(tdf.Name).Fields.Count
                           rsResults!mem_Desc = rsResults!mem_Desc & rs.Fields(j).Name & " ==> " & rs.Fields(j).Value & vbCrLf
                           
                           j = j + 1
                           
                           Loop
                           
                           
                           
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
   k = k + 1
 
   Loop

            Finish = Timer    ' Set end time.
            TotalTime = (Finish - Start) / 60  ' Calculate total time.
            minutes = CInt(TotalTime)
            If TotalTime > minutes Then
                seconds = (TotalTime - minutes) * 60
            Else
                seconds = (TotalTime - (minutes - 1)) * 60
            End If
            MsgBox "Search Completed Successfully!" & " Search took " & minutes & " minute(s) and " & seconds & " seconds."

Exit_Sub:
Me.frm_SearchResults.Requery
Me.frm_SearchResults.Visible = True

                                          
                        Set rs = Nothing
                        Set rsResults = Nothing
                        Set oDB = Nothing
                        Set tdf = Nothing
                        Set AutoField = Nothing
   DoCmd.Hourglass False
   On Error GoTo 0
   
                    Exit Sub

Err_Handler:
   DoCmd.Hourglass False
   Me.frm_SearchResults.Visible = False
   MsgBox Err.Number & " - " & Err.Description & vbNewLine & _
           "In procedure: btn_Search_Click in Form_frm_Search"
    MsgBox "Search Not Completed - Please debug the code or check the excel files"
   Resume Exit_Sub
   
End Sub

Private Function SearchText(k) As String
  
        If k = 1 Then
            If Me.txtSearch1 & "" = "" Then
            SearchText = "NoWordFound"
            Else
            SearchText = Me.txtSearch1
            
            Exit Function
        End If
        End If
            
        If k = 2 Then
            If Me.txtSearch2 & "" = "" Then
           SearchText = "NoWordFound"
          Else
            SearchText = Me.txtSearch2
            
            
            Exit Function
        End If
        End If
        
        If k = 3 Then
            If Me.txtSearch3 & "" = "" Then
          SearchText = "NoWordFound"
           Else
            SearchText = Me.txtSearch3
                        
            Exit Function
        End If
        End If
        
        If k = 4 Then
            If Me.txtSearch4 & "" = "" Then
          SearchText = "NoWordFound"
           Else
            SearchText = Me.txtSearch4
                        
            Exit Function
        End If
        End If
        
        If k = 5 Then
            If Me.txtSearch5 & "" = "" Then
          SearchText = "NoWordFound"
           Else
            SearchText = Me.txtSearch5
                        
            Exit Function
        End If
        End If
        
        If k = 6 Then
            If Me.txtSearch6 & "" = "" Then
          SearchText = "NoWordFound"
           Else
            SearchText = Me.txtSearch6
                        
            Exit Function
        End If
        End If
            
End Function

Private Function CheckBox(k) As Boolean 'function used for Match cell content box
  
        If k = 1 Then
            If Check1 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
            
        If k = 2 Then
            If Check2 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
        
       If k = 3 Then
            If Check3 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
        
        If k = 4 Then
            If Check4 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
        
        If k = 5 Then
            If Check5 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
        
        If k = 6 Then
            If Check6 = True Then
            CheckBox = True
            Else
            CheckBox = False
            
            Exit Function
        End If
        End If
            
End Function


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


   


