Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =35
    Left =5385
    Top =1755
    Right =12330
    Bottom =5805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =4320
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =2340
                    Width =1739
                    Height =300
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5759
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1245
                    Top =180
                    Width =4755
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="NFV Duplicates Report"
                    LayoutCachedLeft =1245
                    LayoutCachedTop =180
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =2340
                    Width =1740
                    Height =300
                    TabIndex =1
                    Name ="ButtonCanopyGap"
                    Caption ="Duplicate Report"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =2340
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =4140
                    Top =840
                    Width =1860
                    TabIndex =2
                    ColumnInfo ="\"ProjectID\";\"\";\"Project Name\";\"\";\"10\";\"70\""
                    Name ="ProjectID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblProjects].[ProjectID], [tblProjects].[ProjectName] FROM tblProjects O"
                        "RDER BY [ProjectID]; "
                    ColumnWidths ="1008;6192"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =840
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =840
                            Width =2760
                            Height =245
                            FontWeight =700
                            Name ="ProjectID_Label"
                            Caption ="Select a project ID if desired"
                            LayoutCachedLeft =1260
                            LayoutCachedTop =840
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =1085
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4500
                    Top =1320
                    Width =1200
                    Height =239
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="StartDate"
                    Format ="Short Date"
                    InputMask ="99/99/0000;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1559
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1320
                            Width =3120
                            Height =239
                            FontWeight =700
                            Name ="Label38"
                            Caption ="Enter a date range if desired - From"
                            LayoutCachedLeft =1260
                            LayoutCachedTop =1320
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =1559
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4500
                    Top =1800
                    Width =1200
                    Height =239
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EndDate"
                    Format ="Short Date"
                    InputMask ="99/99/0000;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =2039
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =1800
                            Width =360
                            Height =239
                            FontWeight =700
                            Name ="Label40"
                            Caption ="To"
                            LayoutCachedLeft =4020
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =2039
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Button_Close_Click()

    DoCmd.Close

End Sub

Private Sub ButtonCanopyGap_Click()
On Error GoTo Err_CanopyGap_Click

  Dim strSQL As String
  Dim strSQLDup As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim DupSite As DAO.Recordset
  Dim Obsvalues As DAO.Recordset
  Dim ObsDups As DAO.Recordset
  Dim RPD As Double
  Dim NoDupMessage As String
  Dim Response As Integer
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Duplicates"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Duplicates WHERE (StationID = '4951260' OR StationID = '4951265')"    ' Selecting only the two primary site ids
  If Not IsNull(Me!ProjectID) Then
    strSQL = strSQL & " AND ProjectID = '" & Me!ProjectID & "'"
  End If
  If Not IsNull(Me!StartDate) Then
    If IsNull(Me!EndDate) Then
      MsgBox "You must enter an end date.", vbOKOnly, "CAB Report"
      Exit Sub
    Else
      strSQL = strSQL & " AND ([START_DATE] BETWEEN #" & Me!StartDate & "# AND #" & Me!EndDate & "#)"
    End If
  End If
  strSQL = strSQL & " ORDER BY START_DATE, StationID, DISPLAY_NAME"
  
  '  Build SQL statement for dup sites
  strSQLDup = "SELECT * FROM qry_Duplicates WHERE (StationID = '4951261' OR StationID = '4951266')"    ' Selecting only the two dup site ids
  If Not IsNull(Me!ProjectID) Then
    strSQLDup = strSQLDup & " AND ProjectID = '" & Me!ProjectID & "'"
  End If
  If Not IsNull(Me!StartDate) Then
    strSQLDup = strSQLDup & " AND ([START_DATE] BETWEEN #" & Me!StartDate & "# AND #" & Me!EndDate & "#)"
  End If

  DoCmd.Hourglass True
  Set db = CurrentDb
   ' MsgBox strSQL
   Set Obsvalues = db.OpenRecordset(strSQL)
   If Obsvalues.EOF Then
     MsgBox "No valid Result records found.", vbOKOnly, "Duplicate Report"
     Obsvalues.Close
     Set Obsvalues = Nothing
     GoTo Exit_CanopyGap_Click
   End If
   Set ObsDups = db.OpenRecordset(strSQLDup)
   
   '  Get the dup sites out of the way first
   Do Until ObsDups.EOF
     Select Case ObsDups!StationID
       Case "4951261"
         strSQL = "SELECT * FROM qry_Duplicates WHERE StationID = '4951260' AND  DISPLAY_NAME = '" & ObsDups!DISPLAY_NAME & "' AND [START_DATE] = #" & ObsDups!START_DATE & "#"    ' Fetch duplicate record
       Case "4951266"
         strSQL = "SELECT * FROM qry_Duplicates WHERE StationID = '4951265' AND  DISPLAY_NAME = '" & ObsDups!DISPLAY_NAME & "' AND [START_DATE] = #" & ObsDups!START_DATE & "#"    ' Fetch duplicate record
     End Select
     Set DupSite = db.OpenRecordset(strSQL)
     If DupSite.EOF Then
       NoDupMessage = "No matching results found for " & ObsDups!StationID & " " & ObsDups!START_DATE & " " & ObsDups!DISPLAY_NAME & "."
       NoDupMessage = NoDupMessage & Chr(13) & Chr(10) & "Do you want to exit?"
       Response = MsgBox(NoDupMessage, vbYesNo, "Duplicate Missing")
       If Response = vbYes Then
         GoTo Exit_CanopyGap_Click
       End If
     End If
     DupSite.Close
     Set DupSite = Nothing
     ObsDups.MoveNext
   Loop
   ObsDups.Close
   Set ObsDups = Nothing
   
   ' Now the primary sites
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Duplicates")
   Do Until Obsvalues.EOF
     Select Case Obsvalues!StationID
       Case "4951260"
         strSQL = "SELECT * FROM qry_Duplicates WHERE StationID = '4951261' AND  DISPLAY_NAME = '" & Obsvalues!DISPLAY_NAME & "' AND [START_DATE] = #" & Obsvalues!START_DATE & "#"    ' Fetch duplicate record
       Case "4951265"
         strSQL = "SELECT * FROM qry_Duplicates WHERE StationID = '4951266' AND  DISPLAY_NAME = '" & Obsvalues!DISPLAY_NAME & "' AND [START_DATE] = #" & Obsvalues!START_DATE & "#"    ' Fetch duplicate record
     End Select
     Set DupSite = db.OpenRecordset(strSQL)
     If DupSite.EOF Then
       NoDupMessage = "No matching results found for " & Obsvalues!StationID & " " & Obsvalues!START_DATE & " " & Obsvalues!DISPLAY_NAME & "."
       NoDupMessage = NoDupMessage & Chr(13) & Chr(10) & "Do you want to exit?"
       Response = MsgBox(NoDupMessage, vbYesNo, "Duplicate Missing")
       If Response = vbYes Then
         GoTo Exit_CanopyGap_Click
       Else
         GoTo NextRecord
       End If
     End If
     If Obsvalues!Numeric_Result + DupSite!Numeric_Result <> 0 Then
       RPD = Abs(((Obsvalues!Numeric_Result - DupSite!Numeric_Result) / ((Obsvalues!Numeric_Result + DupSite!Numeric_Result) / 2)) * 100)
       If RPD > 30 Then
         ' Write outlier record
         WorkOutput.AddNew
         WorkOutput!ProjectID = Obsvalues!ProjectID  '
         WorkOutput!StationID = Obsvalues!StationID  '
         WorkOutput!StationName = Obsvalues!StationName
         WorkOutput!START_DATE = Obsvalues!START_DATE  '
         WorkOutput!CharacteristicName = Obsvalues!DISPLAY_NAME  '
         WorkOutput!RPD = RPD
         WorkOutput!DetectionCondition = Obsvalues!Detection_Condition  '
         WorkOutput!ResultValue = Obsvalues!Result_Text  '
         WorkOutput!RemarkCode = Obsvalues!Lab_Remarks  '
         WorkOutput!ResultComment = Obsvalues!Result_Comment  '
         WorkOutput!VisitComment = Obsvalues!Visit_Comment  '
         WorkOutput.Update
       End If
     End If  ' End if for divide by zero trap
NextRecord:
     DupSite.Close
     Set DupSite = Nothing
     Obsvalues.MoveNext
   Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
     Obsvalues.Close
     Set Obsvalues = Nothing
     
Exit_CanopyGap_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Duplicates."
    Exit Sub

Err_CanopyGap_Click:
    MsgBox Err.Description & " " & ArrayIndex
    Resume Exit_CanopyGap_Click
End Sub


Private Sub Form_Open(Cancel As Integer)
  DoCmd.Restore
End Sub


Private Sub Visit_Date_AfterUpdate()
  If Not IsNull(Me!Visit_Date) Then
    Me!StartDate = DateSerial(Me!Visit_Date, 10, 1)
    Me!EndDate = DateSerial(Me!Visit_Date + 1, 9, 30)
  End If
End Sub
