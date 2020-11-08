Version =20
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
    Left =6420
    Top =2190
    Right =13365
    Bottom =6240
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
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =975
                    Left =4140
                    Top =1260
                    Width =900
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Date"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Visit_Year.Visit_Year FROM qry_Visit_Year ORDER BY qry_Visit_Year.Vis"
                        "it_Year DESC; "
                    ColumnWidths ="975"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1260
                            Width =2760
                            Height =245
                            FontWeight =700
                            Name ="Select a date if desired_Label"
                            Caption ="Select a year beginning 10/1 of"
                            EventProcPrefix ="Select_a_date_if_desired_Label"
                            LayoutCachedLeft =1260
                            LayoutCachedTop =1260
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =1505
                        End
                    End
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
                    Caption ="Outliers Report"
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
                    TabIndex =2
                    Name ="ButtonCanopyGap"
                    Caption ="Outliers Report"
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
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"16\""
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
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =1800
                    Width =540
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="StartDate"
                    Format ="Short Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =1800
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =2115
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =1800
                    Width =540
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EndDate"
                    Format ="Short Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =2115
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
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Obsvalues As DAO.Recordset
  Dim StationSave As String
  Dim NameSave As String
  Dim RecordCount As Integer
  Dim Save95 As Double
  Dim IndexSave As Integer
  Dim ArrayIndex As Integer
  Dim PercentileArray(1000, 7) As Variant ' Array for percentile calculation
  ' Percentile values array
  ' Column x,0 is station id
  ' Column x,1 is Characteristic Name
  ' Column x,2 is Record count
  ' Column x,3 is 5th percentile count
  ' Column x,4 is 95th percentile count
  ' Column x,5 is 5th percentile value
  ' Column x,6 is 95th percentile value
  
  If IsNull(Me!Visit_Date) Then
    MsgBox "You must enter a beginning water year.", vbOKOnly, "Outliers Report"
    Exit Sub
  End If
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Outliers"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Outlier_Count WHERE 1 = 1"
  If Not IsNull(Me!ProjectID) Then
    strSQL = strSQL & "AND ProjectID = '" & Me!ProjectID & "'"
  End If
 '  strSQL = strSQL & "AND StationID = '5995420' AND Display_Name = 'flow'"
  strSQL = strSQL & " ORDER BY StationID, DISPLAY_NAME, NumericValue"
  ' MsgBox strSQL
  DoCmd.Hourglass True
  Set db = CurrentDb

  ' Count records by station, parameter
   Set Obsvalues = db.OpenRecordset(strSQL)
   If Obsvalues.EOF Then
     MsgBox "No valid Result records found.", vbOKOnly, "Outliers Report"
     Obsvalues.Close
     Set Obsvalues = Nothing
     GoTo Exit_CanopyGap_Click
   End If
   ArrayIndex = 0
   Do Until ArrayIndex > 999           ' Initialize array
     PercentileArray(ArrayIndex, 0) = " "
     PercentileArray(ArrayIndex, 1) = " "
     PercentileArray(ArrayIndex, 2) = 0
     PercentileArray(ArrayIndex, 3) = 0
     PercentileArray(ArrayIndex, 4) = 0
     PercentileArray(ArrayIndex, 5) = -999
     PercentileArray(ArrayIndex, 6) = -999
     ArrayIndex = ArrayIndex + 1
   Loop
   IDSave = Obsvalues!StationID    ' Save necessary fields
   NameSave = Obsvalues!DISPLAY_NAME
   ArrayIndex = 0
   PercentileArray(ArrayIndex, 0) = Obsvalues!StationID  ' set new station
   PercentileArray(ArrayIndex, 1) = Obsvalues!DISPLAY_NAME  ' set new parameter name
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Outliers")
   Do Until Obsvalues.EOF
     If PercentileArray(ArrayIndex, 0) <> Obsvalues!StationID Or PercentileArray(ArrayIndex, 1) <> Obsvalues!DISPLAY_NAME Then  ' New parameter?
       ArrayIndex = ArrayIndex + 1
       PercentileArray(ArrayIndex, 0) = Obsvalues!StationID  ' set new station
       PercentileArray(ArrayIndex, 1) = Obsvalues!DISPLAY_NAME  ' set new parameter name
     End If
     PercentileArray(ArrayIndex, 2) = PercentileArray(ArrayIndex, 2) + 1  ' Count the record
     Obsvalues.MoveNext
   Loop
   IndexSave = ArrayIndex  ' save last index
   ArrayIndex = 0
   Do Until ArrayIndex > IndexSave
     PercentileArray(ArrayIndex, 3) = (0.05 * PercentileArray(ArrayIndex, 2))  ' calculate 5th
     PercentileArray(ArrayIndex, 4) = (0.95 * PercentileArray(ArrayIndex, 2))  ' and 95th percentile counts
     ArrayIndex = ArrayIndex + 1
   Loop
   Obsvalues.MoveFirst  ' Back to start
   RecordCount = 0
   ArrayIndex = 0
   Save95 = -999
   Do Until Obsvalues.EOF
     If PercentileArray(ArrayIndex, 0) <> Obsvalues!StationID Or PercentileArray(ArrayIndex, 1) <> Obsvalues!DISPLAY_NAME Then  ' New parameter?
       ArrayIndex = ArrayIndex + 1  ' increment index
       FifthSave = -999
       RecordCount = 0
     End If
     RecordCount = RecordCount + 1
     If RecordCount > PercentileArray(ArrayIndex, 3) And PercentileArray(ArrayIndex, 5) = -999 Then  ' is this the first record past the fifth percentile count
       PercentileArray(ArrayIndex, 5) = Obsvalues!NumericValue   ' Set the 5th percentile value
     End If
     If RecordCount > PercentileArray(ArrayIndex, 4) And PercentileArray(ArrayIndex, 6) = -999 Then  ' is this the first record past the ninetyfifth percentile count
       PercentileArray(ArrayIndex, 6) = Save95   ' Set the 95th percentile value
     End If
     Save95 = Obsvalues!NumericValue  ' set save 95
     Obsvalues.MoveNext
   Loop
   Obsvalues.MoveFirst  ' Back to start
   ArrayIndex = 0
   
   Do Until Obsvalues.EOF
     If PercentileArray(ArrayIndex, 0) <> Obsvalues!StationID Or PercentileArray(ArrayIndex, 1) <> Obsvalues!DISPLAY_NAME Then  ' New parameter?
       ArrayIndex = ArrayIndex + 1
     End If
     If Obsvalues!NumericValue < PercentileArray(ArrayIndex, 5) Or Obsvalues!NumericValue > PercentileArray(ArrayIndex, 6) Then
       If (IsNull(Me!ProjectID) Or Me!ProjectID = Obsvalues!ProjectID) And (Obsvalues!START_DATE >= Me!StartDate And Obsvalues!START_DATE <= Me!EndDate) Then
         ' Write outlier record
         WorkOutput.AddNew
         WorkOutput!ProjectID = Obsvalues!ProjectID  '
         WorkOutput!StationID = Obsvalues!StationID  '
         WorkOutput!StationName = Obsvalues![Station Name]
         WorkOutput!START_DATE = Obsvalues!START_DATE  '
         WorkOutput!CharacteristicName = Obsvalues!DISPLAY_NAME  '
         WorkOutput!DetectionCondition = Obsvalues!Detection_Condition  '
         WorkOutput!ResultValue = Obsvalues!Result_Text  '
         WorkOutput!RemarkCode = Obsvalues!Lab_Remarks  '
         WorkOutput!ResultComment = Obsvalues!Result_Comment  '
         WorkOutput!VisitComment = Obsvalues!Visit_Comment  '
         WorkOutput!Cutoff_5 = PercentileArray(ArrayIndex, 5)
         WorkOutput!Cutoff_95 = PercentileArray(ArrayIndex, 6)
         WorkOutput!Sample_Size = PercentileArray(ArrayIndex, 2)
         WorkOutput.Update  ' Write outlier record
       End If
     End If
     Obsvalues.MoveNext
   Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
     Obsvalues.Close
     Set Obsvalues = Nothing
     
Exit_CanopyGap_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_outliers."
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
