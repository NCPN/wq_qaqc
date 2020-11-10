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
    Caption ="Field Bias Report"
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
                    Left =1215
                    Top =180
                    Width =4815
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Field Data Bias"
                    LayoutCachedLeft =1215
                    LayoutCachedTop =180
                    LayoutCachedWidth =6030
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
                    Caption ="Bias  Report"
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
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =4140
                    Top =840
                    Width =1860
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"70\""
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
                            Visible = NotDefault
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
  Dim SDWork As DAO.Recordset
  Dim Obsvalues As DAO.Recordset
  Dim StationSave As String
  Dim strStationName
  Dim NameSave As String
  Dim RecordCount As Integer
  Dim ValueSum As Double
  Dim vStdDev As Variant
  Dim dBeginDate As Date
  Dim dEndDate As Date
  Dim MaxValue As Double
  Dim MinValue As Double
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Field_Bias"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Field_Bias WHERE 1 = 1 "
  ' If Not IsNull(Me!ProjectID) Then
    ' strSQL = strSQL & "AND ProjectID = '" & Me!ProjectID & "'"
  ' End If
  If Not IsNull(Me!StartDate) Then
    strSQL = strSQL & "AND START_DATE Between #" & Me!StartDate & "# AND #" & Me!EndDate & "#"
  End If
  strSQL = strSQL & " ORDER BY StationID, DISPLAY_NAME, START_DATE"
  ' MsgBox strSQL
  DoCmd.Hourglass True
  Set db = CurrentDb

  ' Count records by station, parameter
   Set Obsvalues = db.OpenRecordset(strSQL)
   If Obsvalues.EOF Then
     MsgBox "No valid Result records found.", vbOKOnly, "Bias Report"
     Obsvalues.Close
     Set Obsvalues = Nothing
     GoTo Exit_CanopyGap_Click
   End If
   ' Initialize a bunch of fields
   NameSave = Obsvalues!DISPLAY_NAME
   StationSave = Obsvalues!StationID
   strStationName = Obsvalues!StationName
   RecordCount = 0
   dBeginDate = Obsvalues!START_DATE
   MaxValue = Obsvalues!Bias
   MinValue = Obsvalues!Bias
   ValueSum = 0
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_Clear_StdDev"
   DoCmd.SetWarnings True
   Set SDWork = db.OpenRecordset("tbl_wrk_StdDev")
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Field_Bias")
   Do Until Obsvalues.EOF
     If NameSave <> Obsvalues!DISPLAY_NAME Or StationSave <> Obsvalues!StationID Then  ' New parameter?
       SDWork.Close
       Set SDWork = Nothing
       WorkOutput.AddNew
       WorkOutput!StationName = strStationName
       WorkOutput!CharacteristicName = NameSave  '
       If IsNull(Me!StartDate) Then
         WorkOutput!BeginDate = dBeginDate  '
         WorkOutput!EndDate = dEndDate  '
       Else
         WorkOutput!BeginDate = Me!StartDate  '
         WorkOutput!EndDate = Me!EndDate  '
       End If
       WorkOutput!Average = Round(ValueSum / RecordCount, 3)
       vStdDev = DStDev("[Precision]", "tbl_wrk_StdDev")  ' get standard deviation
       If Not IsNull(vStdDev) Then
         WorkOutput!SDeviation = Round(vStdDev, 3)
       End If
       WorkOutput!Minimum = Round(MinValue, 3)  '
       WorkOutput!Maximum = Round(MaxValue, 3)  '
       WorkOutput!RecCount = RecordCount
       WorkOutput.Update  ' Write record
       StationSave = Obsvalues!StationID
       strStationName = Obsvalues!StationName
       NameSave = Obsvalues!DISPLAY_NAME
       RecordCount = 0
       dBeginDate = Obsvalues!START_DATE
       ValueSum = 0
       MaxValue = Obsvalues!Bias
       MinValue = Obsvalues!Bias
       DoCmd.SetWarnings False
       DoCmd.OpenQuery "qry_Clear_StdDev"
       DoCmd.SetWarnings True
       Set SDWork = db.OpenRecordset("tbl_wrk_StdDev")
     End If
     SDWork.AddNew
     SDWork!Precision = Obsvalues!Bias
     SDWork.Update
     RecordCount = RecordCount + 1    ' Count the record
     dEndDate = Obsvalues!START_DATE  ' Set end of date range
     If Obsvalues!Bias < MinValue Then
       MinValue = Obsvalues!Bias
     ElseIf Obsvalues!Bias > MaxValue Then
       MaxValue = Obsvalues!Bias
     End If
     ValueSum = ValueSum + Obsvalues!Bias
     Obsvalues.MoveNext
   Loop
   ' Write last record
   SDWork.Close
   Set SDWork = Nothing
   WorkOutput.AddNew
     WorkOutput!StationName = strStationName
     WorkOutput!CharacteristicName = NameSave  '
     If IsNull(Me!StartDate) Then
       WorkOutput!BeginDate = dBeginDate  '
       WorkOutput!EndDate = dEndDate  '
     Else
       WorkOutput!BeginDate = Me!StartDate  '
       WorkOutput!EndDate = Me!EndDate  '
     End If
     WorkOutput!Average = Round(ValueSum / RecordCount, 3)
     vStdDev = DStDev("[Precision]", "tbl_wrk_StdDev")  ' get standard deviation
     If Not IsNull(vStdDev) Then
       WorkOutput!SDeviation = Round(vStdDev, 3)
     End If
     WorkOutput!Minimum = Round(MinValue, 3)  '
     WorkOutput!Maximum = Round(MaxValue, 3)  '
     WorkOutput!RecCount = RecordCount
   WorkOutput.Update  ' Write record
   WorkOutput.Close
   Set WorkOutput = Nothing
   Obsvalues.Close
   Set Obsvalues = Nothing
     
Exit_CanopyGap_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Field_Bias."
    Exit Sub

Err_CanopyGap_Click:
    MsgBox Err.Description
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
