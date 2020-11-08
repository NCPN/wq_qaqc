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
    Left =5325
    Top =2115
    Right =12270
    Bottom =6165
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    Caption ="Field Data Effectiveness"
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
                    Left =3840
                    Top =2340
                    Width =2039
                    Height =300
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5879
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
                    Left =1185
                    Top =180
                    Width =4875
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Hydrograph Stage"
                    LayoutCachedLeft =1185
                    LayoutCachedTop =180
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1320
                    Top =2340
                    Width =2040
                    Height =300
                    TabIndex =2
                    Name ="ButtonCanopyGap"
                    Caption ="Stage Representativeness"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =2340
                    LayoutCachedWidth =3360
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
                    Width =1020
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="StartDate"
                    Format ="Short Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =1800
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =2115
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =1800
                    Width =1020
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EndDate"
                    Format ="Short Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5400
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
  Dim strStationName
  Dim dBeginDate As Date
  Dim dEndDate As Date
  Dim ArrayIndex As Integer
  Dim StageArray(4, 2) As Variant ' Array for stage counts
  ' Column x,0 is stage description
  ' Column x,1 is visit count

  ' Initialize array descriptions
   StageArray(0, 0) = "Base"
   StageArray(1, 0) = "Falling"
   StageArray(2, 0) = "Peak"
   StageArray(3, 0) = "Rising"
   ArrayIndex = 0
   Do Until ArrayIndex > 3           ' Initialize array values
     StageArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Representativeness_Stage"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Representativeness_Stage WHERE 1 = 1 "
  If Not IsNull(Me!ProjectID) Then
    strSQL = strSQL & "AND ProjectID = '" & Me!ProjectID & "'"
  End If
 ' strSQL = strSQL & "AND StationID = '5995420' "
  If Not IsNull(Me!StartDate) Then
    strSQL = strSQL & "AND START_DATE Between #" & Me!StartDate & "# AND #" & Me!EndDate & "#"
  End If
  strSQL = strSQL & " ORDER BY StationID, START_DATE"

  DoCmd.Hourglass True
  Set db = CurrentDb

  ' Count records by station, parameter
   Set Obsvalues = db.OpenRecordset(strSQL)
   If Obsvalues.EOF Then
     MsgBox "No valid Result records found.", vbOKOnly, "Stage Report"
     Obsvalues.Close
     Set Obsvalues = Nothing
     GoTo Exit_CanopyGap_Click
   End If
   ' Initialize a bunch of fields
   StationSave = Obsvalues!StationID
   strStationName = Obsvalues![Station Name]
   dBeginDate = Obsvalues!START_DATE
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Representativeness_Stage")
   Do Until Obsvalues.EOF
     If StationSave <> Obsvalues!StationID Then  ' New station?
       ArrayIndex = 0
       Do Until ArrayIndex > 3
         WorkOutput.AddNew
         WorkOutput!StationName = strStationName
         WorkOutput!StationID = StationSave
         If IsNull(Me!StartDate) Then
           WorkOutput!BeginDate = dBeginDate  '
           WorkOutput!EndDate = dEndDate  '
         Else
           WorkOutput!BeginDate = Me!StartDate  '
           WorkOutput!EndDate = Me!EndDate  '
         End If
         WorkOutput!Stage = StageArray(ArrayIndex, 0)
         WorkOutput!Visit_Count = StageArray(ArrayIndex, 1)
         WorkOutput.Update  ' Write record
         ArrayIndex = ArrayIndex + 1
       Loop
       StationSave = Obsvalues!StationID
       strStationName = Obsvalues![Station Name]
       dBeginDate = Obsvalues!START_DATE
       ArrayIndex = 0
       Do Until ArrayIndex > 3           ' Initialize array values
         StageArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
     End If
     dEndDate = Obsvalues!START_DATE  ' Set end of date range
     ArrayIndex = 0
     Do Until ArrayIndex > 3
       If Trim(Obsvalues!Stage) = StageArray(ArrayIndex, 0) Then
          StageArray(ArrayIndex, 1) = StageArray(ArrayIndex, 1) + 1   ' Count the visit
          GoTo NextRecord
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     MsgBox "Unknown stage description - " & Obsvalues!Stage
     GoTo Exit_CanopyGap_Click
NextRecord:
     Obsvalues.MoveNext
   Loop
   ' Write last record
       ArrayIndex = 0
       Do Until ArrayIndex > 3
         WorkOutput.AddNew
         WorkOutput!StationName = strStationName
         WorkOutput!StationID = StationSave
         If IsNull(Me!StartDate) Then
           WorkOutput!BeginDate = dBeginDate  '
           WorkOutput!EndDate = dEndDate  '
         Else
           WorkOutput!BeginDate = Me!StartDate  '
           WorkOutput!EndDate = Me!EndDate  '
         End If
         WorkOutput!Stage = StageArray(ArrayIndex, 0)
         WorkOutput!Visit_Count = StageArray(ArrayIndex, 1)
         WorkOutput.Update  ' Write record
         ArrayIndex = ArrayIndex + 1
       Loop
   WorkOutput.Close
   Set WorkOutput = Nothing
   Obsvalues.Close
   Set Obsvalues = Nothing
     
Exit_CanopyGap_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Representativeness_Stage."
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
