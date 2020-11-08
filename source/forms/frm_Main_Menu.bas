Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =9
    ItemSuffix =14
    Left =3900
    Top =1965
    Right =12480
    Bottom =7710
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x81b16fedaecae340
    End
    Caption ="Water Quality QA/QC Summaries"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
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
        Begin Section
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1425
                    Top =360
                    Width =5775
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label0"
                    Caption ="QA/QC Summaries"
                    LayoutCachedLeft =1425
                    LayoutCachedTop =360
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =780
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =2340
                    Width =2370
                    Height =300
                    Name ="ButtonInfestation"
                    Caption ="Duplicates"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5610
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4440
                    Top =4320
                    Width =2370
                    Height =300
                    TabIndex =1
                    Name ="ButtonInfestRoute"
                    Caption ="Cation-anion Balance"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =4320
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =4620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3840
                    Top =4920
                    Width =1035
                    Height =300
                    TabIndex =2
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =4920
                    LayoutCachedWidth =4875
                    LayoutCachedHeight =5220
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =3840
                    Width =2370
                    Height =300
                    TabIndex =3
                    Name ="ButtonInfestSize"
                    Caption ="Effectiveness"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =4140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =1320
                    Width =2370
                    Height =300
                    TabIndex =4
                    Name ="ButtonInfestGrowth"
                    Caption ="Outliers"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5610
                    LayoutCachedHeight =1620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =3360
                    Width =2370
                    Height =300
                    TabIndex =5
                    Name ="ButtonMonitoringTransect"
                    Caption ="Precision"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3360
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =3660
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =3360
                    Width =2370
                    Height =300
                    TabIndex =6
                    Name ="ButtonSpeciesCoover"
                    Caption ="Bias"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =3360
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =3660
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1455
                    Top =1860
                    Width =5760
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label9"
                    Caption ="Data Validation for Lab Data"
                    LayoutCachedLeft =1455
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =2160
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1425
                    Top =840
                    Width =5775
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="Data Validation for Field Data"
                    LayoutCachedLeft =1425
                    LayoutCachedTop =840
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1140
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =2880
                    Width =5775
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Reporting Field Data"
                    LayoutCachedLeft =1440
                    LayoutCachedTop =2880
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =3180
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =3840
                    Width =2370
                    Height =300
                    TabIndex =7
                    Name ="ButtonStage"
                    Caption ="Representativeness - Stage"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =3840
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =4140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =4320
                    Width =2370
                    Height =300
                    TabIndex =8
                    Name ="ButtonFlow"
                    Caption ="Representativeness - Flow"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =4320
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =4620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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


Private Sub ButtonFlow_Click()
On Error GoTo Err_ButtonFlow_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Flow"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonFlow_Click:
    Exit Sub

Err_ButtonFlow_Click:
    MsgBox Err.Description
    Resume Exit_ButtonFlow_Click
    
End Sub

Private Sub ButtonInfestation_Click()
On Error GoTo Err_ButtonInfestation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
     stDocName = "frm_Duplicates"
     DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestation_Click:
    Exit Sub

Err_ButtonInfestation_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestation_Click
    
End Sub
Private Sub ButtonInfestRoute_Click()
On Error GoTo Err_ButtonInfestRoute_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_CAB"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestRoute_Click:
    Exit Sub

Err_ButtonInfestRoute_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestRoute_Click
    
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
Private Sub ButtonInfestSize_Click()
On Error GoTo Err_ButtonInfestSize_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Effectiveness"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestSize_Click:
    Exit Sub

Err_ButtonInfestSize_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestSize_Click
    
End Sub
Private Sub ButtonInfestGrowth_Click()
On Error GoTo Err_ButtonInfestGrowth_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Outliers"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestGrowth_Click:
    Exit Sub

Err_ButtonInfestGrowth_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestGrowth_Click
    
End Sub
Private Sub ButtonMonitoringTransect_Click()
On Error GoTo Err_ButtonMonitoringTransect_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Field_Precision"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMonitoringTransect_Click:
    Exit Sub

Err_ButtonMonitoringTransect_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMonitoringTransect_Click
    
End Sub
Private Sub ButtonSpeciesCoover_Click()
On Error GoTo Err_ButtonSpeciesCoover_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Field_Bias"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonSpeciesCoover_Click:
    Exit Sub

Err_ButtonSpeciesCoover_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSpeciesCoover_Click
    
End Sub

Private Sub ButtonStage_Click()
On Error GoTo Err_ButtonStage_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Stage"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonStage_Click:
    Exit Sub

Err_ButtonStage_Click:
    MsgBox Err.Description
    Resume Exit_ButtonStage_Click
End Sub
