VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} wcDataInput 
   ClientHeight    =   10020
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   11940
   _ExtentX        =   21061
   _ExtentY        =   17674
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   8
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   2
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tmpJournalInput"
         DISPID          =   1280
         Template        =   "JournalInput1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{50CCDCA7-36AA-47A2-AF42-57E504E300EF}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "D:\Data Input\Web Input\JournalInput.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tmpTreatiseInput"
         DISPID          =   1281
         Template        =   "TreatiseInput1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{DCBF8AC1-2FB4-458B-AD57-922E41835C6D}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "D:\Data Input\Web Input\TreatiseInput.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "DataInput"
End
Attribute VB_Name = "wcDataInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Sub tmpJournalInput_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    If TagName = "WC@link" Then
        TagContents = "<a href=" & URLFor(tmpTreatiseInput) & ">Treatise</a>"
    End If
End Sub

Private Sub tmpTreatiseInput_Respond()
    'Response.Write "hi"
    tmpTreatiseInput.WriteTemplate
End Sub

Private Sub WebClass_Start()
    'tmpTreatiseInput.WriteTemplate
    tmpJournalInput.WriteTemplate
    'Write a reply to the user
    'With Response
    '    .Write "<html>"
    '    .Write "<body>"
    '    .Write "<h1><font face=""Arial"">WebClass1's Starting Page</font></h1>"
    '    .Write "<p>This response was created in the Start event of WebClass1.</p>"
    '    .Write "</body>"
    '    .Write "</html>"
    'End With

End Sub
