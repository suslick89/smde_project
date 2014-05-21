Attribute VB_Name = "MGP_Macro_AKT_Peredachi"



Sub Act_peredachi_dizain()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи дизайн.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
         MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
         MGP_IN_Name_DATE = .Bookmarks.Item("a2").Range.Text
         MGP_IN_Name_Customer = .Bookmarks.Item("a3").Range.Text
         MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
         MGP_IN_Name_Company = .Bookmarks.Item("a5").Range.Text
         MGP_IN_Name_UR_Address = .Bookmarks.Item("a6").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
            .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Company2").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
            .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
            .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
            .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 г."
            End With
End Sub

Sub Act_peredachi_KD()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи КД.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
         MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
         MGP_IN_Name_DATE = .Bookmarks.Item("a2").Range.Text
         MGP_IN_Name_Customer = .Bookmarks.Item("a3").Range.Text
         MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
         MGP_IN_Name_Company = .Bookmarks.Item("a5").Range.Text
         MGP_IN_Name_UR_Address = .Bookmarks.Item("a6").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
            .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Company2").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
            .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
            .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
            .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 г."
            End With
End Sub
 
Sub Act_peredachi_konstrukcyi()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи конструкция.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
         MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
         MGP_IN_Name_DATE = .Bookmarks.Item("a2").Range.Text
         MGP_IN_Name_Customer = .Bookmarks.Item("a3").Range.Text
         MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
         MGP_IN_Name_Company = .Bookmarks.Item("a5").Range.Text
         MGP_IN_Name_UR_Address = .Bookmarks.Item("a6").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
            .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Company2").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
            .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
            .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
            .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 г."
            End With
End Sub
Sub Act_peredachi_OK_Varianta()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи Ок.вариант.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
         MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
         MGP_IN_Name_DATE = .Bookmarks.Item("a2").Range.Text
         MGP_IN_Name_Customer = .Bookmarks.Item("a3").Range.Text
         MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
         MGP_IN_Name_Company = .Bookmarks.Item("a5").Range.Text
         MGP_IN_Name_UR_Address = .Bookmarks.Item("a6").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
            .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Company2").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
            .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
            .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
            .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 г."
            End With
End Sub
Sub Act_peredachi_TZ()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи ТЗ.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
         MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
         MGP_IN_Name_DATE = .Bookmarks.Item("a2").Range.Text
         MGP_IN_Name_Customer = .Bookmarks.Item("a3").Range.Text
         MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
         MGP_IN_Name_Company = .Bookmarks.Item("a5").Range.Text
         MGP_IN_Name_UR_Address = .Bookmarks.Item("a6").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
            .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Company2").Range.Text = MGP_IN_Name_Company
            .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
            .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
            .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
            .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 г."
            End With
End Sub
