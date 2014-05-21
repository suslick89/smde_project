Attribute VB_Name = "MGP_Macro_AKT_END_Jobs"
Sub Act_Sdachi_stage_1()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ÿ¿¡ÀŒÕ€\ÕÓ‚˚Â ÿ¿¡ÀŒÕ€\¿ÍÚ˚\¿ÍÚ Ò‰‡˜Ë ‡·ÓÚ ˝Ú 1.dotx"
 Macro_Date
Set WordApp = GetObject("", "word.application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("word.application")
        End If
        WordApp.Visible = True
    
        With ActiveDocument.Range
                MGP_IN_Name_Dog = .Bookmarks.Item("a1").Range.Text
                MGP_IN_Date = .Bookmarks.Item("a2").Range.Text
                MGP_IN_Name_Company = .Bookmarks.Item("a3").Range.Text
                MGP_IN_Name_Product = .Bookmarks.Item("a4").Range.Text
                MGP_IN_Name_adress = .Bookmarks.Item("a5").Range.Text
                MGP_IN_Name_Zag_Dog = .Bookmarks.Item("a6").Range.Text
                MGP_IN_Name_DATE = .Bookmarks.Item("a7").Range.Text
                MGP_IN_1STAGE_cost = .Bookmarks.Item("a8").Range.Text
                MGP_IN_1STAGE_avans = .Bookmarks.Item("a9").Range.Text
                MGP_IN_1STAGE_avans2 = .Bookmarks.Item("a10").Range.Text
                MGP_IN_1STAGE_platej = .Bookmarks.Item("a11").Range.Text
                MGP_IN_1STAGE_platej2 = .Bookmarks.Item("a12").Range.Text
                MGP_IN_1STAGE_3_day = .Bookmarks.Item("a13").Range.Text
                MGP_IN_Name_Customer = .Bookmarks.Item("a14").Range.Text
                MGP_IN_Name_FIO = .Bookmarks.Item("a14").Range.Text
    End With
    Set Word_Doc = WordApp.Documents.Open(DOC)
        With Word_Doc
                .Bookmarks.Item("MGP_OUT_Name_Dog").Range.Text = MGP_IN_Name_Dog
                .Bookmarks.Item("MGP_OUT_Date").Range.Text = """" & Day(date) & """ " & Month_1 & " 2014 „."
                .Bookmarks.Item("MGP_OUT_Name_Company").Range.Text = MGP_IN_Name_Company
                .Bookmarks.Item("MGP_OUT_Name_Product").Range.Text = MGP_IN_Name_Product
                .Bookmarks.Item("MGP_OUT_Name_adress").Range.Text = MGP_IN_Name_adress
                .Bookmarks.Item("MGP_OUT_Name_Zag_Dog").Range.Text = MGP_IN_Name_Zag_Dog
                .Bookmarks.Item("MGP_OUT_Name_DATE").Range.Text = MGP_IN_Name_DATE
                .Bookmarks.Item("MGP_OUT_1STAGE_cost").Range.Text = MGP_IN_1STAGE_cost
                .Bookmarks.Item("MGP_OUT_1STAGE_avans").Range.Text = MGP_IN_1STAGE_avans
                .Bookmarks.Item("MGP_OUT_1STAGE_avans2").Range.Text = MGP_IN_1STAGE_avans2
                .Bookmarks.Item("MGP_OUT_1STAGE_platej").Range.Text = MGP_IN_1STAGE_platej
                .Bookmarks.Item("MGP_OUT_1STAGE_platej2").Range.Text = MGP_IN_1STAGE_platej2
                .Bookmarks.Item("MGP_OUT_1STAGE_3_day").Range.Text = MGP_IN_1STAGE_3_day
                .Bookmarks.Item("MGP_OUT_Name_customer").Range.Text = MGP_IN_Name_Customer
                .Bookmarks.Item("MGP_OUT_Name_FIO").Range.Text = MGP_IN_Name_FIO
            End With
End Sub
