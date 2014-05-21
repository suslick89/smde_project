Attribute VB_Name = "MGP_Macro"
Public I_Counter As Integer
Public Month_1 As Variant

Sub add_new_Bookmarks()

Application.ScreenUpdating = False
     Selection.HomeKey Unit:=wdStory
        Dim I_Counter As Long
        Dim BK_counter As Long
I_Counter = 1
BK_counter = 1
     With ActiveDocument.Content.Find
            .Highlight = True
            .Forward = True
            .Wrap = wdFindStop
Do While (.Execute = True)
     BK_counter = BK_counter + 1
  Loop
End With
    Do
        With Selection.Find
            .Highlight = True
            .Forward = True
            .Wrap = wdFindContinue
            .Execute
             End With
      
        With ActiveDocument.Bookmarks
            .Add Range:=Selection.Range, Name:="a" & I_Counter
        End With
       
       
        I_Counter = I_Counter + 1
Loop While I_Counter < BK_counter
Application.ScreenUpdating = True
End Sub

Sub Del_All_Bookmarks()
      Dim stBookmark As Bookmark
      ActiveDocument.Bookmarks.ShowHidden = True
      If ActiveDocument.Bookmarks.count >= 1 Then
         For Each stBookmark In ActiveDocument.Bookmarks
            stBookmark.Delete
         Next stBookmark
      End If
   End Sub
Sub Scroll_Text_Word_Doc()
Application.ScreenUpdating = False
Dim Success As Boolean
Do
  Selection.GoTo What:=wdGoToBookmark, Name:="a" & I_Counter
Loop While Success = True
Application.ScreenUpdating = True
End Sub


Sub qty_Name_Bookmarks()
Dim BK_counter As Long
BK_counter = 0
Selection.HomeKey Unit:=wdStory
With ActiveDocument.Content.Find
            .Highlight = True
Do While (.Execute = True)
     BK_counter = BK_counter + 1
  Loop
End With
MsgBox ("Найдено " & BK_counter & " количество закладок")
End Sub
Sub Act_peredachi()
Dim MGP_IN_Name_Bookmark As Variant
Dim I_Counter As Integer
DOC = "W:\Templates-ШАБЛОНЫ\Новые ШАБЛОНЫ\Акты\Акт передачи дизайн 13.dotx"
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


 

