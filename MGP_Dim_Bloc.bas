Attribute VB_Name = "MGP_Dim_Bloc"
Public I_Counter As Integer
Public Month_1 As Variant
Sub Macro_Date()
Select Case Month(date)
  Case "1": Month_1 = "Январь"
  Case "2": Month_1 = "Февраля"
  Case "3": Month_1 = "Марта"
  Case "4": Month_1 = "Апреля"
  Case "5": Month_1 = "Мая"
  Case "6": Month_1 = "Июня"
  Case "7": Month_1 = "Июля"
  Case "8": Month_1 = "Августа"
  Case "8": Month_1 = "Сентября"
  Case "8": Month_1 = "Октября"
  Case "8": Month_1 = "Ноября"
  Case "8": Month_1 = "Декабря"
  End Select
End Sub
Sub add_new_Bloc_Bookmarks()

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


Sub Del_All_Bloc_Bookmarks()
      Dim stBookmark As Bookmark
      ActiveDocument.Bookmarks.ShowHidden = True
      If ActiveDocument.Bookmarks.count >= 1 Then
         For Each stBookmark In ActiveDocument.Bookmarks
            stBookmark.Delete
         Next stBookmark
      End If
   End Sub
Sub Bloc_Scroll_Text_Word_Doc()
Application.ScreenUpdating = False
Dim Success As Boolean
Do
  Selection.GoTo What:=wdGoToBookmark, Name:="a" & I_Counter
Loop While Success = True
Application.ScreenUpdating = True
End Sub


Sub qty_Name_Bloc_Bookmarks()
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
