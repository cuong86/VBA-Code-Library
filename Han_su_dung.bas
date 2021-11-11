Attribute VB_Name = "Han_su_dung"
Option Explicit

Sub hansudung()
Dim ngaysudung As Date ' khai bao dinh dang ngay thang
    ngaysudung = DateSerial(2021, 11, 9) ' Thoi han su dung muon dat
    If Now() >= ngaysudung Then ' Ham dieu kien : Neu ngay hom nay >= ngay den han

Retry:
' Hop thong bao thoi gian qua han
Select Case MsgBox(Buttons:=vbAbortRetryIgnore + vbCritical, Prompt:="Chuong trinh da qua han su dung " & ngaysudung, Title:="Chuong trinh da qua han su dung")
Case vbAbort
    GoTo HandleExit
Case vbRetry
    GoTo Retry
Case vbIgnore
    
End Select

    
ThisWorkbook.Close savechanges:=False ' Dong chuong trinh
    End If
    
HandleExit:

End Sub
'Author     : Cuong86
'Description: Dat thoi han su dung cho file
'Date       : 11Nov21

