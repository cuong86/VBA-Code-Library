Attribute VB_Name = "congmaunen"
'########## Tinh tong theo mau nen ##########
Function tongmaunen(omau As Range, vungdulieu As Range) ' khai bao dang cong thuc
Dim tinhtong As Long ' khai bao ct tinh tong
Dim mamau As Integer ' khai bao mamau
mamau = omau.Interior.ColorIndex 'lay dinh dang o chua mau
For Each ocantong In vungdulieu 'vong lap qua tung phan tu
' ------------------Ham if ------------------
If ocantong.Interior.ColorIndex = mamau Then
tinhtong = WorksheetFunction.Sum(ocantong, tinhtong)
' ------------------ket thuc ham IF ------------------
End If
Next ocantong
tongmaunen = tinhtong
End Function


'########## Tinh tong theo mau chu ##########
Function tongmauchu(omau As Range, vungdulieu As Range) ' khai bao dang cong thuc
Dim mamau As Integer ' khai bao ct tinh tong
Dim tinhtong As Long ' khai bao mamau
mamau = omau.Font.ColorIndex 'lay dinh dang o chua mau
For Each ocantong In vungdulieu 'vong lap qua tung phan tu
' ------------------Ham if ------------------
If ocantong.Font.ColorIndex = mamau Then
tinhtong = WorksheetFunction.Sum(ocantong, tinhtong)
' ------------------ket thuc ham IF ------------------
End If
Next ocantong
tongmauchu = tinhtong
End Function
'Author     : Cuong86
'Description: Tinh tong theo mau nen hoac theo mau chu
'Date       : 11Nov21
    
