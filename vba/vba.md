### フォーム内ボタンのクリックイベント
```vb
Private Sub CommandButton1_Click()

    Macro1

End Sub
```
### シート上デザインボタン(ActiveX)のクリックイベント
```vb
Private Sub CommandButton1_Click()

    
    Dim win As New UserForm1
    UserForm1.Show
    

End Sub
```

### シード上フォームボタン => Module1 => Sub Test()
```vb
Sub Test()


    MsgBox ("OK")


End Sub
```
### マクロ登録した Module2 => Sub Macro()
```vb
Sub Macro1()
'
' Macro1 Macro
'

'
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G3").Select
    Selection.AutoFill Destination:=Range("G3:G20"), Type:=xlFillSeries
    Range("G3:G20").Select
End Sub
```


![image](https://user-images.githubusercontent.com/1501327/198270975-c4fe4d80-42c8-43e3-8530-dd622bc3627d.png)

![image](https://user-images.githubusercontent.com/1501327/198271106-fbbb8891-907b-4dca-90fe-88e4034ccb2d.png)
