
Set Cn = WScript.CreateObject("ADODB.Connection")
Cn.CursorLocation = 3

db_path = "C:\app\workspace\販売管理.accdb"

connection_string = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db_path + ";"

Wscript.Echo connection_string

Cn.Open( connection_string )

' --------------------------------------------
Set Rs = Cn.OpenSchema( 4, Array(Empty,Empty,"社員マスタ",Empty) )
Rs.Sort = "ORDINAL_POSITION"
Rs.MoveFirst

For I = 0  to Rs.Fields.Count - 1
    Wscript.Echo Rs.Fields(I).Name
Next

Wscript.Echo "--------"

While NOT Rs.EOF

    Wscript.Echo Rs.Fields("ORDINAL_POSITION").Value
    Wscript.Echo Rs.Fields("COLUMN_NAME").Value
    Wscript.Echo Rs.Fields("DATA_TYPE").Value
    Wscript.Echo Rs.Fields("CHARACTER_MAXIMUM_LENGTH").Value
    Rs.MoveNext()

Wend

Rs.Close()
' --------------------------------------------
Wscript.Echo "--------"
' --------------------------------------------
Set Rs = Cn.OpenSchema( 12, Array(Empty,Empty,Empty,Empty,"社員マスタ") )
Rs.Sort = "ORDINAL_POSITION"
Rs.MoveFirst

For I = 0  to Rs.Fields.Count - 1
    Wscript.Echo Rs.Fields(I).Name
Next

Wscript.Echo "--------"

While NOT Rs.EOF
    if Rs.Fields("ORDINAL_POSITION").Value & "" <> "" then
        Wscript.Echo Rs.Fields("ORDINAL_POSITION").Value
        Wscript.Echo Rs.Fields("COLUMN_NAME").Value
    end if
    Rs.MoveNext()

Wend

Rs.Close()
' --------------------------------------------



Cn.Close()
