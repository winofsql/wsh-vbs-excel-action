
Set Cn = WScript.CreateObject("ADODB.Connection")
Set Rs = WScript.CreateObject("ADODB.Recordset")

db_path = "C:\app\workspace\îÃîÑä«óù.accdb"

connection_string = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db_path + ";"

Wscript.Echo connection_string

Cn.Open( connection_string )

query = "select * from é–àıÉ}ÉXÉ^"

Call Rs.Open( query, Cn )

While NOT Rs.EOF

    Wscript.Echo Rs.Fields("éÅñº").Value

    Rs.MoveNext()

Wend

Cn.Close()
