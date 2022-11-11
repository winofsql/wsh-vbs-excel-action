
Cn = WScript.CreateObject("ADODB.Connection");

db_path = "C:\\app\\workspace\\îÃîÑä«óù.accdb";

connection_string = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db_path + ";";

WScript.Echo( connection_string );

Cn.Open( connection_string );

// --------------------------------------------
Rs = Cn.OpenSchema( 4, [,,"é–àıÉ}ÉXÉ^",] );
Rs.Sort = "ORDINAL_POSITION";
Rs.MoveFirst;

for (index = 0; index < Rs.Fields.Count; index++) {
    WScript.Echo( Rs.Fields(index).Name );
}

while( !Rs.EOF ) {

    WScript.Echo( Rs.Fields("ORDINAL_POSITION").Value );
    WScript.Echo( Rs.Fields("COLUMN_NAME").Value );
    WScript.Echo( Rs.Fields("DATA_TYPE").Value );
    WScript.Echo( Rs.Fields("CHARACTER_MAXIMUM_LENGTH").Value );
    Rs.MoveNext()

}

Rs.Close()
// --------------------------------------------

Cn.Close()
