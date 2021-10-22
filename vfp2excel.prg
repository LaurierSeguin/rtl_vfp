*select 0
*USE j:\stad\usagers\séguin\resultats\ponc_pc_lig\requete0.dbf EXCLUSIVE
CLOSE data
reqpath = "j:\stad\usagers\séguin\resultats\ponc_pc_lig"
reqpath = "c:\temp"
rdbf = "yyy"
rrst = reqpath +"\requete0.rst"

oExcel = Createobject("Excel.Application")
With oExcel
  * .DisplayAlerts = .F.
  .Workbooks.Add
  .Visible = .T.
  With .ActiveWorkBook
    *tcCursorName = 'crsToExcel'
    toSheet = .WorkSheets(1)

	*      VFP2Excel(m.lcCursorName, .WorkSheets(1),"A1" )
	  *Local lcTemp, lcTempRs, loConn As AdoDB.Connection, loRS As AdoDB.Recordset
	 
	  loConn = Createobject("Adodb.connection")
	  loConn.ConnectionString = "Provider=VFPOLEDB;Data Source="+reqpath
	  loConn.Open()
	  *loRS = loConn.Execute("select * from " + rdbf)
	  *loRS.Save(rrst)
	  *loRS.Close
	  *loConn.Close
	  *Erase (m.lcTemp)
	  *loRS.Open(rrst)
	  
	  *loRS = Createobject("Adodb.recordset")
	  *loRS.Open("select * from " + rdbf, loconn)

      oExcel.ActiveSheet.QueryTables.Add( loconn, oExcel.ActiveSheet.Range("A1"), "select * from " + rdbf)
	  loRS.Close
	  Erase (m.lcTempRs)
      
    .WorkSheets(1).Activate
  Endwith
Endwith

