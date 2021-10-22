* agregation.prg
*-----------------------------------------------------------------------------------------
parameters oxmlreq, oxmlclasse
private typfichier, typheures, typecarts, ficfinal
store '' to typfichier, typheures, typecarts
?
? "Exportation de la requête."

*-----

close databases
do getoptions with typfichier, typheures, typecarts

*-----

use tmp_final

if typheures = 'Texte (hh:mm:ss)'
  * ne rien faire
else
  alter table tmp_final add column tmp_heure c(8) null
  do fmt_heures with typheures
  alter table tmp_final drop column tmp_heure
endif

if typecarts = 'Temps (min) / Durées (s)'
  * ne rien faire
else
  alter table tmp_final add column tmp_ecart n(9,2) null
  do fmt_ecarts with typecarts
  alter table tmp_final drop column tmp_ecart
endif

use

*-----

public ExcelExtension
do case
  case typfichier = 'Visual FoxPro'
    ficfinal = "req_final.dbf"
    do tofoxpro
    browse nowait
    modify database nowait
  case typfichier = 'Microsoft Excel'
    if not toexcel3()  && 20111129
      g_bok = .f.
      return
    endif
    ficfinal = "req_final.xls"
  case typfichier = 'DBF'
    ficfinal = "req_final.dbf"
    do toDBF
endcase

*-----

sfichiers = "requete0.bak tmp_final.dbf tmp_final.bak"
declare afichiers[1]
x = split('afichiers', sfichiers, ' ')
for i = 1 to x
  if file(afichiers[i])
    delete file &afichiers[i]
  endif
next

*-----

if typfichier <> 'Microsoft Excel'
  select req_final
  browse nowait
else
  sh = CreateObject("WScript.Shell")
  curver = sh.RegRead("HKEY_CLASSES_ROOT\Excel.Sheet\CurVer\")
  excelpath = Sh.RegRead("HKEY_CLASSES_ROOT\" + curver + "\shell\Open\command\")
  cmd= excelpath + ' ' + dirreq + 'req_final.' + ExcelExtension
  run /n &cmd
endif

? "  input: tmp_final.dbf  output: " + ficfinal 
return

*-----------------------------------------------------------------------------------------

procedure getoptions
  parameters typfichier, typheures, typecarts

  onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='OutFormat']")
  typfichier = onode.text

  onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='HeureFormat']")
  typheures = onode.text

  onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='EcartFormat']")
  typecarts = onode.text
return

*-----------------------------------------------------------------------------------------

procedure toDBF
  sfichiers = "req_final.dbf,req_final.cdx"
  declare afichiers[1]
  x = split('afichiers', sfichiers, ',')
  for i = 1 to x
    if file(afichiers[i])
      delete file &afichiers[i]
    endif
  next
  
  use tmp_final
  copy to req_final type foxplus
  use req_final
return

*-----------------------------------------------------------------------------------------

procedure tofoxpro
  if file('bd_requete.dbc')
    open database bd_requete
  else
    create database bd_requete
  endif

  add table requete0
    
  use tmp_final
  copy to req_final
  add table req_final

  onode = oxmlreq.selectsinglenode("//DefChampSortie/Regroupement")
  if not isnull(onode)
    create sql view vue_final_0 as ;
    select * ;
      from req_final ;
     inner join requete0 ;
        on req_final.grp = requete0.grp
  endif
  
  select tmp_final
  use
  
  if select('req_final') = 0 
    select 0
    use req_final
  else
    select req_final
  endif
  
  do renamefields with 'foxpro', 'req_final'


return

*-----------------------------------------------------------------------------------------

function toexcel
  use tmp_final
  
  if reccount() > 65536
    sreccount = alltrim(str(reccount()))
    messagebox("Ne peut exporter cette requête dans Excel" + ;
               " (plus de 65536 enregistrements): " + sreccount)
    use
    return .f.
  endif
  
  copy to req_final type xl5
  use
  
  oxl = createobject("Excel.Application")  
  oxl.workbooks.open(dirreq + '\req_final.xls')
  do renamefields with 'excel', 'req_final'
  oxl.activeworkbook.save()
  oxl.quit()
return .t.

*-----------------------------------------------------------------------------------------

function toexcel2
  PRIVATE oxl AS Excel
  LOCAL oBook AS Excel.Workbook, oSheet AS OBJECT
  oxl = CREATEOBJECT("Excel.Application")
  oBook = oxl.Workbooks.Add
  oSheet = oBook.Worksheets(1)

  LOCAL oQryTable AS OBJECT, sCURDIR AS STRING
  sCURDIR = SYS(5) + curdir()
  oQryTable = oSheet.QueryTables.Add ;
      ( "OLEDB;Provider=VFPOLEDB.1;Data Source=" + sCURDIR, ;
        oSheet.RANGE("A1"), ;
        "select * from tmp_final" )
  oQryTable.RefreshStyle = 2   && xlInsertEntireRows = 2
  oQryTable.REFRESH(.F.)
  oxl.VISIBLE=.T.

  do renamefields with 'excel', 'req_final'   && 20120110

  LOCAL oSaveFile AS OBJECT
  oSaveFile = oBook.SaveAs(sCURDIR + 'req_final')
  oxl.quit
return .t.

*-----------------------------------------------------------------------------------------

function toexcel3
  PRIVATE oxl AS Excel
  PRIVATE xlFormat
  LOCAL oBook AS Excel.Workbook, oSheet AS OBJECT
  oxl = CREATEOBJECT("Excel.Application")
  if val(oxl.Version) <= 9 
    ExcelExtension = 'xls'
    xlFormat = 56
  else
    xlFormat = 51
    ExcelExtension = 'xlsx'
  endif
  oxl.DisplayAlerts = .F.

  *oBook = oxl.Workbooks.Add
  *oSheet = oBook.Worksheets(1)

  *LOCAL oQryTable AS OBJECT, sCURDIR AS STRING
  *sCURDIR = SYS(5) + curdir()
  *oQryTable = oSheet.QueryTables.Add ;
  *    ( "OLEDB;Provider=VFPOLEDB.1;Data Source=" + sCURDIR, ;
  *      oSheet.RANGE("A1"), ;
  *      "select * from tmp_final" )
  *oQryTable.RefreshStyle = 2   && xlInsertEntireRows = 2
  *oQryTable.REFRESH(.F.)

  sfichiers = "req_final.xls,req_final.xlsx,req_final.csv"
  declare afichiers[1]
  x = split('afichiers', sfichiers, ',')
  for i = 1 to x
    if file(afichiers[i])
      delete file &afichiers[i]
    endif
  next
  
  USE tmp_final
  COPY TO tmp_final csv  
  USE

  LOCAL sCURDIR AS STRING
  sCURDIR = SYS(5) + curdir()
  oBook = oxl.Workbooks.Open(sCURDIR + "tmp_final.csv")
  oxl.VISIBLE=.T.

  do renamefields with 'excel', 'req_final'   && 20120110

  LOCAL oSaveFile AS OBJECT
  if val(oxl.Version) <= 9 
    oSaveFile = oBook.SaveAs(sCURDIR + 'req_final.' + ExcelExtension)
  else
    oSaveFile = oBook.SaveAs(sCURDIR + 'req_final.' + ExcelExtension, xlFormat)
  endif
  oxl.quit
return .t.

*-----------------------------------------------------------------------------------------

procedure fmt_heures
  parameters typheures
  nbchamps = afields(achamps)
  for i = 1 to nbchamps
    schamp = upper(achamps(i,1))
    onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
            "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+schamp+"']")

    if isnull(onode) and substr(schamp, len(schamp) - 1) $ 'MIMXMOTONBETPC'
      * peut-être un champ à compiler
      xchamp = substr(schamp, 1, len(schamp) - 2)
      onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
              "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+xchamp+"']")
    endif

    if not isnull(onode)
      snomchamp = onode.getattributenode("Nom").value
      onode2 = g_oxmlparam.selectsinglenode("//ChampsUnites/Heures/" + snomchamp)
      if not isnull(onode2)      
        update tmp_final set tmp_heure = &schamp
        update tmp_final set &schamp = ''
        if typheures = 'Excel (nombres réels)'
          alter table tmp_final alter column &schamp b(5) null
          update tmp_final set &schamp = ( ;
              val(substr(tmp_heure, 1, 2)) * 3600 + ;
              val(substr(tmp_heure, 4, 2)) * 60 +   ;
              val(substr(tmp_heure, 7, 2))          ) / 86400       
        endif
        if typheures = 'Entier (hhmmss)'
          alter table tmp_final alter column &schamp n(6) null
          update tmp_final set &schamp = ;
              val(substr(tmp_heure, 1, 2)) * 10000 + ;
              val(substr(tmp_heure, 4, 2)) * 100 +   ;
              val(substr(tmp_heure, 7, 2))
        endif
      endif
    endif
  next
return

*-----------------------------------------------------------------------------------------

procedure fmt_ecarts
  parameters typecarts
  nbchamps = afields(achamps)
  for i = 1 to nbchamps
    schamp = upper(achamps(i,1))
    onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
            "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+schamp+"']")

    if isnull(onode) and substr(schamp, len(schamp) - 1) $ 'MIMXMOTONBETPC'
      * peut-être un champ à compiler
      xchamp = substr(schamp, 1, len(schamp) - 2)
      onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
              "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+xchamp+"']")
    endif

    if not isnull(onode)
      snomchamp = onode.getattributenode("Nom").value
      onode2 = g_oxmlparam.selectsinglenode( ;
               "//ChampsUnites/Minutes/" + snomchamp + '|' + ;
               "//ChampsUnites/Secondes/" + snomchamp)
      if not isnull(onode2)      
        update tmp_final set tmp_ecart = &schamp
        update tmp_final set &schamp = 0
        if typecarts = 'Heures décimales'
          alter table tmp_final alter column &schamp b(4) null
          if onode2.parentnode.nodename = 'Secondes'
            update tmp_final set &schamp = tmp_ecart / 3600
          else
            update tmp_final set &schamp = tmp_ecart / 60
          endif
        endif
        if typecarts = 'Minutes décimales'
          alter table tmp_final alter column &schamp b(2) null
          if onode2.parentnode.nodename = 'Secondes'
            update tmp_final set &schamp = tmp_ecart / 60
          else
            update tmp_final set &schamp = tmp_ecart
          endif
        endif
        if typecarts = 'Secondes'
          alter table tmp_final alter column &schamp n(6) null
          if onode2.parentnode.nodename = 'Minutes'
            update tmp_final set &schamp = tmp_ecart * 60
          else
            update tmp_final set &schamp = tmp_ecart
          endif
        endif
      endif
    endif
  next
return

*-----------------------------------------------------------------------------------------
