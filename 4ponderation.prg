* pondedation.prg
*-----------------------------------------------------------------------------------------
parameters hconn, oxmlreq, sparam, sclasse, oxmlclasse
set procedure to LireOnglets2.prg additive

?
? "Pondération des données"
? "  input: requete0.dbf  output: requete0.dbf"

close databases all

open database bd_requete
nb = adbobjects(atables, "TABLE")
if nb > 0
  stables2 = "TOT_NCP,NCP,TOT_NCE,NCE,TOT_ECHANT,ECHANT"
  declare atables2[1]
  x = split('atables2', stables2, ',')
  for i = 1 to x
    if ascan(atables, atables2[i]) > 0
      remove table atables2[i] delete
    endif
  next
endif

*-----

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='assignation']")
fAssign = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='type_service']")
fTypser = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='voiture']")
fVoiture = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='ligne']")
fLigne = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='hre_pre_dep']")
fHrepredep = onode.getattributenode("OutDBF").value

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='sdap_date']")
fSdapdate = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='periode']")
fPeriode = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='direction']")
fDir = onode.getattributenode("OutDBF").value

if sClasse = 'Arrêts' or sClasse = 'Points de contrôle' or sClasse = 'Balises'
  onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
  fPosition = onode.getattributenode("OutDBF").value
endif

sreftable = ''
do case
  case sClasse = 'Courses'
    sreftable = 'ref_course'
  case sClasse = 'Arrêts'
    sreftable = 'ref_course_arret'
  case sClasse = 'Points de contrôle'
    sreftable = 'ref_course_pc'
  otherwise
    messagebox("Ne peut faire de pondération sur cette classe: " + sClasse)
    g_bOK = .f.
    return
endcase

*-----

* générer la clause where
private swhere
private stypser, sbookings, stypjrs, sHeures, sligtra, sGeopoints, sConditions, sChamps
swhere = ''

stypser = ''     && voir commentaire dans LireOnglets.prg
sbookings = ''
do OngletDates with sbookings, stypser
if not empty(sbookings)
  swhere = iif(empty(swhere), sbookings, swhere + ' AND ' + sbookings)
endif

stypjrs = ''
do OngletTypJrs with stypjrs, stypser
if not empty(stypjrs)
  swhere = iif(empty(swhere), stypjrs, swhere + ' AND ' + stypjrs)
endif

sHeures = ''
do OngletHeures with sheures
if not empty(sheures)
  swhere = iif(empty(swhere), sheures, swhere + ' AND ' + sheures)
endif

sligtra = ''
do OngletLigTra with sligtra
if not empty(sligtra)
  swhere = iif(empty(swhere), sligtra, swhere + ' AND ' + sligtra)
endif

* on ne traite PAS l'onglet 'GeoPoints' car les calculs des facteurs de
* pondération se font au niveau de la COURSE

* on ne traite PAS l'onglet 'Indices' pour trouver les courses manquantes

* Laurier; 20050116; on ne tient pas compte des conditions pour l'instant (temporairement)
* car la construction des clauses FROM et WHERE peut être diifficile
*sConditions = ''
*sChamps = '' 
*do OngletConditions with sconditions, schamps
*if not empty(sconditions)
*  swhere = iif(empty(swhere), sconditions, swhere + ' AND ' + sconditions)
*endif

swhere = strtran(swhere, "a.", "")

*-----

private shaving
private sniveaux, sniveauxas, sniveauxshort, sniveauxtot, sniveauxtotas

if empty(sbookings)
  messagebox("Vous devez sélectionner des dates.")
  g_bok = .f.
  return
endif

shaving = sbookings
if not empty(stypjrs)
  shaving = shaving + " AND " + stypjrs
endif
shaving = strtran(shaving, "a.", "")

*if not empty(stypjrs)
  sniveaux = "assignation, type_service, ligne"
  sniveauxtot = "assignation, type_service"
*else
*  sniveaux = "assignation, ligne"
*  sniveauxtot = "assignation"
*endif

if "Période" $ sparam
  sniveaux = sniveaux + ", periode"
endif
if "Direction" $ sparam
  sniveaux = sniveaux + ", direction"
endif

*-----

use requete0

declare aniveaux(1)
xn = split('aniveaux', sniveaux, ', ')
xf = afields(aflds)
bok = .t.
for i = 1 to xn
  sxpath = "//DefChampSortie/Champ[@Nom='" + aniveaux[i] + "']"
  onode = oxmlclasse.selectsinglenode(sxpath)
  flddbf = upper(onode.getattributenode("OutDBF").value)
  if ascan(aflds, flddbf) = 0
    messagebox("Le champ " + upper(aniveaux[i]) + " est manquant pour la " + ;
               "pondération. Veuillez l'ajouter dans les champs complémentaires.")
    bok = .f.
  endif
next
if not bok
  g_bok = .f.
  return
endif

*-----

sniveauxas = sniveaux
sniveauxas = strtran(sniveauxas, "assignation", "assignation as " + fAssign)
sniveauxas = strtran(sniveauxas, "type_service", "type_service as " + fTypser)
sniveauxas = strtran(sniveauxas, "ligne", "ligne as " + fLigne)
sniveauxas = strtran(sniveauxas, "periode", "periode as " + fPeriode)
sniveauxas = strtran(sniveauxas, "direction", "direction as " + fDir)

sniveauxshort = sniveaux
sniveauxshort = strtran(sniveauxshort, "assignation", fAssign)
sniveauxshort = strtran(sniveauxshort, "type_service", fTypser)
sniveauxshort = strtran(sniveauxshort, "ligne", fLigne)
sniveauxshort = strtran(sniveauxshort, "periode", fPeriode)
sniveauxshort = strtran(sniveauxshort, "direction", fDir)

sniveauxtotas = sniveauxtot
sniveauxtotas = strtran(sniveauxtotas, "assignation", "assignation as " + fAssign)
sniveauxtotas = strtran(sniveauxtotas, "type_service", "type_service as " + fTypser)

sniveauxtotshort = sniveauxtot
sniveauxtotshort = strtran(sniveauxtotshort, "assignation", fAssign)
sniveauxtotshort = strtran(sniveauxtotshort, "type_service", fTypser)

*-----

private g_seek_expr
g_seek_expr = ''
declare aniveaux(1)
xn = split('aniveaux', sniveaux, ', ')
xf = afields(aflds)
for i = 1 to xn
  sxpath = "//DefChampSortie/Champ[@Nom='" + aniveaux[i] + "']"
  onode = oxmlclasse.selectsinglenode(sxpath)
  flddbf = upper(onode.getattributenode("OutDBF").value)
  g_seek_expr = g_seek_expr + iif(i=1, '', ' + ')
  if lower(flddbf) = 'ligne'
    g_seek_expr = g_seek_expr + 'str(xxx.ligne,3)'
  else
    g_seek_expr = g_seek_expr + 'xxx.' + flddbf
  endif
next
*?
*? 'g_seek_expr: ', g_seek_expr

*-----

if ascan(aflds, 'PERIODE') > 0
  do calc_periode
endif

*-----

* compilation des courses prévues
tot_ncp = calc_totncp()
do calc_ncp with tot_ncp

* compilation des courses échantillonées
tot_nce = calc_totnce()
do calc_nce with tot_nce

* compilation du nombre d'échantillons et du facteur de pondération (fc)
do calc_totechant
do calc_echant

* calcul du facteur de pondération (fl)
do calc_fl

* report des facteurs de pondération dans le fichier de requete
do report_facteurs

return

*-----------------------------------------------------------------------------------------

function calc_totncp

ssql = "SELECT " + sniveauxtotas + ", count(*) as cnt" + ;
        " FROM ref_course" + ;
       " WHERE " +  swhere + ;
    " GROUP BY " + sniveauxtot + ;
      " HAVING " + shaving
*?
*? 'calc_totncp: ' + ssql

x = sqlexec(hconn, ssql)
select sqlresult
copy to tot_ncp
add table tot_ncp
use tot_ncp

return cnt

*---------------------------------------

procedure calc_ncp
parameters tot_ncp

ssql = "SELECT " + sniveauxas + ", count(*) as ncp" + ;
        " FROM ref_course" + ;
       " WHERE " + swhere + ;
    " GROUP BY " + sniveaux + ;
      " HAVING " + shaving + ;
    " ORDER BY " + sniveaux
*?
*? 'calc_ncp: ' + ssql

x = sqlexec(hconn, ssql)
select sqlresult
copy to ncp
add table ncp
use

alter table ncp add column pcp b(4)
alter table ncp add column plp b(4)

update ncp set pcp = 1 / ncp
update ncp set plp = ncp / tot_ncp
return

*-----------------------------------------------------------------------------------------

procedure calc_totnce

ssql = "SELECT " + sniveauxtotshort + ", count(*) as cnt" + ;
        " FROM requete0" + ;
        " INTO table tot_nce" + ;
       " WHERE not isnull(" + fSdapdate + ")" + ;
    " GROUP BY " + sniveauxtotshort
*?
*? 'calc_totnce: ' + ssql
&ssql
use

add table tot_nce
use tot_nce

return cnt

*---------------------------------------

procedure calc_nce
parameters tot_nce

ssql = "SELECT " + sniveauxshort + ", count(*) as nce" + ;
        " FROM requete0" + ;
        " INTO table nce" + ;
       " WHERE not isnull(" + fSdapdate + ")" + ;
    " GROUP BY " + sniveauxshort + ;
    " ORDER BY " + sniveauxshort
*?
*? 'calc_nce: ' + ssql
&ssql

use
add table nce
alter table nce add column ple b(4)
update nce set ple = nce / tot_nce

return

*-----------------------------------------------------------------------------------------

procedure calc_totechant

ssql = "SELECT " + sniveauxshort + ", count(*) as cnt" + ;
        " FROM requete0" + ;
  " INTO TABLE tot_echant" + ;
       " WHERE not isnull(" + fSdapdate + ")" + ;
    " GROUP BY " + sniveauxshort
*?
*? 'calc_totechant: ' + ssql
&ssql

use
add table tot_echant

return

*---------------------------------------

procedure calc_echant

sGoupBy = fAssign + ', ' + fTypser + ', ' + fLigne + ', ' + fVoiture + ', ' + fHrepredep
if sClasse = 'Arrêts' or sClasse = 'Points de contrôle' or sClasse = 'Balises'
  sGoupBy = sGoupBy + ', ' + fPosition
endif
if "Période" $ sparam
  sGoupBy = sGoupBy + ", " + fPeriode
endif
if "Direction" $ sparam
  sGoupBy = sGoupBy + ", " + fDir
endif


ssql = "SELECT " + sGoupBy + ", count(*) as cnt" + ;
        " FROM requete0" + ;
  " INTO TABLE echant" + ;
       " WHERE not isnull(" + fSdapdate + ")" + ;
    " GROUP BY " + sGoupBy
*?
*? 'calc_echant: ' + ssql
&ssql

use
add table echant
use echant

alter table echant add column pce b(4)
alter table echant add column fc n(6,3)

seek_expr = strtran(g_seek_expr, 'xxx.', '')
select 0
use tot_echant
index on &seek_expr tag idx_niv
reindex
set order to idx_niv
select ncp
index on &seek_expr tag idx_niv
reindex
set order to idx_niv

seek_echant = strtran(g_seek_expr, 'xxx.', 'echant.')

select echant
go top
do while not eof()
  select tot_echant
  seek &seek_echant
  select echant
  if not eof('tot_echant')
    replace pce with cnt / tot_echant.cnt
  else
    do handle_eof with 'calc_echant (tot_echant)'
  endif

  select ncp
  *seek &seek_expr
  seek &seek_echant
  select echant
  if not eof('ncp')
    if pce > 0 
      *? seek_echant
      *? &seek_echant
      *?assign, typ_ser, ligne, periode, dir, ncp.pcp, pce, ncp.pcp / pce
      replace fc with ncp.pcp / pce
      *if typ_ser = 'SE'
      *  suspend
      *endif
    else
      replace fc with 0
    endif
  else
    do handle_eof with 'calc_echant (ncp)'
  endif
  skip
enddo

return

*-----------------------------------------------------------------------------------------

procedure calc_fl

alter table nce add column fl n(6,3)

idx_expr = strtran(g_seek_expr, 'xxx.', '')
select ncp
index on &idx_expr tag idx_niv
reindex

seek_nce = strtran(g_seek_expr, 'xxx.', 'nce.')
select nce
go top
do while not eof()
  select ncp
  seek &seek_nce
  select nce
  if not eof('ncp')
    if ple > 0 
      replace fl with ncp.plp / ple
    else
      replace fl with 0
    endif
  else
    do handle_eof with 'calc_fl'
  endif
  skip
enddo

return

*-----------------------------------------------------------------------------------------

procedure report_facteurs

alter table requete0 add column fc n(6,3)
alter table requete0 add column fl n(6,3)
alter table requete0 add column fg n(6,3)

idx_expr = strtran(g_seek_expr, 'xxx.', '')
select nce
index on &idx_expr tag idx_niv
set order to idx_niv
reindex
seek_nce = strtran(g_seek_expr, 'xxx.', 'requete0.')
*?
*? 'report_facteurs(nce): ', seek_nce

sEchantPK = 'xxx.' + fAssign + ;
         ' + xxx.' + fTypser + ;
         ' + xxx.' + fVoiture + ;
         ' + str(xxx.' + fLigne + ',3)' + ;
         ' + xxx.' + fHrepredep
if sClasse = 'Arrêts' or sClasse = 'Points de contrôle' or sClasse = 'Balises'
  sEchantPK = sEchantPK + ' + str(xxx.' + fPosition + ",5)"
endif
idx_expr = strtran(sEchantPK, 'xxx.', '')
select echant
index on &idx_expr tag idx_niv
set order to idx_niv
reindex
seek_echant = strtran(sEchantPK, 'xxx.', 'requete0.')
*?
*? 'report_facteurs (echant): ', seek_echant

select requete0
set filter to not isnull(&fSdapdate)
go top
do while not eof()
  select echant
  seek &seek_echant
  select requete0
  if not eof('echant')
    replace fc with echant.fc
  else
    do handle_eof with 'report_facteurs (echant)'
  endif
  
  select nce
  seek &seek_nce
  select requete0
  if not eof('nce')
    replace fl with nce.fl
  else
    do handle_eof with 'report_facteurs (nce)'
  endif
  
  replace fg with fc * fl
  skip
enddo
set filter to

return

*-----------------------------------------------------------------------------------------
*-----------------------------------------------------------------------------------------

procedure handle_eof
parameters pmess

smess = "un problème de EOF() s'est produit dans la procédure: " + pmess
? smess
messagebox(smess)
suspend

return

*-----------------------------------------------------------------------------------------
