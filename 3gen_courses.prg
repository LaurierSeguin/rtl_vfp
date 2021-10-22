* gen_courses.prg
*-----------------------------------------------------------------------------------------
parameters hconn, oxmlreq, oxmlclasse
set procedure to LireOnglets2.prg additive

?
? "Ajout des courses prévues manquantes"
? "  input: requete0.dbf  output: requete0.dbf"

onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='Classe']")
sclasse = onode.text

*-----

* champs communs aux différents types de requête
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

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='direction']")
fDirection = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='trace']")
fTrace = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='periode']")
fPeriode = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='sdap_date']")
fSdapdate = onode.getattributenode("OutDBF").value

*-----

close databases all
select 0
use requete0

* modifier sdap_date not null
x = afields(aflds)
*if ascan(aflds, upper(fSdapdate)) > 0
*  alter table requete0 alter column &fSdapdate null
*endif
for i = 1 to x
  if not aflds[i,5]
    champ = aflds[i,1]  
    alter table requete0 alter column &champ null
  endif
next

*-----

* générer la clause where
private swhere
private stypser, sbookings, stypjrs, sHeures, sligtra
private sGeopoints, sGeotroncons, sConditions, sChamps
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

sGeopoints = ''
do OngletGeoPoints with sgeopoints, sclasse
if not empty(sgeopoints)
  swhere = iif(empty(swhere), sgeopoints, swhere + ' AND ' + sgeopoints)
endif

* on ne traite PAS l'onglet 'Indices' pour trouver les courses manquantes

sConditions = ''
sChamps = '' 
do OngletConditions with sconditions, schamps
if not empty(sconditions)
  swhere = iif(empty(swhere), sconditions, swhere + ' AND ' + sconditions)
endif

select requete0
maxrequete0 = reccount()

* générer les enregistrements manquants
if sclasse = 'Courses'
  strsql = makefrom('ref_course', schamps, swhere)
  do genCourses with strsql
endif

if sclasse = 'Arrêts'
  schamps = iif(empty(schamps), '', ' ' + schamps)
  schamps = 'ref_course.direction ref_course.trace' + schamps
  strsql = makefrom('ref_course_arret', schamps, swhere)
  do genCoursesArrets with strsql
endif
if sclasse = 'Points de contrôle'
  schamps = iif(empty(schamps), '', ' ' + schamps)
  schamps = 'ref_course.direction ref_course.trace' + schamps
  strsql = makefrom('ref_course_pc', schamps, swhere)
  do genCoursesPCs with strsql
endif

if sclasse = 'Balises'
  schamps = iif(empty(schamps), '', ' ' + schamps)
  schamps = 'ref_course.direction ref_course.trace' + schamps
  strsql = makefrom('v_ref_course_balise', schamps, swhere)
  do genCoursesBalises with strsql
endif

sTableTroncon = ''
if sclasse = 'Tronçons'
  onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='Auteur']")
  sTableTroncon = 'ref_troncon_' +onode.text

  schamps = iif(empty(schamps), '', ' ' + schamps)
  schamps = 'ref_course.hre_pre_arr ref_course.trace' + schamps
  strsql = makefrom(sTableTroncon, schamps, swhere)
  do genCoursesTroncons with strsql
endif

*-----

do fillRefFields

*-----

select requete0
if ascan(aflds, 'PERIODE') > 0
  do calc_periode
endif

*-----
select requete0
go top
*brow

return

*-----------------------------------------------------------------------------------------
*-----------------------------------------------------------------------------------------

procedure genCourses
parameters strsql
? "  requête 'Courses'"

select requete0
index on &fAssign+&fTypser+str(&fLigne,3)+&fVoiture+&fHrepredep ;
      tag refprimary
set order to refprimary
x = sqlexec(hconn, strsql, 'rs')

select rs
do while not eof()
  select requete0
  seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.voiture + rs.hre_pre_dep
  if eof()

*    append blank
*    replace &fAssign    with rs.assignation, ; 
*            &fTypser    with rs.type_service, ;
*            &fLigne     with rs.ligne, ;
*            &fVoiture   with rs.voiture, ;
*            &fHrepredep with rs.hre_pre_dep, ;
*            &fDirection with rs.direction, ;
*            &fTrace     with rs.trace
    insert into requete0 ;
        (&fAssign, &fTypser, &fLigne, &fVoiture, &fHrepredep, &fDirection, &fTrace) ;
      values ;
        (rs.assignation, rs.type_service, rs.ligne, rs.voiture, rs.hre_pre_dep, ;
         rs.direction, rs.trace)
  endif
  select rs
  skip
enddo
use

return

*-----------------------------------------------------------------------------------------

procedure genCoursesArrets
parameters strsql
? "  requête 'Arrêts'"

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='hre_pre_arret']")
fHreprearret = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
fPosition = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='no_arret']")
fNoarret = onode.getattributenode("OutDBF").value

select requete0
index on &fAssign+&fTypser+str(&fLigne,3)+&fVoiture+&fHrepredep+&fHreprearret ;
      tag refprimary
set order to refprimary
x = sqlexec(hconn, strsql, 'rs')

select rs
do while not eof()
  select requete0
  seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.voiture + ;
       rs.hre_pre_dep + rs.hre_pre_arret
  if eof()
*    append blank
*    replace &fAssign      with rs.assignation, ; 
*            &fTypser      with rs.type_service , ;
*            &fLigne       with rs.ligne, ;
*            &fVoiture     with rs.voiture, ;
*            &fHrepredep   with rs.hre_pre_dep, ;
*            &fHreprearret with rs.hre_pre_arret, ;
*            &fPosition    with rs.position, ;
*            &fDirection   with rs.direction, ;
*            &fTrace       with rs.trace, ;
*            &fNoarret     with rs.no_arret
    insert into requete0 ;
        (&fAssign, &fTypser, &fLigne, &fVoiture, &fHrepredep, &fHreprearret, ;
         &fPosition, &fDirection, &fTrace, &fNoarret ) ;
      values ;
        (rs.assignation, rs.type_service, rs.ligne, rs.voiture, rs.hre_pre_dep, ;
         rs.hre_pre_arret, rs.position, rs.direction, rs.trace, rs.no_arret)
  endif
  select rs
  skip
enddo
use

return

*-----------------------------------------------------------------------------------------

procedure genCoursesPCs
parameters strsql
? "  requête 'Points de contrôle'"

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='hre_pre_pc']")
fHreprepc = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
fPosition = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='lieu']")
fLieu = onode.getattributenode("OutDBF").value

select requete0
index on &fAssign+&fTypser+str(&fLigne,3)+&fVoiture+&fHrepredep+&fHreprepc ;
      tag refprimary
set order to refprimary
x = sqlexec(hconn, strsql, 'rs')

select rs
do while not eof()
  select requete0
  seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.voiture + ;
       rs.hre_pre_dep + rs.hre_pre_pc
  if eof()
*    append blank
*    replace &fAssign      with rs.assignation, ; 
*            &fTypser      with rs.type_service , ;
*            &fLigne       with rs.ligne, ;
*            &fVoiture     with rs.voiture, ;
*            &fHrepredep   with rs.hre_pre_dep, ;
*            &fHreprepc    with rs.hre_pre_pc, ;
*            &fPosition    with rs.position, ;
*            &fDirection   with rs.direction, ;
*            &fTrace       with rs.trace, ;
*            &fLieu        with rs.lieu
    insert into requete0 ;
        (&fAssign, &fTypser, &fLigne, &fVoiture, &fHrepredep, &fHreprepc, ;
         &fPosition, &fDirection, &fTrace, &fLieu) ;
      values ;
        (rs.assignation, rs.type_service, rs.ligne, rs.voiture, rs.hre_pre_dep, ;
         rs.hre_pre_pc, rs.position, rs.direction, rs.trace, upper(rs.lieu))
  endif
  select rs
  skip
enddo
use

return

*-----------------------------------------------------------------------------------------

procedure genCoursesBalises
parameters strsql
? "  requête 'Balises'"

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
fPosition = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='no_balise']")
fBalise = onode.getattributenode("OutDBF").value

select requete0
index on &fAssign+&fTypser+str(&fLigne,3)+&fVoiture+&fHrepredep+str(&fPosition,5) ;
      tag refprimary
set order to refprimary
x = sqlexec(hconn, strsql, 'rs')

select rs
do while not eof()
  select requete0
  seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.voiture + ;
       rs.hre_pre_dep + str(rs.position,5)
  if eof()
*    append blank
*    replace &fAssign      with rs.assignation, ; 
*            &fTypser      with rs.type_service , ;
*            &fLigne       with rs.ligne, ;
*            &fVoiture     with rs.voiture, ;
*            &fHrepredep   with rs.hre_pre_dep, ;
*            &fPosition    with rs.position, ;
*            &fDirection   with rs.direction, ;
*            &fTrace       with rs.trace, ;
*            &fBalise      with rs.no_balise
    insert into requete0 ;
        (&fAssign, &fTypser, &fLigne, &fVoiture, &fHrepredep, ;
         &fPosition, &fDirection, &fTrace, &fBalise) ;
      values ;
        (rs.assignation, rs.type_service, rs.ligne, rs.voiture, rs.hre_pre_dep, ;
         rs.position, rs.direction, rs.trace, rs.no_balise)
  endif
  select rs
  skip
enddo
use

return

*-----------------------------------------------------------------------------------------

procedure genCoursesTroncons
parameters strsql
? "  requête 'Tronçons'"

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='no_troncon']")
fTroncon = onode.getattributenode("OutDBF").value

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='pos_deb']")
fPosdeb = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='no_point_deb']")
fEvnmdeb = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='pos_fin']")
fPosfin = onode.getattributenode("OutDBF").value
onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='no_point_fin']")
fEvnmfin = onode.getattributenode("OutDBF").value

select requete0
index on &fAssign + &fTypser + str(&fLigne, 3) + &fVoiture + &fHrepredep + str(&fTroncon, 3) ;
      tag refprimary
set order to refprimary
x = sqlexec(hconn, strsql, 'rs')

select rs
do while not eof()
  select requete0
  seek rs.assignation + rs.type_service + str(rs.ligne, 3) + rs.voiture + ;
       rs.hre_pre_dep + str(rs.no_troncon, 3) 
  if eof()
    insert into requete0 ;
        (&fAssign,  &fTypser,    &fLigne,  &fVoiture, &fHrepredep, &fTroncon, ;
         &fPosdeb,  &fEvnmdeb,   &fPosfin, &fEvnmfin, ;
         &fPeriode, &fDirection, &fTrace) ;
      values ;
        (rs.assignation, rs.type_service, rs.ligne, rs.voiture, rs.hre_pre_dep, rs.no_troncon, ;
         rs.pos_deb,     rs.no_point_deb, rs.pos_fin, rs.no_point_fin, ;
         rs.periode,     rs.direction,    rs.trace)
  endif
  select rs
  skip
enddo
use

return

*-----------------------------------------------------------------------------------------
*-----------------------------------------------------------------------------------------

function makefrom
parameters smaintable, schamps, swhere

if empty(schamps)
  sret = 'SELECT a.* FROM ' + smaintable + ' AS a'
  
ELSE
  if at('ref_noeudtc', schamps) > 0  
     * 20100415; conditions de ref_noeudtc dans l'onglet 'Conditions logiques'
    if sclasse = 'Arrêts'
      schamps = schamps + ' ref_arret.no_noeudtc'
    else && 'Points de contrôle'
      schamps = schamps + ' ref_pc.no_noeudtc'
    endif
  endif

  * trouver les tables à joindre à partir des champs
  declare achamps[1]
  x = split('achamps', schamps, ' ')
  stables = ''
  for i = 1 to x
    tmptable = "'" + substr(achamps[i], 1, at('.', achamps[i]) - 1) + "'"
    if at(tmptable, stables) = 0
      stables = iif(empty(stables), tmptable, stables + ' ' + tmptable)
    endif
  next
  declare atables[1]
  stables = strtran(stables, "'", "")
  nbtables = split('atables', stables, ' ')
  i = ascan(atables, smaintable)
  if i > 0
    * ne pas inclure la table principale dans les tables à joindre
    x = adel(atables, i)
    nbtables = nbtables - 1
  endif
  if nbtables = 0 
    ? "Une erreur s'est produite: gen_courses.prg function makefrom; nbtables = 0"
    suspend
  endif

  oxmlref = createobject("Msxml2.DOMDocument.4.0")
  oxmlref.load(getenv('STAD_HOME') + '\source_prg\reference.xml')

  * generer la clause from
  sfrom = ''
  for i = 1 to nbtables
    sfrom = sfrom + " INNER JOIN " + atables[i] + " ON "
    onode = oxmlref.selectsinglenode("//" + atables[i])
    onodes = onode.selectnodes("./ancestor-or-self::*")
    for j = 0 to onodes.length - 1
      declare pks[1]
      x = split('pks', onodes.item(j).getattributenode('pk').value, ' ')
      for k = 1 to x
        if not (j = 0 and k = 1)
          sfrom = sfrom + ' AND '
        endif
        if 'ref_course_' $ smaintable and (pks[k] = 'direction' or pks[k] = 'trace')
          * pour les liens entre les tables ref_course_* et ref_trace_*, les champs
          * ref_course.direction, ref_course.trace DOIVENT être les PREMIERS éléments
          * du paramètre schamps
          alias = 'ref_course'
        else
          IF 'ref_noeudtc' $ atables[i]
            alias = atables[i]
          else
            alias = 'a'  && smaintable
          endif
        endif
        sfrom = sfrom + alias + '.' + pks[k] + ' = ' + atables[i] + '.' + pks[k]
      next
    next
  next
  
  sret = 'SELECT a.*, ' + strtran(schamps, ' ', ', ') + ;
         ' FROM ' + smaintable + ' AS a' + sfrom
endif

* generer le SQL complet
if not empty(swhere)
  sret = sret + ' WHERE ' + swhere
endif
sret = strtran(sret, smaintable + '.', 'a.')
do savesql with sret, 'gen_courses'
*? sret

return sret

*-----------------------------------------------------------------------------------------

procedure fillRefFields
* 2006/10/17
* remplir les champs prévus pour les courses prévues non-échantillonnées
 

inbtables = 18
if not empty(sTableTroncon)
  inbtables = 19
endif

declare aRefTables[inbtables]
declare a_SQL[inbtables]
aRefTables[1] = "ref_arret"
a_SQL[1] = "select * from ref_arret where assignation = '%1' and no_arret = %2"
aRefTables[2] = "ref_assignation"
a_SQL[2] = "select * from ref_assignation where assignation = '%1'"
aRefTables[3] = "ref_balise"
a_SQL[3] = "select * from ref_balise where assignation = '%1' and no_balise = %2"
aRefTables[4] = "ref_calendrier"
a_SQL[4] = "select * from ref_calendrier where assignation = '%1' and type_service = %2"
aRefTables[5] = "ref_course"
a_SQL[5] = "select * from ref_course       where assignation = '%1' and type_service = '%2' and ligne = %3 and voiture = '%4' and hre_pre_dep = '%5'"
aRefTables[6] = "ref_course_arret"
a_SQL[6] = "select * from ref_course_arret where assignation = '%1' and type_service = '%2' and ligne = %3 and voiture = '%4' and hre_pre_dep = '%5' and hre_pre_arret = '%6'"
aRefTables[7] = "ref_course_pc"
a_SQL[7] = "select * from ref_course_pc    where assignation = '%1' and type_service = '%2' and ligne = %3 and voiture = '%4' and hre_pre_dep = '%5' and hre_pre_pc = '%6'"
aRefTables[8] = "ref_course_vres"
a_SQL[8] = "select * from ref_course_vres  where assignation = '%1' and type_service = '%2' and ligne = %3 and voiture = '%4' and hre_pre_dep = '%5' and no_vres = %6"
aRefTables[9] = "ref_ligne"
a_SQL[9] = "select * from ref_ligne where assignation = '%1' and ligne = %2"
aRefTables[10] = "ref_pc"
a_SQL[10] = "select * from ref_pc where assignation = '%1' and lieu = '%2'"
aRefTables[11] = "ref_trace"
a_SQL[11] = "select * from ref_trace        where assignation = '%1' and ligne = %2 and direction = '%3' and trace = %4"
aRefTables[12] = "ref_trace_arret"
a_SQL[12] = "select * from ref_trace_arret  where assignation = '%1' and ligne = %2 and direction = '%3' and trace = %4 and position = %5"
aRefTables[13] = "ref_trace_balise"
a_SQL[13] = "select * from ref_trace_balise where assignation = '%1' and ligne = %2 and direction = '%3' and trace = %4 and position = %5"
aRefTables[14] = "ref_trace_pc"
a_SQL[14] = "select * from ref_trace_pc     where assignation = '%1' and ligne = %2 and direction = '%3' and trace = %4 and position = %5"
aRefTables[15] = "ref_trace_vres"
a_SQL[15] = "select * from ref_trace_vres   where assignation = '%1' and ligne = %2 and direction = '%3' and trace = %4 and no_vres = %5"
aRefTables[16] = "ref_voiture"
a_SQL[16] = "select * from ref_voiture where assignation = '%1' and type_service = '%2' and voiture = '%4'"
aRefTables[17] = "ref_vres"
a_SQL[17] = "select * from ref_vres where assignation = '%1' and no_vres = %2"
aRefTables[18] = "ref_noeudtc"
if sclasse = 'Arrêts'  && 20100415
  a_SQL[18] = "select a.no_arret, b.* from ref_noeudtc b inner join ref_arret a on a.assignation = b.assignation and a.no_noeudtc = b.no_noeudtc where a.assignation = '%1' and a.no_arret = %2"
else  && Points de contrôle
  a_SQL[18] = "select a.lieu, b.* from ref_noeudtc b inner join ref_pc a on a.assignation = b.assignation and a.no_noeudtc = b.no_noeudtc where a.assignation = '%1' and a.lieu = %2"
endif
if not empty(sTableTroncon)
  aRefTables[19] = 'ref_troncon_*USER*'
  a_SQL[19] = "select * from " + sTableTroncon + ;
             " where assignation = '%1' and type_service = '%2' and ligne = %3 and voiture = '%4' and hre_pre_dep = '%5' and no_troncon = %6"
endif

private i, schamps, schampsdbf, sxpath, onodes, j, sfld, sxpath2, onodfld, etable, sflddbf
private xx, yy, ssql, j, sfrom, sto
declare achamps[1]
declare achampsdbf[1]

?
? "Remplir les champs prévus inclus dans Regroupement et Complement"
select requete0
set order to

for i = 1 to inbtables
  schamps = ''
  schampsdbf = ''

  for k = 1 to 2
    sclassechamps = iif(k=1,"Regroupement", "Complement")
    sxpath = "//DefChampSortie/" + sclassechamps + "/Champ"
    onodes = oxmlreq.selectnodes(sxpath)
    for j = 0 to onodes.length - 1
      sfld = lower(onodes.item(j).getattributenode('Nom').value)
*      if not (sfld+' ' $ "assignation type_service voiture ligne hre_pre_dep " + ;
*                         "hre_pre_arret hre_pre_pc position direction trace " + ;
*                         "no_arret lieu no_balise no_vres ")
      sxpath2 = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
      onodfld = oxmlclasse.selectsinglenode(sxpath2)
      oblig = onodfld.getattributenode("Obligatoire").value
      if oblig = 'N'
        sbasetable = onodfld.getattributenode("BaseTable").value
        sflddbf = onodfld.getattributenode("OutDBF").value
        if 'v_sdap' $ sbasetable
          sbasetable = onodfld.getattributenode("SrcTable").value
        endif
        if sbasetable = aRefTables[i]
          schamps = iif(empty(schamps), sfld, schamps + ',' + sfld)
          schampsdbf = iif(empty(schampsdbf), sflddbf, schampsdbf + ',' + sflddbf)
        endif
      endif
    next 
  next 
  
  if not empty(schamps)
    ? '  ' + aRefTables[i] + ' : ' + schamps
    
    xx = split('achampsdbf', schampsdbf, ',')
    yy = split('achamps', schamps, ',')
    select requete0
    if maxrequete0 > 0
      go maxrequete0
      skip
    else
      go top
    endif
    do while not eof()
      ssql = a_SQL[i]
      do case
        case aRefTables[i] = 'ref_arret' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(no_arr)))
        case aRefTables[i] = 'ref_assignation' 
          ssql = strtran(ssql, "%1", assign)
        case aRefTables[i] = 'ref_balise' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(no_bal)))
        case aRefTables[i] = 'ref_calendrier' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
        case aRefTables[i] = 'ref_course' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", trim(str(ligne)))
          ssql = strtran(ssql, "%4", voiture)
          ssql = strtran(ssql, "%5", HPcrsDP)          
        case aRefTables[i] = 'ref_course_arret' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", trim(str(ligne)))
          ssql = strtran(ssql, "%4", voiture)
          ssql = strtran(ssql, "%5", HPcrsDP)
          ssql = strtran(ssql, "%6", HParrAR)          
        case aRefTables[i] = 'ref_course_pc' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", trim(str(ligne)))
          ssql = strtran(ssql, "%4", voiture)
          ssql = strtran(ssql, "%5", HPcrsDP)
          ssql = strtran(ssql, "%6", HPpcAR)
        case aRefTables[i] = 'ref_course_vres' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", trim(str(ligne)))
          ssql = strtran(ssql, "%4", voiture)
          ssql = strtran(ssql, "%5", HPcrsDP)
          ssql = strtran(ssql, "%6", trim(str(no_vres)))
        case aRefTables[i] = 'ref_ligne' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
        case aRefTables[i] = 'ref_pc' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", lieu)
        case aRefTables[i] = 'ref_trace' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
          ssql = strtran(ssql, "%3", dir)
          ssql = strtran(ssql, "%4", trace)
        case aRefTables[i] = 'ref_trace_arret' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
          ssql = strtran(ssql, "%3", dir)
          ssql = strtran(ssql, "%4", trace)
          ssql = strtran(ssql, "%5", trim(str(no_arr)))
        case aRefTables[i] = 'ref_trace_balise' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
          ssql = strtran(ssql, "%3", dir)
          ssql = strtran(ssql, "%4", trace)
          ssql = strtran(ssql, "%5", trim(str(no_bal)))
        case aRefTables[i] = 'ref_trace_pc' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
          ssql = strtran(ssql, "%3", dir)
          ssql = strtran(ssql, "%4", trace)
          ssql = strtran(ssql, "%5", lieu)
        case aRefTables[i] = 'ref_trace_vres' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(ligne)))
          ssql = strtran(ssql, "%3", dir)
          ssql = strtran(ssql, "%4", trace)
          ssql = strtran(ssql, "%5", trim(str(no_vres)))
        case aRefTables[i] = 'ref_voiture' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", voiture)
        case aRefTables[i] = 'ref_vres' 
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", trim(str(no_vres)))
        case aRefTables[i] = 'ref_noeudtc' && 20100415
          ssql = strtran(ssql, "%1", assign)
          if sclasse = 'Arrêts'  && 20100415
            ssql = strtran(ssql, "%2", trim(str(no_arr)))
          else
            ssql = strtran(ssql, "%2", "'" + trim(lieu) + "'")
          endif
        case aRefTables[i] = 'ref_troncon_*USER*'
          ssql = strtran(ssql, "%1", assign)
          ssql = strtran(ssql, "%2", typ_ser)
          ssql = strtran(ssql, "%3", trim(str(ligne)))
          ssql = strtran(ssql, "%4", voiture)
          ssql = strtran(ssql, "%5", HPcrsDP)
          ssql = strtran(ssql, "%6", trim(str(no_tron)))
      endcase 
      
      x = sqlexec(hconn, ssql, 'rs')
      select requete0
      for j = 1 to xx
        sto = achampsdbf[j]
        sfrom = 'rs.' + achamps[j]
        if (' '+sto) $ ' mois semaine jour_sem heure dheure qheure'
          do fillRefFields_Champscalcules
        else
          replace &sto with &sfrom
        endif
      next
      skip
    enddo
  endif
  
next 

return

procedure fillRefFields_Champscalcules
  if sto = 'mois' or sto = 'semaine' or sto = 'jour_sem'
    * on ne peut calculer ces chamsp de reference de courses manquantes
    return
  endif

  if sto = 'heure' or  sto = 'dheure' or sto = 'qheure' 
    private schamp1, hre, min, dhre, qhre
    do case 
      case sClasse = 'Courses'
        sChamp1 = 'HPcrsDP'
      case sClasse = 'Arrêts'
        sChamp1 = 'HParrAR'
      case sClasse = 'Points de contrôle'
        sChamp1 = 'HPpcAR'
      case sClasse = 'Tronçons'
        sChamp1 = 'HPcrsDP'
    endcase

    do case 
      case sto = 'heure'
        replace &sto with substr(&schamp1, 1, 2) + ':00:00'
      case sto = 'dheure'
        hre = substr(&schamp1, 1, 3)
        min = substr(&schamp1, 4, 2)
        if min <= '29'
          dhre = hre + '00:00'
        else  
          dhre = hre + '30:00'
        endif              
        replace &sto with dhre
      case sto = 'qheure'
        hre = substr(&schamp1, 1, 3)
        min = substr(&schamp1, 4, 2)
        if min <= '14'
          qhre = hre + '00:00'
        endif              
        if min >= '15' and min <= '29'
          qhre = hre + '15:00'
        endif              
        if min >= '30' and min <= '44'
          qhre = hre + '30:00'
        endif              
        if min >= '45'
          qhre = hre + '45:00'
        endif              
        replace &sto with qhre
    endcase
  endif

return


*-----------------------------------------------------------------------------------------
