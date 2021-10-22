* imputation.prg
*-----------------------------------------------------------------------------------------
parameters hconn, oxmlreq, oxmlclasse
private sclasse, fulldbfparam, dbfparam, def_centile
private def_fs, pNbAnnees, plagedef, intermin, intermax, pmax, qmin_g, qmin_a, tamponhre
private NOCENTIL, DateDebutEte, DateFinEte 
NOCENTIL = 999999 

?
? "Imputation des données"
? "  input: req_agr.dbf  output: req_imp.dbf"
close databases all

if not file('req_agr.dbf')
  messagebox("La table req_agr.dbf n'existe pas !")
  g_bOK = .f.
  return
endif

*****
local onode, soper, sxpath

onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='Classe']")
sclasse = onode.text

onode = g_oxmlparam.selectsinglenode("//CentileDefaut")
def_centile = val(onode.text) / 100
onode = g_oxmlparam.selectsinglenode("//Imputation/FsDefaut")
def_fs = val(onode.text) / 100
onode = g_oxmlparam.selectsinglenode("//Imputation/NbAnnees")
pNbAnnees = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/PlageDef")
plagedef = val(onode.text) 
onode = g_oxmlparam.selectsinglenode("//Imputation/IntervalleMin")
intermin = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/IntervalleMax")
intermax = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/PlafonnementNbArrets")
pmax = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/CoteGlobalMimimum")
qmin_g = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/CoteAchalandageMimimum")
qmin_a = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/TamponHeure")
tamponhre = val(onode.text)
onode = g_oxmlparam.selectsinglenode("//Imputation/DateDebutEte")
DateDebutEte = onode.text
onode = g_oxmlparam.selectsinglenode("//Imputation/DateFinEte")
DateFinEte = onode.text

*****
private methodCentile, userdef_centile
local sflds2

soper = "Imputation de données"
sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
onode = oxmlreq.SelectSingleNode(sxpath)
if substr(onode.text,1,5) = 'Table'
  methodCentile = 'Table'
  fulldbfparam = onode.text
  if file(fulldbfparam)
    dbfparam = substr(fulldbfparam, rat('\', fulldbfparam) + 1)
  else
    dbfparam = 'table_imput.dbf'
  endif
  if not file(dbfparam)
    if file(fulldbfparam)
      copy file &fulldbfparam to &dbfparam
    else
      copy file (getenv('stad_home') + '\config\table_imput.dbf') to &dbfparam
    endif
  endif

  select 0
  use &dbfparam alias table_imput
  x = afields(aflds1)
  sflds2 = 'LIGNE,PERIODE,DIRECTION,P_DEB,P_FIN,P_AUT'
  for i = 1 to 12
    sflds2 = sflds2 + ',FS' + alltrim(str(i))
  next
  declare aflds2[1]
  x = split('aflds2', sflds2, ',')
  for i = 1 to 6
    if ! (aflds1[i, 1] = aflds2[i])
      messagebox("Le fichier de paramètres a une structure invalide.")
      g_bok = .f.
      return
    endif
  next

  index on str(ligne,3) + periode + direction tag pkimput
  set order to pkimput
else
  * Centile
  methodCentile = 'Constante'
  userdef_centile = val(substr(onode.text, at('[', onode.text) + 1, ;
                        at(']',  onode.text) - at('[', onode.text) - 1))
  userdef_centile = userdef_centile / 100                      
endif

*****
local stables2

open database bd_requete

nb = adbobjects(atables, "TABLE")
if nb > 0
  stables2 = "CANDIDATS,IMPUTATION,REQ_IMP,REQ_IMPCOURSES,IMP_INDICATEURS"
  declare atables2[1]
  x = split('atables2', stables2, ',')
  for i = 1 to x
    if ascan(atables, atables2[i]) > 0
      remove table atables2[i] delete
    endif
  next
endif

select 0
use req_agr
copy to req_imp database bd_requete
use

******
* trouver et valider les champs de compilation
private sChampCharge, ChampsCompiles, fMont, fDesc, fChrg
local onodes, sfld, onodfld, sflddbf, onodes2, sop, iop, sflddbf2, schamps

if sclasse = 'Courses'
  sChampCharge = 'charge_max'
else
  sChampCharge = 'charge'
endif
ChampsCompiles = ''
store '' to fMont, fDesc, fChrg
sxpath = "//DefChampSortie/Compilation/Champ"
onodes = oxmlreq.selectnodes(sxpath)
for i = 0 to onodes.length - 1
  sfld = onodes.item(i).getattributenode('Nom').value
  if sfld = 'montants' or sfld = 'descendants' or sfld = sChampCharge
    sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
    onodfld = oxmlclasse.selectsinglenode(sxpath)
    sflddbf = onodfld.getattributenode("OutDBF").value
    sxpath = "./Parametre"
    onodes2 = onodes.item(i).selectnodes(sxpath)
    for j = 0 to onodes2.length - 1
      sop = onodes2.item(j).getattributenode('Nom').value
      iop = ascan(g_agregs, sop)
      sflddbf2 = lower(sflddbf + g_agregs(iop+1))
      alter table req_imp alter &sflddbf2 b
      if j = 0
        if sfld = 'montants'
          ChampsCompiles = ChampsCompiles + 'M'
          fMont = sflddbf2
        endif
        if sfld = 'descendants'
          ChampsCompiles = ChampsCompiles + 'D'
          fDesc = sflddbf2
        endif
        if sfld = sChampCharge
          ChampsCompiles = ChampsCompiles + 'C'
          fChrg = sflddbf2
        endif
      endif
    next
  endif
next
if len(ChampsCompiles) <> 3
  schamps = "montants, descendants, " + sChampCharge
  messagebox("Les champs " + schamps + " doivent être inclus dans la requête.")
  g_bOK = .f.
  return
endif

******
private fAssign, fTypser, fLigne, fHrepredep, fDir, fPos, fVoiture, fDate

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='assignation']")
fAssign = onode.getattributenode("OutDBF").value
if lower(fAssign ) <> 'assignation'
  alter table req_imp rename column &fAssign to assignation
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='type_service']")
fTypser = onode.getattributenode("OutDBF").value
if lower(fTypser) <> 'type_service'
  alter table req_imp rename column &fTypser to type_service
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='ligne']")
fLigne = onode.getattributenode("OutDBF").value
if lower(fLigne) <> 'ligne'
  alter table req_imp rename column &fLigne to ligne
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='hre_pre_dep']")
fHrepredep = onode.getattributenode("OutDBF").value
if lower(fHrepredep) <> 'hre_pre_dep'
  alter table req_imp rename column &fHrepredep to hre_pre_dep
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='direction']")
fDir = onode.getattributenode("OutDBF").value
if lower(fDir) <> 'direction'
  alter table req_imp rename column &fDir to direction
endif

if sclasse = 'Arrêts'
  onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
  fPos = onode.getattributenode("OutDBF").value
  if lower(fPos) <> 'position'
    alter table req_imp rename column &fPos to position
  endif
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='voiture']")
fVoiture = onode.getattributenode("OutDBF").value

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='sdap_date']")
fDate = onode.getattributenode("OutDBF").value

******
local idx_expr

select req_imp
if sclasse <> 'Arrêts'
  index on assignation + type_service + str(ligne, 3) + direction + hre_pre_dep ;
        tag pklong && candidate && 20060220: enlever le candidate

else
  index on assignation + type_service + str(ligne, 3) + direction + hre_pre_dep + ;
           str(position, 5) ;
        tag pklong && candidate && 20060220: enlever le candidate
endif

******

do MarkAgreg

if sclasse = 'Arrêts'
  do FillNull
endif

do FillAgreg

do CreateImputation

do Imputation1
  * selectCourses
  * getProfil()
  * ImputABC()
    * fCentile
  * putImputCourse
  * putImputArrets

do Imputation2
  * getProfil()
  * fCentile
  * putImputCourse
  * putImputArrets

do putImpcod

do Indicateurs

******

select req_imp
on error do ignore
delete tag pklong
on error

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='assignation']")
fAssign = onode.getattributenode("OutDBF").value
if lower(fAssign ) <> 'assignation'
  alter table req_imp rename column assignation to &fAssign
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='type_service']")
fTypser = onode.getattributenode("OutDBF").value
if lower(fTypser) <> 'type_service'
  alter table req_imp rename column type_service to &fTypser
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='ligne']")
fLigne = onode.getattributenode("OutDBF").value
if lower(fLigne) <> 'ligne'
  alter table req_imp rename column ligne to &fLigne
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='hre_pre_dep']")
fHrepredep = onode.getattributenode("OutDBF").value
if lower(fHrepredep) <> 'hre_pre_dep'
  alter table req_imp rename column hre_pre_dep to &fHrepredep
endif

onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='direction']")
fDir = onode.getattributenode("OutDBF").value
if lower(fDir) <> 'direction'
  alter table req_imp rename column direction to &fDir
endif

if sclasse = 'Arrêts'
  onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@Nom='position']")
  fPos = onode.getattributenode("OutDBF").value
  if lower(fPos) <> 'position'
    alter table req_imp rename column position to &fPos
  endif
endif

******
local idx_expr

select req_imp
idx_expr = fAssign+' + '+fTypser+' + str('+fLigne+', 3) + '+fDir+' + '+fHrepredep
if sclasse = 'Arrêts'
  idx_expr = idx_expr+' + str('+fPos+', 5)'
endif
index on &idx_expr tag pk && candidate && 20060220: enlever le candidate

return

*-----------------------------------------------------------------------------------------

procedure MarkAgreg
  * indiquer la source des enregistrements de la table REQ_IMPCOURSES
  * dans le champ GENERE
  *   AGR : cette course provient de la BDU
  *         cette course a été agrégé
  *   GEN : cette course a été généré par le processus GEN_COURSES
  *         cette course est à imputer
  *
  * plus tard dans le traitement, GEN est remplacé par I1 ou I2
  *   I1  : cette course a été rempli par la méthode d'imputation des courses 1
  *         (à partir d'autres comptages)
  *   I2  : cette course a été rempli par la méthode d'imputation des courses 2
  *         (â partir des courses préc. et suiv.)
  local idx_expr

  if sclasse = 'Courses'
    select assignation, type_service, ligne, periode, voiture, hre_pre_dep, ;
           direction, &fMont as mont, &fDesc as desc, &fChrg as chrg ;
      from req_imp ;
      into table req_impcourses database bd_requete
* 20060220: enlever le group by
*     group by assignation, type_service, ligne, periode, voiture, hre_pre_dep, ;
*           direction
  else
    select assignation, type_service, ligne, periode, voiture, hre_pre_dep, ;
           direction, sum(&fMont) as mont, sum(&fDesc) as desc, max(&fChrg) as chrg ;
      from req_imp ;
      into table req_impcourses database bd_requete ;
     group by assignation, type_service, ligne, periode, voiture, hre_pre_dep, ;
           direction
  endif

  alter table req_impcourses add column genere c(3)
  select req_impcourses
  go top
  do while not eof()
    select count(*) as mcnt ;
      from req_imp ;
     where assignation = req_impcourses.assignation ;
       and type_service = req_impcourses.type_service ;
       and ligne = req_impcourses.ligne ;
       and voiture = req_impcourses.voiture ;
       and hre_pre_dep = req_impcourses.hre_pre_dep ;
       and not isnull(&fMont) ;
      into cursor rs
    select req_impcourses
    if rs.mcnt = 0
      replace genere with 'GEN'
    else
      replace genere with 'AGR'
    endif
    skip
  enddo

  alter table req_impcourses add column nb_arr n(3)
  alter table req_impcourses add column capacite n(3)
  alter table req_impcourses add column cote_bs b
  alter table req_impcourses add column fs b

  select 0
  use requete0
  idx_expr = fDate+' + str('+fLigne+', 3) + '+fHrepredep+' + '+fDir
  if sclasse = 'Arrêts'
    idx_expr = idx_expr+' + str('+fPos+', 5)'
  endif
  index on &idx_expr tag req0_gen
  set order to req0_gen

return

*-----------------------------------------------------------------------------------------

procedure FillNull
  *  remplace les valeurs 0 par NULL pour les arrêts générés pour les courses agrégés

  select req_impcourses
  go top
  do while not eof()
    if genere = 'AGR'
      *----
      * déterminer les champs de compilation
      sxpath = "//DefChampSortie/Compilation/Champ"
      onodes = oxmlreq.selectnodes(sxpath)
      for i = 0 to onodes.length - 1
        sfld = onodes.item(i).getattributenode('Nom').value
        if sfld = 'montants' or sfld = 'descendants' or sfld = 'charge'
          sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
          onodfld = oxmlclasse.selectsinglenode(sxpath)
          sflddbf = onodfld.getattributenode("OutDBF").value
          sxpath = "./Parametre"
          onodes2 = onodes.item(i).selectnodes(sxpath)
          for j = 0 to onodes2.length - 1
            sop = onodes2.item(j).getattributenode('Nom').value
            iop = ascan(g_agregs, sop)
            sflddbf2 = lower(sflddbf + g_agregs(iop+1))

            update req_imp ;
               set &sflddbf2 = 0 ;
             where assignation = req_impcourses.assignation ;
               and type_service = req_impcourses.type_service ;
               and ligne = req_impcourses.ligne ;
               and voiture = req_impcourses.voiture ;
               and hre_pre_dep = req_impcourses.hre_pre_dep ;
               and isnull(&sflddbf2)
          next
        endif
      next
    endif

    select req_impcourses
    skip
  enddo
return

*-----------------------------------------------------------------------------------------

procedure FillAgreg
  * pour les courses 'AGR' seulement:
  *   remplir les champs nb_arr, capacite, cote_bs, fs

  select req_impcourses
  set filter to genere = 'AGR'
  go top
  do while not eof()
    ssql = "select nb_arret" + ;
           "  from ref_course" + ;
           " where assignation = '" + assignation + "'" + ;
           "   and type_service = '" + type_service + "'" + ;
           "   and ligne = " + str(ligne,3) + ;
           "   and voiture = '" + voiture + "'" + ;
           "   and hre_pre_dep = '" + hre_pre_dep + "'"
    x = sqlexec(hConn, ssql)
    select req_impcourses
    replace nb_arr with sqlresult.nb_arret

    do selectCourses with 'AGR', type_service, ligne, direction, hre_pre_dep, ;
                          hre_pre_dep, assignation, nb_arr

    select req_impcourses
    skip
  enddo
  set filter to
return

*-----------------------------------------------------------------------------------------

procedure CreateImputation
  * creer la table des imputations

  create table imputation ;
   ( assignation  c(8), ;
     type_service c(2), ;
     ligne        n(3), ;
     voiture      c(8), ;
     hre_pre_dep  c(8), ;
     direction    c(1), ;
     trace        n(2), ;
     periode      c(2), ;
     imput_date   c(8), ;
     type_course  c(1), ;
     cod_imp      c(2), ;
     cod_av       n(1), cod_ap       n(1), ;
     cod_profil   c(2), ;
     m            b,    ;
     m_av         n(3), m_ap         n(3), ;
     d            b,    ;
     d_av         n(3), d_ap         n(3), ;
     ch           b,    ;
     ch_av        n(3), ch_ap        n(3), ;
     hre_pre_1    c(8), ;
     hre_pre_2    c(8), ;
     hre_pre_av   c(8), hre_pre_ap   c(8), ;
     nb_arr       n(3), ;
     nb_arr_av    n(3), nb_arr_ap    n(3), ;
     capacite     n(3), ;
     capacite_av  n(3), capacite_ap  n(3), ;
     cote_bs_av   n(4), cote_bs_ap   n(4), ;
     fs_av        b,    fs_ap        b,    ;
     int_1        n(3), ;
     int_2        n(3), ;
     int_av       n(3), int_ap       n(3), ;
     plage_deb    c(8), ;
     plage_fin    c(8), ;
     nb_courses   n(4), ;
     fs           b,    ;
     ass_a        c(8), ass_b        c(8), ass_c        c(8), ;
     date_a       c(8), date_b       c(8), date_c       c(8), ;
     hre_a        c(8), hre_b        c(8), hre_c        c(8), ;
     m_a          n(3), m_b          n(3), m_c          n(3), ;
     d_a          n(3), d_b          n(3), d_c          n(3), ;
     ch_a         n(3), ch_b         n(3), ch_c         n(3), ;
     nb_arr_a     n(3), nb_arr_b     n(3), nb_arr_c     n(3), ;
     capacite_a   n(3), capacite_b   n(3), capacite_c   n(3), ;
     cote_bs_a    n(4), cote_bs_b    n(4), cote_bs_c    n(4), ;
     fs_a         b,    fs_b         b,    fs_c         b     ;
   )

return

*-----------------------------------------------------------------------------------------

procedure Imputation1
  * essayer le première méthode d'imputation pour les courses à imputer
  local i, cntCourses, ssql, x, courseprec, coursesuiv, mdiff

  select req_impcourses
  set filter to genere = 'GEN'
  count to cntCourses
  ? '  Imputation 1'

  i = 1
  go top
  do while not eof()

    * -----
    * création de la course dans imputation.dbf

    select imputation
    append blank
    replace assignation  with req_impcourses.assignation,  ;
            type_service with req_impcourses.type_service, ;
            ligne        with req_impcourses.ligne,        ;
            voiture      with req_impcourses.voiture,      ;
            hre_pre_dep  with req_impcourses.hre_pre_dep,  ;
            direction    with req_impcourses.direction,    ;
            periode      with req_impcourses.periode
    replace cod_imp    with .null., ;
            cod_profil with .null., ;
            fs_av      with .null., ;
            fs_ap      with .null.

    * -----
    * calcul des champs de base pour la construction de la requête

    ssql = "select trace, nb_arret, periode" + ;
           "  from ref_course" + ;
           " where assignation = '" + imputation.assignation + "'" + ;
           "   and type_service = '" + imputation.type_service + "'" + ;
           "   and ligne = " + str(imputation.ligne,3) + ;
           "   and voiture = '" + imputation.voiture + "'" + ;
           "   and hre_pre_dep = '" + imputation.hre_pre_dep + "'"
    x = sqlexec(hConn, ssql)
    select imputation
    replace trace   with sqlresult.trace,    ;
            nb_arr  with sqlresult.nb_arret, ;
            periode with sqlresult.periode

    courseprec = ''
    ssql = "select hre_pre_dep" + ;
           "  from ref_course" + ;
           " where type_service = '" + imputation.type_service + "'" + ;
           "   and ligne = " + str(imputation.ligne,3) + ;
           "   and direction = '" + imputation.direction + "'" + ;
           "   and hre_pre_dep < '" + imputation.hre_pre_dep + "'" + ;
           " order by hre_pre_dep desc"
    x = sqlexec(hConn, ssql)

    if not eof('sqlresult')
      courseprec = sqlresult.hre_pre_dep
      select imputation
      replace hre_pre_1 with courseprec
      mdiff = (hre2sec(hre_pre_dep) - hre2sec(courseprec))
      if periode = 'AM' or periode = 'PM'
        mdiff = min(mdiff, intermax * 60)
      endif
      if mdiff < intermin * 60    && 20081209; calculer une intervalle minimum
        mdiff = intermin * 60
      endif
      replace int_1 with mdiff / 60
      replace plage_deb with sec2hre(hre2sec(hre_pre_dep) - mdiff / 2)
      replace type_course with 'I'
    else
      select imputation
      replace hre_pre_1 with .null.
      replace int_1 with plagedef
      replace plage_deb with sec2hre(hre2sec(hre_pre_dep) - plagedef * 60)
      replace type_course with 'P'
    endif
 
    coursesuiv = ''
    ssql = "select hre_pre_dep" + ;
           "  from ref_course" + ;
           " where type_service = '" + imputation.type_service + "'" + ;
           "   and ligne = " + str(imputation.ligne,3) + ;
           "   and direction = '" + imputation.direction + "'" + ;
           "   and hre_pre_dep > '" + imputation.hre_pre_dep + "'" + ;
           " order by hre_pre_dep asc"
    x = sqlexec(hConn, ssql)
    if not eof('sqlresult')
      coursesuiv = sqlresult.hre_pre_dep
      select imputation
      replace hre_pre_2 with coursesuiv
      mdiff = (hre2sec(coursesuiv) - hre2sec(hre_pre_dep))
      if periode = 'AM' or periode = 'PM'
        mdiff = min(mdiff, intermax * 60)
      endif
      if mdiff < intermin * 60    && 20081209; calculer une intervalle minimum
        mdiff = intermin * 60
      endif
      replace int_2 with mdiff / 60
      replace plage_fin with sec2hre(hre2sec(hre_pre_dep) + mdiff / 2)
      replace type_course with 'I'
    else
      select imputation
      replace hre_pre_2 with .null.
      replace int_2 with plagedef
      replace plage_fin with sec2hre(hre2sec(hre_pre_dep) + plagedef * 60)
      replace type_course with 'D'
    endif

    select imputation
    ? str(i, 4) + ' / ' + str(cntcourses, 4), ;
      assignation, type_service, ligne, direction, hre_pre_dep

    * -----
    do selectCourses with 'IMP', type_service, ligne, direction, plage_deb, ;
                          plage_fin, assignation, nb_arr

    * -----
    local datdeb, datfin, datmilieu, sdatmilieu

*    select sdap_date as dt;
*      from candidats ;
*      into cursor rs ;
*     where assignation = imputation.assignation ;
*     order by sdap_date
* 20060220: enlever le where

    select sdap_date as dt;
      from candidats ;
      into cursor rs ;
     order by sdap_date
    select rs
    go top
    datdeb = ctod(substr(dt,1,4)+'.'+substr(dt, 5, 2)+'.'+substr(dt,7,2))
    go bottom
    datfin = ctod(substr(dt,1,4)+'.'+substr(dt, 5, 2)+'.'+substr(dt,7,2))
    datmilieu = datdeb + (datfin - datdeb) / 2
    use
* 20080123: modification de imput_date; modifié aussi la source de dtimput
*    select imputation
*    replace imput_date with dtos(datmilieu)

    * -----
    local datimput, datsdap

    select imputation
*    dtimput = imput_date
*    dtimput = dtos(datmilieu)
    dtimput = dtos(datfin)     && 20081209; changer la date d'imputation
    dtimput = ctod(substr(dtimput,1,4)+'.'+substr(dtimput,5,2)+'.'+substr(dtimput,7,2))

    select candidats
    go top
    do while not eof()
      dtsdap = sdap_date
      dtsdap = ctod(substr(dtsdap,1,4)+'.'+substr(dtsdap,5,2)+'.'+substr(dtsdap,7,2))
      replace proximite with abs(dtsdap - dtimput)
      skip
    enddo

    index on str(ind_assign, 1) + str(proximite, 5) tag proximite
    set order to proximite

    * -----

    if ImputABC()
      thecourse = getProfil('I1')
      if thecourse <> 'x'
        if sclasse = 'Courses'
          do putImputCourse with .t.
        else
          do putImputCourse with .f.
          do putImputArrets with 'I1', thecourse
        endif
        replace genere with 'I1' in 'req_impcourses'
        ?? ' I'
      endif
    endif

    select candidats
    use
    remove table candidats delete
    i = i + 1

    select req_impcourses
    skip
  enddo

  select req_impcourses
  set filter to
return

*-----------------------------------------------------------------------------------------

procedure selectCourses
  * si flag = 'AGR', pour cette course AGRÉGÉE
  *   calculer les champs capacite. cote_bs, fs
  * si flag = 'IMP', pour cette course à IMPUTER
  *   trouver les courses candidates
  *   aussi, trouver capacite et nb_courses de la course à imputer

  parameters flag, typser, nligne, mdir, hremin, hremax, sassign, nbarr
  local datemin, filtrebooking, filtretypser, nbarrmax, ssql, x

  * sélectionner dans les x dernières années seulement
  datemin = str(year(date()) - pNbAnnees, 4) + substr(str(100 + month(date()), 3), 2, 2)

  * filtrer sur des bookings compatibles
  filtrebooking = "(substring(sdap_course.sdap_date,5,4) between '" + DateDebutEte + "' and '" + DateFinEte + "')"
  if not (substr(sassign,5,4) >= DateDebutEte and substr(sassign,5,4) <= DateFinEte)
    filtrebooking = " not " + filtrebooking
  endif
  * 20200618 Supprimer temporairement ce filtre
  filtrebooking = " 1=1 "

  * filtrer sur des types de service compatibles
  if substr(typser,1,1) = 'F'
    filtretypser = " <> 'SE'"
  else
    filtretypser = " = '" + typser + "'"
  endif

  * filtrer sur des tracés compatibles
  nbarrmax = int(nbarr * pmax)

  * filtrer sur une qualité minimum
  * paramètres qmin_g, qmin_a

  ssql = ;
    "SELECT ref_course.assignation, ref_course.type_service, ref_course.ligne," + ;
    "       ref_course.voiture, ref_course.hre_pre_dep, ref_course.direction," + ;
    "       ref_course.trace, ref_course.periode, sdap_course.montants," + ;
    "       sdap_course.descendants, sdap_course.charge_max, ref_course.nb_arret," + ;
    "       pil_autobus.capacite, sdap_course.cotebs_global, sdap_course.sdap_date," + ;
    "       sdap_course.cotebs_achal, 00000 as proximite," + ;
    "       substring(sdap_course.sdap_date,5,2) as smois," + ;
    "       case sdap_course.assignation" + ;
    "         when '" + sassign + "' then 0 else 1 end as ind_assign" + ;
    "  FROM sdap_course" + ;
    " INNER JOIN ref_course" + ;
    "    ON sdap_course.assignation = ref_course.assignation" + ;
    "   AND sdap_course.type_service = ref_course.type_service" + ;
    "   AND sdap_course.ligne = ref_course.ligne" + ;
    "   AND sdap_course.voiture = ref_course.voiture" + ;
    "   AND sdap_course.hre_pre_dep = ref_course.hre_pre_dep" + ;
    " INNER JOIN pil_autobus" + ;
    "    ON sdap_course.assignation = pil_autobus.assignation" + ;
    "   AND sdap_course.no_bus = pil_autobus.no_bus"
  if flag = 'IMP'
    * 20081209; changer filtrebooking
    ssql = ssql + ;
      " WHERE sdap_course.assignation >= '" + datemin + "'" + ;
      "   AND " + filtrebooking + ;
      "   AND sdap_course.type_service" + filtretypser + ;
      "   AND sdap_course.ligne = " + str(nligne,3) + ;
      "   AND (sdap_course.hre_pre_dep between '" + hremin + "' and '" + hremax + "')" + ;
      "   AND ref_course.direction = '" + mdir + "'" + ;
      "   AND (ref_course.nb_arret between "+str(nbarr,3)+" and "+str(nbarrmax,3)+")"+;
      "   AND (not (sdap_course.cotebs_achal is null))" + ;
      "   AND sdap_course.cotebs_global > " + alltrim(str(qmin_g)) + ;
      "   AND sdap_course.cotebs_achal > " + alltrim(str(qmin_a))
    do savesql with ssql, 'candidats'
  else
    ssql = ssql + ;
      " WHERE sdap_course.assignation = '" + sassign + "'" + ;
      "   AND sdap_course.type_service = '" + typser + "'" + ;
      "   AND sdap_course.ligne = " + str(nligne,3) + ;
      "   AND sdap_course.hre_pre_dep ='" + hremin + "'" + ;
      "   AND ref_course.direction = '" + mdir + "'"
  endif
  ssql = ssql + "   AND (sdap_course.montants is not null)"
  x = sqlexec(hConn, ssql)

  if flag = 'IMP'
    do savesql with ssql, 'candidats'
    
    SELECT sqlresult
    COPY TO candidats database bd_requete

    *-----
    select capacite, count(*) as mcnt ;
      from candidats ;
      into cursor rs ; 
     group by capacite ;
     order by mcnt desc

    select imputation
    replace capacite   with rs.capacite

    select rs
    use

    *-----
    select count(*) as mcnt ;
      from candidats ;
      into cursor rs ;

    select imputation
    replace nb_courses with rs.mcnt

    select rs
    use

  else

    select capacite, count(*) as mcnt ;
      from sqlresult ;
      into cursor rs ;
     group by capacite ;
     order by mcnt desc

    select rs
    mcapacite = capacite
    use

    *-----
    select avg(cotebs_global) as avg_cote_bs ;
      from sqlresult ;
      into cursor rs

    select rs
    mcotebs = avg_cote_bs
    use

    *-----
    select smois, count(*) as mcnt ;
      from sqlresult ;
      into cursor rs ;
     group by smois ;
     order by mcnt desc

    select rs
    mmois = smois
    use

    if methodCentile = 'Table'
      select table_imput
      seek str(sqlresult.ligne, 3) + ;
           getperiode(sqlresult.hre_pre_dep) + ;
           sqlresult.direction
      if not eof()
        fld_fs = 'fs' +  iif(substr(mmois,1,1)='0', substr(mmois,2,1), mmois)
        mfs = &fld_fs
      else
        mfs = def_fs
      endif
    else
      mfs = def_fs
    endif

    *-----
    select req_impcourses
    replace capacite with mcapacite, ;
            cote_bs  with mcotebs,   ;
            fs       with mfs
  endif

return

*-----------------------------------------------------------------------------------------

function ImputABC()
  * exécuter la première méthode d'imputation
  local i, bsucces, indcou, champ
  i = 0
  bsucces = .f.

  if methodCentile = 'Table'
    select table_imput
    seek str(imputation.ligne, 3) + ;
         getperiode(imputation.hre_pre_dep) + ;
         imputation.direction
  endif

  declare aDates(3) && 20080123: modification de imput_date
  select candidats
  go top
  do while i < 3 and not eof()
    i = i + 1
    indcou = iif(i=1, 'a', iif(i=2, 'b', 'c'))
    select imputation
    champ = 'ass_' + indcou
    replace &champ with candidats.assignation
    champ = 'date_' + indcou
    replace &champ with candidats.sdap_date
    aDates(i) = candidats.sdap_date
    champ = 'hre_' + indcou
    replace &champ with candidats.hre_pre_dep
    champ = 'm_' + indcou
    replace &champ with candidats.montants
    champ = 'd_' + indcou
    replace &champ with candidats.descendants
    champ = 'ch_' + indcou
    replace &champ with candidats.charge_max
    champ = 'nb_arr_' + indcou
    replace &champ with candidats.nb_arret
    champ = 'capacite_' + indcou
    replace &champ with candidats.capacite
    champ = 'cote_bs_' + indcou
    replace &champ with candidats.cotebs_global

    champout = 'fs_' + indcou
    schamp = 'imputation.date_' + indcou
    smois = substr(&schamp,5,2)
    if substr(smois,1,1) = '0'
      smois = substr(smois,2,1)
    endif
    if methodCentile = 'Table'
      if not eof('table_imput')
        schamp2 = 'table_imput.fs' + smois
        replace &champout with &schamp2
      else
        replace &champout with def_fs
      endif
    else
      replace &champout with def_fs
    endif

    select candidats
    skip
  enddo
  if i > 0
    bsucces = .t.
    nb = i
  endif
  

  if bsucces
    select imputation
    replace cod_imp with 'I1'
    
    * 20080123: modification de imput_date
    if i = 1 or i = 2
      replace imput_date with aDates(1)
    else
      replace imput_date with aDates(2)
    endif

    smois = substr(imput_date,5,2)
    if substr(smois,1,1) = '0'
      smois = substr(smois,2,1)
    endif
    if methodCentile = 'Table'
      schamp = 'table_imput.fs' + smois
      if not eof('table_imput')
        replace fs with &schamp
      else
        replace fs with def_fs
      endif
    else
      replace fs with def_fs
    endif

    fldcentil = iif(type_course='P', 'p_deb', iif(type_course='D', 'p_fin','p_aut'))
    fldcentil = 'table_imput.' + fldcentil
    if methodCentile = 'Table'
      if not eof('table_imput')
        centile = &fldcentil / 100
      else
        centile = def_centile
      endif
    else
      centile = userdef_centile     && 20080602: typo
    endif

    declare centi[nb]
    external array acenti
    declare a[nb], n[nb], c[nb], fs[nb], bs[nb]
    store 0 to totn, totc, totbs
    for j = 1 to 3
      sj = iif(j=1, 'm', iif(j=2, 'd', 'ch'))
      for i = 1 to nb
        si = iif(i=1, 'a', iif(i=2, 'b', 'c'))
        champ = sj + '_' + si
        a[i] = &champ
        champ = 'nb_arr_' + si
        n[i] = &champ
        champ = 'capacite_' + si
        c[i] = &champ
        champ = 'fs_' + si
        fs[i] = &champ
        
        centi[i] = NOCENTIL
        if not isnull(n[i]) and not isnull(c[i]) and not isnull(fs[i])
          if n[i] > 0 and c[i] > 0 and fs[i] > 0
             centi[i] = a[i] * (nb_arr / n[i]) * (sqrt(capacite / c[i])) * (fs / fs[i]) && 20090113
          endif
        endif
             
        champ = 'cote_bs_' + si
        bs[i] = &champ
        if j = 1
          totn = totn + n[i]
          totc = totc + c[i]
          totbs = totbs + bs[i]
        endif
      next
      replace &sj with fCentile(@centi, centile)
    next


    select req_impcourses
    replace mont     with imputation.m,  ;
            desc     with imputation.d,  ;
            chrg     with imputation.ch, ;
            nb_arr   with totn / nb,     ;
            capacite with totc / nb,     ;
            cote_bs  with totbs / nb
    replace fs     with imputation.fs && Laurier 20141103; bonne solution
  endif

return bsucces

*-----------------------------------------------------------------------------------------

procedure Imputation2
  * exécuter la deuxième méthode d'imputation
  local i, cntCourses, fldcentil, centile

  select imputation
  * courses non imputées avec la premiêre méthode
  set filter to isnull(cod_imp) or isnull(cod_profil)
  * 2006/12/20; ajouter la condition sur cod_profil
  count to cntCourses
  ? '  Imputation 2'

  i = 1
  go top
  do while not eof()
    ? str(i, 4) + ' / ' + str(cntcourses, 4), ;
      assignation, type_service, ligne, direction, hre_pre_dep
    replace cod_av with 0 , cod_ap with 0

    *20080123: modification de imput_date
    replace imput_date with null  && ????

    *-----

    if methodCentile = 'Table'
      select table_imput
      seek str(imputation.ligne, 3) + ;
           getperiode(imputation.hre_pre_dep) + ;
           imputation.direction

      select imputation
      fldcentil = iif(type_course='P', 'p_deb', iif(type_course='D', 'p_fin','p_aut'))
      fldcentil = 'table_imput.' + fldcentil
      if not eof('table_imput')
        centile = &fldcentil / 100
      else
        centile = def_centile
      endif
    else
      centile = userdef_centile     && 20080602: typo
    endif

    *-----

    select * ;
      from req_impcourses ;
      into cursor rs_av ;
     where assignation  = imputation.assignation  ;
       and type_service = imputation.type_service ;
       and ligne        = imputation.ligne        ;
       and direction    = imputation.direction    ;
       and hre_pre_dep  < imputation.hre_pre_dep  ;
     order by hre_pre_dep desc

    select rs_av
    if not eof()
      hre1 = hre_pre_dep
      *int1 = ( hre2sec(imputation.hre_pre_dep) - hre2sec(hre1) ) / 60
      *select imputation
      *replace hre_pre_1 with hre1, ;
      *        int_1     with int1
      select rs_av
      do while genere = 'GEN' && généré mais non imputé
        skip
      enddo
      if not eof()
        hreav = hre_pre_dep
        intav = ( hre2sec(imputation.hre_pre_dep) - hre2sec(hreav) ) / 60
        select imputation
        if intav <= tamponhre
          replace hre_pre_av  with hreav,          ;
                  m_av        with rs_av.mont,     ;
                  d_av        with rs_av.desc,     ;
                  ch_av       with rs_av.chrg,     ;
                  nb_arr_av   with rs_av.nb_arr,   ;
                  capacite_av with rs_av.capacite, ;
                  cote_bs_av  with rs_av.cote_bs,  ;
                  fs_av       with rs_av.fs,       ;
                  int_av      with intav,          ;
                  cod_av      with iif(hreav = hre1, 1, 2)
        endif
      endif
    endif

    *-----

    select * ;
      from req_impcourses ;
      into cursor rs_ap ;
     where assignation  = imputation.assignation  ;
       and type_service = imputation.type_service ;
       and ligne        = imputation.ligne        ;
       and direction    = imputation.direction    ;
       and hre_pre_dep  > imputation.hre_pre_dep  ;
     order by hre_pre_dep asc

    select rs_ap
    if not eof()
      hre2 = hre_pre_dep
      *int2 = ( hre2sec(hre2) - hre2sec(imputation.hre_pre_dep) ) / 60
      *select imputation
      *replace hre_pre_2 with hre2, ;
      *        int_2     with int2
      select rs_ap
      do while genere = 'GEN' && généré mais non imputé
        skip
      enddo
      if not eof()
        hreap = hre_pre_dep
        intap = ( hre2sec(hreap) - hre2sec(imputation.hre_pre_dep) ) / 60
        select imputation
        if intap <= tamponhre
          replace hre_pre_ap  with hreap,          ;
                  m_ap        with rs_ap.mont,     ;
                  d_ap        with rs_ap.desc,     ;
                  ch_ap       with rs_ap.chrg,     ;
                  nb_arr_ap   with rs_ap.nb_arr,   ;
                  capacite_ap with rs_ap.capacite, ;
                  cote_bs_ap  with rs_ap.cote_bs,  ;
                  fs_ap       with rs_ap.fs,       ;
                  int_ap      with intap,          ;
                  cod_ap      with iif(hreap = hre2, 1, 2)
        endif
      endif
    endif

    *-----

    bsucces = .f.
    
*    if not eof('rs_av') and not eof('rs_ap')
*    if not empty('hre_pre_av') and not empty('hre_pre_ap')
    declare centi[2]
    external array acenti

    select imputation
    for j = 1 to 3
      sj = iif(j=1, 'm', iif(j=2, 'd', 'ch'))
      champav = sj + '_av'
      champap = sj + '_ap'

      centi[1] = NOCENTIL
      centi[2] = NOCENTIL
      if nb_arr_av > 0
        if capacite = 0
         replace capacite with capacite_av, ;
                 fs       with fs_av
        endif
        centi[1] = &champav*(nb_arr/nb_arr_av)*(capacite/capacite_av)*(fs/fs_av)
      endif  
      if nb_arr_ap > 0
        if capacite = 0
         replace capacite with capacite_ap, ;
                 fs       with fs_ap
        endif
        centi[2] = &champap*(nb_arr/nb_arr_ap)*(capacite/capacite_ap)*(fs/fs_ap)
      endif  
      replace &sj with fCentile(@centi, centile)
    next

    thecourse = getProfil('I2')
    if thecourse <> 'x'
      if sclasse = 'Courses'
        do putImputCourse with .t.
      else
        do putImputCourse with .f.
        do putImputArrets with 'I2', thecourse
      endif
      replace genere with 'I2' in 'req_impcourses'
      ?? ' I'
      select imputation
      replace cod_imp with 'I2'
      bsucces = .t.
    endif
*    endif
*    endif

    if not bsucces
      select imputation
      replace cod_imp with 'NI',   ;
              m       with .null., ;
              d       with .null., ;
              ch      with .null.
    endif

    *-----

    select imputation
    skip
    i = i + 1
  enddo

  set filter to
return

*-----------------------------------------------------------------------------------------

procedure putImpcod
  * transférer le code d'imputation dans la table req_imp
  alter table req_imp add column cod_imp c(3)

  select imputation
  go top
  do while not eof()
    update req_imp ;
       set cod_imp = imputation.cod_imp ;
     where assignation = imputation.assignation ;
       and type_service = imputation.type_service ;
       and ligne = imputation.ligne ;
       and voiture = imputation.voiture ;
       and hre_pre_dep = imputation.hre_pre_dep
    select imputation
    skip
  enddo

  update req_imp ;
     set cod_imp = 'AGR' ;
   where isnull(cod_imp)
return

*-----------------------------------------------------------------------------------------

procedure Indicateurs
  * calculer les indicateurs liés à l'imputation

  create table imp_indicateurs ;
   ( assignation  c(8),   ;
     type_service c(2),   ;
     ligne        n(3),   ;
     direction    c(1),   ;
     periode      c(2),   ;
     ni           n(5),   ;
     nv           n(5),   ;
     nt           n(5),   ;
     pc_ni        n(5,1), ;
     pc_nv        n(5,1), ;
     mi           b,      ;
     mt           b,      ;
     pc_mi        n(5,1), ;
     di           b,      ;
     dt           b,      ;
     pc_di        n(5,1)  )

  *----

  select assignation, type_service, ligne, direction, periode ;
    from imputation ;
    into cursor rs ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    append blank
    replace assignation  with rs.assignation,  ;
            type_service with rs.type_service, ;
            ligne        with rs.ligne,        ;
            direction    with rs.direction,    ;
            periode      with rs.periode,      ;
            ni with 0, nv with 0, pc_ni with 0, pc_nv with 0, ;
            mi with 0, mt with 0, pc_mi with 0, ;
            di with 0, dt with 0, pc_di with 0
    select rs
    skip
  enddo

  *----

  select imp_indicateurs
  index on assignation + type_service + str(ligne,3) + direction + periode tag pk
  set order to pk

  *----

  select assignation, type_service, ligne, direction, periode, count(*) as mcnt ;
    from req_impcourses ;
    into cursor rs ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.direction + rs.periode
    replace nt with rs.mcnt
    select rs
    skip
  enddo

  *----

  select assignation, type_service, ligne, direction, periode, count(*) as mcnt ;
    from imputation ;
    into cursor rs ;
   where cod_imp = 'I1' or cod_imp = 'I2' ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.direction + rs.periode
    replace ni with rs.mcnt
    select rs
    skip
  enddo

  *----

  select assignation, type_service, ligne, direction, periode, count(*) as mcnt ;
    from imputation ;
    into cursor rs ;
   where cod_imp = 'NI' ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.direction + rs.periode
    replace nv with rs.mcnt
    select rs
    skip
  enddo

  *----

  select assignation, type_service, ligne, direction, periode, ;
         sum(mont) as totm, sum(desc) as totd ;
    from req_impcourses ;
    into cursor rs ;
   where genere <> 'AGR' ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.direction + rs.periode
    replace mi with rs.totm, di with rs.totd
    select rs
    skip
  enddo

  select assignation, type_service, ligne, direction, periode, ;
         sum(mont) as totm, sum(desc) as totd ;
    from req_impcourses ;
    into cursor rs ;
   group by assignation, type_service, ligne, direction, periode

  select rs
  go top
  do while not eof()
    select imp_indicateurs
    seek rs.assignation + rs.type_service + str(rs.ligne,3) + rs.direction + rs.periode
    replace mt with rs.totm, dt with rs.totd
    select rs
    skip
  enddo

  use
  *----

  select imp_indicateurs
  replace pc_ni with (ni / nt) * 100 for nt > 0
  replace pc_nv with (nv / nt) * 100 for nt > 0
  replace pc_mi with (mi / mt) * 100 for mt > 0
  replace pc_di with (di / dt) * 100 for dt > 0
return

*-----------------------------------------------------------------------------------------
*-----------------------------------------------------------------------------------------

function fCentile
  parameters acenti, centile
  local x, rangmax
  x = asort(acenti)
  rangmax = alen(acenti)
  local valeur1, valeur2, i, k, p_ent_k, p_dec_k
  valeur1 = 0
  valeur2 = 0
  
  for i = rangmax to 1 step -1
    if acenti[i] = NOCENTIL
      rangmax = rangmax - 1
    endif
  next

  for i = 1 to rangmax
    k = 1 + centile * (rangmax - 1)

    p_ent_k = int(k)
    p_dec_k = k - p_ent_k
    if p_ent_k < rangmax
      do case
        case i = p_ent_k
          valeur1 = acenti[i]
        case i = p_ent_k + 1
          valeur2 = acenti[i]
          return valeur1 + p_dec_k * (valeur2 - valeur1)
      endcase
    else
      return acenti[i]
    endif
  next
return -1

*-----------------------------------------------------------------------------------------

function getProfil
  parameters methode

  select imputation
  sdte = imput_date
  datimput = ctod(substr(sdte,1,4)+'.'+substr(sdte,5,2)+'.'+substr(sdte,7,2))

  thecourse = 'x'
  if methode = 'I1'
    declare adiff1[3]
    store 999 to adiff1[1], adiff1[2], adiff1[3]
    for i = 1 to 3
      fdte = 'date_' + iif(i=1, 'a', iif(i=2, 'b', 'c'))
      if not empty(&fdte)
        ddte = ctod(substr(&fdte,1,4)+'.'+substr(&fdte,5,2)+'.'+substr(&fdte,7,2))
        adiff1[i] = abs(ddte - datimput)
      else
        adiff1[i] = 999
      endif
      shre = 'hre_' + iif(i=1, 'a', iif(i=2, 'b', 'c'))
    next
    mindiff = min(adiff1[1], adiff1[2], adiff1[3])
    thei = iif(mindiff = adiff1[1], 1, iif(mindiff = adiff1[2], 2, 3))
    if adiff1[thei] < 999
      thecourse = iif(mindiff = adiff1[1], '_a', iif(mindiff = adiff1[2], '_b', '_c'))
      snbarr = 'nb_arr' + thecourse
      if &snbarr <> nb_arr
        thecourse = 'x'
      endif
    endif
  endif

  if methode = 'I2'

    declare adiff2[2]
    store 999 to adiff2[1], adiff2[2]
    for i = 1 to 2
      fproxy = 'hre_pre_' + iif(i=1, 'av', 'ap')
      if not empty(&fproxy)
        adiff2[i] = abs((hre2sec(&fproxy) - hre2sec(hre_pre_dep)) / 60)
      endif
    next
    mindiff = min(adiff2[1], adiff2[2])
    thei = iif(mindiff = adiff2[1], 1, 2)
    if adiff2[thei] < 999
      thecourse = iif(mindiff = adiff2[1], '_av', '_ap')
      snbarr = 'nb_arr' + thecourse
      *if &snbarr <> nb_arr  && 2006/10/01
      if not between(&snbarr, nb_arr*(1/pmax), nb_arr*pmax)
        thecourse = 'x'
      endif
    endif
  endif

  if thecourse <> 'x'
    replace cod_profil with substr(thecourse, 2)
  endif
return thecourse

*-----------------------------------------------------------------------------------------

procedure putImputCourse
  parameters bIndcourse

  if bIndcourse
    select req_imp
  else
    select req_impcourses
  endif

  locate for assignation  = imputation.assignation  ;
         and type_service = imputation.type_service ;
         and ligne        = imputation.ligne        ;
         and voiture      = imputation.voiture      ;
         and hre_pre_dep  = imputation.hre_pre_dep

  *----
  * déterminer les champs de compilation

  if bIndcourse
    sxpath = "//DefChampSortie/Compilation/Champ"
    onodes = oxmlreq.selectnodes(sxpath)
    for i = 0 to onodes.length - 1
      sfld = onodes.item(i).getattributenode('Nom').value
      if sfld = 'montants' or sfld = 'descendants' or sfld = 'charge_max'
        sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
        onodfld = oxmlclasse.selectsinglenode(sxpath)
        sflddbf = onodfld.getattributenode("OutDBF").value
        sxpath = "./Parametre"
        onodes2 = onodes.item(i).selectnodes(sxpath)
        for j = 0 to onodes2.length - 1
          sop = onodes2.item(j).getattributenode('Nom').value
          iop = ascan(g_agregs, sop)
          sflddbf2 = lower(sflddbf + g_agregs(iop+1))

          if sfld = 'montants'
            replace &sflddbf2 with imputation.m
          endif
          if sfld = 'descendants'
            replace &sflddbf2 with imputation.d
          endif
          if sfld = 'charge_max'
            replace &sflddbf2 with imputation.ch
          endif
        next
      endif
    next
  else
    replace mont with imputation.m
    replace desc with imputation.d
    replace chrg with imputation.ch
  endif

return

*-----------------------------------------------------------------------------------------

procedure putImputArrets
  parameters methode, thecourse

  select imputation
  fldmont = 'm' + thecourse
  fm = m / &fldmont
  flddesc = 'd' + thecourse
  if &fldmont = &flddesc
    fd = fm
  else
    if &flddesc > 0 
      fd = d / &flddesc
    else
      fd = 1
    endif
  endif

  * 2006/12/20; facteur d'ajustement de la charge initiale sur la course
  if &flddesc - &fldmont <> 0 
    k = (d - m) / (&flddesc - &fldmont)
  else
    k = 1
  endif

  if methode = 'I1'
    thedte = 'imputation.date' + thecourse
    thehre = 'imputation.hre' + thecourse

* 2007/10/04; remplacer voiture par direction pour trouver la course correspondantes
*             puique les voitures ne sont plus constantes
*    ssql = "select *" + ;
*           "  from sdap_course_arret" + ;
*           " where sdap_date = '" + &thedte + "'" + ;
*           "   and ligne = " + str(imputation.ligne,3) + ;
*           "   and voiture = '" + imputation.voiture + "'" + ;
*           "   and hre_pre_dep  = '" +  &thehre + "'" + ;
*           " order by position"
    ssql = "select *" + ;
           "  from sdap_course_arret" + ;
           " inner join v_sdap_course" + ; 
           "    on sdap_course_arret.sdap_date = v_sdap_course.sdap_date" + ;
           "   and sdap_course_arret.ligne = v_sdap_course.ligne" + ;
           "   and sdap_course_arret.voiture = v_sdap_course.voiture" + ;
           "   and sdap_course_arret.hre_pre_dep = v_sdap_course.hre_pre_dep" + ;
           " where sdap_course_arret.sdap_date = '" + &thedte + "'" + ;
           "   and sdap_course_arret.ligne = " + str(imputation.ligne,3) + ;
           "   and v_sdap_course.direction = '" + imputation.direction + "'" + ;
           "   and sdap_course_arret.hre_pre_dep  = '" +  &thehre + "'" + ;
           " order by position"

    x = sqlexec(hConn, ssql, 'rscandidat')
    scondition = ssql
  endif
  if methode = 'I2'
    thehre = 'imputation.hre_pre' + thecourse

    select *, &fMont as montants, &fDesc as descendants, &fChrg as charge ;
      from req_imp ;
      into cursor rscandidat ;
     where assignation  = imputation.assignation  ;
       and type_service = imputation.type_service ;
       and ligne        = imputation.ligne        ;
       and direction    = imputation.direction    ;
       and hre_pre_dep  = &thehre            ;
     order by position
     * 2006/12/20; changer la condition sur la voiture par une condition sur la direction
     * la voiture peut avoir changer, mais on ne veut pas une course de l'autre direction
     scondition = 'assignation  = ' + imputation.assignation + ;
                 ' type_service = ' + imputation.type_service + ;
                 ' ligne = ' + str(imputation.ligne,4) + ;
                 ' direction = ' + imputation.direction + ;
                 ' hre_pre_dep = ' + &thehre
  endif

  select rscandidat
  go top
  * 2007/10/04; vérifier que des candidats ont été trouvés
  if eof()
    ? 'Ne peut trouver de candidats pour la condition suivante (methode: ' + methode + '):'
    ? '  ' + scondition
    return
  endif
  
  * 2006/12/20; voir plus haut la note sur k
  lastcharge = charge * k + descendants - montants

  som_mont = 0
  som_desc = 0

  select req_imp
  set filter to assignation  = imputation.assignation  ;
            and type_service = imputation.type_service ;
            and ligne        = imputation.ligne        ;
            and direction    = imputation.direction    ;
            and hre_pre_dep  = imputation.hre_pre_dep
  set order to pklong

  go top
  do while assignation  = imputation.assignation  ;
       and type_service = imputation.type_service ;
       and ligne        = imputation.ligne        ;
       and direction    = imputation.direction    ;
       and hre_pre_dep  = imputation.hre_pre_dep

    select rscandidat

    * 2007/10/04; vérifier que des candidats ont été trouvés
    if eof()
      ? "Le nombre d'arrêts ne correspond pas."
      return
    endif
    
    arr_mont = iif(! isnull(montants), montants * fm, 0)
    arr_desc = iif(! isnull(descendants), descendants * fd, 0)
    arr_chrg = lastcharge + arr_mont - arr_desc
    if arr_chrg < 0
      arr_desc = arr_desc + arr_chrg
      arr_chrg = 0
    endif
    som_mont = som_mont + arr_mont
    som_desc = som_desc + arr_desc
    skip
    if eof()
      * 2006/12/20; balancement non assurée par voyage
      *if som_mont <> som_desc
      *  arr_desc = lastcharge
      *  arr_chrg = 0
      *endif
    endif
    skip -1

    if position <> req_imp.position
      ? 'Arrêt pas à la même position', position, req_imp.position, ;
         position - req_imp.position
      if abs(position - req_imp.position) > 20
        ?? ' ; ATTENTION: différence > 20'
        *suspend
      endif
      *if abs(position - req_imp.position) = 239
      *  suspend
      *endif
    endif

    select req_imp

    *----
    * déterminer les champs de compilation
    sxpath = "//DefChampSortie/Compilation/Champ"
    onodes = oxmlreq.selectnodes(sxpath)
    for i = 0 to onodes.length - 1
      sfld = onodes.item(i).getattributenode('Nom').value
      if sfld = 'montants' or sfld = 'descendants' or sfld = 'charge'
        sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
        onodfld = oxmlclasse.selectsinglenode(sxpath)
        sflddbf = onodfld.getattributenode("OutDBF").value
        sxpath = "./Parametre"
        onodes2 = onodes.item(i).selectnodes(sxpath)
        for j = 0 to onodes2.length - 1
          sop = onodes2.item(j).getattributenode('Nom').value
          iop = ascan(g_agregs, sop)
          sflddbf2 = lower(sflddbf + g_agregs(iop+1))

          if sfld = 'montants'
            replace &sflddbf2 with arr_mont
          endif
          if sfld = 'descendants'
            replace &sflddbf2 with arr_desc
          endif
          if sfld = 'charge'
            replace &sflddbf2 with arr_chrg
          endif
        next
      endif
    next

    lastcharge = arr_chrg
    select rscandidat
    skip
    select req_imp
    skip
  enddo
  
  if .f.
    select rscandidat
    go top
    browse nowait
    select req_imp
    go top
    browse nowait
    suspend
  endif

  ?? fm, fd

  select rscandidat
  use
  select req_imp
  set filter to

   *----

  select sum(&fDesc) as totdesc, max(&fChrg) as maxchrg ;
    from req_imp ;
    into cursor rs ;
   where assignation  = imputation.assignation  ;
     and type_service = imputation.type_service ;
     and ligne        = imputation.ligne        ;
     and direction    = imputation.direction    ;
     and hre_pre_dep  = imputation.hre_pre_dep
  select imputation
  replace d  with rs.totdesc, ;
          ch with rs.maxchrg
  select rs
  use

return

*-----------------------------------------------------------------------------------------
