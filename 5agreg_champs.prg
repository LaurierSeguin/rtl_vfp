* agreg_champs.prg
*-------------------------------------------------------------------------------
parameters stypagreg, oxmlreq, oxmlclasse
store '' to sdbfsrc, sdbfout

close databases all
?

onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='Classe']")
sclasse = onode.text

agregoper = ''
if stypagreg = 'COURSES'
  ? "Agrégation des observations"

  soper = "Agrégation des observations"
  sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
  onode = oXMLreq.SelectSingleNode(sxpath)
  agregoper = onode.text
  if substr(agregoper,1,7) = 'Centile'
    posdeb = at('[', agregoper) + 1
    posfin = at(']', agregoper) - 1
    centil = val(substr(agregoper, posdeb, posfin - posdeb + 1)) / 100
    agregoper = 'Centile'
  endif
  
  sxpath = "//DefChampSortie/Champ[@Nom='ligne']"
  onode = oxmlclasse.selectsinglenode(sxpath)
  fLigne = onode.getattributenode("OutDBF").value
  sxpath = "//DefChampSortie/Champ[@Nom='direction']"
  onode = oxmlclasse.selectsinglenode(sxpath)
  fDir = onode.getattributenode("OutDBF").value
  sxpath = "//DefChampSortie/Champ[@Nom='hre_pre_dep']"
  onode = oxmlclasse.selectsinglenode(sxpath)
  fHrepredep = onode.getattributenode("OutDBF").value
  if sclasse <> 'Courses' and sclasse <> 'Tronçons'
    sxpath = "//DefChampSortie/Champ[@Nom='position']"
    onode = oxmlclasse.selectsinglenode(sxpath)
    fPos = onode.getattributenode("OutDBF").value
  endif

  open database bd_requete
  nb = adbobjects(atables, "TABLE")
  if nb > 0
    stables2 = "REQ_AGR"
    declare atables2[1]
    x = split('atables2', stables2, ',')
    for i = 1 to x
      if ascan(atables, atables2[i]) > 0
        remove table atables2[i] delete
      endif
    next
  endif
  
else
  ? "Agrégation des champs"
endif

do findinputoutput
? "  input: " + sdbfsrc + ".dbf  output: " + sdbfout + ".dbf"

select 0
use &sdbfsrc

onode = oxmlreq.selectsinglenode("//DefChampSortie/Compilation")
if isnull(onode)
  copy to &sdbfout
  *use &sdbfout
  return
endif

*----
* déterminer les champs de regroupement
* créer un champ de regroupement, le remplir
sgroup = ''
if stypagreg = 'COURSES'
  sniveaux = ;
      'assignation,type_service,ligne,voiture,hre_pre_dep,periode,direction'
  if sclasse <> 'Courses' and sclasse <> 'Tronçons'
    sniveaux = sniveaux + ',position'
  endif

  onode = oxmlreq.selectsinglenode("//DefChampSortie/Regroupement")
  if not isnull(onode)
    onodes = onode.childnodes
    for i = 0 to onodes.length - 1
      sfld = onodes.item(i).getattributenode('Nom').value
      sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
      onodfld = oxmlclasse.selectsinglenode(sxpath)
      *sflddbf = onodfld.getattributenode("OutDBF").value
      if not (sfld $ sniveaux)
        sniveaux = sniveaux + ',' + sfld
      endif
    next
  endif

  select &sdbfsrc
  declare aniveaux[1]
  xn = split('aniveaux', sniveaux, ',')
  xf = afields(aflds)
  bok = .t.
  for i = 1 to xn
    sxpath = "//DefChampSortie/Champ[@Nom='" + aniveaux[i] + "']"
    onodfld = oxmlclasse.selectsinglenode(sxpath)
    flddbf = onodfld.getattributenode("OutDBF").value
    if ascan(aflds, upper(flddbf)) = 0
      messagebox("Le champ " + aniveaux[i] + " (" + flddbf + ;
                 ") n'existe pas dans la table " + sdbfsrc + ".dbf")
      bok = .f.
    else
      sgroup = iif(i=1, '', sgroup + ' ')
      sgroup = sgroup + upper(flddbf)
    endif
  next
  if not bok
    g_bOK = .f.
    return
  endif

else
  onode = oxmlreq.selectsinglenode("//DefChampSortie/Regroupement")
  if not isnull(onode)
    onodes = onode.childnodes
    for i = 0 to onodes.length - 1
      sfld = onodes.item(i).getattributenode('Nom').value
      sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
      onodfld = oxmlclasse.selectsinglenode(sxpath)
      sflddbf = onodfld.getattributenode("OutDBF").value
      sgroup = iif(i=0, '', sgroup + ' ')
      sgroup = sgroup + upper(sflddbf)
    next
  endif
endif
bakexact = set('exact')
set exact on && pour les ascan()

select &sdbfsrc
x = afields(aflds)
declare agroup[1]
nwords = split('agroup', sgroup, ' ')
sindxpr = ''
for i = 1 to nwords
  sindxpr = sindxpr + iif(i=1, '' , '+')
  irow = asubscript(aflds, ascan(aflds, agroup[i]), 1)
  if aflds(irow, 2) = 'C' 
    sindxpr = sindxpr + aflds[irow,1]
  endif
  if aflds(irow, 2) $ 'NFIB'
    sindxpr = sindxpr+'STR('+aflds[irow,1]+','+ltrim(str(aflds[irow,3]))+')'
  endif
next
*? sindxpr

* 20060220: ajouter le IF
if ascan(aflds, 'GRP') = 0
  alter table &sdbfsrc add column grp i
endif
* 20080602: ajouter le IF
if ascan(aflds, 'RECNUM') = 0
  alter table &sdbfsrc add column recnum i
endif
* 20080602: ajouter le IF
if ascan(aflds, 'INDXPR') = 0
  alter table &sdbfsrc add column indxpr c(100)
endif

replace all indxpr with &sindxpr
if not empty(sindxpr)
  index on indxpr tag fldsgrp
  reindex
  set order to fldsgrp
  bakgrp = 'xxx'
  i = 0
  go top 
  do while not eof()
    if &sindxpr <> bakgrp
      bakgrp = &sindxpr
      i = i + 1
    endif
    replace grp with i
    skip
  enddo
else
  replace all grp with 1
endif

*----
* déterminer les champs de compilation
scomps = ''
onodes = oxmlreq.selectsinglenode("//DefChampSortie/Compilation").childnodes
for i = 0 to onodes.length - 1
  sfld = onodes.item(i).getattributenode('Nom').value
  sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
  onodfld = oxmlclasse.selectsinglenode(sxpath)
  sflddbf = onodfld.getattributenode("OutDBF").value

  if stypagreg = 'COURSES'
    if agregoper = 'Moyenne' or agregoper = 'Maximum'
      soper = iif(agregoper = 'Moyenne', 'avg', 'max')
    else  && Centile
      soper = 'count'
    endif
    onodes2 = onodes.item(i).childnodes
    for j = 0 to onodes2.length - 1
      sop = onodes2.item(j).getattributenode('Nom').value
      iop = ascan(g_agregs, sop)
      sfld2 = sflddbf + g_agregs(iop+1)
      sfld3 = soper + '(' + sfld2 + ') as ' + sfld2
      scomps = scomps + ', ' + sfld3
    next
  else
    onodes2 = onodes.item(i).childnodes
    for j = 0 to onodes2.length - 1
      sop = onodes2.item(j).getattributenode('Nom').value
      iop = ascan(g_agregs, sop)
      sfld2 = sflddbf + g_agregs(iop+1)
      sfld3 = g_agregs(iop+2) + '(' + sfld2 + ') as ' + sfld2
      scomps = scomps + ', ' + sfld3
    next
  endif
next

if stypagreg = 'COURSES'
  scomps = scomps + ', count(*) as nb_courses'
endif

soper = "Pondération des données"
sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
oNode = oXMLreq.SelectSingleNode(sxpath)
if not isnull(oNode) AND ascan(aflds, 'FG') > 0
  if stypagreg = 'COURSES'
    if agregoper = 'Moyenne' or agregoper = 'Maximum'
      soper = iif(agregoper = 'Moyenne', 'avg', 'max')
    else  && Centile
      soper = 'count'
    endif
    scomps = scomps + ', ' + soper + '(fg) as fg'
  else
    scomps = scomps + ', sum(fg) as tot_fg'
  endif
endif
set exact off && pour les seek

*----
* exécuter la requête
sgroup2 = lower(strtran(sgroup, ' ', ', '))
sgroup2 = 'grp' + iif(not empty(sgroup2), ', ' + sgroup2, '')
ssql = 'SELECT ' + sgroup2 + scomps + ;
       ' INTO dbf ' + sdbfout + ' FROM ' + sdbfsrc + ' GROUP BY ' + sgroup2
*? '  ' + ssql
do savesql with ssql, 'agregation'
&ssql       

*----
* calculer l'écart-type (s'il y a lieu) 
sxpath = "//DefChampSortie/Compilation/Champ/Parametre[@Nom='Écart Type']"
onodes = oxmlreq.selectnodes(sxpath)
if onodes.length > 0
  select &sdbfsrc
  index on grp tag idxgrp
  set order to idxgrp

  for i = 0 to onodes.length - 1
    sfld = onodes.item(i).parentnode.getattributenode('Nom').value
    sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
    onodfld = oxmlclasse.selectsinglenode(sxpath)
    sflddbf = onodfld.getattributenode("OutDBF").value
    do ecarttype with sflddbf + 'et'
  next
endif

*----
* calculer le percentile (s'il y a lieu) 
if stypagreg = 'COURSES'
  if agregoper = 'Centile'
  
    sxpath = "//DefChampSortie/Compilation/Champ"
    onodes = oxmlreq.selectnodes(sxpath)
    if onodes.length > 0
      alter table &sdbfsrc add column rang_i i
      alter table &sdbfout add column rang_max i

      for i = 0 to onodes.length - 1
        sfld = onodes.item(i).getattributenode('Nom').value
        sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
        onodfld = oxmlclasse.selectsinglenode(sxpath)
        sflddbf = onodfld.getattributenode("OutDBF").value
        
        onodes2 = onodes.item(i).childnodes
        for j = 0 to onodes2.length - 1
          sop = onodes2.item(j).getattributenode('Nom').value
          iop = ascan(g_agregs, sop)
          sfld2 = sflddbf + g_agregs(iop+1)
            do percentile with sfld2, centil
        next        
      next

      if not isnull(oNode) AND ascan(aflds, 'FG') > 0
        do percentile with 'fg', centil
      endif

      * pour avoir le détail, ne pas détruire ces champs
      *alter table &sdbfsrc drop column rang_i
      *alter table &sdbfout drop column rang_max
    endif
  endif
else
  sxpath = "//DefChampSortie/Compilation/Champ/Parametre[@Nom='Percentile']"
  onodes = oxmlreq.selectnodes(sxpath)
  if onodes.length > 0
    alter table &sdbfsrc add column rang_i i
    alter table &sdbfout add column rang_max i

    for i = 0 to onodes.length - 1
      centil = val(onodes.item(i).text) / 100
      sfld = onodes.item(i).parentnode.getattributenode('Nom').value
      sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
      onodfld = oxmlclasse.selectsinglenode(sxpath)
      sflddbf = onodfld.getattributenode("OutDBF").value
      do percentile with sflddbf + 'pc', centil
    next

    * pour avoir le détail, ne pas détruire ces champs
    *alter table &sdbfsrc drop column rang_i
    *alter table &sdbfout drop column rang_max
    
  endif
endif

*----
*alter table &sdbfout drop column grp

select &sdbfsrc
on error do ignore
*delete tag idxpcentil
delete tag idxgrp
on error
*alter table &sdbfsrc drop column grp

select &sdbfsrc
use
select &sdbfout
use

open database bd_requete
if sdbfsrc = 'req_agr'
  add table req_agr
endif

set exact &bakexact
return

procedure ignore
return

*-------------------------------------------------------------------------------

procedure findinputoutput
  * déterminer les table input et output

  if stypagreg = 'COURSES'
    sdbfsrc = 'requete0'
    sdbfout = 'req_agr'
  else
    sdbfsrc = 'requete0'
    sdbfout = 'tmp_final'
    
    soper = "Agrégation des observations"
    sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
    oNode = oXMLreq.SelectSingleNode(sxpath)
    if not isnull(oNode)
      sdbfsrc = 'req_agr'  
    endif

    soper = "Imputation de données"
    sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
    oNode = oXMLreq.SelectSingleNode(sxpath)
    if not isnull(oNode)
      sdbfsrc = 'req_imp'  
    endif
  endif  
return

*-------------------------------------------------------------------------------

procedure ecarttype
  * calculer l'écart-type pour des enregistrements à regrouper
  parameters sflddbf
  private mavg, mcnt, mstddev
  
  select &sdbfsrc
  set filter to not isnull(&sflddbf)
  
  alter table &sdbfout alter column &sflddbf n(16, 2) null
  select &sdbfout
  replace all &sflddbf with 0     && le calcul fait dans le SELECT est dummy
  
  go top
  do while not eof()
    mstddev = .null.
    select &sdbfsrc
    seekexpr = sdbfout+'.grp'
    seek str(&seekexpr, 10)
    if not eof()
      calculate avg(&sflddbf), cnt() ;
          while grp = &seekexpr ;
             to mavg, mcnt
      if mcnt > 1
        mstddev = 0
        seek str(&seekexpr, 10)
        do while grp = &seekexpr
          mstddev = mstddev + (&sflddbf - mavg) ** 2
          skip
        enddo
        mstddev = (mstddev / (mcnt - 1 )) ** 0.5
      endif
      if mcnt = 1
        mstddev = 0
      endif
    endif
    select &sdbfout
    replace &sflddbf with mstddev
    skip
  enddo
  
  select &sdbfsrc
  set filter to
return

*-------------------------------------------------------------------------------

procedure percentile
  * calculer le percentile pour des enregistrements à regrouper
  parameters sflddbf, centil
  private bakgrp, i, percentil
  *? sflddbf, centil
          
  *----
  * grp, &sflddbf et rang_i sont des champs de &sdbfsrc
  select &sdbfsrc
  replace all recnum with recno()
  replace all rang_i with 0
  sort to tmp_centile on grp, &sflddbf
  index on recnum tag recnum
  reindex
  set order to recnum
  
  select 0
  use tmp_centile
  set filter to not isnull(&sflddbf)
  set relation to recnum into &sdbfsrc
  bakgrp = 0
  i = 0
  go top 
  do while not eof()
    if grp <> bakgrp
      bakgrp = grp
      i = 1
    else
      i = i + 1
    endif
    replace rang_i with i in &sdbfsrc
    skip
  enddo
  use

  select &sdbfsrc
  index on str(grp, 10) + str(rang_i, 10) tag idxpcentil
  reindex
  set order to idxpcentil
  
  *----
  * grp, &sflddbf et rang_max sont des champs de &sdbfout
  
  select &sdbfout
  replace all rang_max with &sflddbf, &sflddbf with 0
  alter table &sdbfout alter column &sflddbf n(16, 2) null
  go top
  do while not eof()
    percentil = f_percentile(sflddbf, centil, grp, rang_max)
    select &sdbfout
    replace &sflddbf with percentil
    skip
  enddo
  
  select &sdbfsrc
  set filter to
return

*---------------------------------------

function f_percentile
  * grp, rang_i et &sflddbf sont des champs de &sdbfsrc
  parameters sflddbf, centi, mgrp, rangmax
  private valeur1, valeur2, k, p_ent_k, p_dec_k
  valeur1 = 0
  valeur2 = 0

  select &sdbfsrc
  seek str(mgrp, 10)
  do while grp = mgrp
    k = 1 + centi * (rangmax - 1)
    p_ent_k = int(k)
    p_dec_k = k - p_ent_k
    if p_ent_k < rangmax
      do case
        case rang_i = p_ent_k
          valeur1 = &sflddbf
        case rang_i = p_ent_k + 1
          valeur2 = &sflddbf
          return valeur1 + p_dec_k * (valeur2 - valeur1)
      endcase
    ELSE
      IF STR(&sflddbf) = '**********'
        * Laurier 20111222
        * un overflow mets la valeur '**********'; on la remplace par .null.
        * pourquoi l'overflow: probablement un calcul sur des valeurs nulles
        RETURN .null.
      endif
      return &sflddbf
    endif
    skip
  enddo
return .null.

*-------------------------------------------------------------------------------
