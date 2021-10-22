* LireOnglets.prg

*-----------------------------------------------------------------------------------------

procedure OngletDates
* trouver les assignations
parameters sbookings, stypser
sdatmin = '99999999'
sdatmax = '00000000'
onode = oxmlreq.selectsinglenode("//Onglet[@Nom='Dates']")

onodes = onode.selectnodes("./Condition[@Type='Dates']")
for i = 0 to onodes.length - 1
  *ssql = iif(i=0, ssql, ssql + ' OR ')
  onode3 = onodes.item(i).selectsinglenode("./Du")
  *ssql = ssql + " (assignation <= '" + onode3.text + "'"
  onode4 = onodes.item(i).selectsinglenode("./Au")
  *ssql = ssql + " and date_fin >= '" + onode4.text + "')"
  if onode3.text < sdatmin
    sdatmin = onode3.text
  endif
  if onode4.text > sdatmax
    sdatmax = onode4.text
  endif

  * dans le cas d'intervalles de dates, on génère une liste de type de services
  * (stypser), qui servira SI l'usager ne choisit pas de type de journées
  ssql2 = "SELECT sdap_course.type_service AS typser, COUNT(*) AS cnt" + ;
          "  FROM sdap_course INNER JOIN ref_course" + ;
          "    ON sdap_course.assignation = ref_course.assignation" + ;
          "   AND sdap_course.type_service = ref_course.type_service" + ;
          "   AND sdap_course.ligne = ref_course.ligne" + ;
          "   AND sdap_course.voiture = ref_course.voiture" + ;
          "   AND sdap_course.hre_pre_dep = ref_course.hre_pre_dep" + ;
          " WHERE sdap_course.sdap_date >= '" + onode3.text + "'" + ;
          "   AND sdap_course.sdap_date <= '" + onode4.text + "'" + ;
          " GROUP BY sdap_course.assignation, sdap_course.type_service"
  x = sqlexec(hconn, ssql2)
  select sqlresult
  go top
  do while not eof()
    if not (typser $ stypser)
      stypser = stypser + "'" + typser + "'"
    endif
    skip
  enddo
  use
next

if onodes.length > 0
  ssql = "select *" + ;
         "  from ref_assignation" + ;
         " where assignation <= '" + sdatmin + "' " + ;
         " order by assignation desc"
  x = sqlexec(hconn, ssql)
  select sqlresult
  if not eof()
    assmin = assignation
  else
    ssql = "select min(assignation) as expr" + ;
           "  from ref_assignation"
    x = sqlexec(hconn, ssql)
    assmin = expr  
  endif

  ssql = "select *" + ;
         "  from ref_assignation" + ;
         " where date_fin >= '" + sdatmax + "' " + ;
         " order by date_fin asc"
  x = sqlexec(hconn, ssql)
  select sqlresult
  if not eof()
    assmax = date_fin
  else
    ssql = "select max(date_fin) as expr" + ;
           "  from ref_assignation"
    x = sqlexec(hconn, ssql)
    assmax = expr
  endif

  ssql = "select *" + ;
         "  from ref_assignation" + ;
         " where assignation >= '" + assmin + "' " + ;
         "   and date_fin <= '" + assmax + "' "
  x = sqlexec(hconn, ssql)
  select sqlresult
  go top
  do while not eof()
    sbookings = iif(empty(sbookings), '', sbookings + ', ')
    sbookings = sbookings + "'" + assignation + "'"
    skip
  enddo
  use
endif

onodes = onode.selectnodes('./Condition[@Type="Période d'+"'"+'assignation"]')
for i = 0 to onodes.length - 1
  sbookings = iif(empty(sbookings), '', sbookings + ', ')
  sbookings = sbookings + "'" + onodes.item(i).text + "'"
next

if not empty(sbookings)
  sbookings = '(a.assignation in (' + sbookings + '))'
endif

*? sbookings, stypser
* on ne traite pas l'option "Derniers Comptages"

return

*-----------------------------------------------------------------------------------------

procedure OngletTypJrs
* trouver les types de journées
parameters stypjrs, stypser
stypjrs = ''
onode = oxmlreq.selectsinglenode("//Onglet[@Nom='Types de journée']")

onodes = onode.selectnodes("./Condition[@Type='Type de service']")
for i = 0 to onodes.length - 1
  stypjrs = iif(empty(stypjrs), '', stypjrs + ', ')
  stypjrs = stypjrs + "'" + onodes.item(i).text + "'"
next
 
onodes = onode.selectnodes("./Condition[@Type='Jour de la semaine']")
for i = 0 to onodes.length - 1
  if onodes.item(i).text $ 'LundiMardiMercrediJeudiVendredi'
    if not ("'SE'" $ stypjrs)
      stypjrs = iif(empty(stypjrs), '', stypjrs + ', ')
      stypjrs = stypjrs + "'SE'"
    endif
  endif
  if onodes.item(i).text = 'Samedi'
    if not ("'SA'" $ stypjrs)
      stypjrs = iif(empty(stypjrs), '', stypjrs + ', ')
      stypjrs = stypjrs + "'SA'"
    endif
  endif      
  if onodes.item(i).text = 'Dimanche'
    if not ("'DI'" $ stypjrs)
      stypjrs = iif(empty(stypjrs), '', stypjrs + ', ')
      stypjrs = stypjrs + "'DI'"
    endif
  endif
next
 
if not empty(stypjrs)
  stypjrs = '(a.type_service in (' + stypjrs + '))'
else
  if not empty(stypser)
    stypjrs = '(a.type_service in (' + stypser + '))'
  endif
endif

return

*-----------------------------------------------------------------------------------------

procedure OngletHeures
* trouver les heures
parameters sheures
sheures = ''
onode = oxmlreq.selectsinglenode("//Onglet[@Nom='Heures']")

declare ahres[2]
ahres[1] = "Heure de départ planifiée" 
ahres[2] = "Heure d'arrivée planifiée"
declare afhres[2]
afhres[1] = 'a.hre_pre_dep' 
afhres[2] = 'ref_course.hre_pre_arr' 
shres = ''
for i = 1 to 2
  onodes = onode.selectnodes('./Condition[@Type="'+ ahres[i] + '"]')
  for j = 0 to onodes.length - 1
    if onodes.item(j).getattributenode("Periode").value = "Heures"
      shres = iif(empty(shres), shres, shres + ' OR ')
      onode3 = onodes.item(j).selectsinglenode("./De")
      shres = shres + "(" + afhres[i] + " >= '" + onode3.text + "'"
      onode4 = onodes.item(j).selectsinglenode("./A")
      shres = shres + " AND " + afhres[i] + " <= '" + onode4.text + "')"
    endif
    if onodes.item(j).getattributenode("Periode").value = "Heure précise"
      shres = iif(empty(shres), shres, shres + ' OR ')
      shres = shres + afhres[i] + " = '" + onodes.item(j).text + "'"
    endif
    if onodes.item(j).getattributenode("Periode").value = "Période standard"
      shres = iif(empty(shres), shres, shres + ' OR ')
      onod = g_oxmlparam.selectsinglenode("//Periode[@nom='"+onodes.item(j).text+"']")
      shres = shres+"("+afhres[i]+" >= '"+onod.getattributenode("hde").value+"'"
      shres = shres+" AND "+afhres[i]+" <= '"+onod.getattributenode("ha").value+"')"
    endif
  next
next

* on ne traite pas les heures réelles pour les courses

if not empty(shres)
  sheures = '(' + shres + ')'  
  
  onode2 = onode.selectsinglenode('./Condition[@Type="Option"]')
  if not isnull(onode2)
    if onode2.text = 'Exclure'
    sheures = '(not ' + sheures + ')'
    endif
  endif
endif

return

*-----------------------------------------------------------------------------------------

procedure OngletLigTra
* trouver les lignes
parameters sligtra
sligtra = ''
onode = oxmlreq.selectsinglenode("//Onglet[@Nom='Lignes / Tracés']")

slignes = ''
onodes = onode.selectnodes("./Condition[@Type='Ligne']")
for i = 0 to onodes.length - 1
  slignes = iif(empty(slignes), '', slignes + ', ')
  slignes = slignes + onodes.item(i).text
next
if not empty(slignes)
  slignes = 'a.ligne in (' + slignes + ')'
endif

onodes = onode.selectnodes("./Condition[@Type='Ligne et direction']")
ld = ''
for i = 0 to onodes.length - 1
  ld = iif(i=0, '', ld + ' OR ')
  slig = substr(onodes.item(i).text, 1, at('-', onodes.item(i).text) - 1)
  sdir = substr(onodes.item(i).text, at('-', onodes.item(i).text) + 1)
  ld = ld + "(a.ligne = " + slig + " AND ref_course.direction = '" + sdir + "')"
next

onodes = onode.selectnodes("./Condition[@Type='Tracé']")
ldt = ''
for i = 0 to onodes.length - 1
  ldt = iif(i=0, '', ldt + ' OR ')
  slig = substr(onodes.item(i).text, 1, at('-', onodes.item(i).text) - 1)
  stmp = substr(onodes.item(i).text, at('-', onodes.item(i).text) + 1)
  sdir = substr(stmp, 1, at('-', stmp) - 1)
  stra = substr(stmp, at('-', stmp) + 1)
  ldt = ldt+"(a.ligne = "+slig+ ;
        " AND ref_course.direction = '"+sdir+"' AND ref_course.trace = "+stra+")"
next

if not (empty(slignes) and empty(ld) and empty(ldt))
  sligtra = '('
  if not empty(slignes)
    sligtra = sligtra + slignes
  endif
  if not empty(ld)
    sligtra = iif(len(sligtra)=1, sligtra + ld, sligtra + ' OR ' + ld)
  endif
  if not empty(ldt)
    sligtra = iif(len(sligtra)=1, sligtra + ldt, sligtra + ' OR ' + ldt)
  endif
  sligtra = sligtra + ')'

  onode2 = onode.selectsinglenode('./Condition[@Type="Option"]')
  if not isnull(onode2)
    if onode2.text = 'Exclure'
    sligtra = '(not ' + sligtra + ')'
    endif
  endif
endif

return

*-----------------------------------------------------------------------------------------

procedure OngletGeoPoints
* trouver les points (et les heures de ces points)
parameters sgeopoints, sclasse
sgeopoints = ''
onode0 = oxmlreq.selectsinglenode("//Onglet[@Nom='Géo. Points']")

if sclasse = 'Arrêts' 
  fpoint = 'a.no_arret'
  fhre = 'a.hre_pre_arret'
endif
if sclasse = 'Points de contrôle'
  fpoint = 'a.lieu'
  fhre = 'a.hre_pre_pc'
endif
if sclasse = 'Balises'
  fpoint = 'no_balise'
  * pas d'heures prévues pour les balises
endif

spoints = ''
onode1 = onode0.selectsinglenode("./Points")
if not isnull(onode1)
  onodes = onode1.selectnodes("./Condition")
  for i = 0 to onodes.length - 1
    spoints = iif(empty(spoints), '', spoints + ', ')
    if fpoint = 'a.lieu'
      spoints = spoints + "'" + onodes.item(i).text + "'"
    else
      spoints = spoints + onodes.item(i).text
    endif
  next
  if not empty(spoints)
    sgeopoints = fpoint + ' in (' + spoints + ')'
  endif
endif

shres = ''
onode1 = onode0.selectsinglenode("./Heures")
if not isnull(onode1)
  onodes = onode1.selectnodes("./Condition[@Type='Heure planifiée']")
  for j = 0 to onodes.length - 1
    if onodes.item(j).getattributenode("Periode").value = "Heures"
      shres = iif(empty(shres), shres, shres + ' OR ')
      onode3 = onodes.item(j).selectsinglenode("./De")
      shres = shres + "(" + fhre + " >= '" + onode3.text + "'"
      onode4 = onodes.item(j).selectsinglenode("./A")
      shres = shres + " AND " + fhre + " <= '" + onode4.text + "')"
    endif
    if onodes.item(j).getattributenode("Periode").value = "Heure précise"
      shres = iif(empty(shres), shres, shres + ' OR ')
      shres = shres + fhre + " = '" + onodes.item(j).text + "'"
    endif
  next
  if not empty(shres)
    if not empty(sgeopoints)
      sgeopoints = sgeopoints + ' AND (' + shres + ')'
    else
      sgeopoints = shres
    endif
  endif
  * on ne traite pas les heures réelles pour les points 
endif

if not empty(sgeopoints)
  sgeopoints = '(' + sgeopoints + ')'

  onode2 = onode0.selectsinglenode('./Condition[@Type="Option"]')
  if not isnull(onode2)
    if onode2.text = 'Exclure'
    sgeopoints = '(not ' + sgeopoints + ')'
    endif
  endif
endif

return

*-----------------------------------------------------------------------------------------

procedure OngletConditions
* trouver les conditions logiques
* schamps servira à définir à faire la clause from
parameters sconditions, schamps
sconditions = ''
onode = oxmlreq.selectsinglenode("//Onglet[@Nom='Conditions logiques']")

onodes = onode.childnodes
for i = 0 to onodes.length - 1
  condition = onodes.item(i).text

  * on ne traite que les conditions provenant de la facette référence
  if lower(substr(condition, 1, 4)) = 'ref_'
    sconditions = iif(empty(sconditions), condition, sconditions + ' AND ' + condition)

    champ = substr(condition, 1, at(' ', condition) - 1) 
    if not champ $ schamps
      schamps = iif(empty(schamps), champ, schamps + ' ' + champ)
    endif
  endif
next

if not empty(sconditions)
  sconditions = '(' + sconditions + ')'
endif

return

*-----------------------------------------------------------------------------------------
