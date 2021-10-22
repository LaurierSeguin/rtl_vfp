* utils.prg

*-------------------------------------------------------------------------------

function hre2sec
  parameters sheure

  ihre = val(substr(sheure,1,2))
  if ihre < 4 
    ihre = ihre + 24
  endif
  imin = val(substr(sheure,4,2))
  isec = val(substr(sheure,7,2))

return ihre * 3600 + imin * 60 + isec

*-------------------------------------------------------------------------------

function sec2hre
  parameters secs
  
  if secs > 0
    heures = int(100 + secs / 3600)
    minssecs = secs % 3600
    if minssecs > 0
      minutes = 100 + int(minssecs / 60)
      secondes = 100 + minssecs % 60
      return substr(str(heures,3),2,2) + ':' + ;
             substr(str(minutes,3),2,2) + ':' + ;
             substr(str(secondes,3),2,2)
    else
      return substr(str(heures,3),2,2) + ':00:00'
    endif
  endif
return '00:00:00'

*-------------------------------------------------------------------------------

function split
  parameters cTableau, cString, cSeparator

  if empty(cString)
    return 0
  endif
  
  nWords = occurs(cSeparator, cString) + 1
  declare &cTableau[nWords]
  
  nWordNo = 1
  do while nWordNo < nWords
    nPos = at(cSeparator, cString)
    &cTableau[nWordNo] = left(cString, nPos - 1)
    cString = substr(cString, nPos + len(cSeparator))
    nWordNo = nWordNo + 1
  enddo
  &cTableau[nWordNo] = cString
  
return nWords

*-------------------------------------------------------------------------------

procedure savesql
  parameters strsql, fic
  private ssql, crlf, handle, x
  ssql = strsql

  crlf = chr(13) + chr(10)
  ssql = strtran(ssql, ' FROM',  crlf + '  FROM')
  ssql = strtran(ssql, ' INNER', crlf + ' INNER')
  ssql = strtran(ssql, ' ON',    crlf + '    ON')
  ssql = strtran(ssql, ' AND',   crlf + '   AND')
  ssql = strtran(ssql, ' WHERE', crlf + ' WHERE')
  ssql = strtran(ssql,'       ', crlf +'       ') 

  handle = fcreate(fic + '.sql')
  x=fputs(handle, ssql)
  x=fclose(handle)
return

*-----------------------------------------------------------------------------------------

procedure renamefields
  parameters sformat, stable
  bcompil = .f.

  declare agregs[7,3]
  agregs[1,1] = 'Minimum'
  agregs[1,2] = 'mi'
  agregs[1,3] = 'min_'

  agregs[2,1] = 'Maximum'
  agregs[2,2] = 'mx'
  agregs[2,3] = 'max_'

  agregs[3,1] = 'Moyenne'
  agregs[3,2] = 'mo'
  agregs[3,3] = 'moy_'

  agregs[4,1] = 'Total'
  agregs[4,2] = 'to'
  agregs[4,3] = 'tot_'

  agregs[5,1] = 'Nombre'
  agregs[5,2] = 'nb'
  agregs[5,3] = 'occ_'

  agregs[6,1] = 'Écart Type'
  agregs[6,2] = 'et'
  agregs[6,3] = 'ect_'

  agregs[7,1] = 'Percentile'
  agregs[7,2] = 'pc'
  agregs[7,3] = 'pct_' 

  *-----
  
  sxpath = "//DefChampSortie/Compilation/Champ"
  onodes = oxmlreq.selectnodes(sxpath)
  for i = 0 to onodes.length - 1
    bcompil = .t.
    sfld = onodes.item(i).getattributenode('Nom').value
    sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
    onodfld = oxmlclasse.selectsinglenode(sxpath)
    sflddbf = onodfld.getattributenode("OutDBF").value
    sxpath = "./Parametre"
    onodes2 = onodes.item(i).selectnodes(sxpath)
    for j = 0 to onodes2.length - 1
      sop = onodes2.item(j).getattributenode('Nom').value
      iop = ascan(agregs, sop)
      sflddbf2 = lower(sflddbf + agregs(iop+1))
      sfld2 = lower(agregs(iop+2) + sfld)
      do renamefield with sformat, sflddbf2, sfld2, stable
    next
  next
  
  sxpath = "//DefChampSortie/Regroupement/Champ"
  onodes = oxmlreq.selectnodes(sxpath)
  for i = 0 to onodes.length - 1
    sfld = lower(onodes.item(i).getattributenode('Nom').value)
    sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
    onodfld = oxmlclasse.selectsinglenode(sxpath)
    sflddbf = lower(onodfld.getattributenode("OutDBF").value)

    do renamefield with sformat, sflddbf, sfld, stable
  next

  if not bcompil or stable <> 'req_final'
    sxpath = "//DefChampSortie/Complement/Champ"
    onodes = oxmlreq.selectnodes(sxpath)
    for i = 0 to onodes.length - 1
      if not bcompil && Laurier; 20060306; ajouter le IF
        sfld = onodes.item(i).getattributenode('Nom').value
      else
        sfld = lower(onodes.item(i).getattributenode('Nom').value)
      endif
      sxpath = "//DefChampSortie/Champ[@Nom='" + sfld + "']"
      onodfld = oxmlclasse.selectsinglenode(sxpath)
      sflddbf = lower(onodfld.getattributenode("OutDBF").value)
      do renamefield with sformat, sflddbf, sfld, stable
    next
  endif
return

*-------------------

procedure renamefield
  parameters sformat, sflddbf, sfld, stable
  private i

  if lower(sflddbf) <> lower(sfld)
    if sformat = 'foxpro'
      alter table &stable rename column &sflddbf to &sfld
    endif
  endif
  
  if sformat = 'excel'
    i = 1
    do while not empty(oxl.activesheet.cells(1,i).value)
      if oxl.activesheet.cells(1,i).value = sflddbf
        oxl.activesheet.cells(1,i).value = sfld
        exit
      endif
      i = i + 1
    enddo  
  endif
return

*-------------------------------------------------------------------------------

function getperiode
  parameters sheure
  shre = alltrim(str(hre2sec(sheure)))
  * Xpath semble ne pas fonctionner avec >, >=, <, <= sur des chaines
  *   de caractères; alors, on convertit l'heure en secondes
  sxpath = "//Periode[@de <= " + shre + " and @a >= " + shre + "]"
  onode = g_oxmlparam.selectsinglenode(sxpath)
  if not isnull(onode)
    return substr(onode.getattributenode("nom").value, 1, 2)
  endif
return 'xx'

*-------------------------------------------------------------------------------

procedure calc_periode
  select requete0
  go top
  *set filter to isnull(&fsdapdate)

  do while not eof()
    shre = alltrim(str(hre2sec(&fHrepredep)))
    * Xpath semble ne pas fonctionner avec >, >=, <, <= sur des chaines de caractères
    *   alors, on convertit l'heure en secondes
    sxpath = "//Periode[@de <= " + shre + " and @a >= " + shre + "]"
    onode = g_oxmlparam.selectsinglenode(sxpath)
    if not isnull(onode)
      replace periode with substr(onode.getattributenode("nom").value, 1, 2)
    endif
    skip
  enddo

  *set filter to
return

*-----------------------------------------------------------------------------------------
