* calcul_base.prg
*-----------------------------------------------------------------------------------------
parameters sclasse, oxmlclasse
?
? "Calcul des champs du fichier de base"
? "  input: requete0.dbf  output: requete0.dbf"

close databases all
select 0
use requete0

nbchamps = afields(achamps)
for i = 1 to nbchamps

  select requete0
  go top

  schamp = upper(achamps(i,1))
  
  *onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[@OutDBF='"+schamp+"']")
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
    onode = oxmlclasse.selectsinglenode("//DefChampRequete/DefChamp[@Nom='"+snomchamp+"']")

    if not isnull(onode)

      for j = 0 to onode.childnodes.length - 1
        sj = ltrim(str(j+1))
        schamp&sj = onode.childnodes.item(j).getattributenode("OutDBF").value
      next

      * ----------------------------------------------------------------------------------

      if sclasse = 'Courses'
        do case 
          case snomchamp = 'tmps_tot_reel'
            alter table requete0 alter &schamp b(2) null
            do while not eof()
              if not isnull(&schamp3)
                replace &schamp with (hre2sec(&schamp1) - hre2sec(&schamp2)) + &schamp3
              else
                replace &schamp with (hre2sec(&schamp1) - hre2sec(&schamp2))
              endif
              skip
            enddo
          case snomchamp == 'tmps_par_reel'
            replace all &schamp with hre2sec(&schamp1) - hre2sec(&schamp2)
          case snomchamp == 'tmps_par_reel_hp'
            replace all &schamp with (hre2sec(&schamp1) - hre2sec(&schamp2)) - &schamp3
          case snomchamp = 'tmps_lib'
            replace all &schamp with (hre2sec(&schamp1) - hre2sec(&schamp2)) - &schamp3
          case snomchamp = 'dist_reel'
            replace all &schamp with &schamp1 - &schamp2
          case snomchamp = 'delta_bat'
            alter table requete0 alter &schamp b
            replace all &schamp with &schamp1 - &schamp2
          case snomchamp = 'cla_delta_temps'
            do while not eof()
              replace &schamp with iif(dl_t < 0, -1, iif(dl_t = 0, 0, 1))
              skip
            enddo
          case snomchamp = 'cla_delta_bat'
            do while not eof()
              delta = &schamp1 - &schamp2
              replace &schamp with iif(delta<0, -1, iif(delta=0, 0, 1))
              skip
            enddo
          case snomchamp = 'ponc_dep' or snomchamp = 'ponc_arr'
            replace all &schamp with hre2sec(&schamp1) - hre2sec(&schamp2)
          case snomchamp = 'cla_ponc_dep' or snomchamp = 'cla_ponc_arr'
            do while not eof()
              ponc = ltrim(str((hre2sec(&schamp1) - hre2sec(&schamp2)) / 60, 9, 2))              
              sxpath = "//ClassesPonctualite/Classe[@min<="+ponc+" and "+ponc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              replace &schamp with val(onode.getattributenode("valeur").value)
              skip
            enddo
          case snomchamp = 'grp_ponc_dep' or snomchamp = 'grp_ponc_arr'
            alter table requete0 alter &schamp c(1)
            do while not eof()
              ponc = ltrim(str((hre2sec(&schamp1) - hre2sec(&schamp2)) / 60, 9, 2))              
              sxpath = "//ClassesPonctualite/Classe[@min<="+ponc+" and "+ponc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              clpnc = onode.getattributenode("valeur").value
              sxpath = "//GroupesPonctualite/Groupe[@min<="+clpnc+" and "+clpnc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              grpponc = onode.getattributenode("valeur").value
              replace &schamp with grpponc
              skip
            enddo

          case snomchamp = 'tmps_tot_pre'
            replace all &schamp with ;
                (hre2sec(&schamp1) - hre2sec(&schamp2)) + (&schamp3 * 60)
          case snomchamp = 'tmps_par_pre'
            replace all &schamp with hre2sec(&schamp1) - hre2sec(&schamp2)
          case snomchamp = 'dist_pre'
            replace all &schamp with &schamp1 - &schamp2
        endcase
      endif

      * ----------------------------------------------------------------------------------

      if sclasse = 'Points de contrôle' or sclasse = 'Arrêts' or sclasse = 'Balises'
        do case 
          case snomchamp = 'delta_temps'
            replace all &schamp with &schamp1 - (hre2sec(&schamp2) - hre2sec(&schamp3))
          case snomchamp = 'delta_dist'
            replace all &schamp with &schamp1 - (&schamp2 - &schamp3)
          case snomchamp = 'cla_delta_temps'
            do while not eof()
              delta = &schamp1 - (hre2sec(&schamp2) - hre2sec(&schamp3))
              replace &schamp with iif(delta<0, -1, iif(delta=0, 0, 1))
              skip
            enddo
          case snomchamp = 'cla_ponc'
            do while not eof()
              *sponc = alltrim(str(ponc))
              *if not isnull(ponc)
                sponc = alltrim(str(ponc/60, 9, 2))
                sxpath = "//ClassesPonctualite/Classe[@min<="+sponc+" and "+sponc+"<@max]"
                onode = g_oxmlparam.selectsinglenode(sxpath)
                replace &schamp with val(onode.getattributenode("valeur").value)
              *endif
              skip
            enddo
          case snomchamp = 'grp_ponc'
            alter table requete0 alter &schamp c(1)
            do while not eof()
              *sponc = alltrim(str(ponc))
              sponc = alltrim(str(ponc/60, 9, 2))
              sxpath = "//ClassesPonctualite/Classe[@min<="+sponc+" and "+sponc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              clpnc = onode.getattributenode("valeur").value
              sxpath = "//GroupesPonctualite/Groupe[@min<="+clpnc+" and "+clpnc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              grpponc = onode.getattributenode("valeur").value
              replace &schamp with grpponc
              skip
            enddo
          case snomchamp = 'tmps_par_pre_ech'
            replace all &schamp with hre2sec(&schamp1) - hre2sec(&schamp2)
          case snomchamp = 'dist_pre_ech'
            replace all &schamp with &schamp1 - &schamp2
        endcase
      endif

      * ----------------------------------------------------------------------------------

      if sclasse = 'Points de contrôle'
        do case 
        * ajout NT 2007-02-23 ajout du calcul de temps parcours arrondi
          case snomchamp = 'tmps_par_reel_ent'
              replace all &schamp with ROUND(&schamp1/60,0)
            case snomchamp = 'tmps_cum_reel_ent'
              replace all &schamp with ROUND(&schamp1,0)
            case snomchamp = 'tp_cum_ent'
              replace all &schamp with ROUND((&schamp1 + &schamp2/60),0)
        endcase
      endif

      * ----------------------------------------------------------------------------------

      if sclasse = 'Arrêts'
        do case 
          case snomchamp = 'charge_arr'
            replace all &schamp with &schamp1 + &schamp2 - &schamp3
          case snomchamp = 'pass_min_arr'
            replace all &schamp with (&schamp1 + &schamp2 - &schamp3) * &schamp4
          case snomchamp = 'pass_min_dep'
            replace all &schamp with &schamp1 * &schamp2
          case snomchamp = 'pass_min'
            replace all &schamp with &schamp1 * &schamp2 / 60
          case snomchamp = 'pass_m'
            replace all &schamp with &schamp1 * &schamp2
        endcase
      endif

      * ----------------------------------------------------------------------------------

      if sclasse = 'Tronçons'
        do case 
          case snomchamp == 'tmps_par_reel'
            replace all &schamp with hre2sec(&schamp2) - hre2sec(&schamp1)
          case snomchamp == 'tmps_par_reel_hp'
            replace all &schamp with (hre2sec(&schamp2) - hre2sec(&schamp1)) - &schamp3
          case snomchamp == 'dist_reel'
            replace all &schamp with &schamp2 - &schamp1
          case snomchamp == 'delta_temps'
            replace all &schamp with (hre2sec(&schamp2) - hre2sec(&schamp1)) - ;
                                     (hre2sec(&schamp4) - hre2sec(&schamp3))
          case snomchamp == 'delta_dist'
            replace all &schamp with (&schamp2 - &schamp1) - (&schamp4 - &schamp3)
          case snomchamp == 'cla_delta_temps'
            do while not eof()
              delta = (hre2sec(&schamp2) - hre2sec(&schamp1)) - ;
                      (hre2sec(&schamp4) - hre2sec(&schamp3))
              replace &schamp with iif(delta<0, -1, iif(delta=0, 0, 1))
              skip
            enddo
          case snomchamp == 'cla_ponc_deb' or snomchamp == 'cla_ponc_fin'
            do while not eof()
              ponc = ltrim(str(&schamp1 / 60, 9, 2))
              sxpath = "//ClassesPonctualite/Classe[@min<="+ponc+" and "+ponc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              replace &schamp with val(onode.getattributenode("valeur").value)
              skip
            enddo
          case snomchamp == 'grp_ponc_deb' or snomchamp == 'grp_ponc_fin'
            alter table requete0 alter &schamp c(1)
            do while not eof()
              ponc = ltrim(str(&schamp1 / 60, 9, 2))
              sxpath = "//ClassesPonctualite/Classe[@min<="+ponc+" and "+ponc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              clpnc = onode.getattributenode("valeur").value
              sxpath = "//GroupesPonctualite/Groupe[@min<="+clpnc+" and "+clpnc+"<@max]"
              onode = g_oxmlparam.selectsinglenode(sxpath)
              grpponc = onode.getattributenode("valeur").value
              replace &schamp with grpponc
              skip
            enddo
          case snomchamp == 'tmps_lib'
            replace all &schamp with (hre2sec(&schamp2) - hre2sec(&schamp1)) - &schamp3

          case snomchamp == 'tmps_par_pre'
            replace all &schamp with hre2sec(&schamp2) - hre2sec(&schamp1)

          case snomchamp == 'delta_pos_deb'
            replace all &schamp with &schamp2 - &schamp1
          case snomchamp == 'delta_pos_fin'
            replace all &schamp with &schamp2 - &schamp1
          case snomchamp == 'dist_pre'
            replace all &schamp with &schamp2 - &schamp1
        endcase
      endif

      * ----------------------------------------------------------------------------------
      * pour toutes les classes

      do case 
        case snomchamp = 'mois'
          replace all &schamp with val(substr(&schamp1, 5, 2))
        case snomchamp = 'semaine'
          do while not eof()
            sdate = substr(datecab,1,4)+'.'+substr(datecab,5,2)+'.'+substr(datecab,7,2)
            ddate = ctod(sdate)
            replace &schamp with week(ddate)
            skip
          enddo
        case snomchamp = 'jour_sem'
          alter table requete0 alter &schamp c(1)
          do while not eof()
            sdate = substr(datecab,1,4)+'.'+substr(datecab,5,2)+'.'+substr(datecab,7,2)
            ddate = ctod(sdate)
            jrsem = ''
            do case
              case dow(ddate, 2) = 1
                jrsem = 'L'
              case dow(ddate, 2) = 2
                jrsem = 'M'
              case dow(ddate, 2) = 3
                jrsem = 'W'
              case dow(ddate, 2) = 4
                jrsem = 'J'
              case dow(ddate, 2) = 5
                jrsem = 'V'
              case dow(ddate, 2) = 6
                jrsem = 'S'
              case dow(ddate, 2) = 7
                jrsem = 'D'
            endcase
            replace &schamp with jrsem
            skip
          enddo
        case snomchamp = 'per_usag'
          * à faire
        case snomchamp = 'heure'
          alter table requete0 alter &schamp c(8)
          replace all &schamp with substr(&schamp1, 1, 2) + ':00:00'
        case snomchamp = 'dheure'
          alter table requete0 alter &schamp c(8)
          do while not eof()
            hre = substr(&schamp1, 1, 3)
            min = substr(&schamp1, 4, 2)
            if min <= '29'
              dhre = hre + '00:00'
            else  
              dhre = hre + '30:00'
            endif              
            replace &schamp with dhre
            skip
          enddo
        case snomchamp = 'qheure'
          alter table requete0 alter &schamp c(8)
          do while not eof()
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
            replace &schamp with qhre
            skip
          enddo
        case snomchamp = 'grp_lig_usag'
          * à faire
      endcase

      * ----------------------------------------------------------------------------------

    endif

  endif
next

for i = 1 to nbchamps
  schamp = upper(achamps(i,1))
  onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
          "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+schamp+"']")

  if isnull(onode) and substr(schamp, len(schamp) - 1) $ 'MIMXMOTOETPC'
    * peut-être un champ à compiler
    xchamp = substr(schamp, 1, len(schamp) - 2)
    onode = oxmlclasse.selectsinglenode("//DefChampSortie/Champ[translate(@OutDBF,"+ ;
            "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"+xchamp+"']")
  endif

  if not isnull(onode)
    snomchamp = onode.getattributenode("Nom").value
    do sec2min with snomchamp, schamp
  endif
next

return

*-----------------------------------------------------------------------------------------

procedure sec2min
  parameters snomchamp, schamp
  private sxpath, onode

  sxpath = "//ChampsUnites/Minutes/" + snomchamp + "[@sec2min='O']"
  onode = g_oxmlparam.selectsinglenode(sxpath)
  if not isnull(onode)
    alter table requete0 alter &schamp b(2) null
    update requete0 set &schamp = &schamp / 60 where not isnull(&schamp)
  endif
return

*-----------------------------------------------------------------------------------------
