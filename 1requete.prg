* "%vfp%" %stad_home%\source_prg\requete.prg REQDIR UID PWD
*   REQDIR est le r�pertoire de sauvegarde choisi par l'usager
*     exemple: %stad_home%\requetes\requete1
*   UID est le nom de l'usager SQL Server
*   PWD est son mot de passe
* DO j:\stad\source_prg\requete with 'J:\STAD\usagers\S�guin\resultats\test2', "x", "x"
* DO d:\projets\stad_home\source_prg\requete with 'd:\projets\stad_home\usagers\S�guin\resultats\test2', "x", "x"
* DO j:\stad\source_prg\requete with 'J:\STAD\usagers\S�guin\test_troncons_bon\', "stad", "stad"
*-----------------------------------------------------------------------------------------
PARAMETERS dirreq, uid, pwd, bBat
PUBLIC oXMLini, g_bOK, g_oxmlparam
oXMLini = .f.
g_bOK = .t.

CLEAR
? "Initialisation du traitement FOXPRO"

PUBLIC g_bDev
g_bDev = .f.
IF FILE('d:\laurier\isstad.txt')
*  g_bDev = .t.
ENDIF

CLOSE ALL
SET NULL ON
SET TALK OFF
SET DATE ANSI
SET CENTURY ON
SET SAFETY OFF
*SET PATH TO GETENV('STAD_HOME') + '\source_prg'
programpath = substr(sys(16), 1, rat('\', sys(16)) - 1)
SET PATH TO &programpath
SET PROCEDURE TO 'utils2.prg'

IF validparam(PCOUNT())
  SET DEFAULT TO (dirreq)
  SET ALTERNATE TO requete.log
  SET ALTERNATE ON
  ? '  D�but: ' + TIME()

  DO process

  ?
  ? '  Fin: ' + TIME()
  IF g_bOK
    ? 'Fin normale de requete.prg'
    ?
    * QUIT
  ELSE
    ? 'Fin ANORMALE de requete.prg'
    ?
  ENDIF
  CLOSE ALTERNATE
ENDIF

RETURN

*-----------------------------------------------------------------------------------------

FUNCTION validparam
  PARAMETERS pcnt

  *IF pcnt <> 3
  *  MESSAGEBOX("Usage: requete.prg <repertoire> <uid> <pwd>")
  *  RETURN .F.
  *ENDIF
  IF TYPE('dirreq') <> 'C' OR TYPE('uid') <> 'C' OR TYPE('pwd') <> 'C'
    MESSAGEBOX("Un des param�tres est invalide.")
    RETURN .F.
  ENDIF

  dirreq = STRTRAN(dirreq, '|', ' ')
  uid = STRTRAN(uid, '|', ' ')
  pwd = STRTRAN(pwd, '|', ' ')

  IF NOT DIRECTORY(dirreq)
    MESSAGEBOX("Le r�pertoire " + dirreq + " n'existe pas.")
    RETURN .F.
  ENDIF
  IF NOT FILE(dirreq + '\requete.sql')
    MESSAGEBOX("Le fichier " + dirreq + "\requete.sql n'existe pas.")
    RETURN .F.
  ENDIF
  IF EMPTY(GETENV('STAD_HOME'))
    MESSAGEBOX("La variable syst�me %STAD_HOME% n'existe pas.")
    RETURN .F.
  ENDIF

  isError = .f.
  ON ERROR isError = .t.
  oXMLini = CREATEOBJECT("Msxml2.DOMDocument.4.0")
  IF isError
    MESSAGEBOX("MSXML 4.0 n'est pas install� sur cette machine.")
    RETURN .F.
  ENDIF
  ON ERROR

RETURN .T.

*-----------------------------------------------------------------------------------------

PROCEDURE process

  sficparam = getenv('stad_home') + '\config\parametres.xml'
  g_oxmlparam = createobject("Msxml2.DOMDocument.4.0")
  x = g_oxmlparam.load(sficparam)

  *-----

  oXMLreq = CREATEOBJECT("Msxml2.DOMDocument.4.0")
  oXMLreq.LOAD('requete.xml')

  *----

  public array g_agregs[7,3]
  g_agregs[1,1] = 'Minimum'
  g_agregs[1,2] = 'mi'
  g_agregs[1,3] = 'min'

  g_agregs[2,1] = 'Maximum'
  g_agregs[2,2] = 'mx'
  g_agregs[2,3] = 'max'

  g_agregs[3,1] = 'Moyenne'
  g_agregs[3,2] = 'mo'
  g_agregs[3,3] = 'avg'

  g_agregs[4,1] = 'Total'
  g_agregs[4,2] = 'to'
  g_agregs[4,3] = 'sum'

  g_agregs[5,1] = 'Nombre'
  g_agregs[5,2] = 'nb'
  g_agregs[5,3] = 'count'

  g_agregs[6,1] = '�cart Type'
  g_agregs[6,2] = 'et'
  g_agregs[6,3] = 'count'         && dummy;  l'�cart-type est recalcul�

  g_agregs[7,1] = 'Percentile'
  g_agregs[7,2] = 'pc'
  g_agregs[7,3] = 'count'         && custom; le count sert � calculer rang_max

  *-----

  IF FILE('bd_requete.dbc')
  *  OPEN DATABASE bd_requete
  *  x = ADBOBJECTS(atables, "TABLE")
  *  FOR i = x TO 1 STEP -1
  *    REMOVE TABLE atables[i] DELETE
  *  NEXT
  *ELSE
  *  CREATE DATABASE bd_requete
    DELETE DATABASE bd_requete DELETETABLES
  ENDIF
  IF FILE('requete0.dbf')
    DELETE FILE requete0.dbf
  ENDIF
  IF FILE('requete0.cdx')
    DELETE FILE requete0.cdx
  ENDIF
  CREATE DATABASE bd_requete
  CLOSE DATABASES

  *-----

  oNode = oXMLreq.SelectSingleNode("//ParametresGeneraux/Parametre[@Nom='Classe']")
  sClasse = oNode.text
  oNode = oXMLreq.SelectSingleNode("//ParametresGeneraux/Parametre[@Nom='Type']")
  sType = oNode.text

  sClasseXml = ''
  DO CASE
    CASE sClasse = 'Courses' AND stype = 'Achalandage'
      sClasseXml = 'Course-Achalandage'
    CASE sClasse = 'Courses' AND stype = 'Ponctualit�'
      sClasseXml = 'Course-Ponctualit�'
    CASE sClasse = 'Courses' AND stype = 'Temps de parcours'
      sClasseXml = 'Course-Temps de parcours'

    CASE sClasse = 'Points de contr�le' AND stype = 'Ponctualit�'
      sClasseXml = 'PC-Ponctualit�'
    CASE sClasse = 'Points de contr�le' AND stype = 'Temps de parcours'
      sClasseXml = 'PC-Temps de parcours'

    CASE sClasse = 'Arr�ts' AND stype = 'Achalandage'
      sClasseXml = 'Arret-Achalandage'
    CASE sClasse = 'Arr�ts' AND stype = 'Ponctualit�'
      sClasseXml = 'Arret-Ponctualit�'
    CASE sClasse = 'Arr�ts' AND stype = 'Temps de parcours'
      sClasseXml = 'Arret-Temps de parcours'

    CASE sClasse = 'Tron�ons' AND stype = 'Achalandage'
      sClasseXml = 'Tron�ons-Achalandage'
    CASE sClasse = 'Tron�ons' AND stype = 'Ponctualit�'
      sClasseXml = 'Tron�ons-Ponctualit�'
    CASE sClasse = 'Tron�ons' AND stype = 'Temps de parcours'
      sClasseXml = 'Tron�ons-Temps de parcours'

    CASE sClasse = 'Balises' AND stype = 'Temps de parcours'
      sClasseXml = 'Balise-Temps de parcours'
  ENDCASE
  sClasseXml = GETENV('STAD_HOME') + '\config\D�finition ' + sClasseXml + '.xml'
  ? '  Fichier XML: ' + sClasseXml
  oXmlClasse = CREATEOBJECT("Msxml2.DOMDocument.4.0")
  IF NOT OXMLCLASSE.LOAD(sClasseXml)
    MESSAGEBOX("MSXML ne peut charger " + sClasseXml)
    g_bOK = .f.
    RETURN
  ENDIF

  oXMLini.LOAD(GETENV('STAD_HOME') + '\config\stad.ini')
  oNode = oXMLini.SelectSingleNode("//FoxProConnectString")
  IF ISNULL(oNode)
    ? "'FoxProConnectString' est ind�fini dans stad.ini"
    g_bOK = .f.
    RETURN
  ENDIF
  sConn = oNode.TEXT
  sConn = STRTRAN(sConn, "xxxx", uid)
  sConn = STRTRAN(sConn, "yyyy", pwd)
  *? '  ' + sConn
  hConn = SQLSTRINGCONNECT(sConn)
  IF hConn = 0
    ? "Connexion impossible: " + sConn
    g_bOK = .f.
    RETURN
  ENDIF
  *LS:  driver={SQL Server};server=TITAN;database=stad_dev;uid=stad;pwd=stad
  *rtl: driver={SQL Server};server=SQL2\MSSSQL2PRD;database=stad;uid=stad;pwd=stad

  *-----

  ?
  ? "Requ�te  initiale."
  ? "  input: BD STAD  output: requete0.dbf"

  handle = FOPEN('requete.sql')
  ssql = ""
  DO WHILE NOT FEOF(handle)
    sline = FGETS(handle)
    *? sline
    ssql = ssql + sline
  ENDDO
  x=FCLOSE(handle)

  x = sqlexec(hConn, ssql)
  IF x <= 0
    ? "  Erreur dans la requ�te."
    g_bOK = .f.
    suspend
    RETURN
  ENDIF

  SELECT sqlresult
  COPY TO requete0
  USE

  *-----

  DO calcul_base WITH sClasse, oXmlClasse
  IF ! g_bOK
    ? "  Erreur dans calcul_base."
    RETURN
  ENDIF

  *-----
  * Traitements avanc�es des donn�es

  bGenCourses = .f.
  soper = "Ajout des courses pr�vues non-�chantillonn�es"
  sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
  oNode = oXMLreq.SelectSingleNode(sxpath)
  IF NOT ISNULL(oNode)
    bGenCourses = .t.
    DO gen_courses WITH hConn, oXMLreq, oXmlClasse
    IF ! g_bOK
    ? "  Erreur dans gen_courses."
      RETURN
    ENDIF
  ENDIF

  soper = "Pond�ration des donn�es"
  sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
  oNode = oXMLreq.SelectSingleNode(sxpath)
  IF NOT ISNULL(oNode)
    IF sClasse = 'Courses' or sClasse = 'Points de contr�le' or sClasse = 'Arr�ts'
      DO ponderation WITH hConn, oXMLreq, oNode.text, sClasse, oXmlClasse
      IF ! g_bOK
        ? "  Erreur dans ponderation."
        RETURN
      ENDIF
    ELSE
      MESSAGEBOX("Ne peut effectuer des pond�rations pour ce type de requ�te")
      g_bOK = .f.
      RETURN
    ENDIF
  ENDIF

  soper = "Agr�gation des observations"
  sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
  oNode = oXMLreq.SelectSingleNode(sxpath)
  IF NOT ISNULL(oNode)
    DO agreg_champs WITH 'COURSES', oXMLreq, oXmlClasse
    IF ! g_bOK
      ? "  Erreur dans agreg_champs."
      RETURN
    ENDIF
  ENDIF

  soper = "Imputation de donn�es"
  sxpath = "//DefChampSortie/Traitement/Operation[@Nom='" + soper + "']"
  oNode = oXMLreq.SelectSingleNode(sxpath)
  IF NOT ISNULL(oNode)
    IF g_bDev
      DO imputationDev WITH hConn, oXMLreq, oXmlClasse
    ELSE
      DO imputation WITH hConn, oXMLreq, oXmlClasse
    ENDIF
    IF ! g_bOK
       ? "  Erreur dans imputation."
      RETURN
    ENDIF
  ENDIF

  *-----

  DO agreg_champs WITH 'CHAMPS', oXMLreq, oXmlClasse
  IF ! g_bOK
    ? "  Erreur dans agreg_champs."
    RETURN
  ENDIF

  *-----

  DO export_requete WITH oXMLreq, oXmlClasse
  IF ! g_bOK
    ? "  Erreur dans export_requete."
    RETURN
  ENDIF
  
  *-----
  
  * pour les requ�tes troncons, cleanup
  if sclasse = 'Tron�ons'
    onode = oxmlreq.selectsinglenode("//ParametresGeneraux/Parametre[@Nom='Auteur']")
    
    sTableTroncon = 'sdap_troncon_' + onode.text
    x = sqlexec(hConn, 'select * from ' + sTableTroncon)
    COPY TO sdap_troncon DATABASE bd_requete
    USE
    x = sqlexec(hConn, 'drop table ' + sTableTroncon)
    if bGenCourses
      sTableTroncon = 'ref_troncon_' + onode.text
      x = sqlexec(hConn, 'select * from ' + sTableTroncon)
      COPY TO ref_troncon DATABASE bd_requete
      USE
      x = sqlexec(hConn, 'drop table ' + sTableTroncon)
    endif    
  endif
  
  *-----
  
*  ! del *.bak
   erase *.bak
  x = SQLDISCONNECT(hConn)
  IF TYPE('BBAT') = 'C'
    if bBat = 'quit'
      quit
    ENDIF
  ENDIF
RETURN

*-----------------------------------------------------------------------------------------
