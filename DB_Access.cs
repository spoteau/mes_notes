using System;
using System.Collections.Generic;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Collections;
using System.Web;
using System.Net;
using System.Net.Mail;

/// <summary>
/// Helper class to access Database
/// </summary>
public class DB_Access: IDisposable
{
    private String mConnectionString;

    // Add property for error reporting:
    /// <summary>
    /// Constructeur
    /// </summary>
    public DB_Access()
    {
        mConnectionString = ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="connectionString"></param>
    public DB_Access(String connectionString)
    {
        mConnectionString = connectionString;
    }

    /// <summary>
    /// get the windows user
    /// </summary>
    /// <param name="sEmail"></param>
    /// <returns>string</returns>
    public string GetWindowsUser()
    {
        string sWindowsUser = HttpContext.Current.User.Identity.Name;
        sWindowsUser = sWindowsUser.Substring(sWindowsUser.IndexOf('\\') + 1);

        if (sWindowsUser == string.Empty || sWindowsUser == "DefaultAppPool")
        {
            sWindowsUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            sWindowsUser = sWindowsUser.Substring(sWindowsUser.IndexOf('\\') + 1);
        }
        if (sWindowsUser == string.Empty || sWindowsUser == "DefaultAppPool")
        {
            sWindowsUser = Environment.UserName;
        }
        
        return sWindowsUser;
    }

    /// <summary>
    /// Check if a user is part of the Admin group
    /// </summary>
    /// <returns>string</returns>
    public bool isCurrentWindowsUserAdmin()
    {
        bool retValue = false;
        try
        {
            System.Security.Principal.WindowsIdentity identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(identity);
            retValue = principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }

        return retValue;
    }
    /// <summary>
    /// Check if a user is part of the Admin group
    /// </summary>
    /// <param name="connectionString"></param>
    /// <returns>string</returns>
    public bool isCustomAdminMember(String sUserName, SqlConnection sqlConn)
    {
        bool retVal = false;

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;
        try
        {
            string sSql = String.Format("SELECT COUNT(UserName) AS Trouve FROM CustomAdmin WHERE UserName = '{0}'", sUserName);

            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtTemp.Read())
            {
                int iTrouve = int.Parse(dtTemp["Trouve"].ToString());
                if (iTrouve > 0)
                {
                    retVal = true;
                }
            }
            dtTemp.Close();
            dtTemp = null;
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }

        return retVal;
    }

    public double getMoyenneEdutiant(String PersonneID, SqlConnection sqlConn)
    {
        String sQuery = String.Format("SELECT AVG(NoteSurCent) FROM CoursPris WHERE PersonneID = '{0}' AND NoteSurCent <> 0", PersonneID);
        return GetScalarDouble(sQuery, sqlConn);
    }
    /// <summary>
    /// Check if a user is a student
    /// </summary>
    /// <param name="sEmail"></param>
    /// <returns>string</returns>
    //public bool isCurrentWindowsInThisGroup(String sGroup)
    //{
    //    bool retValue = false;
    //    try
    //    {
    //        retValue = System.Security.Principal.WindowsPrincipal.Current.IsInRole(sGroup);
    //    }
    //    catch (Exception ex)
    //    {
    //        Debug.WriteLine(ex.Message);
    //    }

    //    return retValue;
    //}
    /// <summary>
    /// Check for unwanted characters before inserting in Database
    /// </summary>
    /// <param name="toScrub"></param>
    /// <returns>string</returns>
    public string ScrubStringForDB(String toScrub)
    {
        String sReturnedString = toScrub.Replace("'", "''").Trim();

        return sReturnedString;
    }
    public String GetHeaderPrintString()
    {
        /*****************************************************************************************************
        Created By: Ferdous Md. Jannatul, Sr. Software Engineer
        Created On: 10 December 2005
        Last Modified: 13 April 2006
        ******************************************************************************************************/
        //Generating Pop-up Print Preview page
        return " " +
        " function getPrint(print_area)" +
        " {	" +
        " 	//Creating new page" +
        " 	var pp = window.open();" +
        " 	//Adding HTML opening tag with <HEAD> … </HEAD> portion " +
        " 	pp.document.writeln('<HTML><HEAD><title>Print Preview</title><LINK href=Styles.css  type='text/css' rel='stylesheet'>')" +
        " 	pp.document.writeln('<LINK href=PrintStyle.css  type='text/css' rel='stylesheet' media='print'><base target='_self'></HEAD>')" +
        " 	//Adding Body Tag" +
        " 	pp.document.writeln('<body MS_POSITIONING='GridLayout' bottomMargin='0' leftMargin='0' topMargin='0' rightMargin='0'>');" +
        " 	//Adding form Tag" +
        " 	pp.document.writeln('<form  method='post'>');" +
        " 	//Creating two buttons Print and Close within a table" +
        " 	pp.document.writeln('<TABLE width=100%><TR><TD></TD></TR><TR><TD align=right><INPUT ID='PRINT' type='button' value='Print' onclick='window.print();'><INPUT ID='CLOSE' type='button' value='Close' onclick='window.close();'></TD></TR><TR><TD></TD></TR></TABLE>');" +
        " 	//Writing print area of the calling page" +
        " 	pp.document.writeln(document.getElementById(print_area).innerHTML);" +
        " 	//Ending Tag of </form>, </body> and </HTML> " +
        " 	pp.document.writeln('</form></body></HTML>'); " +
        "}		 ";

    }

    /*********************** UEspoir specific stuffs *************************/

    /// <summary>
    /// Check if Student already is registered for this class
    /// </summary>
    /// <param name="PersonneID"></param>
    /// <param name="NumeroCours"></param>
    /// <returns></returns>
    public bool EtudiantDejaInscrit(string PersonneID, string NumeroCours, SqlConnection sqlConn)
    {
        bool retValue = false;
       
        SqlCommand SqlCmd;
        try
        {
            string sSql = String.Format("SELECT COUNT(PersonneID) AS Trouve FROM CoursPris WHERE PersonneID = '{0}' " +  
                        " AND NumeroCours = '{1}' AND IsNull(DateReussite, '') = ''", PersonneID, NumeroCours);

            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtTemp.Read())
            {
                int iTrouve = int.Parse(dtTemp["Trouve"].ToString());
                if (iTrouve > 0)
                {
                    retValue = true;
                }
            }
            dtTemp.Close();
            dtTemp = null;

            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
        return retValue;
    }

    public int NombreDeCoursCetteSession(String sPersonneID, SqlConnection sqlConn)
    {
        int retValue = 0;
        
        SqlCommand SqlCmd;
        try
        {
            string sSql = String.Format("SELECT COUNT(CP.NumeroCours) AS Trouve FROM CoursPris CP, Cours C WHERE CP.NumeroCours = C.NumeroCours " +
            " AND C.ExamenEntree = 0 " +
            " AND PersonneID =  '{0}'  " +
            " AND CoursOffertID IN (SELECT CoursOffertID FROM CoursOfferts WHERE ACTIF = 1) ", sPersonneID);

            sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            sqlConn.Open();

            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;
            SqlCmd.CommandText = sSql;
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dtTemp.Read())
            {
                retValue = int.Parse(dtTemp["Trouve"].ToString());
                dtTemp.Close();
                dtTemp = null;              
            }
        }
        catch(Exception ex)
        {
            Debug.WriteLine(ex.Message);
            
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }        

        return retValue;
    }

    /// <summary>
    /// Facturer pour Obligation quand montant non disponible
    /// </summary>
    /// <param name="sObligation"></param>
    /// <param name="p"></param>
    public void FacturerPourObligation(string sObligationLabel, string sPersonneID, SqlConnection sqlConn)
    {
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;

            // Facturer
            SqlCmd.CommandText = String.Format("INSERT INTO MontantsDus (PersonneID, CodeObligation, Montant) SELECT '{0}', '{1}', Montant " +
                    " FROM Obligations WHERE Code = '{1}'",
                sPersonneID, sObligationLabel);
            SqlCmd.ExecuteNonQuery();

            SqlCmd = null;
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
    }

    /// <summary>
    /// Facturer pour Obligation quand montant disponible
    /// </summary>
    /// <param name="sObligation"></param>
    /// <param name="p"></param>
    public void FacturerPourObligation(string sObligationLabel, string sPersonneID, String sMontant, SqlConnection sqlConn)
    {
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;

            // Facturer
            SqlCmd.CommandText = String.Format("INSERT INTO MontantsDus (PersonneID, CodeObligation, Montant) VALUES ('{0}', '{1}', {2}) ",
                sPersonneID, sObligationLabel, sMontant);
            SqlCmd.ExecuteNonQuery();

            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
    }
    /// <summary>
    /// Si c'est le premier cours inscrit pour la session, l'étudiant est facturer le montant de la session
    /// Tout ajustement de temps partiel doit être réglé par la direction administrive
    /// </summary>
    /// <param name="sPersonneID"></param>
    /// <param name="sNumeroCours"></param>
    public void Facturer(String sPersonneID, String sNumeroCours, int NombreDeMoisParSession, SqlConnection sqlConn)
    {
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;
        Hashtable htExamenDePlacement = GetExamenDePlacement(sqlConn);
        if (htExamenDePlacement.Contains(sNumeroCours))
        {
            // Pas encore de facture
            htExamenDePlacement = null;
            return;
        }

        try
        {
            //string sSql = String.Format("SELECT COUNT(NumeroCours) AS Trouve FROM CoursPris WHERE ISNULL(DateReussite, '') = '' AND PersonneID =  '{0}' ", sPersonneID);
            //string sSql = String.Format("SELECT COUNT(NumeroCours) AS Trouve FROM CoursPris WHERE PersonneID =  '{0}' " +
            //    " AND CoursOffertID IN (SELECT CoursOffertID FROM CoursOfferts WHERE ACTIF = 1)", sPersonneID);

            string sSql = String.Format("SELECT COUNT(CP.NumeroCours) AS Trouve FROM CoursPris CP, Cours C WHERE CP.NumeroCours = C.NumeroCours " +
	                " AND C.ExamenEntree = 0 " +
	                " AND PersonneID =  '{0}'  " +
                    " AND CoursOffertID IN (SELECT CoursOffertID FROM CoursOfferts WHERE ACTIF = 1) ", sPersonneID);

            sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            sqlConn.Open();

            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;
            SqlCmd.CommandText = sSql;
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dtTemp.Read())
            {
                int iTrouve = int.Parse(dtTemp["Trouve"].ToString());

                dtTemp.Close();     // Sinon on ne peut pas réutiliser la connection (variable)
                dtTemp = null;
                sqlConn.Open();
                if (iTrouve == 1) // Première classe inscrite qui n'est pas une classe de placement
                {
                    // Facturer d'Entrée
                    SqlCmd.CommandText = String.Format("INSERT INTO MontantsDus (PersonneID, CodeObligation, Montant) SELECT '{0}', '{1}', " +
                            " Montant FROM Obligations WHERE Code = '{1}'", sPersonneID, "FE");
                    SqlCmd.ExecuteNonQuery();
                }

                // A partir de 4 classes, no more facturation
                int iMaxClasses = Int16.Parse(ConfigurationManager.AppSettings["MaximumClassesFacturees"].ToString());
                if (iTrouve <= iMaxClasses)
                {
                    // Facturer pour 1 cours
                    SqlCmd.CommandText = String.Format("INSERT INTO MontantsDus (PersonneID, CodeObligation, Montant) SELECT '{0}', '{1}', " +
                            " Montant * {2} FROM Obligations WHERE Code = '{1}'", sPersonneID, "FM", NombreDeMoisParSession);
                    SqlCmd.ExecuteNonQuery();
                }
                //sqlConn.Close();
            }
            sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
        htExamenDePlacement = null;
        sqlConn = null;
    }
    public String getNoteForEdutiant(String sPersonneID, String sCoursOffertID, SqlConnection sqlConn)
    {
        String sReturnedValue = "Error...";
        try
        {
            string sSql = String.Format("SELECT NoteSurCent FROM CoursPris WHERE PersonneID = '{0}' AND CoursOffertID = '{1}'", sPersonneID, sCoursOffertID);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    sReturnedValue = dt["NoteSurCent"].ToString();
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }

        return sReturnedValue;
    }
    public String GetMontantObligation(String sObligationCode, SqlConnection sqlConn)
    {
        String sReturnedValue = "Error...";
        try
        {
            string sSql = String.Format("SELECT Montant FROM Obligations WHERE Code = '{0}'", sObligationCode);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    sReturnedValue = double.Parse(dt["Montant"].ToString()).ToString("F");
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }

        return sReturnedValue;
    }
    /// <summary>
    /// Check if Student already is registered for this class
    /// </summary>
    /// <param name="PersonneID"></param>
    /// <param name="NumeroCours"></param>
    /// <returns></returns>
    public bool EtudiantDejaReussi(string PersonneID, string NumeroCours, SqlConnection sqlConn)
    {
        bool retValue = false;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            string sSql = String.Format("SELECT COUNT(PersonneID) AS Trouve FROM CoursPris WHERE PersonneID = '{0}'  AND NumeroCours = '{1}' AND (NoteSurCent >= NotePassage OR Waiver = 1)", PersonneID, NumeroCours);
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dtTemp.Read())
            {
                int iTrouve = int.Parse(dtTemp["Trouve"].ToString());
                if (iTrouve > 0)
                {
                    retValue = true;
                }
            }
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
        return retValue;
    }
    /// <summary>
    /// Check if Student already is registered for this class OR already took this class
    /// </summary>
    /// <param name="PersonneID"></param>
    /// <param name="NumeroCours"></param>
    /// <returns></returns>
    public bool EtudiantDejaInscritOuReussi(string PersonneID, string NumeroCours, SqlConnection sqlConn)
    {
        bool retValue = false;
        SqlCommand SqlCmd;
        
        try
        {
            string sSql = String.Format("SELECT COUNT(PersonneID) AS Trouve FROM CoursPris WHERE PersonneID = '{0}'  AND NumeroCours = '{1}' AND (NoteSurCent >= NotePassage OR Waiver = 1) " +
                " UNION " +
                " SELECT COUNT(PersonneID) AS Trouve FROM CoursPris WHERE PersonneID = '{0}'  AND NumeroCours = '{1}' AND IsNull(DateReussite, '') = ''", PersonneID, NumeroCours);

            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();
            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlDataReader dtTemp = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dtTemp.Read())
            {
                int iTrouve = int.Parse(dtTemp["Trouve"].ToString());
                if (dtTemp.Read())     // Lis la 2e ligne du Résultat
                {
                    iTrouve += int.Parse(dtTemp["Trouve"].ToString());
                }
                if (iTrouve > 0)
                {
                    retValue = true;
                }
            }
            dtTemp.Close();
            SqlCmd = null;
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }
        return retValue;
    }
    public Hashtable GetEtudiantsAlreadyIn(String sNumeroCours, SqlConnection sqlConn)
    {
        Hashtable ht = new Hashtable();

        // Loop through all records
        try
        {
            String hashSQL = String.Format("SELECT PersonneID FROM CoursPris WHERE CoursOffertID IN (SELECT DISTINCT CO.CoursOffertID FROM CoursOfferts CO, " +
                " LesSessions LS WHERE CO.SessionID = LS.SessionID AND CO.NumeroCours = '{0}' AND SessionCourante = 1)", sNumeroCours);

            SqlDataReader dt = GetDataReader(hashSQL, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    ht.Add(dt["PersonneID"].ToString(), dt["PersonneID"].ToString());
                    while (dt.Read())
                    {
                        if (!ht.ContainsKey(dt["PersonneID"].ToString()))
                        {
                            ht.Add(dt["PersonneID"].ToString(), dt["PersonneID"].ToString());
                        }
                        else
                        {
                            Debug.WriteLine(dt["PersonneID"].ToString()); // So I can check for duplicate while testing
                        }
                    }
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return ht;
    }
    /// <summary>
    /// Les examens de placement sous forme de cours inscrits et une note est assignée
    /// </summary>
    /// <returns></returns>
    public Hashtable GetExamenDePlacement(SqlConnection sqlConn)
    {
        Hashtable ht = new Hashtable();

        // Loop through all records
        try
        {
            String hashSQL = String.Format("SELECT NumeroCours FROM Cours WHERE ExamenEntree = 1");

            SqlDataReader dt = GetDataReader(hashSQL, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    ht.Add(dt["NumeroCours"].ToString(), dt["NumeroCours"].ToString());
                    while (dt.Read())
                    {
                        if (!ht.ContainsKey(dt["NumeroCours"].ToString()))
                        {
                            ht.Add(dt["NumeroCours"].ToString(), dt["NumeroCours"].ToString());
                        }
                        else
                        {
                            Debug.WriteLine(dt["NumeroCours"].ToString()); // So I can check for duplicate while testing
                        }
                    }
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return ht;
    }
    /// <summary>
    /// Valeur retournée est le Montant total des obligations de l'étudiant
    /// </summary>
    /// <param name="sPersonneID"></param>
    /// <returns></returns>
    public Double FraisTotalEtudiant(string sPersonneID, SqlConnection sqlConn)
    {
        Double doubleReturned = 0.00;

        try
        {
            string sSql = String.Format("SELECT Isnull(SUM(Montant), 0) AS MontantDu FROM MontantsDus WHERE PersonneID = '{0}'", sPersonneID);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    doubleReturned = Double.Parse(dt["MontantDu"].ToString());
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return doubleReturned;
    }
    /// <summary>
    /// Valeur retournée est le montant total des paiements reçus de l'étudiant
    /// </summary>
    /// <param name="sPersonneID"></param>
    /// <returns></returns>
    public Double FraisTotalRecuEtudiant(string sPersonneID, SqlConnection sqlConn)
    {
        Double doubleReturned = 0.00;

        // Loop through all records
        try
        {
            string sSql = String.Format("SELECT Isnull(SUM(Montant), 0) AS MontantRecu FROM MontantsRecus WHERE PersonneID = '{0}'", sPersonneID);

            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    doubleReturned = Double.Parse(dt["MontantRecu"].ToString());
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return doubleReturned;
    }
    /// <summary>
    /// Valeur retournée est la balance: Obligations Totales MOINS Paiements Totaux
    /// </summary>
    /// <param name="sPersonneID"></param>
    /// <returns></returns>
    public Double BalanceEtudiant(string sPersonneID, SqlConnection sqlConn)
    {
        return FraisTotalEtudiant(sPersonneID, sqlConn) - FraisTotalRecuEtudiant(sPersonneID, sqlConn);
    }
    /// <summary>
    /// Base function returning the Start Date of the current Session
    /// </summary>
    /// <returns>DateTime</returns>
    public DateTime GetStartDateOfCurrentSession(SqlConnection sqlConn)
    {
        DateTime retValue = new DateTime();

        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT * FROM LesSessions WHERE Actif = 1 AND SessionCourante = 1", sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = DateTime.Parse(dtReader["SessionDateDebut"].ToString());
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            retValue = DateTime.Today.Subtract(TimeSpan.FromDays(30));
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
            //sqlConn = null;
        }

        return retValue;
    }
    /// <summary>
    /// Base function returning the End Date of the current Session
    /// </summary>
    /// <returns>DateTime</returns>
    public DateTime GetEndDateOfCurrentSession(SqlConnection sqlConn)
    {
        DateTime retValue = new DateTime();

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT * FROM LesSessions WHERE Actif = 1 AND SessionCourante = 1", sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = DateTime.Parse(dtReader["SessionDateFin"].ToString());
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            retValue = DateTime.Today.Subtract(TimeSpan.FromDays(30));
            Debug.WriteLine(ex.Message);
            //Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return retValue;
    }
    /// <summary>
    /// Base function returning the Start Date of the current Session
    /// </summary>
    /// <returns>DateTime</returns>
    public DateTime GetStartDateOfASession(String sSessionID, SqlConnection sqlConn)
    {
        DateTime retValue = new DateTime();

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT SessionDateDebut FROM LesSessions WHERE SessionID = " + sSessionID, sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = DateTime.Parse(dtReader["SessionDateDebut"].ToString()); ;
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            retValue = DateTime.Today.Subtract(TimeSpan.FromDays(30));
            Debug.WriteLine(ex.Message);
            ////Debug.WriteLine(ex.Message);
            ////if (sqlConn.State == ConnectionState.Open)
            ////    sqlConn.Close();
        }

        return retValue;
    }
    /// <summary>
    /// Base function returning the End Date of the current Session
    /// </summary>
    /// <returns>DateTime</returns>
    public DateTime GetEndDateOfASession(String sSessionID, SqlConnection sqlConn)
    {
        DateTime retValue = new DateTime();

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT SessionDateFin FROM LesSessions WHERE SessionID = " + sSessionID, sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = DateTime.Parse(dtReader["SessionDateFin"].ToString()); ;
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            retValue = DateTime.Today.Subtract(TimeSpan.FromDays(30));
            Debug.WriteLine(ex.Message);
            //Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return retValue;
    }
    /// <summary>
    /// 
    /// </summary>
    /// <returns>int</returns>
    public int GetIdOfSessionCourante(SqlConnection sqlConn)
    {
        int retValue = 0;

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT SessionID FROM LesSessions WHERE SessionCourante = 1", sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = Int16.Parse(dtReader["SessionID"].ToString()); ;
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return retValue;
    }
    /// <summary>
    /// 
    /// </summary>
    /// <returns>String</returns>
    public String GetIdOfSessionCouranteString(SqlConnection sqlConn)
    {
        String retValue = "";

        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT SessionID FROM LesSessions WHERE SessionCourante = 1", sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                retValue = dtReader["SessionID"].ToString(); ;
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return retValue;
    }
    public bool CoursAppartientASessionCourante(String sCoursOffertID, SqlConnection sqlConn)
    {
        bool bRetVal = false;
        
        SqlCommand SqlCmd;
        try
        {
            String sSql = String.Format("SELECT COUNT(CoursOffertID) As Found FROM CoursOfferts " +
                " WHERE CoursOffertID = {0} AND SessionID IN (SELECT SessionID FROM LesSessions WHERE SessionCourante = 1)", sCoursOffertID);
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                bRetVal = int.Parse(dtReader["Found"].ToString()) > 0; ;
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return bRetVal;
    }
    /// <summary>
    /// Check if amount is right format
    /// </summary>
    /// <param name="sCheck"></param>
    /// <returns></returns>
    public bool ChiffreSeulement(String sCheck)
    {
        char[] sAllChar = sCheck.ToCharArray();

        foreach (char ch in sAllChar)
        {
            if (!Char.IsNumber(ch) && ch != '.')
            {
                return false;
            }
        }
        return true;
    }
    // CCPAP Finance Stuff
    /// <summary>
    /// 
    /// </summary>
    /// <param name="sCheckDate"></param>
    /// <returns></returns>
    public bool EstQueDateOK(String sCheckDate)
    {
        if (sCheckDate == String.Empty)
            return false;

        char[] sAllChar = sCheckDate.ToCharArray();

        foreach (char ch in sAllChar)
        {
            if (!Char.IsNumber(ch) && ch.CompareTo('/') != 0)
            {
                return false;
            }
        }

        DateTime datechoisie = DateTime.Parse(sCheckDate);
        String dt = datechoisie.Year.ToString();
        if (dt == "1")
        {
            return false;
        }


        if (datechoisie > DateTime.Today)
        {
            return false;
        }

        return true;
    }
    /// <summary>
    /// Get the CoursOffertID for the current session with the corresponding numeros
    /// </summary>
    /// <param name="NumeroCours"></param>
    /// <returns></returns>
    public int GetCoursOffertID(String NumeroCours, SqlConnection sqlConn)
    {
        int dtRetVal = 0;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(string.Format("SELECT CO.CoursOffertID FROM CoursOfferts CO, LesSessions LS WHERE " +
                "CO.SessionID = LS.SessionID AND NumeroCours = '{0}' AND LS.SessionCourante = 1", NumeroCours), sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                dtRetVal = int.Parse(dtReader["CoursOffertID"].ToString());
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            dtRetVal = 0;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return dtRetVal;
    }
    /// <summary>
    /// Trouver le nom de la Discipline en utilisant le ID
    /// </summary>
    /// <param name="sID"></param>
    /// <returns></returns>
    public string GetDisciplineFromID(string sID, SqlConnection sqlConn)
    {
        String sReturnedValue = "Error...";
        try
        {
            string sSql = String.Format("SELECT DisciplineNom FROM Disciplines WHERE DisciplineID = '{0}'", sID);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    sReturnedValue = dt["DisciplineNom"].ToString();
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return sReturnedValue;
    }
    /// <summary>
    /// Trouver le ID de la Discipline en utilisant le Nom
    /// </summary>
    /// <param name="iID"></param>
    /// <returns></returns>
    public int GetDisciplineIDFromName(int iID, SqlConnection sqlConn)
    {
        int iReturnedValue = 0;
        try
        {
            string sSql = String.Format("SELECT DisciplineID FROM Disciplines WHERE DisciplineNom = '{0}'", iID);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    iReturnedValue = int.Parse(dt["DisciplineNom"].ToString());
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return iReturnedValue;
    }
    /// <summary>
    /// Trouver le nom de la Discipline en utilisant le ID
    /// </summary>
    /// <param name="sID"></param>
    /// <returns></returns>
    public string GetDisciplineEtudiant(String PersonneID, SqlConnection sqlConn)
    {
        String sReturnedValue = "Error...";
        try
        {
            string sSql = String.Format("SELECT D.DisciplineID, DisciplineNom FROM Personnes P, Disciplines D WHERE P.DisciplineID = D.DisciplineID AND PersonneID = '{0}'", PersonneID);
            SqlDataReader dt = GetDataReader(sSql, sqlConn);
            if (dt != null)
            {
                if (dt.Read())
                {
                    sReturnedValue = dt["DisciplineNom"].ToString();
                }
                dt.Close();
                dt = null;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return sReturnedValue;
    }
    /// <summary>
    /// Effacer et Remplir la table de controle tblControleFinanceSesCour
    /// </summary>
    /// <param name="i"></param>
    /// <param name="s"></param>
    public bool PrepareTableDeControle(int i, string s, SqlConnection sqlConn)
    {
        bool retVal = false;
        try
        {
            string sSql = "INSERT INTO TblControleSesCour (PersonneID, Prenom, Nom, DDN, EtudiantID, Actif, NIF) " +
                " SELECT PersonneID, Prenom, Nom, Isnull(DDN, '1900-1-1') AS DDN, EtudiantID, Actif, Isnull(NIF, '1900-1-1') AS NIF  " +
                " FROM Personnes WHERE PersonneID IN " +
                " (SELECT PersonneID FROM CoursPris WHERE CoursOffertID IN ( " +
                " SELECT CoursOffertID FROM CoursOfferts C, LesSessions L  " +
                " WHERE C.SessionID=L.SessionID AND SessionCourante = 1))";
            //SqlConnection sqlConn = null;
            SqlCommand SqlCmd;
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(sSql, sqlConn);
            SqlCmd.ExecuteNonQuery();

            sSql = "SELECT COUNT(C.PersonneID) NDC, C.PersonneID FROM CoursPris C, TblControleSesCour T " +
                " WHERE C.PersonneID = T.PersonneID AND NumeroCours <> 'FRA001' AND NumeroCours <> 'MAT001' AND NumeroCours <> 'ANG001'  " +
                " AND CoursOffertID IN (SELECT CoursOffertID FROM CoursOfferts C, LesSessions L  " +
                " WHERE C.SessionID=L.SessionID AND SessionCourante = 1) " +
                " GROUP BY C.PersonneID ORDER BY C.PersonneID ";

            SqlDataReader dtReader = SqlCmd.ExecuteReader();
            int NbreDeCours = 0;
            //String PersonneID = "";
            if (dtReader.Read())
            {
                NbreDeCours = int.Parse(dtReader["CoursOffertID"].ToString());
                dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            }

            //sqlConn = null;

        }
        catch (Exception ex)
        {
            Debug.WriteLine("ERROR: " + ex.Message);
            retVal = false;
        }

        return retVal;
    }

    /************************ End UEspoir specific stuffs *****************************/

    /// <summary>
    /// Base function used by other functions to get Taux Gourdes/Dollar US
    /// Return a double or 0.0 if error
    /// </summary>
    /// <returns>double</returns>
    public double GetTauxGourdesDollarUS(SqlConnection sqlConn)
    {
        double dtRetVal = 0.0;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand("SELECT * FROM Parametres WHERE ID = 1000", sqlConn);
            SqlDataReader dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            if (dtReader.Read())
            {
                dtRetVal = double.Parse(dtReader["Taux"].ToString());
            }

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            dtRetVal = 0.0;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return dtRetVal;
    }

    /// <summary>
    /// Base function used by other functions to get Taux Gourdes/Dollar US
    /// Return a double or 0.0 if error
    /// </summary>
    /// <returns>double</returns>
    public double GetTauxGourdesDollarUS(int DepenseID, SqlConnection sqlConn)
    {
        double dtRetVal = 0.0;

        // Get Taux from Depense entry
        dtRetVal = GetTauxGourdesDollarUSFromDepense(DepenseID, sqlConn);
        if (dtRetVal > 0.0)
        {
            return dtRetVal;
        }

        // Get Taux from General Settings
        return GetTauxGourdesDollarUS(sqlConn);        // ToDo : ????????????????
    }

    /// <summary>
    /// 
    /// Return a double or 0.0 if error
    /// </summary>
    /// <returns>double</returns>
    public double GetTauxGourdesDollarUSFromDepense(int DepenseID, SqlConnection sqlConn)
    {
        double dtRetVal = 0.0;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;
            SqlCmd.CommandText = "SELECT Taux_Interet FROM Depenses WHERE DepenseID = " + DepenseID;
            dtRetVal = Double.Parse(SqlCmd.ExecuteScalar().ToString());

            //sqlConn.Close();
            //sqlConn = null;
            SqlCmd = null;
        }
        catch (Exception ex)
        {
            dtRetVal = 0.0;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }
        return dtRetVal;
    }

    /// <summary>
    /// Base function used by other functions to access the database and return a DataReader
    /// </summary>
    /// <returns></returns>
    public SqlDataReader GetDataReader(string csSQL, SqlConnection sqlConn)
    {
        SqlDataReader dtRetVal = null;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(csSQL, sqlConn);
            dtRetVal = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

            //sqlConn = null;
        }
        catch (Exception ex)
        {
            dtRetVal = null;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return dtRetVal;
    }

    /// <summary>
    /// Base function used by other functions to access the database and return a DataReader
    /// </summary>
    /// <returns></returns>
    public SqlDataReader GetDataReaderWithParams(string csSQL, SqlConnection sqlConn, params SqlParameter[] paramList)
    {
        SqlDataReader dtRetVal = null;
        SqlCommand SqlCmd = null;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(csSQL, sqlConn);
            // Add parameters
            for (int iIndex = 0; iIndex < paramList.Length; iIndex++)
            {
                SqlCmd.Parameters.Add(paramList[iIndex]);
            }
            dtRetVal = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            SqlCmd.Parameters.Clear();
            SqlCmd = null;
            
        }
        catch (Exception ex)
        {
            dtRetVal = null;
            if (SqlCmd != null)
                SqlCmd.Parameters.Clear();

            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return dtRetVal;
    }

    /// <summary>
    /// Helps to issue query command
    /// </summary>
    /// <returns></returns>
    public bool IssueCommand(string csSQL, SqlConnection sqlConn)
    {
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;
        bool bRetVal = false;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(csSQL, sqlConn);
            SqlCmd.ExecuteNonQuery();
            SqlCmd = null;
            //sqlConn.Close();
            //sqlConn = null;
            bRetVal = true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return bRetVal;
    }

    /// <summary>
    /// Helps to issue query command
    /// </summary>
    /// <returns></returns>
    public bool IssueCommandWithParams(string csSQL, SqlConnection sqlConn, params SqlParameter[] paramList)
    {
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd = null;
        bool bRetVal = false;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(csSQL, sqlConn);
            // Add parameters
            for (int iIndex = 0; iIndex < paramList.Length; iIndex++)
            {
                SqlCmd.Parameters.Add(paramList[iIndex]);
            }

            SqlCmd.ExecuteNonQuery();
            bRetVal = true;
            SqlCmd.Parameters.Clear();
            SqlCmd = null;
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            if (SqlCmd != null)
                SqlCmd.Parameters.Clear();
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();

            Debug.WriteLine(ex.Message);
        }

        return bRetVal;
    }

    /// <summary>
    /// Retrieve the scalar from query paramater
    /// </summary>
    /// <param name="csSql"></param>
    /// <returns></returns>
    public int GetScalar(string csSql, SqlConnection sqlConn)
    {
        int iRetVal = 0;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;
            SqlCmd.CommandText = csSql;
            iRetVal = int.Parse(SqlCmd.ExecuteScalar().ToString());

            //sqlConn.Close();
            //sqlConn = null;
            SqlCmd = null;
        }
        catch (Exception ex)
        {
            iRetVal = -1;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }
        return iRetVal;
    }

    /// <summary>
    /// Retrieve the scalar from query paramater
    /// </summary>
    /// <param name="csSql"></param>
    /// <returns></returns>
    public double GetScalarDouble(string csSql, SqlConnection sqlConn)
    {
        double dRetVal = 0;
        //SqlConnection sqlConn = null;
        SqlCommand SqlCmd;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand();
            SqlCmd.Connection = sqlConn;
            SqlCmd.CommandText = csSql;
            dRetVal = double.Parse(SqlCmd.ExecuteScalar().ToString());

            //sqlConn.Close();
            //sqlConn = null;
            SqlCmd = null;
        }
        catch (Exception ex)
        {
            dRetVal = -1;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }
        return dRetVal;
    }

    /// <summary>
    /// Helps to issue query command and returns identity value created
    /// Date: Avril 2013
    /// </summary>
    /// <returns></returns>
    public int GetScalarWithParams(string csSQL, SqlConnection sqlConn, params SqlParameter[] paramList)
    {
        SqlCommand SqlCmd = null;
        SqlCommand SqlCmdNewID = null;
        //SqlConnection sqlConn = null;
        Int32 bRetVal = 0;

        try
        {
            //sqlConn = new System.Data.SqlClient.SqlConnection(mConnectionString);
            //sqlConn.Open();

            SqlCmd = new SqlCommand(csSQL, sqlConn);
            SqlCmdNewID = new SqlCommand("SELECT @@IDENTITY", sqlConn);
            // Add parameters
            for (int iIndex = 0; iIndex < paramList.Length; iIndex++)
            {
                SqlCmd.Parameters.Add(paramList[iIndex]);
            }

            SqlCmd.ExecuteScalar();
            bRetVal = Convert.ToInt32(SqlCmdNewID.ExecuteScalar().ToString());
            SqlCmd.Parameters.Clear();
            //SqlCmd = null;
        }
        catch (Exception ex)
        {
            if (SqlCmd != null)
                SqlCmd.Parameters.Clear();

            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        return bRetVal;
    }

    /// <summary>
    /// Retrieve the scalar from query paramater
    /// </summary>
    /// <param name="csSql"></param>
    /// <returns></returns>
    public int GetOneIntegerWithParams(string csSql, SqlConnection sqlConn, params SqlParameter[] paramList)
    {
        int iRetVal = 0;
        SqlDataReader dtReader = null;
        SqlCommand SqlCmd = null;

        try
        {
            SqlCmd = new SqlCommand(csSql, sqlConn);
            // Add parameters
            for (int iIndex = 0; iIndex < paramList.Length; iIndex++)
            {
                SqlCmd.Parameters.Add(paramList[iIndex]);
            }

            dtReader = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dtReader != null)
            {
                if (dtReader.Read())
                {
                    iRetVal = Int32.Parse(dtReader[dtReader.GetName(0)].ToString());
                }

                dtReader.Close();
            }
            //sqlConn.Close();
            SqlCmd.Parameters.Clear();
            dtReader = null;
            //sqlConn = null;
            SqlCmd = null;
        }
        catch (Exception ex)
        {
            iRetVal = -1;
            if (SqlCmd != null)
                SqlCmd.Parameters.Clear();

            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }
        return iRetVal;
    }

    /// <summary>
    /// Get a dataset from DB
    /// </summary>
    /// <returns></returns>
    public DataSet GetDataSet(String sqlParam, SqlConnection sqlConn)
    {
        //SqlConnection sqlConn = null;
        string sSql = sqlParam;
        DataSet ds = null;

        bool bRet = true;

        try
        {
            ds = new DataSet();
            sSql = sSql.Trim();

            // Get TSL_List table
            SqlDataAdapter da = new SqlDataAdapter(sSql, sqlConn);
            da.Fill(ds, "TABLE1");
            //sqlConn.Close();
            //sqlConn = null;
        }
        catch (Exception ex)
        {
            bRet = false;
            Debug.WriteLine(ex.Message);
            //if (sqlConn.State == ConnectionState.Open)
            //    sqlConn.Close();
        }

        if (!bRet)
            return null;

        return ds;

    }

    /// <summary>
    /// Change Note de Passage
    /// </summary>
    /// <param name="error"></param>
    /// <param name="csCoursOffertID"></param>
    /// <param name="csNotePassage"></param>
    /// <returns></returns>
    public bool UpdateNotePassage(ref string error, string csCoursOffertID, string csNotePassage, SqlConnection myConnection)
    {
        bool retVal = true;
        error = string.Empty;

        SqlTransaction transaction = null;
        try
        {
            SqlCommand cmdUpdateCoursOfferts = new SqlCommand();
            SqlCommand cmdUpdateCoursPris = new SqlCommand();

            transaction = myConnection.BeginTransaction();  // Pas de sauvegarder si l'une des transactions n'est pas à succès


            //Update CoursOfferts
            cmdUpdateCoursOfferts.CommandText = string.Format("UPDATE CoursOfferts SET NotePassage = {0}, CreeParUserName = '{2}' WHERE CoursOffertID = {1}", 
                csNotePassage, csCoursOffertID, GetWindowsUser());
            cmdUpdateCoursOfferts.Connection = myConnection;
            cmdUpdateCoursOfferts.Transaction = transaction;

            //Update CoursPris
            cmdUpdateCoursPris.CommandText = string.Format("UPDATE CoursPris SET NotePassage = {0}, CreeParUserName = '{2}' WHERE CoursOffertID = {1}", 
                csNotePassage, csCoursOffertID, GetWindowsUser());
            cmdUpdateCoursPris.Connection = myConnection;
            cmdUpdateCoursPris.Transaction = transaction;

            try
            {
                cmdUpdateCoursOfferts.ExecuteNonQuery();
                cmdUpdateCoursPris.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                retVal = false;
                error = ex.Message;
            }
            //myConnection.Close();
            //myConnection = null;
        }
        catch (Exception ex)
        {
            retVal = false;
            error = ex.Message;
        }
        return retVal;
    }

    /// <summary>
    /// Telecharger Table Notations
    /// </summary>
    /// <returns></returns>
    public HashSet<Notation> FillNotations(SqlConnection sqlConn)
    {
        HashSet<Notation> hs = new HashSet<Notation>();
        SqlDataReader dtTemp = GetDataReader("SELECT * FROM Notation WHERE Max <> 0", sqlConn);
        if (dtTemp != null)
            while (dtTemp.Read())
            {
                Notation notation = new Notation();
                notation.Min = int.Parse(dtTemp["Min"].ToString());
                notation.Max = int.Parse(dtTemp["Max"].ToString());
                notation.Lettre = dtTemp["Lettre"].ToString();
                notation.NotationLiterale = dtTemp["NotationLiterale"].ToString();
                notation.Moyenne = float.Parse(dtTemp["Moyenne"].ToString());
                hs.Add(notation);
            }

        return hs;
    }

    // Email Stuff
    /// <param name="sEmailTo"></param>
    /// <param name="sBodyMessage"></param>
    /// <param name="sSubject"></param>
    /// <param name="AttachmentPaths"></param>
    /// <returns></returns>
    public bool SendMailToUser(string sEmailTo, string sBodyMessage, string sSubject, List<Attachment> AttachmentPaths)
    {
        bool bRetVal = false;
        SmtpClient smtp = null;
        //Detailed Method
        try
        {
            string sHost = ConfigurationManager.AppSettings["SMTP_HOST"].ToString();
            string sPassWord = ConfigurationManager.AppSettings["PASSWORD"].ToString();
            int iPort = int.Parse(ConfigurationManager.AppSettings["PORT"].ToString());
            string sReturnAddressEmail = ConfigurationManager.AppSettings["RETURN_EMAIL"].ToString();

            MailAddress mailfrom = new MailAddress(sReturnAddressEmail);
            MailAddress mailto = new MailAddress(sEmailTo);
            MailAddress replyto = new MailAddress(sReturnAddressEmail);
            MailMessage newmsg = new MailMessage(mailfrom, mailto);

            newmsg.Subject = sSubject;
            newmsg.Body = sBodyMessage;
            //newmsg.From = replyto;
            newmsg.ReplyTo = replyto;
            newmsg.IsBodyHtml = true;
            
            //For File Attachments
            if (AttachmentPaths != null)
            foreach (Attachment att in AttachmentPaths)
            {                
                newmsg.Attachments.Add(att);
            }
            
            smtp = new SmtpClient(sHost, iPort);
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(sReturnAddressEmail, sPassWord);
            smtp.EnableSsl = true;

            //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(newmsg);
            bRetVal = true;
            
        }
        catch (Exception ex)
        {
            bRetVal = false;
            Debug.WriteLine(ex.Message);
        }
        finally
        {
            if (smtp != null)
            {
                smtp.Dispose();
            }
        }
        return bRetVal;
    }

    // Calendar Stuff

    /// <summary>
    /// Get all events for a given month of a given year
    /// </summary>
    /// <param name="iMonth"></param>
    /// <param name="iYear"></param>
    /// <returns></returns>
    public Hashtable GetMonthEvents(int iMonth, int iYear, SqlConnection sqlConn)
    {
        Hashtable htRetVal = null;
        try
        {
            int iStartDate = 0, iEndDate = 0;
            htRetVal = new Hashtable();

            string sSql = "SELECT * FROM Events WHERE EventYear = @EventYear AND EventMonth = @EventMonth ORDER BY StartDate";
            SqlParameter paramEventYear = new SqlParameter("@EventYear", SqlDbType.Int);
            paramEventYear.Value = iYear;
            SqlParameter paramEventMonth = new SqlParameter("@EventMonth", SqlDbType.Int);
            paramEventMonth.Value = iMonth;

            SqlDataReader dtTemp = GetDataReaderWithParams(sSql, sqlConn, paramEventYear, paramEventMonth);
            while (dtTemp.Read())
            {
                string EventName = dtTemp["eventname"].ToString();
                string LastDay = string.Empty, LastMonth = string.Empty;

                iStartDate = DateTime.Parse(dtTemp["StartDate"].ToString()).Day;
                iEndDate = DateTime.Parse(dtTemp["EndDate"].ToString()).Day;
                LastMonth = DateTime.Parse(dtTemp["EndDate"].ToString()).Month.ToString();

                // Loop through the recordset to fill the hashtable
                for (int iDay = iStartDate; iDay <= iEndDate; iDay++)
                {
                    LastDay = iDay.ToString();

                    if (htRetVal.Contains(LastMonth + LastDay))	// Entry already exist in hastable for this day
                    {
                        // Add more info to the entry
                        EventName = htRetVal[LastMonth + LastDay].ToString() + "\n" + dtTemp["eventname"].ToString();
                        htRetVal.Remove(LastMonth + LastDay);			// delete previous entry
                        htRetVal.Add(LastMonth + LastDay, EventName);	// add entry back
                    }
                    else
                    {
                        // First time writing entry for this day
                        EventName = dtTemp["eventname"].ToString();
                        htRetVal.Add(LastMonth + LastDay, EventName);
                    }
                    //LastDay = dtTemp["day"].ToString();
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            htRetVal = null;
        }

        return htRetVal;
    }
    /// <summary>
    /// 
    /// </summary>
    /// <param name="connectionString"></param>
    ~DB_Access()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            // clean up managed ressources
        }
        // Clean up unmanaged ressources
        //if (sqlConn != null)
        //{
        //    if (sqlConn.State == ConnectionState.Open)
        //        sqlConn.Close();
        //    sqlConn = null;
        //}

    }
}

