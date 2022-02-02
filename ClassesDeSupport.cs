using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Collections;
using System.Collections.Specialized;
using System.Globalization;                 // CultureInfo

// Impersonation
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Threading;

/// <summary>
/// Summary description for ClassesDeSupport
/// </summary>
public class ClassesDeSupport
{
    public ClassesDeSupport()
    {
    }
}

public class ReleveNotes
{
    private string _PersonneID, _NumeroCours, _DateReussite, _NoteSurCent, _NotePassage,
        _NomCours, _Credits, _EtudiantIdPlus, _Nom, _Prenom, _LettreNote, _PeriodeNoteObtenue, _NoteMoyenne;

    public string PersonneID
    {
        get { return _PersonneID; }
        set { _PersonneID = value; }
    }

    public string NumeroCours
    {
        get { return _NumeroCours; }
        set { _NumeroCours = value; }
    }

    public string DateReussite
    {
        get { return _DateReussite; }
        set { _DateReussite = value; }
    }
    public string NoteSurCent
    {
        get { return _NoteSurCent; }
        set { _NoteSurCent = value; }
    }
    public string NotePassage
    {
        get { return _NotePassage; }
        set { _NotePassage = value; }
    }
    public string NomCours
    {
        get { return _NomCours; }
        set { _NomCours = value; }
    }
    public string Credits
    {
        get { return _Credits; }
        set { _Credits = value; }
    }
    public string EtudiantIdPlus
    {
        get { return _EtudiantIdPlus; }
        set { _EtudiantIdPlus = value; }
    }
    public string Nom
    {
        get { return _Nom; }
        set { _Nom = value; }
    }
    public string Prenom
    {
        get { return _Prenom; }
        set { _Prenom = value; }
    }
    public string LettreNote
    {
        get { return _LettreNote; }
        set { _LettreNote = value; }
    }

    public string PeriodeNoteObtenue
    {
        get { return _PeriodeNoteObtenue; }
        set { _PeriodeNoteObtenue = value; }
    }

    public string NoteMoyenne
    {
        get { return _NoteMoyenne; }
        set { _NoteMoyenne = value; }
    }
    public ReleveNotes(String unePersonneID, String unNom, String unPrenom, String unEtudiantIdPlus, String uneDateReussite, String unNomCours, String unNumeroCours, 
        String uneNoteSurCent, String uneNotePassage, String desCredits, String uneLettreNote, String unePeriodeNoteObtenue, String uneNoteMoyenne)
    {
        PersonneID = unePersonneID;
        NumeroCours = unNumeroCours; DateReussite= uneDateReussite; NoteSurCent= uneNoteSurCent;
        NotePassage = uneNotePassage; NomCours = unNomCours; Credits= desCredits;
        EtudiantIdPlus = unEtudiantIdPlus; Nom= unNom; Prenom = unPrenom; LettreNote = uneLettreNote;
        PeriodeNoteObtenue = unePeriodeNoteObtenue;
        NoteMoyenne = uneNoteMoyenne;
    }

}

public class Utils
{
    public static String FixDate(String sDate)
    {
        try
        {
            if (sDate.Trim() == String.Empty)
                return "";
            DateTime dt = DateTime.Parse(sDate);

            String st = dt.Day.ToString() + "-" + dt.ToString("m", CultureInfo.CreateSpecificCulture("en-fr")).Substring(0, 3) + "-" + dt.Year.ToString();
            return st;
            //return sDate.Substring(0, sDate.IndexOf(" "));
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            return sDate;
        }
    }

    public static String DateFormattee(String sDate, out int iOutYear)
    {
        iOutYear = 9999;
        try
        {
            if (sDate.Trim() == String.Empty)
                return "";
            DateTime dt = DateTime.Parse(sDate);
            String Periode = "";
            int iMonth = dt.Month;
            int iYear = dt.Year;
            if (iYear == 9999)
            {
                return "Session Courante";
            }
            if (iMonth < 4)
                Periode = "Hiver";
            else if (iMonth < 7)
                Periode = "Printemps";
            else if (iMonth < 10)
                Periode = "Eté";
            else
                Periode = "Automne";
            String st = Periode + " " + iYear.ToString();
            iOutYear = iYear;

            return st;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            return sDate;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="NoteSurCent"></param>
    /// <param name="moyenne"></param>
    /// <returns></returns>
    public static String ObtenirLettre(int NoteSurCent, HashSet<Notation> getNotesScheme, out float moyenne)
    {
        String retVal = "";
        moyenne = 0.0f;

        try
        {
            foreach (Notation n in getNotesScheme)
            {
                if (NoteSurCent >= n.Min && NoteSurCent <= n.Max)
                {
                    moyenne = n.Moyenne;
                    retVal = n.Lettre;
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
        }
        return retVal;

    }

}

public class Notation
{
    public Notation()
    {
    }

    private int _Min;
    public int Min
    {
        get { return _Min; }
        set { _Min = value; }
    }

    private int _Max;
    public int Max
    {
        get { return _Max; }
        set { _Max = value; }
    }

    private string _Lettre;
    public string Lettre
    {
        get { return _Lettre; }
        set { _Lettre = value; }
    }

    private string _NotationLiterale;
    public string NotationLiterale
    {
        get { return _NotationLiterale; }
        set { _NotationLiterale = value; }
    }

    private float _Moyenne;
    public float Moyenne
    {
        get { return _Moyenne; }
        set { _Moyenne = value; }
    }
}

public class EtudiantInfo
{
    public EtudiantInfo()
    {
        _BalanceDue = 0.00;
    }

    private string _PersonneID;

    public string PersonneID
    {
        get { return _PersonneID; }
        set { _PersonneID = value; }
    }
    private string _Prenom;

    public string Prenom
    {
        get { return _Prenom; }
        set { _Prenom = value; }
    }
    private string _Nom;

    public string Nom
    {
        get { return _Nom; }
        set { _Nom = value; }
    }
    private string _Telephone1;

    public string Telephone1
    {
        get { return _Telephone1; }
        set { _Telephone1 = value; }
    }
    private string _Telephone2;

    public string Telephone2
    {
        get { return _Telephone2; }
        set { _Telephone2 = value; }
    }
    private string _DDN;

    public string DDN
    {
        get { return _DDN; }
        set { _DDN = value; }
    }

    private string _AdresseRue;

    public string AdresseRue
    {
        get { return _AdresseRue; }
        set { _AdresseRue = value; }
    }
    private string _AdresseExtra;

    public string AdresseExtra
    {
        get { return _AdresseExtra; }
        set { _AdresseExtra = value; }
    }
    private string _Ville;

    public string Ville
    {
        get { return _Ville; }
        set { _Ville = value; }
    }
    private string _Pays;

    public string Pays
    {
        get { return _Pays; }
        set { _Pays = value; }
    }
    private DateTime _DateCreee;

    public DateTime DateCreee
    {
        get { return _DateCreee; }
        set { _DateCreee = value; }
    }
    private string _Remarque;

    public string Remarque
    {
        get { return _Remarque; }
        set { _Remarque = value; }
    }
    private int _Etudiant;

    public int Etudiant
    {
        get { return _Etudiant; }
        set { _Etudiant = value; }
    }
    private int _AdminStaff;

    public int AdminStaff
    {
        get { return _AdminStaff; }
        set { _AdminStaff = value; }
    }
    private string _Photo;

    public string Photo
    {
        get { return _Photo; }
        set { _Photo = value; }
    }
    private string _UserNameAttribue;

    public string UserNameAttribue
    {
        get { return _UserNameAttribue; }
        set { _UserNameAttribue = value; }
    }
    private string _CreeParUsername;

    public string CreeParUsername
    {
        get { return _CreeParUsername; }
        set { _CreeParUsername = value; }
    }
    private string _NumeroRecu;

    public string NumeroRecu
    {
        get { return _NumeroRecu; }
        set { _NumeroRecu = value; }
    }
    private string _NIF;

    public string NIF
    {
        get { return _NIF; }
        set { _NIF = value; }
    }
    private string _email;

    public string Email
    {
        get { return _email; }
        set { _email = value; }
    }
    private string _DisciplineInitiale;

    public string DisciplineInitiale
    {
        get { return _DisciplineInitiale; }
        set { _DisciplineInitiale = value; }
    }
    private string _ExamEntree_Fran;

    public string ExamEntree_Fran
    {
        get { return _ExamEntree_Fran; }
        set { _ExamEntree_Fran = value; }
    }
    private string _ExamEntree_Math;

    public string ExamEntree_Math
    {
        get { return _ExamEntree_Math; }
        set { _ExamEntree_Math = value; }
    }
    private int _DisciplineID;

    public int DisciplineID
    {
        get { return _DisciplineID; }
        set { _DisciplineID = value; }
    }

    private Double _BalanceDue;
    public Double BalanceDue
    {
        get { return _BalanceDue; }
        set { _BalanceDue = value; }
    }

    private int _Actif;
    public int Actif 
    {
        get { return _Actif; }
        set { _Actif = value; }
    }

    private string _Discipline;

    public string Discipline
    {
        get { return _Discipline; }
        set { _Discipline = value; }
    }

    public void GetInfoEtudiant(string sPersonneID)
    {
        EtudiantInfo returnedInfo = new EtudiantInfo();

        DB_Access db = new DB_Access();
        // Loop through all records
        try
        {
            string sSql = String.Format("SELECT Nom, Prenom, IsNull(Telephone1, '') AS Telephone1, IsNull(Telephone2, '') AS Telephone2, IsNull(DDN, '') AS DDN," +
                " IsNull(AdresseRue, '') AS AdresseRue, IsNull(AdresseExtra, '') AS AdresseExtra, IsNull(Ville, '') AS Ville, Pays, IsNull(DateCreee, '') AS DateCreee, IsNull(Remarque, '') AS Remarque, " +
                " Etudiant, AdminStaff, IsNull(Photo, 0) AS Photo, IsNull(UserNameAttribue, '') AS UserNameAttribue, IsNull(CreeParUserName, '') AS CreeParUserName, IsNull(NumeroRecu, '') AS NumeroRecu, IsNull(NIF, '') AS NIF, " +
                " IsNull(email, '') AS email, IsNull(DisciplineInitiale, '') AS DisciplineInitiale, IsNull(ExamEntree_Fran, '') AS ExamEntree_Fran, IsNull(ExamEntree_Math, '') AS ExamEntree_Math, " +
                " Personnes.DisciplineID, Actif, DisciplineNom FROM Personnes, Disciplines WHERE Personnes.DisciplineID = Disciplines.DisciplineID AND PersonneID = '{0}'", sPersonneID);
            using (SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection(ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString()))
            {
                sqlConn.Open();
                SqlDataReader dt = db.GetDataReader(sSql, sqlConn);
                if (dt != null)
                {
                    if (dt.Read())
                    {
                        PersonneID = sPersonneID;
                        Nom = dt["Nom"].ToString();
                        Prenom = dt["Prenom"].ToString();
                        Telephone1 = dt["Telephone1"].ToString();
                        Telephone2 = dt["Telephone2"].ToString();
                        DDN = dt["DDN"].ToString();
                        AdresseRue = dt["AdresseRue"].ToString();
                        AdresseExtra = dt["AdresseExtra"].ToString();
                        Ville = dt["Ville"].ToString();
                        Pays = dt["Pays"].ToString();
                        //
                        if (dt["DateCreee"].ToString() != String.Empty)
                        {
                            returnedInfo.DateCreee = DateTime.Parse(dt["DateCreee"].ToString());
                        }
                        Remarque = dt["Remarque"].ToString();
                        Etudiant = Int16.Parse(dt["Etudiant"].ToString());
                        AdminStaff = Int16.Parse(dt["AdminStaff"].ToString());

                        //returnedInfo.Photo = dt[""].ToString();
                        UserNameAttribue = dt["UserNameAttribue"].ToString();
                        CreeParUsername = dt["CreeParUserName"].ToString();
                        NumeroRecu = dt["NumeroRecu"].ToString();
                        NIF = dt["NIF"].ToString();
                        Email = dt["email"].ToString();
                        DisciplineInitiale = dt["DisciplineInitiale"].ToString();
                        ExamEntree_Fran = dt["ExamEntree_Fran"].ToString();
                        ExamEntree_Math = dt["ExamEntree_Math"].ToString();
                        DisciplineID = Int16.Parse(dt["DisciplineID"].ToString());
                        Actif = Int16.Parse(dt["Actif"].ToString());
                        Discipline = dt["DisciplineNom"].ToString();

                        //
                        BalanceDue = db.BalanceEtudiant(PersonneID, sqlConn);
                    }
                }
            }
            db = null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            db = null;
        }
    }
}

public class SystemCalls
{
    public const int LOGON32_LOGON_INTERACTIVE = 2;
    public const int LOGON32_PROVIDER_DEFAULT = 0;

    public static WindowsImpersonationContext impersonationContext;

    [DllImport("advapi32.dll")]
    public static extern int LogonUserA(String lpszUserName,
        String lpszDomain,
        String lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        ref IntPtr phToken);
    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int DuplicateToken(IntPtr hToken,
        int impersonationLevel,
        ref IntPtr hNewToken);

    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool RevertToSelf();

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern bool CloseHandle(IntPtr handle);
    protected static string Messages = String.Empty;
    public static bool impersonateValidUser(String userName, String domain, String password)
    {
        WindowsIdentity tempWindowsIdentity;
        IntPtr token = IntPtr.Zero;
        IntPtr tokenDuplicate = IntPtr.Zero;

        if (RevertToSelf())
        {
            if (LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE,
                LOGON32_PROVIDER_DEFAULT, ref token) != 0)
            {
                if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                {
                    tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                    impersonationContext = tempWindowsIdentity.Impersonate();
                    if (impersonationContext != null)
                    {
                        CloseHandle(token);
                        CloseHandle(tokenDuplicate);
                        return true;
                    }
                }
            }
        }
        if (token != IntPtr.Zero)
            CloseHandle(token);
        if (tokenDuplicate != IntPtr.Zero)
            CloseHandle(tokenDuplicate);
        return false;
    }

    public static void undoImpersonation()
    {
        impersonationContext.Undo();
    }
}
