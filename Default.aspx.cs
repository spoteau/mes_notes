using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;

// Impersonation
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections;
namespace mes_notes
{
    public partial class Default : System.Web.UI.Page
    {
        
        
        HashSet<Notation> getNotesScheme = null;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {                   
                litBody.Text = ProcessInfo();
            }
        }

        String ProcessInfo()
        {
            String sRetString = ""; // String.Format("<div style=\'page-break-after:always;\'></div>");    // Start with page break in order not to print the button 'print'
            float noteMoyenne = 0.0f, gpa = 0.0f;
            int iYear = 9999, iNbreDeNotes = 0;
            int creditsTotal = 0;
            string credits, sPersonneID = "";
            string username;

            // Get username      
            DB_Access db0 = new DB_Access();
            username = db0.GetWindowsUser();
            if (username.Trim() == String.Empty)
            {
                sRetString += "<br> ERREUR 0 : UserName est vide!";
                return sRetString;
            }

            SqlParameter ParamNumerUsername = new SqlParameter("@Username", SqlDbType.NVarChar);
            ParamNumerUsername.Value = username;
            String sSql = String.Format("SELECT PersonneID FROM Personnes WHERE usernameattribue = '{0}'", username);
            
            using (SqlConnection sqlConn0 = new SqlConnection(ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString()))
            {
                try
                {
                    sqlConn0.Open();
                    //SqlDataReader dtTemp = db0.GetDataReaderWithParams(sSql, sqlConn0, ParamNumerUsername);                    
                    SqlDataReader dtTemp = db0.GetDataReader(sSql, sqlConn0);                    

                    if (dtTemp.Read())
                    {
                        sPersonneID = dtTemp["PersonneID"].ToString();
                        if (sPersonneID == String.Empty)
                        {
                            sRetString += "<br> ERREUR 0 : ID-DB Key Etudiant est vide : ";
                            sRetString += username;
                            return sRetString;
                        }
                    }
                    sqlConn0.Close();
                }
                catch(Exception ex)
                {
                    sRetString += "<br> ERREUR 1 : ID Etudiant!";
                    return sRetString;
                }
            }

            if (sPersonneID == String.Empty)
            {
                sRetString += string.Format("<br> ERREUR 2 : ID Etudiant est vide! ({0})", username);
                return sRetString;
            }

            using (SqlConnection sqlConn0 = new SqlConnection(ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString()))
            {                
                try
                {
                    sqlConn0.Open();
                    getNotesScheme = db0.FillNotations(sqlConn0);                    // Get the different notations
                    sqlConn0.Close();
                }
                catch
                {
                    sRetString += "<br> ERREUR 0 : ProcessInfo!";
                    return sRetString;
                }
            }

            string sDisciplineDeclaree;
            using (SqlConnection sqlConn1 = new SqlConnection(ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString()))
            {
                try
                {
                    DB_Access db1 = new DB_Access();
                    sqlConn1.Open();
                    sDisciplineDeclaree = db1.GetDisciplineEtudiant(sPersonneID, sqlConn1);
                }
                catch
                {
                    sRetString += "<br> ERREUR 1 : ProcessInfo!";
                    return sRetString;
                }
            }

            Dictionary<string, ReleveNotes> dictionaryNotes = new Dictionary<string, ReleveNotes>();
            sSql = String.Format("SELECT CP.NumeroCours, IsNull(CP.DateReussite, '9999-01-01') DateReussite, CP.NoteSurCent, " +
                " CP.NotePassage, C.NomCours, C.Credits, P.EtudiantIdPlus, P.Nom, P.Prenom FROM CoursPris CP, Cours C, Personnes P " +
                " WHERE CP.PersonneID = '{0}' AND CP.NumeroCours = C.NumeroCours AND CP.PersonneID = P.PersonneID  AND C.ExamenEntree = 0 " +
                " AND C.Credits > 0 ORDER BY DateReussite DESC", sPersonneID);
            DB_Access db = new DB_Access();
            using (SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["uespoir_connectionString"].ToString()))
            {
                try
                {
                    sqlConn.Open();

                    SqlDataReader dtTemp = db.GetDataReader(sSql, sqlConn);
                    String PeriodeNoteObtenue = "", PeriodeNoteObtenueOld = "", lettreNote = "";

                    if (dtTemp.Read())
                    {
                        do
                        {
                            PeriodeNoteObtenue = Utils.DateFormattee(dtTemp["DateReussite"].ToString(), out iYear);
                            if (iYear == 9999)
                                continue;

                            lettreNote = Utils.ObtenirLettre((int)float.Parse(dtTemp["NoteSurCent"].ToString()), getNotesScheme, out noteMoyenne);
                            if (lettreNote.ToUpper() != "I" && lettreNote.ToUpper() != "E" && lettreNote.ToUpper() != "F")
                                credits = dtTemp["credits"].ToString();
                            else
                            {
                                credits = "0";
                                //continue;
                            }
                            string NomCours = dtTemp["NomCours"].ToString();
                            float noteSurCent = float.Parse(dtTemp["NoteSurCent"].ToString());
                            float notePassage = float.Parse(dtTemp["NotePassage"].ToString());


                            ReleveNotes releveNote = new ReleveNotes(sPersonneID, dtTemp["Nom"].ToString().ToUpper(), dtTemp["Prenom"].ToString(), dtTemp["EtudiantIdPlus"].ToString(),
                                dtTemp["DateReussite"].ToString(), NomCours, dtTemp["NumeroCours"].ToString(), noteSurCent.ToString("F"),
                                notePassage.ToString("F"), credits, lettreNote, PeriodeNoteObtenue, noteMoyenne.ToString("F"));

                            if (!dictionaryNotes.ContainsKey(NomCours))
                            {
                                dictionaryNotes.Add(dtTemp["NomCours"].ToString(), releveNote);
                            }
                            else
                            {
                                ReleveNotes temp = dictionaryNotes[NomCours];
                                // Utiliser la note la plus grande
                                if (noteSurCent > float.Parse(temp.NoteSurCent))
                                {
                                    // Nouvelle note est supérieure, alors remplacez
                                    dictionaryNotes.Remove(NomCours);
                                    dictionaryNotes.Add(dtTemp["NomCours"].ToString(), releveNote);
                                }
                                noteSurCent.ToString("F");
                            }
                        }
                        while (dtTemp.Read());

                        List<string> keys = new List<string>(dictionaryNotes.Keys);
                        bool needHeader = true;

                        for (int i = 0; i < keys.Count; i++)
                        {
                            ReleveNotes releveNote = dictionaryNotes[keys[i]];

                            if (needHeader)
                            {
                                sRetString += String.Format("<TABLE style='width:80%;align:center'>");
                                sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:center;font-weight:bold;font-size:14px'>Université Espoir</TD></TR>");
                                sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:center;font-weight:bold;font-size:14px'>Historique des Cours<div style='font-color:red'>{0}({1})</div></TD></TR>",
                                    releveNote.Prenom + " " + releveNote.Nom.ToUpper() + " ", releveNote.EtudiantIdPlus);
                                sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:center;font-weight:bold;font-size:14px'>Discipline Déclarée : {0}</TD></TR>",
                                    sDisciplineDeclaree);
                                sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:center;font-weight:bold;font-size:14px'>Date d'Impression: {0}</TD></TR>", DateTime.Today.Date.ToString("dd-MMM-yyyy"));
                                sRetString += String.Format("<TR><TD Colspan='7' width:'80%'><hr style='background-color:#669999;' size='3'/></TD></TR>");
                                sRetString += String.Format("<TR><TD width:'40%' style='text-align:left;font-weight:bold;font-size:14px'>Nom du Cours</TD>" +
                                    "<TD></TD>" +
                                    "<TD style='text-align:center;font-weight:bold;font-size:14px'>Cours</TD>" +
                                    "<TD style='text-align:center;font-weight:bold;font-size:14px'>Note Obtenue</TD>" +
                                    "<TD style='text-align:center;font-weight:bold;font-size:14px'>Lettre Equivalente</TD>" +
                                    "<TD style='text-align:center;font-weight:bold;font-size:14px'>Moyenne (GPA)</TD>" +
                                    "<TD width:'12%' style='text-align:center;font-weight:bold;font-size:14px'>Crédits</TD></TR>");
                                sRetString += String.Format("<TR><TD Colspan='7' width:'80%'><hr style='background-color:#669999;' size='3'/></TD></TR>");
                                needHeader = false;
                            }
                            Boolean firstPass = true;

                            PeriodeNoteObtenue = releveNote.PeriodeNoteObtenue;
                            if (PeriodeNoteObtenue != PeriodeNoteObtenueOld)
                            {
                                if (!firstPass)
                                {
                                    sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:left;font-weight:bold;font-size:14px'></TD></TR>");
                                }
                                sRetString += String.Format("<TR><TD Colspan='7' style='width:80%;text-align:left;font-weight:bold;font-size:14px'>{0}</TD></TR>", PeriodeNoteObtenue);
                                PeriodeNoteObtenueOld = PeriodeNoteObtenue;

                                firstPass = false;
                            }
                            lettreNote = Utils.ObtenirLettre((int)float.Parse(releveNote.NoteSurCent), getNotesScheme, out noteMoyenne);
                            //if (lettreNote.ToUpper() != "I" && lettreNote.ToUpper() != "E" && lettreNote.ToUpper() != "F")
                            //    credits = releveNote.Credits;
                            //else
                            //    credits = "0";
                            if (float.Parse(releveNote.NoteSurCent) >= float.Parse(releveNote.NotePassage))
                                credits = releveNote.Credits;
                            else
                            {
                                credits = "0";
                                //continue;
                            }
                            sRetString += String.Format("<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;{0}</TD><TD></TD><TD style='text-align:center;'>{1}</TD>" +
                            "<TD style='text-align:center;'>{2}</TD>" +
                            "<TD style='text-align:center;'>{3}</TD>" +
                            "<TD style='text-align:center;'>{4}</TD>" +
                            "<TD style='text-align:center;'>{5}</TD></TR>",
                              releveNote.NomCours,
                              releveNote.NumeroCours,
                              releveNote.NoteSurCent,
                              lettreNote,
                              noteMoyenne.ToString("F"),
                              credits);

                            if (noteMoyenne >= 0)
                            {
                                gpa += noteMoyenne;
                                creditsTotal += int.Parse(credits);
                                iNbreDeNotes++;
                            }
                            //}
                            //while (dtTemp.Read());
                        }
                        //db = null;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    sRetString += "<br> ERREUR 2 : ProcessInfo!";
                    //db = null;
                }
            }
            if (iNbreDeNotes > 0)
                gpa = gpa / iNbreDeNotes;
            else
                gpa = 0.0f;
            //sRetString += String.Format("<TR><TD Colspan='6' width:'80%'><hr style='background-color:#669999;' size='2' width='100%'/></TD></TR>");
            //sRetString += String.Format("<TR><TD Colspan='6' style='width:80%;text-align:left;font-weight:bold;font-size:14px'>Moyenne Générale (GPA) : {0}</TD></TR>", gpa.ToString("F"));
            //sRetString += String.Format("<TR><TD Colspan='6' width:'80%'><hr style='background-color:#669999;' size='2' width='100%'/></TD></TR>");


            sRetString += String.Format("<TR><TD Colspan='7' width:'80%'><hr style='background-color:#669999;' size='2' width='100%'/></TD></TR>");
            sRetString += String.Format("<TR><TD Colspan='4' style='width:80%;text-align:left;font-weight:bold;font-size:14px'>Moyenne Générale (GPA) : {0}</TD>", gpa.ToString("F"));
            sRetString += String.Format("<TD Colspan='2' style='width:80%;text-align:right;font-weight:bold;font-size:14px'>Nombre de Crédits :</TD>");
            sRetString += String.Format("<TD style='width:80%;text-align:center;font-weight:bold;font-size:14px'>{0}</TD></TR>", creditsTotal);
            sRetString += String.Format("<TD Colspan='7' width:'80%'><hr style='background-color:#669999;' size='2' width='100%'/></TD></TR>");

            sRetString += "</TABLE>";
            return sRetString;
        }
    }

}
