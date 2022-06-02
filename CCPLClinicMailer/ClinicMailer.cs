using CCPLClinicMailer.DAL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CCPLClinicMailer
{
    partial class ClinicMailer : ServiceBase
    {
        DB_Helper help = new DB_Helper();
        DataTable dt = new DataTable();
        DataTable dtEmail = new DataTable();
        DataTable dtEmailTo = new DataTable();
        DataTable dtEmailCC = new DataTable();
        string Query, EmailModule, EmailSubject, EmailMsg, EmailDocNo, EmailSiteID;
        string MailServer, MailUsername, MailPass, MailPortNo, MailName, EmailFooter;
        bool IsMailed, IsSSL;
        MailMessage objMailMessage = new MailMessage();

        public ClinicMailer()
        {
            InitializeComponent();
        }

        private void App_Log(string Error)
        {
            EventLog eventLog = new EventLog("Application");
            eventLog.Source = "Application";
            eventLog.WriteEntry(Error, EventLogEntryType.Error, 101, 1);
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            Modules();
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            Modules();
        }
        void Refresh_Controls()
        {
            help = new DB_Helper();
            dt = new DataTable();
            dtEmail = new DataTable();
            dtEmailTo = new DataTable();
            dtEmailCC = new DataTable();

            Query = string.Empty;

            EmailModule = string.Empty;
            //EmailSubject = string.Empty;
            EmailMsg = string.Empty;
            //EmailDocNo = string.Empty;
            //EmailSiteID = string.Empty;

            MailServer = string.Empty;
            MailUsername = string.Empty;
            MailPass = string.Empty;
            MailPortNo = string.Empty;
            MailName = string.Empty;
            IsSSL = new bool();

            EmailFooter = string.Empty;
            IsMailed = new bool();

            //objMailMessage = new MailMessage();
        }
        public void MailServerSetting()
        {
            try
            {
                Refresh_Controls();
                Query = @"SELECT  * From EmailSetup";
                dt = new DataTable();
                dt = help.Return_DataTable_Query(Query);
                if (dt.Rows.Count > 0)
                {
                    MailServer = dt.Rows[0]["MailSMTP"].ToString().Trim();
                    MailUsername = dt.Rows[0]["MailID"].ToString().Trim();
                    MailPass = dt.Rows[0]["MailPassword"].ToString().Trim();
                    MailPortNo = dt.Rows[0]["MailPortNo"].ToString().Trim();
                    MailName = dt.Rows[0]["MailName"].ToString().Trim();
                    IsSSL = bool.Parse(dt.Rows[0]["IsSSL"].ToString().Trim());
                }
                else
                {
                    //===========Error Log===
                    //help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", "Email Server Setting", "", "CCPL Mailer Server Setting Values Null");
                    App_Log("CCPL Mailer Server Setting Values Null");
                    MailServer = null;
                    MailUsername = null;
                    MailPass = null;
                    MailPortNo = null;
                    MailName = null;
                    IsSSL = false;
                    return;
                }

            }
            catch (Exception ex)
            {
                EventLog eventLog = new EventLog("Application");
                eventLog.Source = "Application";
                eventLog.WriteEntry("CCPL Mailer Service (Mailing Server Configuration Error) : " + ex.Message, EventLogEntryType.Error, 101, 1);
                help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", "Email Server Setting", "", ex.Message);
            }
        }
        private void Generate_EmailFooter(int TabWidth)
        {
            EmailFooter = "";
            EmailFooter += "<br/>";
            EmailFooter += "<br/>";
            EmailFooter += "<table style='width:" + TabWidth + "%;' cellspacing='1' cellpadding='1' border='1'>";
            EmailFooter += "<tr><td style='background-color:purple; color:gold;text-align:center;font-size:10px;'><b>System Generated Email - Powered By I.T Department</td></tr>";
            EmailFooter += "</table>";
        }

        private void Get_Email_List(string EmailModule, string SiteID)
        {
            try
            {
                Query = @"SELECT * FROM VW_Email_Active_List where Module='" + EmailModule + "' AND SiteID='" + SiteID + "'";
                //Query = @"SELECT TOP 1 * FROM VW_Email_Active_List where EmailID = 'usman.khalid@cornerclinic.net' ";
                dtEmail = new DataTable();
                dtEmail = help.Return_DataTable_Query(Query);
                if (dtEmail.Rows.Count > 0)
                {
                    dtEmailTo = new DataTable();
                    Query = @"SELECT * FROM VW_Email_Active_List WHERE Module='" + EmailModule + "' AND IsCC = 0 AND SiteID = '" + SiteID + "'";
                    //Query = @"SELECT TOP 1 * FROM VW_Email_Active_List where EmailID = 'usman.khalid@cornerclinic.net' ";
                    dtEmailTo = help.Return_DataTable_Query(Query);

                    dtEmailCC = new DataTable();
                    Query = @"SELECT * FROM VW_Email_Active_List WHERE Module='" + EmailModule + "' AND IsCC = 1 AND SiteID='" + SiteID + "'";
                    dtEmailCC = help.Return_DataTable_Query(Query);
                }
                else
                {
                    //Nothing Here............
                }
            }
            catch (Exception ex)
            {
                EventLog eventLog = new EventLog("Application");
                eventLog.Source = "Application";
                eventLog.WriteEntry("GP Mailer Service (Getting Email List) : " + ex.Message, EventLogEntryType.Error, 101, 1);
                help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", "Getting Email List", "", ex.Message);
            }
        }
        public void SendEmail(string EmailMessage)
        {
            try
            {
                objMailMessage = new MailMessage();
                bool pass = true;
                if (dtEmailTo == null)
                {
                    sendElse(EmailMessage);
                }
                else if (pass && dtEmailTo.Rows.Count > 0)
                {
                    IsMailed = false;
                    for (int i = 0; i <= dtEmailTo.Rows.Count - 1; i++)
                    {
                        objMailMessage.To.Add(new MailAddress(dtEmailTo.Rows[i]["EmailID"].ToString().Trim(), ""));
                    }

                    for (int i = 0; i <= dtEmailCC.Rows.Count - 1; i++)
                    {
                        objMailMessage.CC.Add(new MailAddress(dtEmailCC.Rows[i]["EmailID"].ToString().Trim(), ""));
                    }
                    //objMailMessage.To.Add(new MailAddress("usman.khalid@colonytextiles.com", ""));
                    Send_Email_SMTP(EmailMessage);
                }
                else
                {
                    sendElse(EmailMessage);

                }
            }
            catch (Exception ex)
            {
                IsMailed = false;
                App_Log("CCPL Mailer Service (Send Email) Error :" + EmailModule + ex.Message);
                help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", "Send Email", "", ex.Message);
            }
        }

        private void sendElse(string EmailMessage)
        {
            if (dtEmailCC.Rows.Count > 0)
            {
                for (int i = 0; i <= dtEmailCC.Rows.Count - 1; i++)
                {
                    objMailMessage.To.Add(new MailAddress(dtEmailCC.Rows[i]["EmailID"].ToString().Trim(), ""));
                }
            }
            Send_Email_SMTP(EmailMessage);
        }
        public void Send_Email_SMTP(string Msg)
        {
            try
            {
                if (Msg != "")
                {
                    MailServerSetting();
                    NetworkCredential loginInfo = new NetworkCredential(MailUsername, MailPass);
                    objMailMessage.From = new MailAddress(MailUsername, MailName);
                    objMailMessage.IsBodyHtml = true;
                    objMailMessage.Subject = EmailSubject;
                    objMailMessage.Body = Msg;
                    SmtpClient objSmtp = new SmtpClient();
                    objSmtp.Host = MailServer;
                    objSmtp.EnableSsl = IsSSL;
                    objSmtp.Port = Convert.ToInt32(MailPortNo);
                    objSmtp.Credentials = loginInfo;
                    objSmtp.Send(objMailMessage);
                    objMailMessage.To.Clear();
                    objMailMessage.CC.Clear();
                    objMailMessage.Attachments.Clear();
                    objSmtp.Timeout = 10000;
                    IsMailed = true;
                    //Refresh_Controls();
                    EmailMsg = "";
                }
                else
                {
                    IsMailed = false;
                }
            }
            catch (Exception ex)
            {
                IsMailed = false;
                App_Log("CCPL Mailer Service (Send Email SMTP) Error :" + EmailModule + ex.Message);
                help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", "Send Email SMTP", "", ex.Message);
            }
        }


        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Modules();
        }

        void Modules()
        {
            timer1.Enabled = false;
            Sale_Summary_Night();
            Sale_Summary_Morning();
            timer1.Enabled = true;
        }
        private bool Check_Mail_Time(string MailModule, string SiteID)
        {
            try
            {
                Query = @"SELECT * FROM EmailShoot_Log WHERE Module = '" + MailModule + "' AND SiteID = '" + SiteID + "' AND MailDate = '" + DateTime.Now.ToShortDateString() + "'";
                dt = new DataTable();
                dt = help.Return_DataTable_Query(Query);
                if (dt.Rows.Count > 0)
                {
                    //Nothing to do
                    return false;
                }
                else
                {
                    Query = @"SELECT * FROM Email_Timer WHERE Module='" + MailModule + "' AND SiteID = '" + SiteID + "' ";
                    dt = new DataTable();
                    dt = help.Return_DataTable_Query(Query);
                    if (dt.Rows.Count > 0)
                    {
                        TimeSpan NowTime = DateTime.Now.TimeOfDay;
                        TimeSpan StartTime = DateTime.Parse(dt.Rows[0]["TimeFrom"].ToString()).TimeOfDay;
                        TimeSpan EndTime = DateTime.Parse(dt.Rows[0]["TimeTo"].ToString()).TimeOfDay;
                        if (StartTime <= NowTime && EndTime >= NowTime)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                            //Nothing to do
                        }
                    }
                    else
                    {
                        return false;
                        //Nothing to do
                    }
                }
            }
            catch (Exception ex)
            {
                App_Log("Clinic Mailer Service (Check Email Timer Error) :" + ex.Message);
                //help.DatabaseName = "DYNAMICS";
                //help.ExecuteParameterizedProcedure("UD_Error_ForMail", "@DocNo,@Company,@Module,@Error", "", "", "Check Email Time", ex.Message);
                return false;
            }
        }

        private void Send_Schedule_Email(string MailModule, string SiteID)
        {
            try
            {
                //help.DatabaseName = "CML";
                Query = @"SELECT * FROM EmailShoot_Log WHERE Module = '" + MailModule + "' AND SiteID = '" + SiteID + "' AND MailDate = '" + DateTime.Now.ToShortDateString() + "'";
                dt = new DataTable();
                dt = help.Return_DataTable_Query(Query);
                if (dt.Rows.Count > 0)
                {
                    //Nothing to do
                }
                else
                {
                    Query = @"SELECT * FROM Email_Timer WHERE Module='" + MailModule + "' AND SiteID = '" + SiteID + "' ";
                    dt = new DataTable();
                    dt = help.Return_DataTable_Query(Query);
                    if (dt.Rows.Count > 0)
                    {
                        TimeSpan NowTime = DateTime.Now.TimeOfDay;
                        TimeSpan StartTime = DateTime.Parse(dt.Rows[0]["TimeFrom"].ToString()).TimeOfDay;
                        TimeSpan EndTime = DateTime.Parse(dt.Rows[0]["TimeTo"].ToString()).TimeOfDay;
                        if (StartTime <= NowTime && EndTime >= NowTime)
                        {
                            SendEmail(EmailMsg);
                            //=====Update IsMailed==========
                            if (IsMailed)
                            {
                                Query = @"INSERT INTO EmailShoot_Log (Module,SiteID) VALUES ('" + MailModule + "','" + SiteID + "') ";
                                help.Execute_Query(Query);
                            }
                        }
                        else
                        {
                            //Nothing to do
                        }
                    }
                    else
                    {
                        //Nothing to do
                    }

                }
            }
            catch (Exception ex)
            {
                App_Log("Clinic Mailer Service (Check Schedule Email By (" + MailModule + ") Error) :" + EmailModule + ex.Message);
                //help.DatabaseName = "DYNAMICS";
                //help.ExecuteParameterizedProcedure("UD_Error_ForMail", "@DocNo,@Company,@Module,@Error", "", "CML", "Check Schedule Email", ex.Message);
                //=========Save History=============
                //help.DatabaseName = "DYNAMICS";
                //help.ExecuteParameterizedProcedure("UD_GPMailer_ErroLog", "@DocNo,@Module,@IsError,@IsMailed,@MailBody,@IsAttachment,@ErrorMsg,@CompanyID", "", MailModule, 1, 0, EmailMsg, 0, ex.Message, "CML");
            }
        }

        public void Sale_Summary_Night()
        {
            try
            {
                Refresh_Controls();
                EmailModule = "Sale Summary Night";
                EmailSiteID = "HO";
                if(Check_Mail_Time(EmailModule, EmailSiteID))
                {
                    DateTime Today = DateTime.Parse(DateTime.Now.ToShortDateString());
                    Query = "SELECT * FROM VW_Sale_Summary_DateWise_ForMail WHERE TokenDate = '" + Today.ToString("MM/dd/yyyy") + "' ORDER BY TokenDate,SiteID,StaffShift,TypeID";
                    dt = new DataTable();
                    dt = help.Return_DataTable_Query(Query);
                    if (dt.Rows.Count > 0)
                    {
                        Get_Email_List(EmailModule, EmailSiteID);
                        //Get_Email_List("Test", EmailSiteID);
                        EmailMsg = "";
                        string SiteID = "";
                        int TabWidth = 50;
                        string Sub = "Clinic Sale Summary of : " + Convert.ToDateTime(dt.Rows[0]["TokenDate"]).ToString("dd-MMM-yyyy").Trim();
                        EmailSubject = Sub;
                        EmailMsg += "<table style='width:" + TabWidth + "%;' cellspacing='1' cellpadding='1' border='1'>";
                        EmailMsg += "<tr><td colspan='4' style='background-color:purple; color:gold;text-align:center;'><b>" + Sub + "</td></tr>";

                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th>#</th>";
                        EmailMsg += "<th>Type</th>";
                        EmailMsg += "<th>Shift</th>";
                        EmailMsg += "<th>Amount</th>";
                        EmailMsg += "</tr>";
                        double TotalAmount = 0;
                        double GrandTotal = 0;
                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th colspan='4' style='text-align:left;'>" + dt.Rows[0]["SiteName"].ToString().Trim() + " [ " + dt.Rows[0]["SiteID"].ToString().Trim() + " ] </th>";
                        EmailMsg += "</tr>";


                        int SNo = 0;
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {
                            if (!SiteID.Equals(""))
                            {
                                if (SiteID.Equals(dt.Rows[i]["SiteID"].ToString().Trim()))
                                {
                                    TotalAmount += double.Parse(dt.Rows[i]["TotalAmount"].ToString().Trim());
                                    GrandTotal += double.Parse(dt.Rows[i]["TotalAmount"].ToString().Trim());
                                    EmailMsg += "<tr>";
                                    EmailMsg += "<td>" + SNo + "</td>";
                                    EmailMsg += "<td>" + dt.Rows[i]["TypeID"].ToString().Trim() + "</td>";
                                    EmailMsg += "<td>" + dt.Rows[i]["StaffShift"].ToString().Trim() + "</td>";
                                    EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["TotalAmount"]).ToString("n0").Trim() + "</td>";
                                    EmailMsg += "</tr>";
                                    SNo++;
                                }
                                else
                                {
                                    SNo = 1;
                                    EmailMsg += "<tr style='background-color:purple; color:white;'>";
                                    EmailMsg += "<th colspan='3'>Total of [ " + SiteID + " ]</th>";
                                    EmailMsg += "<th style='text-align:right;'>" + TotalAmount.ToString("N0") + "</th>";
                                    EmailMsg += "</tr>";
                                    SiteID = dt.Rows[i]["SiteID"].ToString().Trim();
                                    TotalAmount = 0;
                                    EmailMsg += "<tr style='background-color:purple; color:white;'>";
                                    EmailMsg += "<th colspan='4' style='text-align:left;'>" + dt.Rows[i]["SiteName"].ToString().Trim() + " [ " + dt.Rows[i]["SiteID"].ToString().Trim() + " ] </th>";
                                    EmailMsg += "</tr>";

                                    i--;
                                }
                            }
                            else
                            {
                                SNo = 1;
                                SiteID = dt.Rows[i]["SiteID"].ToString();
                                i--;
                            }
                        }

                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th colspan='3'>Total of [ " + SiteID + " ]</th>";
                        EmailMsg += "<th style='text-align:right;'>" + TotalAmount.ToString("N0") + "</th>";
                        EmailMsg += "</tr>";

                        //=========Grand Total=============
                        EmailMsg += "<tr style='background-color:purple; color:gold;'>";
                        EmailMsg += "<th colspan='3'>Grand Total </th>";
                        EmailMsg += "<th style='text-align:right;'>" + GrandTotal.ToString("N0") + "</th>";
                        EmailMsg += "</tr>";
                        EmailMsg += "</table>";
                        Generate_EmailFooter(TabWidth);
                        EmailMsg += EmailFooter;
                        //SendEmail(EmailMsg);
                        Send_Schedule_Email(EmailModule, EmailSiteID);

                    }
                    else
                    {
                        //Do Nothing..........
                    }
                }
                
            }
            catch (Exception ex)
            {
                App_Log("CCPL Pharmacy Mailer Service (Sale Summary Night) Error :" + EmailModule + ex.Message);
                //help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", EmailModule, EmailSiteID, ex.Message);
            }
        }


        public void Sale_Summary_Morning()
        {
            try
            {
                Refresh_Controls();
                EmailModule = "Sale Summary Morning";
                EmailSiteID = "HO";
                if (Check_Mail_Time(EmailModule, EmailSiteID))
                {
                    //=======================Consultancy=========================
                    #region Consultancy
                    DateTime Today = DateTime.Parse(DateTime.Now.AddDays(-1).ToShortDateString());
                    Query = "SELECT * FROM VW_Sale_Summary_Clinic_ForMail WHERE TokenDate = '" + Today.ToString("MM/dd/yyyy") + "' ORDER BY TokenDate, SiteID";
                    dt = new DataTable();
                    dt = help.Return_DataTable_Query(Query);
                    if (dt.Rows.Count > 0)
                    {
                        Get_Email_List(EmailModule, EmailSiteID);
                        EmailMsg = "";
                        string SiteID = "";
                        string Sub = "Clinic Sale Detail of : " + Today.ToString("dd-MMM-yyyy").Trim();
                        EmailSubject = Sub;
                        int TabWidth = 70;
                        EmailMsg += "<table style='width:" + TabWidth + "%;' cellspacing='1' cellpadding='1' border='1'>";
                        EmailMsg += "<tr><td colspan='7' style='background-color:purple; color:gold;text-align:center;'><b>" + Sub + "</td></tr>";
                        EmailMsg += "<tr><td colspan='7' style='background-color:purple; color:gold;text-align:center;'><b> Consultancy </td></tr>";

                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th>#</th>";
                        EmailMsg += "<th>Doctor</th>";
                        EmailMsg += "<th>Total Patient</th>";
                        EmailMsg += "<th>Total</th>";
                        EmailMsg += "<th>Discount</th>";
                        EmailMsg += "<th>Refund</th>";
                        EmailMsg += "<th>Net Amount</th>";
                        EmailMsg += "</tr>";
                        //===========TOtal====
                        double TotalAmount = 0;
                        double TotalPatient = 0;
                        double TotalDiscount = 0;
                        double TotalRefund = 0;
                        double TotalPayable = 0;
                        //=========Grand Total======
                        double GrandTotal = 0;
                        double GrandPatient = 0;
                        double GrandDiscount = 0;
                        double GrandRefund = 0;
                        double GrandPayable = 0;


                        double TotalConsultancy = 0;
                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th colspan='7' style='text-align:left;'>" + dt.Rows[0]["SiteName"].ToString().Trim() + " [ " + dt.Rows[0]["SiteID"].ToString().Trim() + " ] </th>";
                        EmailMsg += "</tr>";


                        int SNo = 0;
                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                        {
                            if (!SiteID.Equals(""))
                            {
                                if (SiteID.Equals(dt.Rows[i]["SiteID"].ToString().Trim()))
                                {
                                    TotalConsultancy += double.Parse(dt.Rows[i]["TotalPaid"].ToString().Trim());
                                    //===========TOtal========
                                    TotalAmount += double.Parse(dt.Rows[i]["TotalAmount"].ToString().Trim());
                                    TotalPatient += double.Parse(dt.Rows[i]["TotalPatient"].ToString().Trim());
                                    TotalDiscount += double.Parse(dt.Rows[i]["Discount"].ToString().Trim());
                                    TotalRefund += double.Parse(dt.Rows[i]["RefundFee"].ToString().Trim());
                                    TotalPayable += double.Parse(dt.Rows[i]["TotalPaid"].ToString().Trim());
                                    //===========Grand Total===========
                                    GrandTotal += double.Parse(dt.Rows[i]["TotalAmount"].ToString().Trim());
                                    GrandPatient += double.Parse(dt.Rows[i]["TotalPatient"].ToString().Trim());
                                    GrandDiscount += double.Parse(dt.Rows[i]["Discount"].ToString().Trim());
                                    GrandRefund += double.Parse(dt.Rows[i]["RefundFee"].ToString().Trim());
                                    GrandPayable += double.Parse(dt.Rows[i]["TotalPaid"].ToString().Trim());

                                    EmailMsg += "<tr>";
                                    EmailMsg += "<td>" + SNo + "</td>";
                                    EmailMsg += "<td>" + dt.Rows[i]["DrName"].ToString().Trim() + "</td>";
                                    if (double.Parse(dt.Rows[i]["TotalPatient"].ToString().Trim()) > 0)
                                    {
                                        EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["TotalPatient"]).ToString("n0").Trim() + "</td>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (double.Parse(dt.Rows[i]["TotalAmount"].ToString().Trim()) > 0)
                                    {
                                        EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["TotalAmount"]).ToString("n0").Trim() + "</td>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (double.Parse(dt.Rows[i]["Discount"].ToString().Trim()) > 0)
                                    {
                                        EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["Discount"]).ToString("n0").Trim() + "</td>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (double.Parse(dt.Rows[i]["RefundFee"].ToString().Trim()) > 0)
                                    {
                                        EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["RefundFee"]).ToString("n0").Trim() + "</td>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (double.Parse(dt.Rows[i]["TotalPaid"].ToString().Trim()) > 0)
                                    {
                                        EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[i]["TotalPaid"]).ToString("n0").Trim() + "</td>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }
                                    EmailMsg += "</tr>";
                                    SNo++;
                                }
                                else
                                {
                                    SNo = 1;
                                    EmailMsg += "<tr style='background-color:purple; color:white;'>";
                                    EmailMsg += "<th colspan='2'>Total of [ " + SiteID + " ]</th>";
                                    if (TotalPatient > 0)
                                    {
                                        EmailMsg += "<th style='text-align:right;'>" + TotalPatient.ToString("N0") + "</th>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (TotalAmount > 0)
                                    {
                                        EmailMsg += "<th style='text-align:right;'>" + TotalAmount.ToString("N0") + "</th>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (TotalDiscount > 0)
                                    {
                                        EmailMsg += "<th style='text-align:right;'>" + TotalDiscount.ToString("N0") + "</th>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }

                                    if (TotalRefund > 0)
                                    {
                                        EmailMsg += "<th style='text-align:right;'>" + TotalRefund.ToString("N0") + "</th>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }


                                    if (TotalPayable > 0)
                                    {
                                        EmailMsg += "<th style='text-align:right;'>" + TotalPayable.ToString("N0") + "</th>";
                                    }
                                    else
                                    {
                                        EmailMsg += "<td style='text-align:right;'> - </td>";
                                    }
                                    EmailMsg += "</tr>";
                                    SiteID = dt.Rows[i]["SiteID"].ToString().Trim();
                                    
                                    EmailMsg += "<tr style='background-color:purple; color:white;'>";
                                    EmailMsg += "<th colspan='7' style='text-align:left;'>" + dt.Rows[i]["SiteName"].ToString().Trim() + " [ " + dt.Rows[i]["SiteID"].ToString().Trim() + " ] </th>";
                                    EmailMsg += "</tr>";

                                    i--;

                                    //===TOtal====
                                    TotalPatient = 0;
                                    TotalAmount = 0;
                                    TotalDiscount = 0;
                                    TotalRefund = 0;
                                    TotalPayable = 0;
                                }
                            }
                            else
                            {
                                SNo = 1;
                                SiteID = dt.Rows[i]["SiteID"].ToString();
                                i--;
                            }
                        }

                        EmailMsg += "<tr style='background-color:purple; color:white;'>";
                        EmailMsg += "<th colspan='2'>Total of [ " + SiteID + " ]</th>";
                        if (TotalPatient > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + TotalPatient.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }

                        if (TotalAmount > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + TotalAmount.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }

                        if (TotalDiscount > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + TotalDiscount.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }

                        if (TotalRefund > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + TotalRefund.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }


                        if (TotalPayable > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + TotalPayable.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        EmailMsg += "</tr>";

                        //=========Grand Total=============
                        EmailMsg += "<tr style='background-color:purple; color:gold;'>";
                        EmailMsg += "<th colspan='2'>Grand Total [ Consultancy ] </th>";
                        if (GrandPatient > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + GrandPatient.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        if (GrandTotal > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + GrandTotal.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        if (GrandDiscount > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + GrandDiscount.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        if (GrandRefund > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + GrandRefund.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        if (GrandPayable > 0)
                        {
                            EmailMsg += "<th style='text-align:right;'>" + GrandPayable.ToString("N0") + "</th>";
                        }
                        else
                        {
                            EmailMsg += "<td style='text-align:right;'> - </td>";
                        }
                        EmailMsg += "</tr>";
                        EmailMsg += "</table>";

                        #endregion

                        //=============First Aid=============
                        EmailMsg += "<br/>";
                        Query = "SELECT * FROM VW_Sale_Summary_FirstAid_ForMail WHERE TokenDate = '" + Today.ToString("MM/dd/yyyy") + "' ORDER BY TokenDate, SiteID";
                        dt = new DataTable();
                        dt = help.Return_DataTable_Query(Query);
                        if (dt.Rows.Count>0)
                        {
                            EmailMsg += "<table style='width:" + TabWidth + "%;' cellspacing='1' cellpadding='1' border='1'>";
                            EmailMsg += "<tr><td colspan='6' style='background-color:purple; color:gold;text-align:center;'><b> First Aid </td></tr>";

                            EmailMsg += "<tr style='background-color:purple; color:white;'>";
                            EmailMsg += "<th>#</th>";
                            EmailMsg += "<th>Site ID</th>";
                            EmailMsg += "<th>Total Patient</th>";
                            EmailMsg += "<th>Total</th>";
                            EmailMsg += "<th>Discount</th>";
                            EmailMsg += "<th>Net Amount</th>";
                            EmailMsg += "</tr>";

                            TotalPatient = 0;
                            TotalAmount = 0;
                            TotalDiscount = 0;
                            TotalPayable = 0;

                            GrandTotal = 0;
                            GrandPatient = 0;
                            GrandDiscount = 0;
                            GrandPayable = 0;
                            double TotalFirstAid = 0;
                            for (int a = 0; a <= dt.Rows.Count - 1; a++)
                            {
                                //===========Grand Total===========
                                GrandTotal += double.Parse(dt.Rows[a]["TotalAmount"].ToString().Trim());
                                GrandPatient += double.Parse(dt.Rows[a]["TotalPatient"].ToString().Trim());
                                GrandDiscount += double.Parse(dt.Rows[a]["Discount"].ToString().Trim());
                                GrandPayable += double.Parse(dt.Rows[a]["PaidAmount"].ToString().Trim());

                                TotalFirstAid += double.Parse(dt.Rows[a]["PaidAmount"].ToString().Trim());
                                EmailMsg += "<tr>";
                                EmailMsg += "<td>" + (a + 1) + "</td>";
                                EmailMsg += "<td style='text-align:left;'>" + dt.Rows[a]["SiteName"].ToString().Trim() + " [ " + dt.Rows[a]["SiteID"].ToString().Trim() + " ] </td>";
                                if (double.Parse(dt.Rows[a]["TotalPatient"].ToString().Trim()) > 0)
                                {
                                    EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[a]["TotalPatient"]).ToString("n0").Trim() + "</td>";
                                }
                                else
                                {
                                    EmailMsg += "<td style='text-align:right;'> - </td>";
                                }

                                if (double.Parse(dt.Rows[a]["TotalAmount"].ToString().Trim()) > 0)
                                {
                                    EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[a]["TotalAmount"]).ToString("n0").Trim() + "</td>";
                                }
                                else
                                {
                                    EmailMsg += "<td style='text-align:right;'> - </td>";
                                }

                                if (double.Parse(dt.Rows[a]["Discount"].ToString().Trim()) > 0)
                                {
                                    EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[a]["Discount"]).ToString("n0").Trim() + "</td>";
                                }
                                else
                                {
                                    EmailMsg += "<td style='text-align:right;'> - </td>";
                                }
                                if (double.Parse(dt.Rows[a]["PaidAmount"].ToString().Trim()) > 0)
                                {
                                    EmailMsg += "<td style='text-align:right;'>" + Convert.ToDecimal(dt.Rows[a]["PaidAmount"]).ToString("n0").Trim() + "</td>";
                                }
                                else
                                {
                                    EmailMsg += "<td style='text-align:right;'> - </td>";
                                }
                                EmailMsg += "</tr>";
                            }


                            EmailMsg += "<tr style='background-color:purple; color:gold;'>";
                            EmailMsg += "<th colspan='2'> Grand Total [ First Aid ] </th>";
                            if (GrandPatient > 0)
                            {
                                EmailMsg += "<th style='text-align:right;'>" + GrandPatient.ToString("N0") + "</th>";
                            }
                            else
                            {
                                EmailMsg += "<td style='text-align:right;'> - </td>";
                            }

                            if (GrandTotal > 0)
                            {
                                EmailMsg += "<th style='text-align:right;'>" + GrandTotal.ToString("N0") + "</th>";
                            }
                            else
                            {
                                EmailMsg += "<td style='text-align:right;'> - </td>";
                            }

                            if (GrandDiscount > 0)
                            {
                                EmailMsg += "<th style='text-align:right;'>" + GrandDiscount.ToString("N0") + "</th>";
                            }
                            else
                            {
                                EmailMsg += "<td style='text-align:right;'> - </td>";
                            }

                            if (GrandPayable > 0)
                            {
                                EmailMsg += "<th style='text-align:right;'>" + GrandPayable.ToString("N0") + "</th>";
                            }
                            else
                            {
                                EmailMsg += "<td style='text-align:right;'> - </td>";
                            }
                            EmailMsg += "</tr>";

                            EmailMsg += "</table>";
                        }

                        Generate_EmailFooter(TabWidth);
                        EmailMsg += EmailFooter;
                        //SendEmail(EmailMsg);
                        Send_Schedule_Email(EmailModule, EmailSiteID);

                    }
                    else
                    {
                        //Do Nothing..........
                    }
                }

            }
            catch (Exception ex)
            {
                App_Log("CCPL Pharmacy Mailer Service (Sale Summary Night) Error :" + EmailModule + ex.Message);
                //help.ExecuteParameterizedProcedure("UD_Add_Email_ErrorLog", "@DocNo,@Module,@SiteID,@ErrorMsg", "", EmailModule, EmailSiteID, ex.Message);
            }
        }

    }
}
