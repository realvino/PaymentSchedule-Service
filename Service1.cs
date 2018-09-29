using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Text;

namespace WindowsServiceCS
{
    public partial class Service1 : ServiceBase
    {
        public string EmailId, password, smtpServer;
        public bool useSSL;
        public int Port, i = 0;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }
              

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("Simple Service scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("WatchNet Payment Schedule Test Service"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {

            try
            {
                string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                //SqlConnection con=new SqlConnection(constr);
                DataTable dt = new DataTable();
                string query = "SELECT * FROM ProjectSchedule";

                using (SqlConnection con = new SqlConnection(constr))
                {

                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        //cmd.Parameters.AddWithValue("@Day", DateTime.Today.Day);
                        //cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month);
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
                foreach (DataRow row in dt.Rows)
                {

                    DateTime ST = (DateTime)row["ScheduleDate"];
                     DateTime DT = ST.AddDays(-1);
                     DateTime d = DateTime.Today.Date;
                     this.WriteToFile(" Windows Service" + DT.Date+" "+d);
                    if (DT.Date == d)
                    {
                        string Title = "WatchNet Notification Alert";
                        string SubTitle = string.Empty;
                        string TaskName = string.Empty;
                        string RequestFrom = string.Empty;
                        string RequestTo = string.Empty;
                        string RequestFromRoleName = string.Empty;
                        string ProjectName = string.Empty;
                        string RefNo = string.Empty;                     
                        string ReceiverMailId = string.Empty;
                        string SalesmanId = string.Empty;
                        string CustomerPo = string.Empty;
                        string CustomerCompany = string.Empty;

                        //Sql Start Query Execution
                        SqlConnection con1 = new SqlConnection(constr);
                        con1.Open();
                        string qur = "select * from dbo.Quotations where Id='" + row["QuotationId"] + "'";
                        SqlCommand cmd = new SqlCommand(qur, con1);
                        SqlDataReader dr1 = cmd.ExecuteReader();
                        if (dr1.Read())
                        {
                            SalesmanId = dr1["SalesmanId"].ToString();
                            CustomerPo = dr1["CustomerPONumber"].ToString();
                        }
                        con1.Close();
                        //Sql End Query Execution

                        SqlConnection con10 = new SqlConnection(constr);
                        con10.Open();
                        string qurcom = "select dbo.Companies.CompanyName as CompanyName from dbo.Quotations inner join dbo.Companies on (dbo.Quotations.CompanyId=dbo.Companies.Id) where dbo.Quotations.Id='" + row["QuotationId"] + "'";
                        SqlCommand cmds = new SqlCommand(qurcom, con10);
                        SqlDataReader dr10 = cmds.ExecuteReader();
                        if (dr10.Read())
                        {
                            CustomerCompany = dr10["CompanyName"].ToString();
                        }
                        con10.Close();

                        SubTitle = "Payment Due Alert : " + CustomerCompany + " ( " + CustomerPo + " )";

                        SqlConnection con2 = new SqlConnection(constr);
                        con2.Open();
                        string qur1 = "select dbo.AbpUsers.UserName,dbo.AbpRoles.DisplayName as RoleName from dbo.AbpUserRoles inner join dbo.AbpUsers on (dbo.AbpUsers.Id=dbo.AbpUserRoles.UserId) inner join dbo.AbpRoles on (dbo.AbpRoles.Id=dbo.AbpUserRoles.RoleId) where UserId='" + SalesmanId + "'";
                        SqlCommand cmd1 = new SqlCommand(qur1, con2);
                        SqlDataReader dr2 = cmd1.ExecuteReader();
                        if (dr2.Read())
                        {
                            RequestFrom = dr2["UserName"].ToString();
                            RequestFromRoleName = dr2["RoleName"].ToString();
                            this.WriteToFile("Mail From" + dr2["UserName"] + "");
                            this.WriteToFile("Mail Role" + dr2["RoleName"] + "");

                        }
                        con2.Close();

                        SqlConnection con3 = new SqlConnection(constr);
                        con3.Open();
                        string qur2 = "select UserName,EmailAddress from dbo.AbpUsers where Id='" + SalesmanId + "'";
                        SqlCommand cmd2 = new SqlCommand(qur2, con3);
                        SqlDataReader dr3 = cmd2.ExecuteReader();
                        if (dr3.Read())
                        {
                            RequestTo = dr3["UserName"].ToString();
                            ReceiverMailId = dr3["EmailAddress"].ToString();
                        }
                        con3.Close();

                        SqlConnection con4 = new SqlConnection(constr);
                        con4.Open();
                        string qur3 = "select * from dbo.AbpSettings";
                        SqlCommand cmd3 = new SqlCommand(qur3, con4);
                        DataTable dt1 = new DataTable();
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd3))
                        {
                            sda.Fill(dt1);
                        }
                        foreach (DataRow row1 in dt1.Rows)
                        {
                            if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Port")
                            {
                                Port = Convert.ToInt32(row1["Value"]);
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.UserName")
                            {
                                EmailId = row1["Value"].ToString();
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Password")
                            {
                                password = row1["Value"].ToString();
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.EnableSsl")
                            {
                                useSSL = Convert.ToBoolean(row1["Value"]);
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Host")
                            {
                                smtpServer = row1["Value"].ToString();
                            }
                        }

                        if(Port==0)
                        {
                            Port = 25;
                        }

                        con4.Close();

                        WriteToFile("Trying to send email to: " + ReceiverMailId);

                        using (MailMessage mm = new MailMessage(EmailId, ReceiverMailId))
                        {
                            mm.Subject = SubTitle;

                            //var mailMessage = new StringBuilder();
                            //mailMessage.AppendLine("<b>" + "User Name" + "</b>: " + RequestTo + "<br />");
                            //mailMessage.AppendLine("<b>" + "Notification From" + "</b>: " + RequestFrom + "<br />");
                            //mailMessage.AppendLine("<b>" + "Notification From Role" + "</b>: " + RequestFromRoleName + "<br />");
                            //mailMessage.AppendLine("<b>" + "Project Name" + "</b>: " + ProjectName + "<br />");
                            //mailMessage.AppendLine("<b>" + "Reference Number" + "</b>: " + RefNo + "<br />");
                            //mailMessage.AppendLine("<br />");
                            //mailMessage.AppendLine(NotificationMessage + "<br /><br />");
                            //mailMessage.AppendLine("<a href=\"" + Url + "\">" + Url + "</a>");

                            string body = string.Empty;
                            string path = Directory.GetCurrentDirectory();
                            using (StreamReader reader = new StreamReader("C:\\hostingemail\\PaymentDue.html"))
                            {
                                body = reader.ReadToEnd();
                            }
                            string mailbodycontent = "<b>Dear " + RequestTo + "</b><br /><br /><span>You are receiving this alert for the nearing payment due</span><br /><br /><span>Company Name : <b style='color:blue;font-weight: bolder;'>"+ CustomerCompany +"</b></span><br /><br /><span>Customer PO Number : <b style='color:blue;font-weight: bolder;'>"+ CustomerPo +"</b></span><br /><br />";                            
                            body = body.Replace("{EMAIL_TITLE}", Title);
                            body = body.Replace("{Mail_Content}", mailbodycontent);
                            //body = body.Replace("{EMAIL_SUB_TITLE}", SubTitle);
                            //body = body.Replace("{EMAIL_BODY}", mailMessage.ToString());
                            //body = body.Replace("{Date}", DateTime.Now.Year.ToString());

                            var im = "http://localhost:6240/";
                            var im1 = "http://localhost:6240/";
                            var ipath = "Common/Images/logo.jpg";
                            var ipath1 = "Common/Images/Teamworks%20logo1.png";

                            try
                            {
                                SqlConnection con11 = new SqlConnection(constr);
                                con11.Open();
                                string qur11 = "select Value from dbo.AbpSettings where Name='App.General.WebSiteRootAddress'";
                                SqlCommand cmd11 = new SqlCommand(qur11, con11);
                                SqlDataReader dr11 = cmd11.ExecuteReader();
                                if (dr11.Read())
                                {
                                    im = dr11["Value"].ToString();
                                    im1 = dr11["Value"].ToString();
                                }
                                con11.Close();
                                ////var image = context1.Database.SqlQuery<ret>("select Value from dbo.AbpSettings where Name='App.General.WebSiteRootAddress'").FirstOrDefaultAsync();
                                //if (qur11.Result.Value != null)
                                //{
                                //    im = image.Result.Value;
                                //    im1 = image.Result.Value;
                                //}
                            }
                            catch (Exception ex)
                            {

                            }

                            im = im + ipath;
                            body = body.Replace("{WatchNet_Logo}", im);
                            body = body.Replace("{TeamWorks_logo}", im1 + ipath1);

                            mm.Body = body;
                            mm.IsBodyHtml = true;
                            SmtpClient smtp = new SmtpClient();
                            smtp.Host = smtpServer;
                            smtp.EnableSsl = useSSL;
                            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                            credentials.UserName = EmailId;
                            credentials.Password = password;
                            smtp.UseDefaultCredentials = true;
                            smtp.Credentials = credentials;
                            smtp.Port = Port;
                            smtp.Send(mm);
                            WriteToFile("Email sent successfully to: " + RequestTo + " " + ReceiverMailId);

                            SqlConnection conFinal = new SqlConnection(constr);
                            conFinal.Open();
                            string qurfinal = "update dbo.Notification set IsSent='" + true + "',Status='" + "Send" + "',ActualNotification='" + DateTime.Now + "' where Id='" + Convert.ToInt32(row["Id"]) + "' ";
                            SqlCommand cmdfinal = new SqlCommand(qurfinal, conFinal);
                            cmdfinal.ExecuteNonQuery();
                            conFinal.Close();
                        }
                    }

                }
                this.ScheduleService();
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("WatchNet Payment Schedule Test Service"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void WriteToFile(string text)
        {
            string path = "C:\\hostinglog\\PaymentScheduleTestServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }    
    }
}
