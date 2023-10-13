using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using log4net;
using System.Data.SqlClient;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Quartz;
using Quartz.Impl;
using Quartz.Impl.Triggers;
using System.Configuration;
using System.Net.Http;
using Newtonsoft.Json;

namespace NotifivationCheckOutDate
{
    public partial class Service1 : ServiceBase
    {
        private const string Group1 = "BusinessTasks";
        private const string Job = "Job";
        //cai dat log4net.config is alway copy, AsemblyInfo: 
        //[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config")]
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private static IScheduler _scheduler;
        public Service1()
        {
            InitializeComponent();
        }
        
        protected override void OnStart(string[] args)
        {
            ISchedulerFactory schedulerFactory = new StdSchedulerFactory();
            _scheduler = schedulerFactory.GetScheduler();
            _scheduler.Start();

            Log.Info("Starting Windows Service...");

            AddJobs();
        }

        private void AddJobs()
        {
            SendNotesJob();
            //Execute1();

        }

        public static void SendNotesJob()
        {
            try
            {
                //讀取CronExpresstion
                var CronEx = ConfigurationManager.AppSettings["CronEx"]; //"0 30 7 * * ?";
                const string trigger1 = "DailySendNoteTasksTrigger";
                const string jobName = trigger1 + Job;
                IDoJob myJob = new DailySendNotesJob();
                var jobDetail = new JobDetailImpl(jobName, Group1, myJob.GetType());
                var trigger = new CronTriggerImpl(
                    trigger1,
                    Group1,
                    CronEx)
                { TimeZone = TimeZoneInfo.Local };
                _scheduler.ScheduleJob(jobDetail, trigger);
                var nextFireTime = trigger.GetNextFireTimeUtc();
                if (nextFireTime != null)
                    Log.Info(Group1 + "+" + trigger1 + ":" + nextFireTime.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));



            }
            catch (Exception ex)
            {
                Log.Info("err0: " + ex.Message);
            }
        }
       
        internal class DailySendNotesJob : IDoJob
        {
            public async void Execute(IJobExecutionContext context)
            {
                var bfCheckOutDate = DateTime.Now.AddDays(6).ToString("yyyyMMdd"); //退房日期前一周
                var htmlBody = @"<html><body><style>th {background-color :#006699} table, th, td { border-collapse: collapse; font-size: 15px; padding:7px; text-align: center; } th,td {border: 1px solid black; ; } </style>";
                var htmlBody2 = @"<html><body><style>th {background-color :#006699} table, th, td { border-collapse: collapse; font-size: 15px; padding:7px; text-align: center; } th,td {border: 1px solid black; ; } </style>";
                //var content = @"";

                string connectionString = ConfigurationManager.ConnectionStrings["FPGFlow"].ConnectionString;
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    //生活區門禁卡申請單
                    string query1 = @"select distinct D.DocId,id5,id3 ,id4 ,id8 ,VNW_Fuli from F88PK00116_F88F000184_DOCS D,F88PK00116_F88F000184_MD M, [ro_documentdata] S 
                                     where D.DocId = M.DocId and D.DocId = S.DocId  and REPLACE(M.COLCTRL11,'/','') = @checkOutDate AND isSendNotes = 'N' AND  S.ProcStatusId='S0104' ORDER BY D.DocId ";// sau nhớ bổ sung thêm S.ProcStatusId='S0104'

                    string query2 = @"select COLCTRL1,COLCTRL2 ,COLCTRL3 ,COLCTRL4,COLCTRL5,COLCTRL6 ,
                                COLCTRL7 ,COLCTRL8 ,COLCTRL9 ,COLCTRL10 ,COLCTRL11 ,COLCTRL12 ,COLCTRL13 
                                from F88PK00116_F88F000184_DOCS D,F88PK00116_F88F000184_MD M, [ro_documentdata] S
                                where D.DocId= M.DocId and D.DocId= S.DocId  AND D.DocId =@DocId ";
                    string query3 = @"UPDATE F88PK00116_F88F000184_DOCS set isSendNotes='Y' WHERE DocId =@DocId";

                    //陸籍(廠商)宿舍進住申請單
                    string query4 = @"select distinct D.DocId ,id4,id3 ,VNW_Fuli from F88PK00117_F88F000186_DOCS D,F88PK00117_F88F000186_MD M, [ro_documentdata] S 
                                     where D.DocId = M.DocId and D.DocId = S.DocId  and REPLACE(M.COLCTRL12,'/','') = @checkOutDate AND isSendNotes = 'N' AND S.ProcStatusId='S0104' ORDER BY D.DocId ";// sau nhớ bổ sung thêm S.ProcStatusId='S0104'

                    string query5 = @"select COLCTRL1,COLCTRL2 ,COLCTRL3 ,COLCTRL15,COLCTRL5,COLCTRL6 ,
                                COLCTRL7 ,COLCTRL8 ,COLCTRL9 ,COLCTRL11 ,COLCTRL12 ,COLCTRL14 
                                from F88PK00117_F88F000186_DOCS D,F88PK00117_F88F000186_MD M, [ro_documentdata] S
                                where D.DocId= M.DocId and D.DocId= S.DocId  AND D.DocId =@DocId ";
                    string query6 = @"UPDATE F88PK00117_F88F000186_DOCS set isSendNotes='Y' WHERE DocId =@DocId";


                    try
                    {
                        //生活區門禁卡申請單
                        using (var command = new SqlCommand(query1, connection))
                        {
                            //command.Parameters.AddWithValue("@checkOutDate", bfCheckOutDate);
                            command.Parameters.AddWithValue("@checkOutDate", bfCheckOutDate);// sau nho doi
                            using (var reader = await command.ExecuteReaderAsync())
                            {
                                if (reader.HasRows)
                                {
                                    Log.Info("生活區門禁卡申請單");
                                    while (await reader.ReadAsync())
                                    {
                                        
                                        var docId = reader[0].ToString();
                                        htmlBody += $@"<h4>經盤點眷屬卡資料發現貴廠{reader[2].ToString()}(人員代號:{reader[1].ToString()})申請辦理眷屬門禁卡於{bfCheckOutDate}到期。請速歸還或辦理展延相關手續。</h4><br><label>眷屬人員資料如下：</label><br>";
                                        htmlBody += $@"<table width='100%' border='1'><tr><th>眷屬姓名</th><th>國籍</th><th>性别</th><th>出生日期</th><th>照號碼</th><th>護照有效期</th><th>簽證或暫住證號碼</th><th>簽證效期</th><th>關係</th><th>進住日期</th><th>退宿日期</th><th>房號</th><th>卡號</th></tr>";
                                        var command2 = new SqlCommand(query2, connection);
                                        command2.Parameters.AddWithValue("@DocId", docId);
                                        var reader2 = await command2.ExecuteReaderAsync();
                                        try
                                        {
                                            while (await reader2.ReadAsync())
                                            {
                                                htmlBody += $@"<tr><td>{reader2[0].ToString()}</td><td>{reader2[1].ToString()}</td><td>{reader2[2].ToString()}</td><td>{reader2[3].ToString()}</td><td>{reader2[4].ToString()}</td><td>{reader2[5].ToString()}</td><td>{reader2[6].ToString()}</td><td>{reader2[7].ToString()}</td><td>{reader2[8].ToString()}</td><td>{reader2[9].ToString()}</td><td>{reader2[10].ToString()}</td><td>{reader2[11].ToString()}</td><td>{reader2[12].ToString()}</td></tr>";
                                            }
                                            htmlBody += @"</table></body></html>";
                                            Mail m = new Mail()
                                            {
                                                To = reader[1].ToString() + "@VNFPG",
                                                Cc=reader[5].ToString()+ "@VNFPG", //福利人
                                                Subject = "申請辦理眷屬門禁卡於" + bfCheckOutDate + "到退宿日期",
                                                Content = htmlBody,
                                                CustomFormat = true

                                            };
                                            Log.Info("Note:"+ reader[1].ToString()+", doc ID:"+ docId);
                                            await SendAsync(m);

                                            var command3 = new SqlCommand(query3, connection);
                                            command3.Parameters.AddWithValue("@DocId", docId);
                                            command3.ExecuteScalar();


                                        }
                                        finally
                                        {
                                            reader2.Close();
                                        }
                                    }
                                    reader.Close();
                                }
                            }
                        }

                        //陸籍(廠商)宿舍進住申請單
                        using (var command4 = new SqlCommand(query4, connection))
                        {
                            //command.Parameters.AddWithValue("@checkOutDate", bfCheckOutDate);
                            command4.Parameters.AddWithValue("@checkOutDate", bfCheckOutDate);//
                            using (var reader4 = await command4.ExecuteReaderAsync())
                            {
                                if (reader4.HasRows)
                                {
                                    Log.Info("陸籍(廠商)宿舍進住申請單");
                                    while (await reader4.ReadAsync())
                                    {

                                        var docId2 = reader4[0].ToString();
                                        htmlBody2 += $@"<h4>您註冊的廠商員工即將到退宿日期（{bfCheckOutDate}）。 請快速歸還或辦理展期相關手續。</h4><br><label>廠商人員資料如下：</label><br>";
                                        htmlBody2 += $@"<table width='100%' border='1'><tr><th>中文姓名</th><th>英文姓名</th><th>性别</th><th>國籍</th><th>出入廠卡號碼</th><th>護照號碼</th><th>護照有效期</th><th>簽證或暫住證號碼</th><th>簽證效期</th><th>進住日期</th><th>退住日期</th><th>房號</th>";
                                        var command5 = new SqlCommand(query5, connection);
                                        command5.Parameters.AddWithValue("@DocId", docId2);
                                        var reader5 = await command5.ExecuteReaderAsync();
                                        try
                                        {
                                            while (await reader5.ReadAsync())
                                            {
                                                htmlBody2 += $@"<tr><td>{reader5[0].ToString()}</td><td>{reader5[1].ToString()}</td><td>{reader5[2].ToString()}</td><td>{reader5[3].ToString()}</td><td>{reader5[4].ToString()}</td><td>{reader5[5].ToString()}</td><td>{reader5[6].ToString()}</td><td>{reader5[7].ToString()}</td><td>{reader5[8].ToString()}</td><td>{reader5[9].ToString()}</td><td>{reader5[10].ToString()}</td><td>{reader5[11].ToString()}</td></tr>";
                                            }
                                            htmlBody2 += @"</table></body></html>";
                                            Mail m = new Mail()
                                            {
                                                To = reader4[1].ToString() + "@VNFPG",
                                                Cc = reader4[3].ToString() + "@VNFPG", //福利人
                                               
                                                Subject = "廠商人員" + bfCheckOutDate + "到退宿日期",
                                                Content = htmlBody2,
                                                CustomFormat = true

                                            };
                                            Log.Info("Note:" + reader4[1].ToString() + ", doc ID:" + docId2);
                                            await SendAsync(m);

                                            var command6 = new SqlCommand(query6, connection);
                                            command6.Parameters.AddWithValue("@DocId", docId2);
                                            command6.ExecuteScalar();


                                        }
                                        finally
                                        {
                                            reader5.Close();
                                        }
                                    }
                                    reader4.Close();
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Log.Info("err1:" + ex.Message);
                    }

                }
            }

            
        }
        internal interface IDoJob : IJob
        {

        }
        protected override void OnStop()
        {
        }
        public static async Task SendAsync(Mail mail)
        {
            using (var client = new HttpClient())
            {
                try
                {
                    client.BaseAddress = new Uri("http://10.199.1.32:1234");
                    if (string.IsNullOrEmpty(mail.To) && string.IsNullOrEmpty(mail.Cc)) return;
                    if (string.IsNullOrEmpty(mail.Subject)) return;
                    if (string.IsNullOrEmpty(mail.Content)) return;
                    var json_string = JsonConvert.SerializeObject(mail);
                    var requestContent = new StringContent(json_string, Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("/api/Mail", requestContent);
                    var responseContent = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(responseContent);
                    Log.Info("Note 發送成功！");
                }
                catch (Exception ex) { Log.Info(ex.Message); }
            }
        }
        public class Mail
        {
            public string To { get; set; }
            public string Cc { get; set; }
            public string Subject { get; set; }
            public string From { get; set; }
            public string Content { get; set; }
            public bool CustomFormat { get; set; }
        }
    }
}
