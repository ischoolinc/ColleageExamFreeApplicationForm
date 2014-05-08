using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using JHSchool.Data;
using K12.BusinessLogic;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ColleageExamFreeApplicationForm
{
    public partial class Report : BaseForm
    {
        AccessHelper _A = new AccessHelper();
        QueryHelper _Q = new QueryHelper();
        Dictionary<String, String> _Column2Items;
        Dictionary<String, List<string>> _MappingData;
        Dictionary<String, int> 報名資格;
        Dictionary<String, int> 報名費減免身分;
        Dictionary<String, int> 弱勢身分;
        Dictionary<String, int> 特種生加分類別;
        Dictionary<String, int> 其他比序項目_全民英檢;
        Dictionary<String, int> 其他比序項目_多益測驗;
        BackgroundWorker _BW;
        string _SchoolName;

        public Report()
        {
            InitializeComponent();

            Column2Prepare();

            報名資格 = new Dictionary<string, int>();
            報名費減免身分 = new Dictionary<string, int>();
            弱勢身分 = new Dictionary<string, int>();
            特種生加分類別 = new Dictionary<string, int>();
            其他比序項目_全民英檢 = new Dictionary<string, int>();
            其他比序項目_多益測驗 = new Dictionary<string, int>();

            報名資格.Add("國民中學非應屆畢業生", 0);
            報名資格.Add("同等學歷", 2);

            報名費減免身分.Add("低收入戶", 1);
            報名費減免身分.Add("失業子女", 2);
            報名費減免身分.Add("中低收入戶", 3);

            弱勢身分.Add("低收入戶", 1);
            弱勢身分.Add("失業子女", 2);
            弱勢身分.Add("中低收入戶", 3);
            弱勢身分.Add("特殊境遇家庭", 4);

            特種生加分類別.Add("原住民-未取得原住民文化及語言能力證明", 1);
            特種生加分類別.Add("原住民-取得原住民文化及語言能力證明", 2);
            特種生加分類別.Add("境外-返國(來臺)就讀未滿一學年", 3);
            特種生加分類別.Add("境外-返國(來臺)就讀一學年以上未滿二學年", 4);
            特種生加分類別.Add("境外-返國(來臺)就讀二學年以上未滿三學年", 5);
            特種生加分類別.Add("派外-返國(來臺)就讀未滿一學年", 6);
            特種生加分類別.Add("派外-返國(來臺)就讀一學年以上未滿二學年", 7);
            特種生加分類別.Add("派外-返國(來臺)就讀二學年以上未滿三學年", 8);
            特種生加分類別.Add("蒙藏生", 9);
            特種生加分類別.Add("身障生", 10);
            特種生加分類別.Add("僑生", 11);
            特種生加分類別.Add("退伍軍人-在營服役期間五年以上，退伍後未滿一年", 12);
            特種生加分類別.Add("退伍軍人-在營服役期間五年以上，退伍後一年以上未滿二年", 13);
            特種生加分類別.Add("退伍軍人-在營服役期間五年以上，退伍後二年以上未滿三年", 14);
            特種生加分類別.Add("退伍軍人-在營服役期間五年以上，退伍後三年以上未滿五年", 15);
            特種生加分類別.Add("退伍軍人-在營服役期間四年以上未滿五年，退伍後未滿一年", 16);
            特種生加分類別.Add("退伍軍人-在營服役期間四年以上未滿五年，退伍後一年以上未滿二年", 17);
            特種生加分類別.Add("退伍軍人-在營服役期間四年以上未滿五年，退伍後二年以上未滿三年", 18);
            特種生加分類別.Add("退伍軍人-在營服役期間四年以上未滿五年，退伍後三年以上未滿五年", 19);
            特種生加分類別.Add("退伍軍人-在營服役期間三年以上未滿四年，退伍後未滿一年", 20);
            特種生加分類別.Add("退伍軍人-在營服役期間三年以上未滿四年，退伍後一年以上未滿二年", 21);
            特種生加分類別.Add("退伍軍人-在營服役期間三年以上未滿四年，退伍後二年以上未滿三年", 22);
            特種生加分類別.Add("退伍軍人-在營服役期間三年以上未滿四年，退伍後三年以上未滿五年", 23);
            特種生加分類別.Add("退伍軍人-在營服役期間未滿三年，已達義務役法定役期，且退伍後未滿三年", 24);
            特種生加分類別.Add("退伍軍人-因作戰或因公成殘領有撫卹證明，於免役、除役後未滿五年", 25);
            特種生加分類別.Add("退伍軍人-因病成殘領有撫卹證明，於免役、除役後未滿五年", 26);

            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 初級 初試及格", 1);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 初級 複試及格", 2);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 中級 初試及格", 3);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 中級 複試及格", 4);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 中高級 初試及格", 5);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 中高級 複試及格", 6);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 高級 初試及格", 7);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 高級 複試及格", 8);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 優級 初試及格", 9);
            其他比序項目_全民英檢.Add("全民英語能力分級檢定測驗 GEPT 優級 複試及格", 10);

            其他比序項目_多益測驗.Add("多益測驗 (TOEIC) 聽力 110 以上 閱讀 115 以上", 1);
            其他比序項目_多益測驗.Add("多益測驗 (TOEIC) 聽力 275 以上 閱讀 275 以上", 2);
            其他比序項目_多益測驗.Add("多益測驗 (TOEIC) 聽力 400 以上 閱讀 385 以上", 3);

            _SchoolName = K12.Data.School.ChineseName;

            _BW = new BackgroundWorker();
            _BW.WorkerReportsProgress = true;
            _BW.DoWork += new DoWorkEventHandler(DataBuilding);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ReportBuilding);
            _BW.ProgressChanged += new ProgressChangedEventHandler(BW_Progress);
        }

        private void BW_Progress(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(Global.ReportName + "產生中", e.ProgressPercentage);
        }

        private void ReportBuilding(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(Global.ReportName + " 產生完成");

            EnableForm(true);
            Workbook wb = (Workbook)e.Result;
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = Global.ReportName + ".xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void DataBuilding(object sender, DoWorkEventArgs e)
        {
            _BW.ReportProgress(0);

            SaveSetting();
            //MappingData
            _MappingData = new Dictionary<string, List<string>>();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string tagName = row.Cells[0].Value.ToString();

                    if (!_MappingData.ContainsKey(tagName))
                    {
                        _MappingData.Add(tagName, new List<string>());
                    }

                    if (row.Cells[1].Value != null)
                    {
                        string tagText = row.Cells[1].Value.ToString();
                        if (_Column2Items.ContainsKey(tagText))
                        {
                            string tagId = _Column2Items[tagText];
                            if (!_MappingData[tagName].Contains(tagId))
                            {
                                _MappingData[tagName].Add(tagId);
                            }
                        }
                    }
                }
            }

            Dictionary<string, StudentObj> studentDic = new Dictionary<string, StudentObj>();
            List<string> students = K12.Presentation.NLDPanels.Student.SelectedSource;
            string ids = string.Join("','", students);

            _BW.ReportProgress(10);
            //基本資料
            DataTable dt = _Q.Select("SELECT student.id,student.name,student.id_number,class.class_name,student.seat_no,student.student_number,student.birthdate,student.contact_phone,student.sms_phone,student.mailing_address,student.permanent_address,class.grade_year FROM student LEFT JOIN class ON ref_class_id = class.id WHERE student.id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                StudentObj obj = new StudentObj(row);
                if (!studentDic.ContainsKey(obj.Id))
                {
                    studentDic.Add(obj.Id, obj);
                }
            }

            _BW.ReportProgress(15);
            //基本資料-TagId
            dt = _Q.Select("SELECT ref_student_id,ref_tag_id FROM tag_student WHERE ref_student_id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                string id = row["ref_student_id"].ToString();
                string tagid = row["ref_tag_id"].ToString();
                if (studentDic.ContainsKey(id))
                {
                    if (!studentDic[id].TagIds.Contains(tagid))
                    {
                        studentDic[id].TagIds.Add(tagid);
                    }
                }
            }

            _BW.ReportProgress(20);
            //服務學習紀錄
            dt = _Q.Select("SELECT ref_student_id,hours FROM $k12.service.learning.record WHERE ref_student_id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                string id = row["ref_student_id"].ToString();
                if (studentDic.ContainsKey(id))
                {
                    studentDic[id].ServiceHours += decimal.Parse(row["hours"].ToString());
                }
            }

            _BW.ReportProgress(30);
            //幹部紀錄
            dt = _Q.Select("SELECT studentid,schoolyear,semester FROM $behavior.thecadre WHERE studentid IN ('" + ids + "')");
            List<string> checkList = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                string id = row["studentid"].ToString();
                string schoolyear = row["schoolyear"].ToString();
                string semester = row["semester"].ToString();

                string key = id + "_" + schoolyear + "_" + semester;

                if (!checkList.Contains(key))
                {
                    if (studentDic.ContainsKey(id))
                    {
                        studentDic[id].CadreTimes++;
                        checkList.Add(key);
                    }
                }
            }

            _BW.ReportProgress(40);
            //獎懲紀錄
            List<AutoSummaryRecord> records = AutoSummary.Select(students, null);
            foreach (AutoSummaryRecord record in records)
            {
                string id = record.RefStudentID;
                if (studentDic.ContainsKey(id))
                {
                    studentDic[id].MeritA += record.MeritA;
                    studentDic[id].MeritB += record.MeritB;
                    studentDic[id].MeritC += record.MeritC;
                    studentDic[id].DemeritA += record.DemeritA;
                    studentDic[id].DemeritB += record.DemeritB;
                    studentDic[id].DemeritC += record.DemeritC;
                }
            }

            _BW.ReportProgress(45);
            //取得功過換算比例
            MeritDemeritReduceRecord mdrr = MeritDemeritReduce.Select();
            int MAB = mdrr.MeritAToMeritB.HasValue ? mdrr.MeritAToMeritB.Value : 0;
            int MBC = mdrr.MeritBToMeritC.HasValue ? mdrr.MeritBToMeritC.Value : 0;
            int DAB = mdrr.DemeritAToDemeritB.HasValue ? mdrr.DemeritAToDemeritB.Value : 0;
            int DBC = mdrr.DemeritBToDemeritC.HasValue ? mdrr.DemeritBToDemeritC.Value : 0;

            _BW.ReportProgress(50);
            //獎懲紀錄功過相抵
            foreach (StudentObj obj in studentDic.Values)
            {
                //if (!obj.HasDemeritAB)
                //{
                int merit = ((obj.MeritA * MAB) + obj.MeritB) * MBC + obj.MeritC;
                int demerit = ((obj.DemeritA * DAB) + obj.DemeritB) * DBC + obj.DemeritC;

                int total = merit - demerit;

                if (total > 0)
                {
                    obj.MC = total % MBC;
                    obj.MB = (total / MBC) % MAB;
                    obj.MA = (total / MBC) / MAB;
                }
                else if (total < 0)
                {
                    total *= -1;
                    obj.DC = total % DBC;
                    obj.DB = (total / DBC) % DAB;
                    obj.DA = (total / DBC) / DAB;
                }
                //}
            }

            _BW.ReportProgress(60);
            //體適能
            //先確認UDT存在
            dt = _Q.Select("SELECT name FROM _udt_table where name='ischool_student_fitness'");
            if (dt.Rows.Count > 0)
            {
                dt = _Q.Select("SELECT ref_student_id,sit_and_reach_degree, standing_long_jump_degree, sit_up_degree, cardiorespiratory_degree FROM $ischool_student_fitness WHERE ref_student_id IN ('" + ids + "')");
                foreach (DataRow row in dt.Rows)
                {
                    string id = row["ref_student_id"].ToString();
                    if (studentDic.ContainsKey(id))
                    {
                        //擇優判斷
                        if (GetScore(row) > studentDic[id].SportFitnessScore)
                        {
                            studentDic[id].sit_and_reach_degree = row["sit_and_reach_degree"].ToString();
                            studentDic[id].sit_up_degree = row["sit_up_degree"].ToString();
                            studentDic[id].standing_long_jump_degree = row["standing_long_jump_degree"].ToString();
                            studentDic[id].cardiorespiratory_degree = row["cardiorespiratory_degree"].ToString();
                        }
                    }
                }
            }

            _BW.ReportProgress(65);
            //均衡學習-年級對照
            Dictionary<string, Dictionary<string, string>> SchoolyearSemesteerToGrade = new Dictionary<string, Dictionary<string, string>>();
            foreach (JHSemesterHistoryRecord record in JHSemesterHistory.SelectByStudentIDs(students))
            {
                foreach (SemesterHistoryItem item in record.SemesterHistoryItems)
                {
                    if (!SchoolyearSemesteerToGrade.ContainsKey(item.RefStudentID))
                    {
                        SchoolyearSemesteerToGrade.Add(item.RefStudentID, new Dictionary<string, string>());
                    }

                    string key = item.SchoolYear + "_" + item.Semester;
                    if (!SchoolyearSemesteerToGrade[item.RefStudentID].ContainsKey(key))
                    {
                        if (item.Semester == 1)
                            SchoolyearSemesteerToGrade[item.RefStudentID].Add(key, item.GradeYear + "上");
                        else if (item.Semester == 2)
                            SchoolyearSemesteerToGrade[item.RefStudentID].Add(key, item.GradeYear + "下");
                    }

                }
            }

            _BW.ReportProgress(70);
            //均衡學習-領域分數
            List<JHSemesterScoreRecord> recs = JHSemesterScore.SelectByStudentIDs(students);
            foreach (JHSemesterScoreRecord rec in recs)
            {
                foreach (DomainScore score in rec.Domains.Values)
                {
                    string id = score.RefStudentID;
                    string key = score.SchoolYear + "_" + score.Semester;
                    string grade = "";

                    if (SchoolyearSemesteerToGrade.ContainsKey(id))
                    {
                        if (SchoolyearSemesteerToGrade[id].ContainsKey(key))
                        {
                            grade = SchoolyearSemesteerToGrade[id][key];
                        }
                    }

                    if (studentDic.ContainsKey(id))
                    {
                        string domain = score.Domain;

                        if ((domain == "健康與體育" || domain == "藝術與人文" || domain == "綜合活動") && !string.IsNullOrWhiteSpace(grade))
                        {
                            if (!studentDic[id].DomainScores.ContainsKey(domain))
                            {
                                studentDic[id].DomainScores.Add(domain, new Dictionary<string, decimal>());
                            }

                            if (!studentDic[id].DomainScores[domain].ContainsKey(grade))
                            {
                                decimal value = score.Score.HasValue ? score.Score.Value : 0;
                                studentDic[id].DomainScores[domain].Add(grade, value);
                            }
                        }
                    }
                }
            }

            _BW.ReportProgress(80);
            //排序
            List<StudentObj> list = studentDic.Values.ToList();
            list.Sort(SortStudent);

            int progress = 80;
            decimal per = (decimal)(100 - progress) / studentDic.Count;
            int count = 0;
            //Objects轉Table
            Dictionary<string, int> CloumnIndex = new Dictionary<string, int>();
            CloumnIndex.Add("身分證字統一編號", 0);
            CloumnIndex.Add("學生姓名", 1);
            CloumnIndex.Add("出生年(民國年)", 2);
            CloumnIndex.Add("出生月", 3);
            CloumnIndex.Add("出生日", 4);
            CloumnIndex.Add("年級", 5);
            CloumnIndex.Add("班級", 6);
            CloumnIndex.Add("座號", 7);
            CloumnIndex.Add("報名資格", 8);
            CloumnIndex.Add("郵遞區號", 9);
            CloumnIndex.Add("地址", 10);
            CloumnIndex.Add("市內電話", 11);
            CloumnIndex.Add("行動電話", 12);
            CloumnIndex.Add("特種生加分類別", 13);
            CloumnIndex.Add("報名費減免身分", 14);
            CloumnIndex.Add("競賽", 15);
            CloumnIndex.Add("擔任幹部", 16);
            CloumnIndex.Add("服務時數", 17);
            CloumnIndex.Add("服務學習", 18);
            CloumnIndex.Add("累計嘉獎", 19);
            CloumnIndex.Add("累計小功", 20);
            CloumnIndex.Add("累計大功", 21);
            CloumnIndex.Add("累計警告", 22);
            CloumnIndex.Add("累計小過", 23);
            CloumnIndex.Add("累計大過", 24);
            CloumnIndex.Add("日常生活表現評量", 25);
            CloumnIndex.Add("肌耐力", 26);
            CloumnIndex.Add("柔軟度", 27);
            CloumnIndex.Add("瞬發力", 28);
            CloumnIndex.Add("心肺耐力", 29);
            CloumnIndex.Add("體適能", 30);
            CloumnIndex.Add("多元學習表現", 31);
            CloumnIndex.Add("技藝教育成績", 32);
            CloumnIndex.Add("技藝優良", 33);
            CloumnIndex.Add("弱勢身分", 34);
            CloumnIndex.Add("弱勢積分", 35);
            CloumnIndex.Add("健康與體育", 36);
            CloumnIndex.Add("藝術與人文", 37);
            CloumnIndex.Add("綜合活動", 38);
            CloumnIndex.Add("均衡學習", 39);
            CloumnIndex.Add("家長意見", 40);
            CloumnIndex.Add("導師意見", 41);
            CloumnIndex.Add("輔導教師意見", 42);
            CloumnIndex.Add("適性輔導", 43);
            CloumnIndex.Add("其他比序項目_全民英檢", 44);
            CloumnIndex.Add("合計", 45);
            CloumnIndex.Add("報名「北區」五專學校代碼", 46);
            CloumnIndex.Add("報名「中區」五專學校代碼", 47);
            CloumnIndex.Add("報名「南區」五專學校代碼", 48);
            CloumnIndex.Add("競賽名稱", 49);
            CloumnIndex.Add("其他比序項目_多益測驗", 50);
            int index = 1;
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));
            Cells cs = wb.Worksheets[0].Cells;
            foreach (StudentObj obj in list)
            {
                cs[index, CloumnIndex["身分證字統一編號"]].PutValue(obj.IdNumber);
                cs[index, CloumnIndex["學生姓名"]].PutValue(obj.Name);
                cs[index, CloumnIndex["出生年(民國年)"]].PutValue(obj.Birth_Year.ToString().PadLeft(3,'0'));
                cs[index, CloumnIndex["出生月"]].PutValue(obj.Birth_Month.ToString().PadLeft(2, '0'));
                cs[index, CloumnIndex["出生日"]].PutValue(obj.Birth_Day.ToString().PadLeft(2, '0'));
                cs[index, CloumnIndex["年級"]].PutValue(obj.GradeYear);
                cs[index, CloumnIndex["班級"]].PutValue(obj.ClassName);
                cs[index, CloumnIndex["座號"]].PutValue(obj.SeatNo);
                cs[index, CloumnIndex["報名資格"]].PutValue(CheckTagId(obj.TagIds, 報名資格));
                cs[index, CloumnIndex["郵遞區號"]].PutValue(obj.ZipCode);
                cs[index, CloumnIndex["地址"]].PutValue(obj.Address);
                cs[index, CloumnIndex["市內電話"]].PutValue(obj.Contact_Phone);
                cs[index, CloumnIndex["行動電話"]].PutValue(obj.SMS_Phone);
                cs[index, CloumnIndex["特種生加分類別"]].PutValue(CheckTagId(obj.TagIds, 特種生加分類別));
                cs[index, CloumnIndex["報名費減免身分"]].PutValue(CheckTagId(obj.TagIds, 報名費減免身分));
                cs[index, CloumnIndex["擔任幹部"]].PutValue(obj.CadreTimesScore);
                cs[index, CloumnIndex["服務時數"]].PutValue(obj.ServiceHours);
                cs[index, CloumnIndex["服務學習"]].PutValue(obj.ServiceLearningScore);
                cs[index, CloumnIndex["累計嘉獎"]].PutValue(obj.MeritC);
                cs[index, CloumnIndex["累計小功"]].PutValue(obj.MeritB);
                cs[index, CloumnIndex["累計大功"]].PutValue(obj.MeritA);
                cs[index, CloumnIndex["累計警告"]].PutValue(obj.DemeritC);
                cs[index, CloumnIndex["累計小過"]].PutValue(obj.DemeritB);
                cs[index, CloumnIndex["累計大過"]].PutValue(obj.DemeritA);
                cs[index, CloumnIndex["日常生活表現評量"]].PutValue(obj.MeritDemeritScore);
                cs[index, CloumnIndex["肌耐力"]].PutValue(obj.CheckScore("仰臥起坐"));
                cs[index, CloumnIndex["柔軟度"]].PutValue(obj.CheckScore("坐姿體前彎"));
                cs[index, CloumnIndex["瞬發力"]].PutValue(obj.CheckScore("立定跳遠"));
                cs[index, CloumnIndex["心肺耐力"]].PutValue(obj.CheckScore("心肺適能"));
                cs[index, CloumnIndex["體適能"]].PutValue(obj.SportFitnessScore);

                int x = index + 1;
                string formula = "=IF(P" + x + "+S" + x + "+Z" + x + "+AE" + x + ">16,16,P" + x + "+S" + x + "+Z" + x + "+AE" + x + ")";
                cs[index, CloumnIndex["多元學習表現"]].Formula = formula;

                string[] tag = CheckTagId(obj.TagIds);
                cs[index, CloumnIndex["弱勢身分"]].PutValue(tag[0]);
                cs[index, CloumnIndex["弱勢積分"]].PutValue(tag[1]);

                Dictionary<string, decimal> dic = obj.GetDomainScores();
                cs[index, CloumnIndex["健康與體育"]].PutValue(dic.ContainsKey("健康與體育") ? dic["健康與體育"] : 0);
                cs[index, CloumnIndex["藝術與人文"]].PutValue(dic.ContainsKey("藝術與人文") ? dic["藝術與人文"] : 0);
                cs[index, CloumnIndex["綜合活動"]].PutValue(dic.ContainsKey("綜合活動") ? dic["綜合活動"] : 0);
                cs[index, CloumnIndex["均衡學習"]].PutValue(obj.DomainItemScore);
                cs[index, CloumnIndex["其他比序項目_全民英檢"]].PutValue(CheckTagId(obj.TagIds, 其他比序項目_全民英檢));
                cs[index, CloumnIndex["其他比序項目_多益測驗"]].PutValue(CheckTagId(obj.TagIds, 其他比序項目_多益測驗));

                formula = "=IF(AF" + x + "+AH" + x + "+AJ" + x + "+AN" + x + "+AR" + x + ">30,30,AF" + x + "+AH" + x + "+AJ" + x + "+AN" + x + "+AR" + x + ")";
                cs[index, CloumnIndex["合計"]].Formula = formula;

                index++;
                count++;
                progress += (int)(count * per);
                _BW.ReportProgress(progress);
            }

            //wb.Worksheets[0].AutoFitColumns();
            e.Result = wb;
        }

        private int SortStudent(StudentObj x, StudentObj y)
        {
            string xx = x.ClassName.PadLeft(20, '0');
            xx += x.SeatNo.PadLeft(3, '0');
            xx += x.StudentNumber.PadLeft(20, '0');

            string yy = y.ClassName.PadLeft(20, '0');
            yy += y.SeatNo.PadLeft(3, '0');
            yy += y.StudentNumber.PadLeft(20, '0');

            return xx.CompareTo(yy);
        }


        private int GetScore(DataRow row)
        {
            string sit_and_reach_degree = row["sit_and_reach_degree"].ToString();
            string standing_long_jump_degree = row["standing_long_jump_degree"].ToString();
            string sit_up_degree = row["sit_up_degree"].ToString();
            string cardiorespiratory_degree = row["cardiorespiratory_degree"].ToString();

            int score = 0;

            //sit_and_reach_degree
            if (sit_and_reach_degree == "金牌" || sit_and_reach_degree == "銀牌" || sit_and_reach_degree == "銅牌" || sit_and_reach_degree == "中等")
            {
                score += 2;
            }
            //standing_long_jump_degree
            if (standing_long_jump_degree == "金牌" || standing_long_jump_degree == "銀牌" || standing_long_jump_degree == "銅牌" || standing_long_jump_degree == "中等")
            {
                score += 2;
            }
            //sit_up_degree
            if (sit_up_degree == "金牌" || sit_up_degree == "銀牌" || sit_up_degree == "銅牌" || sit_up_degree == "中等")
            {
                score += 2;
            }
            //cardiorespiratory_degree
            if (cardiorespiratory_degree == "金牌" || cardiorespiratory_degree == "銀牌" || cardiorespiratory_degree == "銅牌" || cardiorespiratory_degree == "中等")
            {
                score += 2;
            }

            if (score > 6)
                score = 6;

            return score;
        }

        //多元學習表現
        private int GetScore(StudentObj obj)
        {
            int score = obj.ServiceLearningScore;
            score += obj.MeritDemeritScore;
            score += obj.SportFitnessScore;

            if (score > 16)
                score = 16;

            return score;
        }

        private int CheckTagId(List<string> list, Dictionary<string, int> refDic)
        {
            int retVal = 0;
            foreach (string sid in list)
            {
                foreach (string tagName in _MappingData.Keys)
                {
                    if (_MappingData[tagName].Contains(sid))
                    {
                        if (refDic.ContainsKey(tagName))
                        {
                            retVal = refDic[tagName];
                            return retVal;
                        }
                    }
                }
            }

            if (refDic.Equals(報名資格))
                return 1;
            else
                return 0;
        }

        private string[] CheckTagId(List<string> list)
        {
            string[] tag = new string[2];
            tag[0] = "0";
            tag[1] = "0";
            foreach (string sid in list)
            {
                foreach (string tagName in _MappingData.Keys)
                {
                    if (_MappingData[tagName].Contains(sid))
                    {
                        if (弱勢身分.ContainsKey(tagName))
                        {
                            tag[0] = 弱勢身分[tagName].ToString();
                            tag[1] = "2";
                            return tag;
                        }
                    }
                }
            }

            return tag;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (_BW.IsBusy)
            {
                MessageBox.Show("系統忙碌中,請稍後再試");
            }
            else
            {
                EnableForm(false);
                _BW.RunWorkerAsync();
            }
        }

        private void Column1Prepare()
        {
            List<string> all = new List<string>();
            DataGridViewRow row;

            foreach (string key in 報名資格.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in 報名費減免身分.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in 弱勢身分.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in 特種生加分類別.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in 其他比序項目_全民英檢.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in 其他比序項目_多益測驗.Keys)
            {
                if (!all.Contains(key))
                    all.Add(key);
            }

            foreach (string key in all)
            {
                row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Cells[0].Value = key;
                dataGridViewX1.Rows.Add(row);
            }

        }

        private void Column2Prepare()
        {
            _Column2Items = new Dictionary<String, String>();

            DataTable dt = _Q.Select("SELECT * FROM tag WHERE category='Student' ORDER BY prefix,name");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();

                string key = "";
                if (string.IsNullOrWhiteSpace(prefix))
                {
                    key = name;
                }
                else
                {
                    key = prefix + ":" + name;
                }

                if (!_Column2Items.ContainsKey(key))
                {
                    _Column2Items.Add(key, id);
                }
            }

            this.Column2.Items.Add("");
            foreach (string item in _Column2Items.Keys)
            {
                this.Column2.Items.Add(item);
            }
        }

        private void SaveSetting()
        {
            List<Setting> UDTlist = _A.Select<Setting>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist.Clear(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value == null) //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                Setting obj = new Setting();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }

        private void LoadSetting()
        {
            List<Setting> UDTlist = _A.Select<Setting>(); //檢查UDT並回傳資料

            UDTlist.Sort(delegate(Setting x, Setting y)
            {
                string xx = x.UID.PadLeft(10, '0');
                string yy = y.UID.PadLeft(10, '0');

                return xx.CompareTo(yy);
            });

            DataGridViewRow row;
            if (UDTlist.Count > 0) //UDT內有設定才做讀取
            {
                for (int i = 0; i < UDTlist.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = UDTlist[i].Target;
                    row.Cells[1].Value = UDTlist[i].Source;
                    dataGridViewX1.Rows.Add(row);
                }
            }
            else
            {
                Column1Prepare();
            }
        }

        private void Report_Load(object sender, EventArgs e)
        {
            LoadSetting();
        }

        private void EnableForm(bool enable)
        {
            this.buttonX1.Enabled = enable;
            this.dataGridViewX1.Enabled = enable;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                List<Setting> UDTlist = _A.Select<Setting>();
                _A.DeletedValues(UDTlist); //清除UDT資料
                dataGridViewX1.Rows.Clear();  //清除datagridview資料
                LoadSetting(); //再次讀入Mapping設定
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
            }

        }
    }
}
