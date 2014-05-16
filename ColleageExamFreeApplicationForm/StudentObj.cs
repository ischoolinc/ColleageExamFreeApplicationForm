using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ColleageExamFreeApplicationForm
{
    class StudentObj
    {
        public string Id, Name, ClassName, IdNumber, SeatNo, StudentNumber, GradeYear, ZipCode, Address, Contact_Phone, SMS_Phone;
        public int Birth_Year, Birth_Month, Birth_Day;
        public decimal ServiceHours;
        public int CadreTimes;
        //原始獎懲紀錄
        public int MeritA, MeritB, MeritC, DemeritA, DemeritB, DemeritC;
        //功過相抵後獎懲紀錄
        public int MA, MB, MC, DA, DB, DC;
        //體適能常模
        public string sit_and_reach_degree, standing_long_jump_degree, sit_up_degree, cardiorespiratory_degree;

        public Dictionary<string, Dictionary<string, decimal>> DomainScores;
        private Dictionary<string, decimal> domainAverageScores;

        public List<string> TagIds;

        public Dictionary<string, decimal> GetDomainScores()
        {
            domainAverageScores.Clear();

            foreach (string domain in DomainScores.Keys)
            {
                int count = 0;

                if (!domainAverageScores.ContainsKey(domain))
                {
                    domainAverageScores.Add(domain, 0);
                }

                foreach (KeyValuePair<string, decimal> kvp in DomainScores[domain])
                {
                    if (kvp.Key == "1上" || kvp.Key == "1下" || kvp.Key == "2上" || kvp.Key == "2下" || kvp.Key == "3上" || kvp.Key == "7上" || kvp.Key == "7下" || kvp.Key == "8上" || kvp.Key == "8下" || kvp.Key == "9上")
                    {
                        domainAverageScores[domain] += kvp.Value;
                        count++;
                    }
                }

                if (count > 0)
                    domainAverageScores[domain] /= count;

                domainAverageScores[domain] = (int)domainAverageScores[domain];
                //domainAverageScores[domain] = Math.Round(domainAverageScores[domain], 2, MidpointRounding.AwayFromZero);
            }

            return domainAverageScores;
        }

        public int DomainItemScore
        {
            get
            {
                int score = 0;

                if (domainAverageScores.ContainsKey("健康與體育"))
                {
                    if (domainAverageScores["健康與體育"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (domainAverageScores.ContainsKey("藝術與人文"))
                {
                    if (domainAverageScores["藝術與人文"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (domainAverageScores.ContainsKey("綜合活動"))
                {
                    if (domainAverageScores["綜合活動"] >= 60)
                    {
                        score += 2;
                    }
                }

                return score;
            }
        }


        //public int ServiceHoursScore
        //{
        //    get
        //    {
        //        int score = (int)(ServiceHours / 8);

        //        if (score > 99)
        //            score = 99;

        //        return score;
        //    }
        //}

        public int CadreTimesScore
        {
            get
            {
                int score = CadreTimes;

                if (score > 6)
                    score = 6;

                return score;
            }
        }

        public int ServiceLearningScore
        {
            get
            {
                int score = (int)(ServiceHours / 8);
                score += CadreTimes;

                if (score > 7)
                    score = 7;

                return score;
            }
        }

        public int MeritDemeritScore
        {
            get
            {
                int score = 0;

                if (DA == 0 && DB == 0 && DC == 0)
                {
                    score = 1;
                }

                if (!HasDemeritAB)
                {
                    if (MA > 0)
                    {
                        score = 4;
                    }
                    else if (MB > 0)
                    {
                        score = 3;
                    }
                    else if (MC > 0)
                    {
                        score = 2;
                    }
                }

                return score;
            }
        }

        public bool HasDemeritAB
        {
            get
            {
                if (DemeritA == 0 && DemeritB == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        public int SportFitnessScore
        {
            get
            {
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
        }

        public int CheckScore(string str)
        {
            string value = "";
            switch (str)
            {
                case "坐姿體前彎":
                    value = sit_and_reach_degree;
                    break;

                case "立定跳遠":
                    value = standing_long_jump_degree;
                    break;

                case "仰臥起坐":
                    value = sit_up_degree;
                    break;

                case "心肺適能":
                    value = cardiorespiratory_degree;
                    break;

                default:
                    value = "";
                    break;
            }

            if (value == "金牌" || value == "銀牌" || value == "銅牌" || value == "中等")
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        private void SetAddress(string str)
        {
            XmlDocument xml = new XmlDocument();
            this.ZipCode = "";
            this.Address = "";

            if (!string.IsNullOrWhiteSpace(str))
            {
                xml.LoadXml(str);
                try
                {
                    this.ZipCode = xml.SelectSingleNode("//ZipCode").InnerText;
                }
                catch
                {
                }

                try
                {
                    this.Address += xml.SelectSingleNode("//County").InnerText;
                }
                catch
                {
                }

                try
                {
                    this.Address += xml.SelectSingleNode("//Town").InnerText;
                }
                catch
                {
                }

                try
                {
                    this.Address += xml.SelectSingleNode("//District").InnerText;
                }
                catch
                {
                }

                try
                {
                    this.Address += xml.SelectSingleNode("//Area").InnerText;
                }
                catch
                {
                }

                try
                {
                    this.Address += xml.SelectSingleNode("//DetailAddress").InnerText;
                }
                catch
                {
                }
            }
            
        }

        public StudentObj(DataRow row)
        {
            this.Id = row["id"].ToString();
            this.Name = row["name"].ToString();
            this.ClassName = row["class_name"].ToString();
            this.IdNumber = row["id_number"].ToString();
            this.SeatNo = row["seat_no"].ToString();
            this.StudentNumber = row["student_number"].ToString();
            try
            {
                DateTime time = Convert.ToDateTime(row["birthdate"].ToString());
                this.Birth_Year = time.Year - 1911;
                this.Birth_Month = time.Month;
                this.Birth_Day = time.Day;
            }
            catch
            {
                this.Birth_Year = 0;
                this.Birth_Month = 0;
                this.Birth_Day = 0;
            }

            this.GradeYear = row["grade_year"].ToString();

            //聯絡地址
            SetAddress(row["mailing_address"].ToString());

            //戶籍地址
            if (string.IsNullOrWhiteSpace(this.Address))
                SetAddress(row["permanent_address"].ToString());

            this.Contact_Phone = row["contact_phone"].ToString();
            this.SMS_Phone = row["sms_phone"].ToString();
            this.ServiceHours = 0;
            this.CadreTimes = 0;
            this.MeritA = 0;
            this.MeritB = 0;
            this.MeritC = 0;
            this.DemeritA = 0;
            this.DemeritB = 0;
            this.DemeritC = 0;
            this.MA = 0;
            this.MB = 0;
            this.MC = 0;
            this.DA = 0;
            this.DB = 0;
            this.DC = 0;
            sit_and_reach_degree = "";
            standing_long_jump_degree = "";
            sit_up_degree = "";
            cardiorespiratory_degree = "";
            DomainScores = new Dictionary<string, Dictionary<string, decimal>>();
            domainAverageScores = new Dictionary<string, decimal>();
            TagIds = new List<string>();
        }
    }
}
