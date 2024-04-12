using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColleageExamFreeApplicationForm
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {   // [ischoolkingdom] Vicky更動，原107年五專(優先)免試入學相關報表，移除"107年"字樣
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            
            #region 優先
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「優先」免試入學集體報名表"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「優先」免試入學集體報名表"].Click += delegate
            {
                Report_priority report = new Report_priority();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.五專優先免試集體報名表權限)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「優先」免試入學集體報名表"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「優先」免試入學集體報名表"].Enable = false;
                }
            }; 
            #endregion

            #region 聯合
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「聯合」免試入學集體報名表"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「聯合」免試入學集體報名表"].Click += delegate
            {
                Report_joint report = new Report_joint();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.五專聯合免試集體報名表權限)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「聯合」免試入學集體報名表"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「聯合」免試入學集體報名表"].Enable = false;
                }
            };
            #endregion

            #region 完全
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「完全」免試入學集體報名表"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「完全」免試入學集體報名表"].Click += delegate
            {
                Report_Completely report = new Report_Completely();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.五專完全免試集體報名表權限)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「完全」免試入學集體報名表"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專「完全」免試入學集體報名表"].Enable = false;
                }
            }; 
            #endregion

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.五專優先免試集體報名表, "五專「優先」免試入學集體報名表"));
            permission.Add(new RibbonFeature(Permissions.五專聯合免試集體報名表, "五專「聯合」免試入學集體報名表"));
            permission.Add(new RibbonFeature(Permissions.五專完全免試集體報名表, "五專「完全」免試入學集體報名表"));
        }
    }
}
