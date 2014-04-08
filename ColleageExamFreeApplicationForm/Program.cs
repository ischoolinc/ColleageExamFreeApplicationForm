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
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專集體報名表"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專集體報名表"].Click += delegate
            {
                Report report = new Report();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.五專集體報名表權限)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專集體報名表"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["五專集體報名表"].Enable = false;
                }
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.五專集體報名表, "五專集體報名表"));
        }
    }
}
