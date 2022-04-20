using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColleageExamFreeApplicationForm
{
    class Permissions
    {
        public static string 五專集體報名表 { get { return "ColleageExamFreeApplicationForm.46CE921D-338A-4980-86BB-2E6C39FBCFA5"; } }
        public static string 南區五專聯合免試入學集體報名表 { get { return "ColleageExamFreeApplicationForm.46CE921D-338A-4980-86BB-2E6C39FBCFA6"; } }
        public static bool 五專集體報名表權限
        {
            get { return FISCA.Permission.UserAcl.Current[五專集體報名表].Executable; }
        }

        public static bool 五專聯合集體報名表權限
        {
            get { return FISCA.Permission.UserAcl.Current[南區五專聯合免試入學集體報名表].Executable; }
        }
    }
}
