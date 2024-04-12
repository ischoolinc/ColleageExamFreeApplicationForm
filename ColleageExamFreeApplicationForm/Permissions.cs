using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColleageExamFreeApplicationForm
{
    class Permissions
    {
        public static string 五專優先免試集體報名表
        {
            get { return "ColleageExamFreeApplicationForm.46CE921D-338A-4980-86BB-2E6C39FBCFA5"; }
        }
        public static string 五專聯合免試集體報名表
        {
            get { return "ColleageExamFreeApplicationForm.46CE921D-338A-4980-86BB-2E6C39FBCFA6"; }
        }

        public static string 五專完全免試集體報名表
        {
            get { return "ColleageExamFreeApplicationForm.BE4F13B6-9963-4D08-94A8-A23A0E84D62F"; }
        }

        public static bool 五專優先免試集體報名表權限
        {
            get { return FISCA.Permission.UserAcl.Current[五專優先免試集體報名表].Executable; }
        }

        public static bool 五專聯合免試集體報名表權限
        {
            get { return FISCA.Permission.UserAcl.Current[五專聯合免試集體報名表].Executable; }
        }

        public static bool 五專完全免試集體報名表權限
        {
            get { return FISCA.Permission.UserAcl.Current[五專完全免試集體報名表].Executable; }
        }
    }
}
