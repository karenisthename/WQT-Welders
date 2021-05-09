using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

namespace WQT_Welders
{
    public static class InitializeData
    {
        static useSQL.useSQL dbCon;
        public static string conStr = File.ReadAllText(@"..\ndtCon.cfg");

        public static void InitWelders(DataTable dtWelder)
        {
            using (dbCon = new useSQL.useSQL(conStr))
            {
                dtWelder = dbCon.PerformQuery("select distinct id,rtrim(stamp) as [stamp],name,subcontractor, issueddate,expdate,remarks,active "
                                                + "from WelderList where stamp not like 'F%' order by stamp ");
            }
        }

        public static void InitJoints(DataTable dtJoints)
        {
            string script = File.ReadAllText(@"..\QUERIES\Joints.sql");

            using (dbCon = new useSQL.useSQL(conStr))
            {
                dtJoints = dbCon.PerformQuery(script);
            }
        }
    }
}
