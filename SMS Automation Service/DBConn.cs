using System;

namespace SMS_Automation_Service
{
    public partial class ReadFilesServices
    {
        public struct DBConn
        {
            public string ConnID;
            public string ConnName;
            public string DSNName;
            public string Username;
            public string Password;
            public string DBType;
            public string Tbl;
            public string SQL;
            public string DateCol;
            public DateTime LastDate;
            public string ServiceName;
            public ReadFilesServices.TempColl[] TempColls;
        }

    }
}
