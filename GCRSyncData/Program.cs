using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GCRSyncData
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            GCRSyncData.ConnectDB();
            GCRSyncData.Start("gcradmin@yourdomain.org");
            //GCRSyncData.Start("anyotherteacher@yourdomain.org");
        }
    }
}
