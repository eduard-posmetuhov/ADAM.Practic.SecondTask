using Adam.Core;
using Adam.Core.Classifications;
using Adam.Core.Fields;
using Adam.Core.Records;
using Adam.Core.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;

namespace First_Task
{
    class Program
    {
        private static Application _app;
        static void Main(string[] args)
        {
            try
            {
                _app = new Application();
                LogIn(_app);
                RecordCollection rc = new RecordCollection(_app);
                rc.Load(new SearchExpression("FieldName('UPC code') = *"));
                Dictionary<string, Guid> dictionaryForSearch = new Dictionary<string, Guid>();
                Dictionary<Guid, string> dictionaryForExcel = new Dictionary<Guid, string>();
                string currentValue = null;               
                foreach (Record r in rc)
                {
                    currentValue = r.Fields.GetField<TextField>("UPC code").Value.TrimStart('0');
                    if (!String.IsNullOrEmpty(currentValue))
                    {                        
                        if (dictionaryForSearch.ContainsKey(currentValue))
                        {
                            dictionaryForExcel.Add(r.Id, currentValue);
                            Guid g = dictionaryForSearch[currentValue];
                            if (!dictionaryForExcel.ContainsKey(g))
                                dictionaryForExcel.Add(g, currentValue); 
                        }
                        else
                            dictionaryForSearch.Add(currentValue, r.Id);                        
                    }
                }
                Console.WriteLine("Count of duplicate assets: " + dictionaryForExcel.Count);
                ExcelWriter writer = new ExcelWriter(ConfigurationSettings.AppSettings["ExcelPath"]);
                writer.Write(dictionaryForExcel,_app);
                Console.WriteLine("Save complite");
                Console.ReadKey();
            }            
            finally
            {
                if(_app.IsLoggedOn)
                _app.LogOff();
            }            
        }

        private static void LogIn(Application app)
        {
            string connection = ConfigurationSettings.AppSettings["AdamRegistrationName"];
            string user = ConfigurationSettings.AppSettings["AdamAdminUserName"];
            string pass = ConfigurationSettings.AppSettings["AdamAdminPassword"];
            LogOnStatus status = app.LogOn(connection, user, pass);
            if (status == LogOnStatus.LoggedOn)
            {
                Console.WriteLine("LogOn complite");
            }
            else Console.WriteLine(status);

        }
    }

    
}
