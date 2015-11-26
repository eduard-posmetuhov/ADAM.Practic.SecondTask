using Adam.Core;
using Adam.Core.Classifications;
using Adam.Core.Fields;
using Adam.Core.Records;
using Adam.Core.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace First_Task
{
    class Program
    {
        static void Main(string[] args)
        {            
            Application app = new Application();
            LogIn(app);            

            Classification c = new Classification(app);
            ClassificationHelper ch = new ClassificationHelper(app);
            Guid? rootGuid= ch.GetId(new SearchExpression("name = 'Luxottica Content*'"));            
            
            RecordCollection rc = new RecordCollection(app);
            rc.Load(new SearchExpression(String.Format("classification = '{0}'", rootGuid)));
            Dictionary<string, Guid> dictionaryForSearch = new Dictionary<string, Guid>();
            Dictionary<Guid, string> dictionaryForExcel = new Dictionary<Guid, string>();            
            string currentValue=null;
            Console.WriteLine(rc.Count);
            foreach (Record r in rc.Take<Record>(100))
            {                
                if (r.Fields.GetField<TextField>("UPC code") != null)
                {
                    if (!String.IsNullOrEmpty(r.Fields.GetField<TextField>("UPC code").Value))
                    {
                        currentValue = r.Fields.GetField<TextField>("UPC code").Value.TrimStart('0');
                        if(dictionaryForSearch.ContainsKey(currentValue)/* .Contains(currentValue)*/)
                        {
                            dictionaryForExcel.Add(r.Id,currentValue);
                            Guid g = dictionaryForSearch[currentValue];
                            if (!dictionaryForExcel.ContainsKey(g))
                            {
                                dictionaryForExcel.Add(g, currentValue);
                            }
                        }
                        else dictionaryForSearch.Add(currentValue,r.Id);
                        Console.WriteLine(currentValue);
                    }
                }
            }
            Console.WriteLine(dictionaryForExcel.Count);
            ExcelWriter writer = new ExcelWriter(@"D:\book1.xls");
            writer.Write(dictionaryForExcel);
            Console.ReadKey();            
        }

        private static void LogIn(Application app)
        {
            LogOnStatus status = app.LogOn("LUXDAM", "Eduard_Pasmetukhau", "P2ssw0rd!");
            if (status == LogOnStatus.LoggedOn)
            {
                Console.WriteLine("Ok");
            }
            else Console.WriteLine(status);

        }
    }

    
}
