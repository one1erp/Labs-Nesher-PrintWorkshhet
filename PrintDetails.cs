using System.Collections.Generic;
using System.Linq;
using DAL;

namespace PrintWorkshhet
{
    public class PrintDetails
    {
        private Result result;






        private List<Result> _results;
        public List<Aliquot> Aliquots { get; set; }

        public PrintDetails()
        {
            _results = new List<Result>();
            Aliquots = new List<Aliquot>();
        }

        public PrintDetails(Result result)
            : this()
        {
            TestTemplateName = result.Test.TestTemplate.Name;
            AddResult(result);
        }

        public void AddResult(Result result)
        {
            Aliquots.Add(result.Test.Aliquot);
            _results.Add(result);
        }
        public string TestTemplateName { get; set; }

        public List<string> GetResultNames()
        {
            return _results.Select(r => r.Name).Distinct().ToList();
        }

        internal bool HasResult(string aliqName, string resultName)
        {
            Aliquot aliq = this.Aliquots.Where(a => a.Name == aliqName).FirstOrDefault();

            if (aliq != null)
            {
                List<Result> results = aliq.Tests.SelectMany(test => test.Results).ToList();

                if (results.ToList().Any(res => res.Name == resultName))
                {
                    return false;
                }
            }
            return true;
        }
    }
}