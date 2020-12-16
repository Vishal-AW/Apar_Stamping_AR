using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apar_Stamping_AR.Model;

namespace Apar_Stamping_AR
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ErrorModel> errorModels = new List<ErrorModel>();

            try
            {
                var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"].ToString();
                var ListName = ConfigurationManager.AppSettings["ListName"].ToString();

                errorModels = CustomSharePointUtility.GetActiveErrorList(siteUrl, ListName);

                if (errorModels.Count > 0)
                {
                    CustomSharePointUtility.UpdateURL(errorModels, siteUrl, ListName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
