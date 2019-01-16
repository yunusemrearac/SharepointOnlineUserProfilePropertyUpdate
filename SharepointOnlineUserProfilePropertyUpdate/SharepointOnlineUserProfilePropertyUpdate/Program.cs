using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Configuration;
using System.Security;

namespace SharepointOnlineUserProfilePropertyUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            bool update = UpdateSharepointOnlineUserProfileProperty("yunusemrearac@yunusemrearac.onmicrosoft.com", "Title", "Test");

            if (update)
            {
                Console.WriteLine("Güncelleme başarılı bir şekilde gerçekleştirilmiştir.");
            }
        }

        public static bool UpdateSharepointOnlineUserProfileProperty(string email, string property, string value)
        {
            bool updateResult = false;
            try
            {
                string tenantAdministrationUrl = ConfigurationManager.AppSettings["TenantAdministrationUrl"].ToString();
                string tenantAdminLoginName = ConfigurationManager.AppSettings["TenantAdminLoginName"].ToString();
                string tenantAdminPassword = ConfigurationManager.AppSettings["TenantAdminPassword"].ToString();

                string UserAccountName = "i:0#.f|membership|" + email;

                using (ClientContext clientContext = new ClientContext(tenantAdministrationUrl))
                {
                    SecureString passWord = new SecureString();

                    foreach (char c in tenantAdminPassword.ToCharArray()) passWord.AppendChar(c);

                    clientContext.Credentials = new SharePointOnlineCredentials(tenantAdminLoginName, passWord);

                    PeopleManager peopleManager = new PeopleManager(clientContext);

                    peopleManager.SetSingleValueProfileProperty(UserAccountName, property, value);

                    clientContext.ExecuteQuery();

                    updateResult = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Sharepoint User Property güncellenirken hata oluştu. User Email : " + email + " Error : ");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

            return updateResult;
        }
    }
}
