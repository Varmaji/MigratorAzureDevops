using System.Collections.Generic;

namespace MigratorAzureDevops.Models.Accounts
{
    public class AccountsResponse
    {
        public class Properties
        {
        }
        public class Value
        {
            public string accountId { get; set; }
            public string accountUri { get; set; }
            public string accountName { get; set; }
            public Properties properties { get; set; }
        }

        public class AccountList
        {
            public int count { get; set; }
            public IList<Value> value { get; set; }
        }
       
    }
    public class ProfileDetails
    {
        public string displayName { get; set; }
        public string publicAlias { get; set; }
        public string emailAddress { get; set; }
        public string id { get; set; }
        public string ErrorMessage { get; set; }
    }
    public class AccessDetails
    {
        public string access_token { get; set; }
        public string token_type { get; set; }
        public string expires_in { get; set; }
        public string refresh_token { get; set; }
    }
    public class LoginModel
    {
        public string AccountName { get; set; }

        public string PAT { get; set; }

        public string Message { get; set; }

        public string Event { get; set; }

        public string name { get; set; }

        public string EnableExtractor { get; set; }
        public string TemplateURL { get; set; }
    }
    
}