namespace terminalServerCore {
    public class Config {
        public string Ipaddress { get; set; }
        public string Database { get; set; }
        public string Port { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public string Customer { get; set; }
        public string Email { get; set; }
        public string DownloadEvery { get; set; }
        public string DeleteFilesAfterDays { get; set; }
        public string DatabaseType { get; set; }
        public string CloseOnlyAutomaticIdles { get; set; }
        public string AddCyclesToOrder { get; set; }
        public string SmtpClient { get; set; }
        public string SmtpUsername { get; set; }
        public string SmtpPassword { get; set; }
        public string SmtpPort { get; set; }

        public  Config() {
            Ipaddress = "zapsidatabase";
            Port = "3306";
            Database = "zapsi2";
            Login = "zapsi_uzivatel";
            Password = "zapsi";
            Customer = "unactivated software";
            Email = "jahoda@zapsi.eu";
            DownloadEvery = "10000";
            DeleteFilesAfterDays = "-20";
            DatabaseType = "mysql";
            CloseOnlyAutomaticIdles = "0";
            SmtpClient = "smtp.forpsi.com";
            SmtpPort = "25";
            SmtpUsername = "support@zapsi.eu";
            SmtpPassword = "support01..";
            AddCyclesToOrder = "0";
        }

    }
}