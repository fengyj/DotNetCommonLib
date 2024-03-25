namespace me.fengyj.CommonLib.Utils.App {
    public class AppSettingManagement {


        public AppSettingManagement(
            string envVarNameForEnv = "Environment",
            string configKeyForEnv = "Environment",
            string userSettingFolder = "ConfigFiles") {

            this.EnvVariableNameForEnv = envVarNameForEnv;
            this.ConfigKeyForEnv = configKeyForEnv;
            this.UserSettingFolder = userSettingFolder;
        }

        public string EnvVariableNameForEnv { get; private set; }
        public string ConfigKeyForEnv { get; private set; }
        public string UserSettingFolder { get; private set; }


    }

    public class Environments {

        public const string Production = "prod";
        public const string Production_DR = "prod-dr";
        public const string UAT = "uat";
        public const string UAT_DR = "uat-dr";
        public const string SIT = "sit";
        public const string QA = "qa";
        public const string Dev = "dev";

        public static readonly List<string> ProductionEnvironments = [Production, Production_DR];
    }
}
