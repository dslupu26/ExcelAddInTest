using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelAddInTest
{
    public static class Config
    {
        //CLU
        public static string CluEndpoint
            => ConfigurationManager.AppSettings["AZURE_CLU_ENDPOINT"];
        public static string CluKey
            => ConfigurationManager.AppSettings["AZURE_CLU_KEY"];

        public static string CluProjectName
            => ConfigurationManager.AppSettings["AZURE_CLU_PROJECT_NAME"];

        public static string CluDeployment
            => ConfigurationManager.AppSettings["AZURE_CLU_DEPLOYMENT_NAME"];

        // Speech
        public static string SpeechKey
            => ConfigurationManager.AppSettings["AZURE_SPEECH_KEY"];

        public static string SpeechRegion
            => ConfigurationManager.AppSettings["AZURE_SPEECH_REGION"];

        public static string SpeechLanguage
            => ConfigurationManager.AppSettings["AZURE_SPEECH_LANGUAGE"];
    }
}
