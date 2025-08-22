using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest
{
    public static class CluDiag
    {
        public static async Task<string> TestCluRestAsync(
        string endpoint, string key, string projectName, string deploymentName, string text)
        {
            // sanitizare
            endpoint = (endpoint ?? "").Trim().TrimEnd('/');
            var url = endpoint + "/language/:analyze-conversations?api-version=2023-04-01";

            var payload = new
            {
                analysisInput = new
                {
                    conversationItem = new
                    {
                        id = "1",
                        participantId = "user",
                        text = (text ?? "").Trim().TrimEnd('.', '!', '?'),
                        modality = "text",
                        language = "en-us"
                    }
                },
                parameters = new
                {
                    projectName = projectName,
                    deploymentName = deploymentName,
                    stringIndexType = "Utf16CodeUnit",
                    verbose = true
                },
                kind = "Conversation"
            };

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
            using (var http = new System.Net.Http.HttpClient())
            {
                var req = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, url);
                req.Headers.Add("Ocp-Apim-Subscription-Key", key.Trim());
                req.Content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");

                var resp = await http.SendAsync(req);
                var body = await resp.Content.ReadAsStringAsync();
                return "[REST] Status=" + (int)resp.StatusCode + " " + resp.StatusCode + "\r\n" + body;
            }
        }

    }
}
