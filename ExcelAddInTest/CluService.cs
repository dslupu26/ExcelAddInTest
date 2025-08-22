using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure;
using Azure.AI.Language.Conversations;
using Azure.Core;
using Newtonsoft.Json.Linq;

namespace ExcelAddInTest
{
    public class NluEntity
    {
        public string Category { get; set; }
        public string Text { get; set; }
    }

    public class NluResult
    {
        public string TopIntent { get; set; }
        public List<NluEntity> Entities { get; set; }
        public string RawJson { get; set; }
    }

    public class CluService
    {
        private readonly ConversationAnalysisClient _client;
        private readonly string _project;
        private readonly string _deployment;

        public CluService(string endpoint, string key, string projectName, string deploymentName)
        {
            if (string.IsNullOrWhiteSpace(endpoint)) throw new ArgumentException("endpoint missing");
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("key missing");

            _client = new ConversationAnalysisClient(new Uri(endpoint.Trim()), new AzureKeyCredential(key.Trim()));
            _project = (projectName ?? "").Trim();
            _deployment = (deploymentName ?? "").Trim();
        }

        public async Task<NluResult> AnalyzeAsync(string text)
        {
            try
            {
                var cleaned = (text ?? "").Trim().TrimEnd('.', '!', '?');

                var payload = new
                {
                    analysisInput = new
                    {
                        conversationItem = new
                        {
                            id = "1",
                            participantId = "user",
                            text = cleaned,
                            modality = "text",
                            language = "en-us" 
                        }
                    },
                    parameters = new
                    {
                        projectName = _project,
                        deploymentName = _deployment,
                        stringIndexType = "Utf16CodeUnit", 
                        verbose = true
                    },
                    kind = "Conversation"
                };

                Response resp = await _client.AnalyzeConversationAsync(RequestContent.Create(payload));
                string raw = resp.Content.ToString(); // JSON brut

                // log minim de erori din corp, dacă există
                var jo = JObject.Parse(raw);
                var error = jo["error"];
                if (error != null)
                {
                    return new NluResult
                    {
                        TopIntent = "None",
                        Entities = new List<NluEntity>(),
                        RawJson = "[CLU ERROR BODY]\r\n" + error.ToString() + "\r\n" + raw
                    };
                }

                // unele versiuni pun direct "prediction", altele sub "result.prediction"
                var prediction = jo["result"] != null ? jo["result"]["prediction"] : jo["prediction"];
                if (prediction == null)
                {
                    return new NluResult
                    {
                        TopIntent = "None",
                        Entities = new List<NluEntity>(),
                        RawJson = "[NO PREDICTION IN RESPONSE]\r\n" + raw
                    };
                }

                string top = (string)(prediction["topIntent"] ?? prediction["top_intent"]) ?? "None";

                var entities = new List<NluEntity>();
                var ents = prediction["entities"] as JArray;
                if (ents != null)
                {
                    foreach (var e in ents)
                    {
                        entities.Add(new NluEntity
                        {
                            Category = (string)(e["category"] ?? e["type"]) ?? "",
                            Text = (string)(e["text"] ?? e["value"] ?? "")
                        });
                    }
                }

                // (opțional) extrage și scorurile pentru debug
                var intentsArr = prediction["intents"] as JArray;
                if (intentsArr != null && intentsArr.Count > 0)
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("[INTENTS]");
                    int k = 0;
                    foreach (var it in intentsArr)
                    {
                        var name = (string)(it["category"] ?? it["intent"] ?? "unknown");
                        var score = (double?)it["confidenceScore"] ?? (double?)it["score"] ?? 0.0;
                        sb.AppendLine(" - " + name + ": " + score.ToString("0.000"));
                        if (++k >= 5) break;
                    }
                    raw += "\r\n" + sb.ToString();
                }

                return new NluResult { TopIntent = top, Entities = entities, RawJson = raw };
            }
            catch (Azure.RequestFailedException ex)
            {
                return new NluResult
                {
                    TopIntent = "None",
                    Entities = new List<NluEntity>(),
                    RawJson = "[REQUEST FAILED] Status=" + ex.Status + " Code=" + ex.ErrorCode + " Msg=" + ex.Message
                };
            }
            catch (Exception ex)
            {
                // dezambalăm AggregateException ca să prindem RequestFailedException din interior
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("[UNHANDLED] " + ex.GetType().Name + ": " + ex.Message);

                Exception cur = ex;
                while (cur != null)
                {
                    sb.AppendLine("INNER -> " + cur.GetType().Name + ": " + cur.Message);
                    var rfe = cur as Azure.RequestFailedException;
                    if (rfe != null)
                        sb.AppendLine("   Status=" + rfe.Status + " Code=" + rfe.ErrorCode);
                    cur = cur.InnerException;
                }

                return new NluResult
                {
                    TopIntent = "None",
                    Entities = new List<NluEntity>(),
                    RawJson = sb.ToString()
                };
            }

        }

    }
}
