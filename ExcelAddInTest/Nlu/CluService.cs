using Azure;
using Azure.AI.Language.Conversations;
using Azure.Core;
using Azure.Core.Pipeline;
using ExcelAddInTest.Nlu;
using ExcelAddInTest.Nlu.NluModels;
using Microsoft.Extensions.Azure;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security.Authentication;
using System.Threading.Tasks;

namespace ExcelAddInTest
{

    public class CluService : INlu
    {
        //ConversationAnalysisClient is used for interacting with Conversation models form Azure Language Studio
        private readonly ConversationAnalysisClient _client;
        private readonly string _project;
        private readonly string _deployment;

        /// <summary>
        /// Initializes CluService that handles the interaction with the NLU (Natural Language Model)
        /// </summary>
        /// <param name="endpoint">The endpoint for the CLU (Conversational Language Understanding) Service</param>
        /// <param name="key">The key for the CLU Service</param>
        /// <param name="projectName">The name of the project</param>
        /// <param name="deploymentName">The deployment name of the model you want to use for the processing of the utterances</param>
        /// <exception cref="ArgumentException"></exception>
        public CluService(string endpoint, string key, string projectName, string deploymentName)
        {
            if (string.IsNullOrWhiteSpace(endpoint)) throw new ArgumentException("endpoint missing");
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("key missing");

            ///Because of the older version of .NET that we use, set the ssl protocol to Tls12 
            ///(needed by the azure services)
            var handler = new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12
            };
            var httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(30)
            };
            ///used to set the option for our ConversationAnalysisClient class 
            ///(we need to use the ssl protocol Tls12)
            var opts = new ConversationsClientOptions
            {
                Retry =
                {
                    Mode = RetryMode.Exponential,
                    Delay = TimeSpan.FromMilliseconds(500),
                    MaxRetries = 0   
                },
                Transport = new HttpClientTransport(httpClient)
            };


            _client = new ConversationAnalysisClient(new Uri(endpoint.Trim()), new AzureKeyCredential(key.Trim()), opts);
            _project = (projectName ?? "").Trim();
            _deployment = (deploymentName ?? "").Trim();
            
        }
     
        /// <summary>
        /// Method analyzes sends the utterance to the clu model deployed and receives 
        /// the top intent and recognized entities.
        /// </summary>
        /// <param name="text">the utterance that is sent</param>
        /// <returns>NluResult entity that contains the information.</returns>
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
                            language = "en-US"
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
                string raw = resp.Content.ToString(); // raw JSON

                // logging the response for debugging purposes 
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

                // some versions use "prediction", others might use "result.prediction"
                // so we treat the both cases
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
               
                // extract entities recognized
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

                // top intents scores (if they exist, for debug)
                JArray intentsArr = prediction["intents"] as JArray;
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
            catch (RequestFailedException ex)
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
                // dismantling AggregateException to catch RequestFailedException from inside
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
