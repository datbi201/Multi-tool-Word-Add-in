// ============================================================================
//  AiClient.cs - Kết nối AI trực tuyến đa nhà cung cấp (OpenAI, Anthropic, Gemini),
//                theo dõi token phiên làm việc và hỗ trợ kiểm thử giả lập (Rev0.9.3)
// ============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    /// <summary>Phản hồi HTTP trừu tượng để tách biệt lớp mạng khỏi logic kiểm thử.</summary>
    public sealed class AiHttpResponse
    {
        public int StatusCode;
        public string Body;
        public string Error;
    }

    /// <summary>Giao diện gửi yêu cầu HTTP (có thể thay thế bằng bản giả lập khi chạy unit test).</summary>
    public interface IAiTransport
    {
        AiHttpResponse Send(string method, string url, Dictionary<string, string> headers, string jsonBody, int timeoutMs);
    }

    /// <summary>Triển khai HTTP mặc định trên .NET Framework 4.x với TLS 1.2 và timeout rõ ràng.</summary>
    public sealed class HttpWebTransport : IAiTransport
    {
        public AiHttpResponse Send(string method, string url, Dictionary<string, string> headers, string jsonBody, int timeoutMs)
        {
            var res = new AiHttpResponse();
            try
            {
                // Bắt buộc bật TLS 1.2 trên .NET 4.x để kết nối Anthropic/OpenAI/Gemini
                try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; }
                catch { }

                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = method;
                req.Timeout = timeoutMs;
                req.ReadWriteTimeout = timeoutMs;
                req.Accept = "application/json";
                if (headers != null)
                {
                    foreach (var kv in headers) req.Headers[kv.Key] = kv.Value;
                }
                if (jsonBody != null && method != "GET")
                {
                    req.ContentType = "application/json; charset=utf-8";
                    byte[] bytes = Encoding.UTF8.GetBytes(jsonBody);
                    req.ContentLength = bytes.Length;
                    using (Stream s = req.GetRequestStream()) s.Write(bytes, 0, bytes.Length);
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    res.StatusCode = (int)resp.StatusCode;
                    res.Body = sr.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    try
                    {
                        using (var hr = (HttpWebResponse)wex.Response)
                        using (var sr = new StreamReader(hr.GetResponseStream(), Encoding.UTF8))
                        {
                            res.StatusCode = (int)hr.StatusCode;
                            res.Body = sr.ReadToEnd();
                        }
                    }
                    catch { }
                }
                if (wex.Status == WebExceptionStatus.Timeout)
                    res.Error = "Hết thời gian chờ kết nối AI (" + (timeoutMs / 1000) + " giây).";
                else if (res.StatusCode == 401 || res.StatusCode == 403)
                    res.Error = "Khóa API không hợp lệ hoặc không đủ quyền truy cập (HTTP " + res.StatusCode + ").";
                else if (res.StatusCode == 429)
                    res.Error = "Vượt giới hạn tốc độ hoặc hết hạn mức sử dụng (HTTP 429).";
                else
                    res.Error = "Lỗi kết nối máy chủ AI: " + wex.Message;
            }
            catch (Exception ex)
            {
                res.Error = "Không gọi được API AI: " + ex.Message;
            }
            return res;
        }
    }

    /// <summary>Bộ đếm token vào/ra trong bộ nhớ cho phiên Word hiện tại (tự reset khi khởi động lại Word).</summary>
    public static class AiSessionUsage
    {
        static readonly object Sync = new object();
        static long tokensIn;
        static long tokensOut;

        public static event Action Changed;

        public static long TokensIn { get { lock (Sync) return tokensIn; } }
        public static long TokensOut { get { lock (Sync) return tokensOut; } }
        public static long TotalTokens { get { lock (Sync) return tokensIn + tokensOut; } }

        public static void Add(int inputTokens, int outputTokens)
        {
            if (inputTokens < 0) inputTokens = 0;
            if (outputTokens < 0) outputTokens = 0;
            lock (Sync)
            {
                tokensIn += inputTokens;
                tokensOut += outputTokens;
            }
            Action h = Changed;
            if (h != null) { try { h(); } catch { } }
        }

        public static void Reset()
        {
            lock (Sync)
            {
                tokensIn = 0;
                tokensOut = 0;
            }
            Action h = Changed;
            if (h != null) { try { h(); } catch { } }
        }

        public static string FormatFooter()
        {
            long i, o;
            lock (Sync) { i = tokensIn; o = tokensOut; }
            return "Token AI trong phiên: vào " + i.ToString("N0", CultureInfo.InvariantCulture)
                 + " / ra " + o.ToString("N0", CultureInfo.InvariantCulture)
                 + " (tổng " + (i + o).ToString("N0", CultureInfo.InvariantCulture) + ")";
        }
    }

    public sealed class AiModelListResult
    {
        public bool Ok;
        public string Provider = "";
        public List<string> Models = new List<string>();
        public string Error = "";
    }

    public sealed class AiChatResult
    {
        public bool Ok;
        public string Provider = "";
        public string Model = "";
        public string Text = "";
        public int TokensIn;
        public int TokensOut;
        public string Error = "";
    }

    /// <summary>Trình phân tích/tạo JSON gọn nhẹ cho C# 5, không phụ thuộc thư viện ngoài.</summary>
    public static class MiniJson
    {
        public static object Parse(string json)
        {
            if (string.IsNullOrEmpty(json)) return null;
            int idx = 0;
            return ParseValue(json, ref idx);
        }

        internal static object ParseStrict(string json)
        {
            int index = 0;
            ValidateValue(json, ref index, 0);
            SkipWs(json, ref index);
            if (index != json.Length) throw new FormatException();
            return Parse(json);
        }

        static readonly Regex JsonPrimitive = new Regex(
            @"\G(?:true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)");

        static void ValidateValue(string json, ref int index, int depth)
        {
            if (depth > 12) throw new FormatException();
            SkipWs(json, ref index);
            if (index >= json.Length) throw new FormatException();
            char token = json[index];
            if (token == '"') { ValidateString(json, ref index); return; }
            if (token != '{' && token != '[')
            {
                Match primitive = JsonPrimitive.Match(json, index);
                if (!primitive.Success) throw new FormatException();
                index += primitive.Length;
                return;
            }
            index++;
            bool obj = token == '{';
            char close = obj ? '}' : ']';
            SkipWs(json, ref index);
            if (index < json.Length && json[index] == close) { index++; return; }
            var keys = new HashSet<string>(StringComparer.Ordinal);
            while (index < json.Length)
            {
                SkipWs(json, ref index);
                if (obj)
                {
                    int keyAt = index;
                    ValidateString(json, ref index);
                    if (!keys.Add(ParseString(json, ref keyAt))) throw new FormatException();
                    SkipWs(json, ref index);
                    if (index == json.Length || json[index++] != ':') throw new FormatException();
                }
                ValidateValue(json, ref index, depth + 1);
                SkipWs(json, ref index);
                if (index == json.Length) break;
                char separator = json[index++];
                if (separator == close) return;
                if (separator != ',') break;
            }
            throw new FormatException();
        }

        static void ValidateString(string json, ref int index)
        {
            if (index >= json.Length || json[index++] != '"') throw new FormatException();
            while (index < json.Length)
            {
                char ch = json[index++];
                if (ch == '"') return;
                if (ch < ' ') throw new FormatException();
                if (ch != '\\') continue;
                if (index == json.Length) break;
                ch = json[index++];
                if (ch == 'u')
                {
                    for (int i = 0; i < 4; i++)
                        if (index == json.Length || !Uri.IsHexDigit(json[index++])) throw new FormatException();
                }
                else if ("\"\\/bfnrt".IndexOf(ch) < 0) throw new FormatException();
            }
            throw new FormatException();
        }

        public static Dictionary<string, object> AsDict(object o) { return o as Dictionary<string, object>; }
        public static List<object> AsList(object o) { return o as List<object>; }

        public static string GetStr(Dictionary<string, object> d, string k)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return "";
            return v.ToString();
        }

        public static int GetInt(Dictionary<string, object> d, string k)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return 0;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        public static string Stringify(object val)
        {
            var sb = new StringBuilder();
            WriteVal(sb, val);
            return sb.ToString();
        }

        static void WriteVal(StringBuilder sb, object val)
        {
            if (val == null) { sb.Append("null"); return; }
            if (val is bool) { sb.Append(((bool)val) ? "true" : "false"); return; }
            if (val is int || val is long) { sb.Append(Convert.ToString(val, CultureInfo.InvariantCulture)); return; }
            if (val is float || val is double || val is decimal)
            {
                sb.Append(Convert.ToDouble(val, CultureInfo.InvariantCulture).ToString("0.######", CultureInfo.InvariantCulture));
                return;
            }
            if (val is string) { WriteStr(sb, (string)val); return; }
            var dict = val as IDictionary;
            if (dict != null)
            {
                sb.Append('{');
                bool first = true;
                foreach (DictionaryEntry de in dict)
                {
                    if (!first) sb.Append(',');
                    first = false;
                    WriteStr(sb, de.Key == null ? "" : de.Key.ToString());
                    sb.Append(':');
                    WriteVal(sb, de.Value);
                }
                sb.Append('}');
                return;
            }
            var list = val as IEnumerable;
            if (list != null)
            {
                sb.Append('[');
                bool first = true;
                foreach (object item in list)
                {
                    if (!first) sb.Append(',');
                    first = false;
                    WriteVal(sb, item);
                }
                sb.Append(']');
                return;
            }
            WriteStr(sb, val.ToString());
        }

        static void WriteStr(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                if (c == '"') sb.Append("\\\"");
                else if (c == '\\') sb.Append("\\\\");
                else if (c == '\b') sb.Append("\\b");
                else if (c == '\f') sb.Append("\\f");
                else if (c == '\n') sb.Append("\\n");
                else if (c == '\r') sb.Append("\\r");
                else if (c == '\t') sb.Append("\\t");
                else if (c < ' ') sb.Append("\\u" + ((int)c).ToString("x4"));
                else sb.Append(c);
            }
            sb.Append('"');
        }

        static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
        }

        static object ParseValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) return null;
            char c = s[i];
            if (c == '{') return ParseObject(s, ref i);
            if (c == '[') return ParseArray(s, ref i);
            if (c == '"') return ParseString(s, ref i);
            if (c == 't' && i + 4 <= s.Length && s.Substring(i, 4) == "true") { i += 4; return true; }
            if (c == 'f' && i + 5 <= s.Length && s.Substring(i, 5) == "false") { i += 5; return false; }
            if (c == 'n' && i + 4 <= s.Length && s.Substring(i, 4) == "null") { i += 4; return null; }
            return ParseNumber(s, ref i);
        }

        static Dictionary<string, object> ParseObject(string s, ref int i)
        {
            var d = new Dictionary<string, object>(StringComparer.Ordinal);
            i++; // '{'
            while (i < s.Length)
            {
                SkipWs(s, ref i);
                if (i >= s.Length) break;
                if (s[i] == '}') { i++; return d; }
                string key = ParseString(s, ref i);
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ':') i++;
                d[key] = ParseValue(s, ref i);
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') i++;
            }
            return d;
        }

        static List<object> ParseArray(string s, ref int i)
        {
            var l = new List<object>();
            i++; // '['
            while (i < s.Length)
            {
                SkipWs(s, ref i);
                if (i >= s.Length) break;
                if (s[i] == ']') { i++; return l; }
                l.Add(ParseValue(s, ref i));
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') i++;
            }
            return l;
        }

        static string ParseString(string s, ref int i)
        {
            if (i >= s.Length || s[i] != '"') return "";
            i++;
            var sb = new StringBuilder();
            while (i < s.Length)
            {
                char c = s[i++];
                if (c == '"') return sb.ToString();
                if (c == '\\' && i < s.Length)
                {
                    char e = s[i++];
                    if (e == '"' || e == '\\' || e == '/') sb.Append(e);
                    else if (e == 'b') sb.Append('\b');
                    else if (e == 'f') sb.Append('\f');
                    else if (e == 'n') sb.Append('\n');
                    else if (e == 'r') sb.Append('\r');
                    else if (e == 't') sb.Append('\t');
                    else if (e == 'u' && i + 4 <= s.Length)
                    {
                        int hex;
                        if (int.TryParse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out hex))
                            sb.Append((char)hex);
                        i += 4;
                    }
                }
                else sb.Append(c);
            }
            return sb.ToString();
        }

        static object ParseNumber(string s, ref int i)
        {
            int start = i;
            if (i < s.Length && s[i] == '-') i++;
            while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '.' || s[i] == 'e' || s[i] == 'E' || s[i] == '+' || s[i] == '-'))
                i++;
            string raw = s.Substring(start, i - start);
            if (raw.IndexOf('.') >= 0 || raw.IndexOf('e') >= 0 || raw.IndexOf('E') >= 0)
            {
                double dv;
                if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out dv)) return dv;
            }
            else
            {
                long lv;
                if (long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out lv)) return lv;
            }
            return 0L;
        }
    }

    /// <summary>Client gọi AI trực tuyến (Anthropic Claude, Google Gemini, OpenAI) với nhận dạng khóa tự động.</summary>
    public static class AiClient
    {
        public const int DefaultTimeoutMs = 30000;
        public const int MaxRewriteWords = 500;
        public static IAiTransport Transport = new HttpWebTransport();

        /// <summary>Tự nhận dạng nhà cung cấp dựa vào tiền tố của khóa API.</summary>
        public static string DetectProvider(string apiKey)
        {
            string k = (apiKey ?? "").Trim();
            if (k.Length < 12) return "";
            if (k.StartsWith("sk-ant-", StringComparison.Ordinal)) return "anthropic";
            if (k.StartsWith("AIza", StringComparison.Ordinal)) return "gemini";
            if (k.StartsWith("sk-", StringComparison.Ordinal)) return "openai";
            return "";
        }

        public static string ProviderDisplayName(string provider)
        {
            if (provider == "anthropic") return "Anthropic Claude";
            if (provider == "gemini") return "Google Gemini";
            if (provider == "openai") return "OpenAI (hoặc tương thích)";
            return "Chưa nhận dạng";
        }

        /// <summary>Ẩn khóa API nếu có lỡ xuất hiện trong thông báo lỗi hay URL.</summary>
        public static string Redact(string text, string apiKey)
        {
            string s = text ?? "";
            if (!string.IsNullOrEmpty(apiKey) && apiKey.Length >= 6)
                s = s.Replace(apiKey, Secret.Mask(apiKey));
            return s;
        }

        /// <summary>Kiểm tra tham số temperature [0.0..2.0] và số token đầu ra tối đa [64..8192].</summary>
        public static bool ValidateParams(double temp, int maxTokens, out string error)
        {
            error = "";
            if (double.IsNaN(temp) || double.IsInfinity(temp) || temp < 0.0 || temp > 2.0)
            {
                error = "Temperature phải nằm trong khoảng 0.0 đến 2.0.";
                return false;
            }
            if (maxTokens < 64 || maxTokens > 8192)
            {
                error = "Số token đầu ra tối đa phải nằm trong khoảng 64 đến 8192.";
                return false;
            }
            return true;
        }

        /// <summary>Đếm số từ theo quy chuẩn phân tách khoảng trắng cho giới hạn 500 từ khi viết lại.</summary>
        public static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            int n = 0;
            foreach (string w in Regex.Split(text.Trim(), @"[\s\u00A0\r\n\t]+"))
                if (w.Length > 0 && w.Any(char.IsLetterOrDigit)) n++;
            return n;
        }

        /// <summary>Nhận dạng nhà cung cấp và truy vấn danh sách model khả dụng.</summary>
        public static AiModelListResult ListModels(string apiKey)
        {
            var res = new AiModelListResult();
            string k = (apiKey ?? "").Trim();
            string prov = DetectProvider(k);
            res.Provider = prov;
            if (string.IsNullOrEmpty(prov))
            {
                res.Error = "Không nhận dạng được nhà cung cấp từ định dạng khóa API (hỗ trợ sk-ant-..., AIza..., sk-...).";
                return res;
            }

            AiHttpResponse http;
            if (prov == "anthropic")
            {
                var h = new Dictionary<string, string> { { "x-api-key", k }, { "anthropic-version", "2023-06-01" } };
                http = Transport.Send("GET", "https://api.anthropic.com/v1/models", h, null, DefaultTimeoutMs);
            }
            else if (prov == "gemini")
            {
                var h = new Dictionary<string, string> { { "x-goog-api-key", k } };
                http = Transport.Send("GET", "https://generativelanguage.googleapis.com/v1beta/models", h, null, DefaultTimeoutMs);
            }
            else
            {
                var h = new Dictionary<string, string> { { "Authorization", "Bearer " + k } };
                http = Transport.Send("GET", "https://api.openai.com/v1/models", h, null, DefaultTimeoutMs);
            }

            if (!string.IsNullOrEmpty(http.Error))
            {
                res.Error = Redact(ExtractErrorMessage(http.Body) ?? http.Error, k);
                return res;
            }
            if (http.StatusCode < 200 || http.StatusCode >= 300)
            {
                res.Error = Redact(ExtractErrorMessage(http.Body) ?? ("Lỗi HTTP " + http.StatusCode), k);
                return res;
            }

            var root = MiniJson.AsDict(MiniJson.Parse(http.Body));
            if (root == null)
            {
                res.Error = "Phản hồi danh sách model từ máy chủ AI không đúng định dạng JSON.";
                return res;
            }

            var models = new List<string>();
            if (prov == "gemini")
            {
                var arr = MiniJson.AsList(root.ContainsKey("models") ? root["models"] : null);
                if (arr != null)
                {
                    foreach (object item in arr)
                    {
                        var md = MiniJson.AsDict(item);
                        if (md == null) continue;
                        var methods = MiniJson.AsList(md.ContainsKey("supportedGenerationMethods") ? md["supportedGenerationMethods"] : null);
                        if (methods != null && !methods.Any(m => string.Equals(m == null ? "" : m.ToString(), "generateContent", StringComparison.OrdinalIgnoreCase)))
                            continue;
                        string name = MiniJson.GetStr(md, "name");
                        if (name.StartsWith("models/", StringComparison.Ordinal)) name = name.Substring(7);
                        if (name.Length > 0 && !models.Contains(name)) models.Add(name);
                    }
                }
            }
            else
            {
                var arr = MiniJson.AsList(root.ContainsKey("data") ? root["data"] : null);
                if (arr != null)
                {
                    foreach (object item in arr)
                    {
                        var md = MiniJson.AsDict(item);
                        string id = MiniJson.GetStr(md, "id");
                        if (id.Length > 0 && !models.Contains(id)) models.Add(id);
                    }
                }
            }

            if (prov == "openai")
            {
                // Đưa các model chat phổ biến lên đầu danh sách cho dễ chọn
                models = models
                    .OrderByDescending(id => id.StartsWith("gpt-", StringComparison.OrdinalIgnoreCase) || id.StartsWith("o", StringComparison.OrdinalIgnoreCase))
                    .ThenBy(id => id, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            if (models.Count == 0)
            {
                res.Error = "Kết nối thành công nhưng tài khoản chưa trả về model nào khả dụng.";
                return res;
            }
            res.Ok = true;
            res.Models = models;
            return res;
        }

        /// <summary>Gửi một lời gọi sinh văn bản tới model đã chọn và ghi nhận token vào bộ đếm phiên.</summary>
        public static AiChatResult Complete(string apiKey, string model, string systemPrompt, string userPrompt, double temp, int maxTokens)
        {
            var res = new AiChatResult { Model = model ?? "" };
            string k = (apiKey ?? "").Trim();
            string prov = DetectProvider(k);
            res.Provider = prov;
            if (string.IsNullOrEmpty(prov))
            {
                res.Error = "Khóa API chưa được cấu hình hoặc không hợp lệ.";
                return res;
            }
            if (string.IsNullOrWhiteSpace(model))
            {
                res.Error = "Chưa chọn model AI. Hãy vào Cài đặt để kiểm tra và chọn model.";
                return res;
            }
            string pErr;
            if (!ValidateParams(temp, maxTokens, out pErr))
            {
                res.Error = pErr;
                return res;
            }

            AiHttpResponse http;
            if (prov == "anthropic")
            {
                var h = new Dictionary<string, string> { { "x-api-key", k }, { "anthropic-version", "2023-06-01" } };
                var bodyObj = new Dictionary<string, object>
                {
                    { "model", model },
                    { "max_tokens", maxTokens },
                    { "temperature", temp },
                    { "system", systemPrompt ?? "" },
                    { "messages", new object[] { new Dictionary<string, object> { { "role", "user" }, { "content", userPrompt ?? "" } } } }
                };
                http = Transport.Send("POST", "https://api.anthropic.com/v1/messages", h, MiniJson.Stringify(bodyObj), DefaultTimeoutMs);
            }
            else if (prov == "gemini")
            {
                var h = new Dictionary<string, string> { { "x-goog-api-key", k } };
                string encodedModel = Uri.EscapeDataString(model.Trim());
                var bodyObj = new Dictionary<string, object>
                {
                    { "systemInstruction", new Dictionary<string, object> { { "parts", new object[] { new Dictionary<string, object> { { "text", systemPrompt ?? "" } } } } } },
                    { "contents", new object[] { new Dictionary<string, object> { { "role", "user" }, { "parts", new object[] { new Dictionary<string, object> { { "text", userPrompt ?? "" } } } } } } },
                    { "generationConfig", new Dictionary<string, object> { { "temperature", temp }, { "maxOutputTokens", maxTokens } } }
                };
                http = Transport.Send("POST", "https://generativelanguage.googleapis.com/v1beta/models/" + encodedModel + ":generateContent", h, MiniJson.Stringify(bodyObj), DefaultTimeoutMs);
            }
            else
            {
                var h = new Dictionary<string, string> { { "Authorization", "Bearer " + k } };
                var bodyObj = new Dictionary<string, object>
                {
                    { "model", model },
                    { "temperature", temp },
                    { "max_tokens", maxTokens },
                    { "messages", new object[]
                        {
                            new Dictionary<string, object> { { "role", "system" }, { "content", systemPrompt ?? "" } },
                            new Dictionary<string, object> { { "role", "user" }, { "content", userPrompt ?? "" } }
                        }
                    }
                };
                http = Transport.Send("POST", "https://api.openai.com/v1/chat/completions", h, MiniJson.Stringify(bodyObj), DefaultTimeoutMs);
            }

            if (!string.IsNullOrEmpty(http.Error))
            {
                res.Error = Redact(ExtractErrorMessage(http.Body) ?? http.Error, k);
                return res;
            }
            if (http.StatusCode < 200 || http.StatusCode >= 300)
            {
                res.Error = Redact(ExtractErrorMessage(http.Body) ?? ("Lỗi HTTP " + http.StatusCode), k);
                return res;
            }

            Dictionary<string, object> root = null;
            try { root = MiniJson.AsDict(MiniJson.ParseStrict(http.Body)); }
            catch (FormatException) { }
            if (root == null)
            {
                res.Error = "Máy chủ AI trả về dữ liệu không phải JSON hợp lệ.";
                return res;
            }

            string stopReason = "";
            bool incomplete;
            if (prov == "anthropic")
            {
                stopReason = MiniJson.GetStr(root, "stop_reason");
                incomplete = stopReason != "" && stopReason != "end_turn" && stopReason != "stop_sequence";
                var content = MiniJson.AsList(root.ContainsKey("content") ? root["content"] : null);
                if (content != null && content.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (object c in content)
                    {
                        var cd = MiniJson.AsDict(c);
                        if (cd != null && MiniJson.GetStr(cd, "type") == "text") sb.Append(MiniJson.GetStr(cd, "text"));
                    }
                    res.Text = sb.ToString().Trim();
                }
                var usage = MiniJson.AsDict(root.ContainsKey("usage") ? root["usage"] : null);
                res.TokensIn = MiniJson.GetInt(usage, "input_tokens");
                res.TokensOut = MiniJson.GetInt(usage, "output_tokens");
            }
            else if (prov == "gemini")
            {
                var cands = MiniJson.AsList(root.ContainsKey("candidates") ? root["candidates"] : null);
                if (cands != null && cands.Count > 0)
                {
                    var c0 = MiniJson.AsDict(cands[0]);
                    stopReason = MiniJson.GetStr(c0, "finishReason");
                    var content = MiniJson.AsDict(c0 != null && c0.ContainsKey("content") ? c0["content"] : null);
                    var parts = MiniJson.AsList(content != null && content.ContainsKey("parts") ? content["parts"] : null);
                    if (parts != null)
                    {
                        var sb = new StringBuilder();
                        foreach (object p in parts)
                        {
                            var pd = MiniJson.AsDict(p);
                            if (pd != null && !(pd.ContainsKey("thought") && object.Equals(pd["thought"], true)))
                                sb.Append(MiniJson.GetStr(pd, "text"));
                        }
                        res.Text = sb.ToString().Trim();
                    }
                }
                incomplete = stopReason != "" && stopReason != "STOP";
                var usage = MiniJson.AsDict(root.ContainsKey("usageMetadata") ? root["usageMetadata"] : null);
                res.TokensIn = MiniJson.GetInt(usage, "promptTokenCount");
                res.TokensOut = MiniJson.GetInt(usage, "candidatesTokenCount");
            }
            else
            {
                var choices = MiniJson.AsList(root.ContainsKey("choices") ? root["choices"] : null);
                if (choices != null && choices.Count > 0)
                {
                    var ch0 = MiniJson.AsDict(choices[0]);
                    stopReason = MiniJson.GetStr(ch0, "finish_reason");
                    var msg = MiniJson.AsDict(ch0 != null && ch0.ContainsKey("message") ? ch0["message"] : null);
                    res.Text = MiniJson.GetStr(msg, "content").Trim();
                }
                var usage = MiniJson.AsDict(root.ContainsKey("usage") ? root["usage"] : null);
                incomplete = stopReason != "" && stopReason != "stop";
                res.TokensIn = MiniJson.GetInt(usage, "prompt_tokens");
                res.TokensOut = MiniJson.GetInt(usage, "completion_tokens");
            }

            AiSessionUsage.Add(res.TokensIn, res.TokensOut);
            if (incomplete)
            {
                res.Text = "";
                res.Error = "Model dừng trước khi trả đủ kết quả. Hãy thử tăng số token tối đa hoặc chọn model khác.";
                return res;
            }
            if (string.IsNullOrEmpty(res.Text))
            {
                res.Error = "Model không trả về nội dung văn bản.";
                return res;
            }
            res.Ok = true;
            return res;
        }

        /// <summary>Kiểm tra khóa và model bằng một lời gọi tối giản trước khi lưu cài đặt.</summary>
        public static AiChatResult TestConnection(string apiKey, string model, double temp, int maxTokens)
        {
            return Complete(apiKey, model,
                "You are a connection test assistant.",
                "Respond ONLY with OK.",
                temp, Math.Min(maxTokens, 64));
        }

        /// <summary>Nhận ngôn ngữ của văn bản: "vi", "en" hoặc "other". Nhận dạng sơ bộ bằng dấu tiếng Việt và chữ Latin.</summary>
        internal static string DetectLang(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "other";
            string t = text.Trim();
            int letters = 0;
            int viMark = 0;
            int latin = 0;
            foreach (char c in t)
            {
                if (char.IsLetter(c))
                {
                    letters++;
                    char lower = char.ToLowerInvariant(c);
                    if (lower > 127 && "àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ".IndexOf(lower) >= 0)
                        viMark++;
                    else if (lower >= 'a' && lower <= 'z')
                        latin++;
                }
            }
            if (letters < 4) return "other";


            if (viMark > 0) return "vi";
            if (latin > 0) return "en";
            return "other";
        }

        /// <summary>Kiểm tra sai khác ngôn ngữ giữa đầu vào và đầu ra (Anh ra Việt hoặc Việt ra Anh).</summary>
        internal static string TextKey(string value)
        {
            return Regex.Replace((value ?? "").Trim().Normalize(NormalizationForm.FormC), @"\s+", " ").ToLowerInvariant();
        }

        internal static bool LanguageMismatch(string input, string output)
        {
            if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output)) return false;
            string inLang = DetectLang(input);
            if (inLang == "other") return false;
            string outLang = DetectLang(output);
            if (outLang == "other") return false;
            return (inLang == "en" && outLang == "vi") || (inLang == "vi" && outLang == "en");
        }

        internal static string MapTone(string tone)
        {
            string t = (tone ?? "").Trim();
            string desc;
            if (t.IndexOf("Trang trọng", StringComparison.OrdinalIgnoreCase) >= 0)
                desc = "Formal and authoritative";
            else if (t.IndexOf("Ngắn gọn", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("súc tích", StringComparison.OrdinalIgnoreCase) >= 0)
                desc = "Concise and direct";
            else if (t.IndexOf("Tự nhiên", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("mạch lạc", StringComparison.OrdinalIgnoreCase) >= 0)
                desc = "Natural and coherent";
            else if (t.IndexOf("Học thuật", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("khoa học", StringComparison.OrdinalIgnoreCase) >= 0)
                desc = "Academic and scholarly";
            else
                desc = "Professional technical";
            return desc + " (apply the tone in the input's own language)";
        }

        internal static string GetLanguageMandate(string lang)
        {
            string inputClause = "";
            if (lang == "en") inputClause = " The input is English.";
            else if (lang == "vi") inputClause = " The input is Vietnamese.";
            return "Write the output in the SAME language as the input text." + inputClause + " NEVER translate. Do not switch language. Keep every number, unit, formula, symbol, citation like [1] and technical term unchanged. Return only the rewritten text, with no preface, quotes, or explanation.";
        }

        internal static Tuple<string, string> BuildRewritePrompt(string tone, string selectedText, string docContext)
        {
            string lang = DetectLang(selectedText);
            string mandate = GetLanguageMandate(lang);
            string mappedTone = MapTone(tone);

            var sbSys = new StringBuilder();
            sbSys.AppendLine("You are an expert technical editor.");
            sbSys.AppendLine(mandate);
            sbSys.AppendLine("Rewrite the user's text following this tone: " + mappedTone + ".");
            sbSys.AppendLine("Mandatory rules:");
            sbSys.AppendLine("1. Preserve the exact number of paragraphs and line breaks (if input has multiple paragraphs, output must have the same number).");
            sbSys.AppendLine("2. Maintain consistency with the overall document context.");
            sbSys.AppendLine("3. The document context is provided strictly as reference material to help you understand the topic. NEVER copy, quote, or include any part of the document context in your output.");
            sbSys.Append("4. Return only the rewritten text without commentary, quotes, or markdown wrappers.");

            int wc = CountWords(selectedText);
            var sbUsr = new StringBuilder();
            sbUsr.AppendLine("--- Text to rewrite (" + wc + " words) ---");
            sbUsr.Append(selectedText ?? "");

            return Tuple.Create(sbSys.ToString(), sbUsr.ToString());
        }

        internal static Tuple<string, string> BuildSplitPrompt(string sentence, string context)
        {
            string lang = DetectLang(sentence);
            string mandate = GetLanguageMandate(lang);

            var sbSys = new StringBuilder();
            sbSys.AppendLine("You are an expert technical document editor.");
            sbSys.AppendLine(mandate);
            sbSys.AppendLine("Rewrite or split the long sentence into 1-2 concise, clear sentences.");
            sbSys.AppendLine("Keep all numbers, units, symbols, and technical terms unchanged.");
            sbSys.AppendLine("The surrounding context is for reference only; do not copy context into output.");
            sbSys.Append("Return only the revised sentence(s) with no explanation or preface.");

            var sbUsr = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(context))
            {
                sbUsr.AppendLine("Surrounding context (reference only):\n" + context);
                sbUsr.AppendLine();
            }
            sbUsr.Append("Sentence to revise:\n" + (sentence ?? ""));

            return Tuple.Create(sbSys.ToString(), sbUsr.ToString());
        }

        internal static Tuple<string, string> BuildSynonymPrompt(string phrase, string context)
        {
            string lang = DetectLang(phrase);
            if (lang == "other" && !string.IsNullOrWhiteSpace(context))
                lang = DetectLang(context);
            string mandate = GetLanguageMandate(lang);

            var sbSys = new StringBuilder();
            sbSys.AppendLine("You are a linguistic expert and technical editor.");
            sbSys.AppendLine(mandate);
            sbSys.AppendLine("Every suggested synonym MUST be in the SAME language as the target phrase/word.");
            sbSys.AppendLine("Propose up to 6 synonyms that can naturally replace the target phrase in the given sentence context.");
            sbSys.Append("Write each suggestion on a separate line, with no bullet points, numbering, quotes, or additional explanation.");

            var sbUsr = new StringBuilder();
            sbUsr.AppendLine("Target phrase/word to replace: \"" + (phrase ?? "") + "\"");
            if (!string.IsNullOrWhiteSpace(context))
                sbUsr.Append("Sentence context: " + context);

            return Tuple.Create(sbSys.ToString(), sbUsr.ToString());
        }

        /// <summary>Yêu cầu AI tách/viết gọn lại câu dài nhưng giữ nguyên thuật ngữ, số liệu và đơn vị.</summary>
        public static AiChatResult SplitLongSentence(AppSettings st, string sentence, string context)
        {
            string key = Secret.Unprotect(st != null ? st.ApiKeyEnc : "");
            string model = st != null ? st.ApiModel : "";
            double temp = st != null ? st.AiTemp : 0.3;
            int maxTok = st != null ? st.AiMaxTokens : 1024;

            Tuple<string, string> prompt = BuildSplitPrompt(sentence, context);
            AiChatResult r = Complete(key, model, prompt.Item1, prompt.Item2, temp, maxTok);
            if (!r.Ok) return r;

            string inLang = DetectLang(sentence);
            if (inLang != "other" && LanguageMismatch(sentence, r.Text))
            {
                string targetLang = inLang == "en" ? "English" : "Vietnamese";
                string retrySys = prompt.Item1 + "\n\nCRITICAL WARNING: Your previous answer was in the wrong language. Answer ONLY in " + targetLang + ". NEVER translate.";
                AiChatResult r2 = Complete(key, model, retrySys, prompt.Item2, temp, maxTok);
                if (!r2.Ok) return r2;
                if (LanguageMismatch(sentence, r2.Text))
                {
                    return new AiChatResult
                    {
                        Ok = false,
                        Provider = r2.Provider,
                        Model = r2.Model,
                        Error = "AI trả lời sai ngôn ngữ (đã thử lại một lần). Không áp dụng.",
                        TokensIn = r.TokensIn + r2.TokensIn,
                        TokensOut = r.TokensOut + r2.TokensOut
                    };
                }
                r = r2;
            }
            return r;
        }

        /// <summary>Yêu cầu AI gợi ý 5-8 từ/cụm từ đồng nghĩa phù hợp với ngữ cảnh trong tài liệu.</summary>
        public static AiChatResult SuggestSynonymsAi(AppSettings st, string phrase, string sentenceContext, out List<string> suggestions)
        {
            suggestions = new List<string>();
            string key = Secret.Unprotect(st != null ? st.ApiKeyEnc : "");
            string model = st != null ? st.ApiModel : "";
            double temp = st != null ? st.AiTemp : 0.3;
            int maxTok = st != null ? st.AiMaxTokens : 512;

            Tuple<string, string> prompt = BuildSynonymPrompt(phrase, sentenceContext);
            AiChatResult r = Complete(key, model, prompt.Item1, prompt.Item2, temp, maxTok);
            if (!r.Ok) return r;
            string origKey = TextKey(phrase);
            string phraseLang = DetectLang(phrase);
            foreach (string raw in r.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string line = Regex.Replace(raw.Trim(), @"^(?:[-*•]+|\d+[.)])\s*", "").Trim().Trim('"', '“', '”', '`');
                if (line.Length == 0 || TextKey(line) == origKey) continue;
                if (phraseLang != "other")
                {
                    string lineLang = DetectLang(line);
                    if (lineLang != "other" && lineLang != phraseLang) continue;
                }
                if (!suggestions.Contains(line)) suggestions.Add(line);
            }
            return r;
        }

        /// <summary>Viết lại đoạn văn đang bôi đen theo tone và ngữ cảnh toàn tài liệu (tối đa 500 từ).</summary>
        public static AiChatResult RewriteText(AppSettings st, string selectedText, string tone, string docContext)
        {
            int wc = CountWords(selectedText);
            if (wc == 0)
                return new AiChatResult { Ok = false, Error = "Chưa có đoạn văn nào được bôi đen." };
            if (wc > MaxRewriteWords)
                return new AiChatResult { Ok = false, Error = "Đoạn đang chọn có " + wc + " từ, vượt quá giới hạn tối đa " + MaxRewriteWords + " từ." };

            string key = Secret.Unprotect(st != null ? st.ApiKeyEnc : "");
            string model = st != null ? st.ApiModel : "";
            double temp = st != null ? st.AiTemp : 0.3;
            int maxTok = st != null ? st.AiMaxTokens : 1024;

            Tuple<string, string> prompt = BuildRewritePrompt(tone, selectedText, docContext);
            AiChatResult r = Complete(key, model, prompt.Item1, prompt.Item2, temp, maxTok);
            if (!r.Ok) return r;

            string inLang = DetectLang(selectedText);
            if (inLang != "other" && LanguageMismatch(selectedText, r.Text))
            {
                string targetLang = inLang == "en" ? "English" : "Vietnamese";
                string retrySys = prompt.Item1 + "\n\nCRITICAL WARNING: Your previous answer was in the wrong language. Answer ONLY in " + targetLang + ". NEVER translate.";
                AiChatResult r2 = Complete(key, model, retrySys, prompt.Item2, temp, maxTok);
                if (!r2.Ok) return r2;
                if (LanguageMismatch(selectedText, r2.Text))
                {
                    return new AiChatResult
                    {
                        Ok = false,
                        Provider = r2.Provider,
                        Model = r2.Model,
                        Error = "AI trả lời sai ngôn ngữ (đã thử lại một lần). Không áp dụng.",
                        TokensIn = r.TokensIn + r2.TokensIn,
                        TokensOut = r.TokensOut + r2.TokensOut
                    };
                }
                r = r2;
            }
            return r;
        }

        // ============================================================ Rev0.9.5.6: Công cụ AI, Viết lại và Dịch
        // Mã, tên tiếng Việt, tên bản ngữ, tên tiếng Anh (dùng trong lời nhắc). Danh sách 10 ngôn ngữ do Muse đề xuất.
        public static readonly string[][] TargetLangs = new string[][]
        {
            new[] { "vi", "Tiếng Việt", "Tiếng Việt", "Vietnamese" },
            new[] { "en", "Tiếng Anh", "English", "English" },
            new[] { "zh-CN", "Tiếng Trung giản thể", "中文 (简体)", "Simplified Chinese" },
            new[] { "ja", "Tiếng Nhật", "日本語", "Japanese" },
            new[] { "ko", "Tiếng Hàn", "한국어", "Korean" },
            new[] { "ru", "Tiếng Nga", "Русский", "Russian" },
            new[] { "fr", "Tiếng Pháp", "Français", "French" },
            new[] { "de", "Tiếng Đức", "Deutsch", "German" },
            new[] { "es", "Tiếng Tây Ban Nha", "Español", "Spanish" },
            new[] { "id", "Tiếng Indonesia", "Bahasa Indonesia", "Indonesian" }
        };

        internal static string LangEnglishName(string code)
        {
            foreach (string[] l in TargetLangs) if (string.Equals(l[0], code, StringComparison.OrdinalIgnoreCase)) return l[3];
            return "English";
        }

        internal static string SourceLangName(string text)
        {
            string d = DetectLang(text);
            return d == "vi" ? "Vietnamese" : (d == "en" ? "English" : "the same language as the input");
        }

        const string PreserveRules = "HARD PRESERVATION RULES: reproduce every number, decimal, unit, formula, mathematical or chemical symbol, standard code, proper name and citation such as [1] or [12] character for character. Keep exactly the same number of paragraphs and the same line breaks as the input. ";
        const string OutputRules = "OUTPUT RULES: output only the result, with no preamble, no explanation, no quotation marks around it and no markdown formatting.";

        /// <summary>Dựng lời nhắc cho ba chế độ: chỉ viết lại, chỉ dịch, vừa dịch vừa viết lại. Trả (system, user).</summary>
        internal static Tuple<string, string> BuildTransformPrompt(string tone, string text, string docContext, bool rewrite, bool translate, string targetCode)
        {
            string src = SourceLangName(text);
            string tgt = LangEnglishName(targetCode);
            string toneEn = MapTone(tone);
            string sys, mode;
            string translationRules = TranslationNumbers.Supports(targetCode)
                ? PreserveRules + "Keep numeric literals and the unit symbols exactly as written, including all trailing zeros, signs and exponents. The add-in will localize the decimal marker and digit grouping. Keep technical values and citations in their original sequence. "
                : PreserveRules;
            if (translate && rewrite)
            {
                sys = "You are a technical translator and editor inside a Word add-in. First translate the user's text from " + src + " into " + tgt
                    + ", then rewrite that translation inside " + tgt + " so it reads naturally and matches the requested tone. The final output must be entirely in " + tgt
                    + ", with nothing left in " + src + " except items covered by the preservation rules. Keep the meaning, facts and level of detail unchanged, do not add or drop information. "
                    + translationRules + "Translate technical terms with their standard equivalent in " + tgt + ", but never alter codes, symbols or formulas. " + OutputRules;
                mode = "Mode: TRANSLATE AND REWRITE.\nSource language detected by the add-in: " + src + "\nTarget language: " + tgt + "\nTarget tone, applied in " + tgt + ": " + toneEn;
            }
            else if (translate)
            {
                sys = "You are a professional technical translator inside a Word add-in. Translate the user's text from " + src + " into " + tgt
                    + ". Translate faithfully and completely. Do NOT rewrite, paraphrase for style, summarise, expand or improve the text, and do NOT change its tone, register or sentence intent. "
                    + translationRules + "Translate technical terms with their standard equivalent in " + tgt + ", but never alter codes, symbols or formulas. "
                    + "Output in " + tgt + " only. " + OutputRules;
                mode = "Mode: TRANSLATE ONLY, do not rewrite and do not change the tone.\nSource language detected by the add-in: " + src + "\nTarget language: " + tgt;
            }
            else
            {
                sys = "You are a technical rewriting engine inside a Word add-in. Rewrite the user's text in " + src
                    + " ONLY. Never translate any part of it into another language; foreign words, names and quoted terms stay exactly as they are. "
                    + "Rewrite for clarity and for the requested tone while keeping the meaning, facts and level of detail unchanged. "
                    + PreserveRules + "Keep technical terms unchanged. " + OutputRules;
                mode = "Mode: REWRITE ONLY, do not translate.\nLanguage of the text and of your answer: " + src + "\nTarget tone: " + toneEn;
            }
            var usr = new StringBuilder();
            usr.AppendLine(mode);
            usr.AppendLine((translate ? "Translate" : "Rewrite") + " the text between the markers:");
            usr.AppendLine("---TEXT START---");
            usr.AppendLine(text ?? "");
            usr.Append("---TEXT END---");
            return Tuple.Create(sys, usr.ToString());
        }

        /// <summary>Câu lệnh thử lại khi AI trả sai ngôn ngữ.</summary>
        internal static string RetryInstruction(bool translate, string targetCode, string sourceName)
        {
            string lang = translate ? LangEnglishName(targetCode) : sourceName;
            return "Retry: your previous output was in the wrong language, so regenerate the output entirely in " + lang + " only, following all the same rules, and output only the result.";
        }

        static int ScriptShare(string text, string code)
        {
            int letters = 0, hit = 0;
            foreach (char c in text ?? "")
            {
                if (!char.IsLetter(c)) continue;
                letters++;
                bool ok;
                switch (code)
                {
                    case "zh-CN": ok = (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0x3400 && c <= 0x4DBF); break;
                    case "ja": ok = (c >= 0x3040 && c <= 0x30FF) || (c >= 0x4E00 && c <= 0x9FFF); break;
                    case "ko": ok = c >= 0xAC00 && c <= 0xD7AF; break;
                    case "ru": ok = c >= 0x0400 && c <= 0x04FF; break;
                    default: ok = true; break;
                }
                if (ok) hit++;
            }
            return letters == 0 ? 100 : hit * 100 / letters;
        }

        /// <summary>Kết quả có sai ngôn ngữ so với mong đợi không. Chỉ kiểm tra được: Việt, Anh (nhận theo dấu và từ), Trung, Nhật, Hàn, Nga (nhận theo chữ viết).</summary>
        internal static bool WrongLanguage(string input, string output, bool translate, string targetCode)
        {
            if (string.IsNullOrWhiteSpace(output)) return false;
            if (!translate) return LanguageMismatch(input, output);
            if (targetCode == "vi" || targetCode == "en")
            {
                string o = DetectLang(output);
                return o != "other" && o != targetCode;
            }
            if (targetCode == "zh-CN" || targetCode == "ja" || targetCode == "ko" || targetCode == "ru")
                return ScriptShare(output, targetCode) < 40;
            return false;   // Pháp, Đức, Tây Ban Nha, Indonesia dùng chữ Latin, không kiểm tra được
        }

        /// <summary>Công cụ AI: viết lại và/hoặc dịch đoạn bôi đen (tối đa 500 từ). Ít nhất một trong hai phải bật.</summary>
        public static AiChatResult TransformText(AppSettings st, string selectedText, string tone, string docContext, bool rewrite, bool translate, string targetCode)
        {
            if (!rewrite && !translate)
                return new AiChatResult { Ok = false, Error = "Hãy chọn Viết lại hoặc Dịch (hoặc cả hai)." };
            int wc = CountWords(selectedText);
            if (wc == 0)
                return new AiChatResult { Ok = false, Error = "Chưa có đoạn văn nào được bôi đen." };
            if (wc > MaxRewriteWords)
                return new AiChatResult { Ok = false, Error = "Đoạn đang chọn có " + wc + " từ, vượt quá giới hạn tối đa " + MaxRewriteWords + " từ." };
            if (translate)
            {
                string inLang = DetectLang(selectedText);
                if ((targetCode == "vi" || targetCode == "en") && inLang == targetCode && !rewrite)
                    return new AiChatResult { Ok = false, Error = "Đoạn đang chọn đã là " + (targetCode == "vi" ? "tiếng Việt" : "tiếng Anh") + ", hãy chọn ngôn ngữ đích khác hoặc chọn Viết lại." };
            }
            string key = Secret.Unprotect(st != null ? st.ApiKeyEnc : "");
            string model = st != null ? st.ApiModel : "";
            double temp = st != null ? st.AiTemp : 0.3;
            int maxTok = st != null ? st.AiMaxTokens : 1024;
            Tuple<string, string> prompt = BuildTransformPrompt(tone, selectedText, docContext, rewrite, translate, targetCode);
            AiChatResult r = Complete(key, model, prompt.Item1, prompt.Item2, temp, maxTok);
            if (!r.Ok) return r;
            if (WrongLanguage(selectedText, r.Text, translate, targetCode))
            {
                string retrySys = prompt.Item1 + "\n\n" + RetryInstruction(translate, targetCode, SourceLangName(selectedText));
                AiChatResult r2 = Complete(key, model, retrySys, prompt.Item2, temp, maxTok);
                if (!r2.Ok) return r2;
                if (WrongLanguage(selectedText, r2.Text, translate, targetCode))
                    return new AiChatResult { Ok = false, Provider = r2.Provider, Model = r2.Model, TokensIn = r.TokensIn + r2.TokensIn, TokensOut = r.TokensOut + r2.TokensOut,
                        Error = "AI trả lời sai ngôn ngữ (đã thử lại một lần). Không áp dụng." };
                r2.TokensIn += r.TokensIn; r2.TokensOut += r.TokensOut;
                r = r2;
            }
            if (translate && TranslationNumbers.Supports(targetCode))
            {
                string localized;
                if (!TranslationNumbers.TryNormalize(selectedText, r.Text, targetCode, out localized))
                    return new AiChatResult { Provider = r.Provider, Model = r.Model, TokensIn = r.TokensIn, TokensOut = r.TokensOut,
                        Error = "Bản dịch không khớp số liệu, đơn vị, số hiệu hoặc trích dẫn của bản gốc." };
                r.Text = localized;
            }
            return r;
        }

        static string ExtractErrorMessage(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;
            try
            {
                var d = MiniJson.AsDict(MiniJson.Parse(body));
                if (d == null) return null;
                if (d.ContainsKey("error"))
                {
                    var errObj = MiniJson.AsDict(d["error"]);
                    if (errObj != null)
                    {
                        string msg = MiniJson.GetStr(errObj, "message");
                        if (msg.Length > 0) return msg;
                    }
                    string s = MiniJson.GetStr(d, "error");
                    if (s.Length > 0) return s;
                }
                string m = MiniJson.GetStr(d, "message");
                if (m.Length > 0) return m;
            }
            catch { }
            return null;
        }
    }
}
