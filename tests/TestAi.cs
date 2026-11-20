// Kiểm thử Rev0.9.3: AI Client (mock HTTP không gọi mạng thật), Token phiên, Giới hạn 500 từ Viết lại,
// Thay từ giữ nguyên định dạng, Cài đặt AI/DPI và Form Per-Monitor V2 DPI.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TroLyKyThuat;

sealed class MockHttpTransport : IAiTransport
{
    public string LastMethod, LastUrl, LastBody;
    public Dictionary<string, string> LastHeaders;
    public int CallCount;
    public Func<string, string, Dictionary<string, string>, string, AiHttpResponse> Handler;

    public AiHttpResponse Send(string method, string url, Dictionary<string, string> headers, string jsonBody, int timeoutMs)
    {
        CallCount++;
        LastMethod = method;
        LastUrl = url;
        LastHeaders = headers;
        LastBody = jsonBody;
        if (Handler != null) return Handler(method, url, headers, jsonBody);
        return new AiHttpResponse { StatusCode = 200, Body = "{}" };
    }
}

sealed class FmtFont : Rec
{
    public string Name = "Arial";
    public float Size = 12f;
    public int Bold = -1;
    public int Italic = -1;
    public int Underline = 0;
    public int Color = 0x112233;
    public int Subscript = 0;
    public int Superscript = 0;
    public FmtFont() : base("font") { }
}

sealed class FmtRange : Rec
{
    public FmtDoc Doc;
    public int Start, End;
    public FmtFont FontObj;
    public FmtRange(FmtDoc doc, int s, int e, FmtFont f) : base("fmtrng") { Doc = doc; Start = s; End = e; FontObj = f; }
    public string Text
    {
        get { return Doc.Buf.ToString().Substring(Start, End - Start); }
        set
        {
            Doc.Buf.Remove(Start, End - Start);
            Doc.Buf.Insert(Start, value);
            End = Start + value.Length;
        }
    }
    public FmtFont Font { get { return FontObj; } }
    public IEnumerable<FmtPara> Paragraphs { get { return new[] { new FmtPara(Doc) }; } }
}

sealed class FmtPara : Rec
{
    public FmtDoc Doc;
    public FmtPara(FmtDoc d) : base("fmtpara") { Doc = d; }
    public int OutlineLevel = 10;
    public FmtRange Range { get { return Doc.Range(0, Doc.Buf.Length); } }
}

sealed class FmtDoc : Rec
{
    public StringBuilder Buf;
    public FmtFont SpanFont = new FmtFont();
    public FmtRange HighlightedRange;
    public int UndoCalls;
    public FmtDoc(string text) : base("fmtdoc") { Buf = new StringBuilder(text); }
    public FmtRange Range(int s, int e)
    {
        var r = new FmtRange(this, s, e, SpanFont);
        HighlightedRange = r;
        return r;
    }
    public FmtRange Content { get { return Range(0, Buf.Length); } }
    public bool Undo(int count) { UndoCalls += count; return true; }
}

sealed class FmtSel : Rec
{
    public FmtDoc D;
    public int S, E;
    public FmtSel(FmtDoc d, int s, int e) : base("fmtsel") { D = d; S = s; E = e; }
    public FmtRange Range { get { return D.Range(S, E); } }
    public void SetRange(int s, int e) { S = s; E = e; }
}

sealed class FmtApp : Rec
{
    public FmtDoc D;
    public FmtSel Sel;
    public FmtApp(FmtDoc d, FmtSel sel) : base("fmtapp") { D = d; Sel = sel; }
    public MDocs Documents { get { return new MDocs(); } }
    public FmtDoc ActiveDocument { get { return D; } }
    public FmtSel Selection { get { return Sel; } }
    public Rec UndoRecord { get { return new Rec("undo"); } }
}

static class TAi
{
    static int fails = 0, total = 0;

    static void Eq(object a, object b, string name)
    {
        total++;
        if (!Equals(a, b))
        {
            fails++;
            if (fails < 80) Console.WriteLine("FAIL " + name + " (nhan [" + a + "], mong [" + b + "])");
        }
    }

    static int Main()
    {
        var mock = new MockHttpTransport();
        AiClient.Transport = mock;
        AiSessionUsage.Reset();

        // 1. Tự nhận dạng nhà cung cấp từ định dạng khóa (không dùng khóa thật)
        Eq(AiClient.DetectProvider("sk-ant-mockKey1234567890"), "anthropic", "nhận dạng Anthropic từ sk-ant-");
        Eq(AiClient.DetectProvider("AIzaMockGeminiKey123456"), "gemini", "nhận dạng Gemini từ AIza");
        Eq(AiClient.DetectProvider("sk-mockOpenAiKey12345678"), "openai", "nhận dạng OpenAI từ sk-");
        Eq(AiClient.DetectProvider("sk-proj-mockKey123456789"), "openai", "nhận dạng OpenAI sk-proj-");
        Eq(AiClient.DetectProvider("short"), "", "khóa quá ngắn không nhận dạng");
        Eq(AiClient.DetectProvider("unknownprefix123456789"), "", "tiền tố lạ trả rỗng");

        // 2. Kiểm tra giới hạn temperature & max tokens
        string pErr;
        Eq(AiClient.ValidateParams(0.3, 1024, out pErr), true, "tham số mặc định hợp lệ");
        Eq(AiClient.ValidateParams(-0.1, 1024, out pErr), false, "temperature âm bị chặn");
        Eq(AiClient.ValidateParams(2.1, 1024, out pErr), false, "temperature > 2.0 bị chặn");
        Eq(AiClient.ValidateParams(0.5, 32, out pErr), false, "max tokens < 64 bị chặn");
        Eq(AiClient.ValidateParams(0.5, 9000, out pErr), false, "max tokens > 8192 bị chặn");

        // 3. ListModels mock cho 3 nhà cung cấp
        mock.Handler = delegate(string m, string url, Dictionary<string, string> h, string b)
        {
            if (url.Contains("anthropic.com/v1/models"))
                return new AiHttpResponse { StatusCode = 200, Body = "{\"data\":[{\"id\":\"claude-sonnet-4-test\"},{\"id\":\"claude-haiku-test\"}]}" };
            if (url.Contains("googleapis.com"))
                return new AiHttpResponse { StatusCode = 200, Body = "{\"models\":[{\"name\":\"models/gemini-2.5-flash\",\"supportedGenerationMethods\":[\"generateContent\"]},{\"name\":\"models/embedding-001\",\"supportedGenerationMethods\":[\"embedContent\"]}]}" };
            if (url.Contains("openai.com/v1/models"))
                return new AiHttpResponse { StatusCode = 200, Body = "{\"data\":[{\"id\":\"text-embedding-3-small\"},{\"id\":\"gpt-4o-mini\"},{\"id\":\"gpt-4.1\"}]}" };
            return new AiHttpResponse { StatusCode = 404, Body = "{}" };
        };

        var mAnt = AiClient.ListModels("sk-ant-mockKey1234567890");
        Eq(mAnt.Ok && mAnt.Provider == "anthropic" && mAnt.Models.Contains("claude-sonnet-4-test"), true, "Anthropic ListModels");
        var mGem = AiClient.ListModels("AIzaMockGeminiKey123456");
        Eq(mGem.Ok && mGem.Models.Count == 1 && mGem.Models[0] == "gemini-2.5-flash", true, "Gemini ListModels lọc generateContent và bỏ tiền tố models/");
        Eq(mock.LastUrl.Contains("AIzaMockGeminiKey"), false, "Gemini dùng header x-goog-api-key, không đưa khóa lên URL");
        var mOai = AiClient.ListModels("sk-mockOpenAiKey12345678");
        Eq(mOai.Ok && mOai.Models.Count == 3 && mOai.Models[0].StartsWith("gpt-"), true, "OpenAI ListModels đưa gpt-* lên đầu");

        // 4. Lỗi HTTP / bảo mật không lộ khóa API trong thông báo lỗi
        string fakeKey = "sk-secretFakeKeyDoNotLeak999";
        mock.Handler = delegate { return new AiHttpResponse { StatusCode = 401, Error = "Sai khóa " + fakeKey, Body = "{\"error\":{\"message\":\"Invalid key " + fakeKey + "\"}}" }; };
        var failRes = AiClient.ListModels(fakeKey);
        Eq(failRes.Ok, false, "HTTP 401 trả Ok = false");
        Eq(failRes.Error.Contains(fakeKey), false, "thông báo lỗi tuyệt đối không lộ nguyên văn khóa API");

        // 5. Gọi Complete, tích lũy token phiên & Sửa câu dài / Đồng nghĩa AI
        AiSessionUsage.Reset();
        mock.Handler = delegate(string m, string url, Dictionary<string, string> h, string b)
        {
            if (url.Contains("anthropic.com/v1/messages"))
                return new AiHttpResponse { StatusCode = 200, Body = "{\"content\":[{\"type\":\"text\",\"text\":\"Câu 1 ngắn gọn. Câu 2 rõ ý.\"}],\"usage\":{\"input_tokens\":120,\"output_tokens\":45}}" };
            if (url.Contains(":generateContent"))
                return new AiHttpResponse { StatusCode = 200, Body = "{\"candidates\":[{\"content\":{\"parts\":[{\"text\":\"vì vậy\\ncho nên\\nbởi lẽ đó\"}]}}],\"usageMetadata\":{\"promptTokenCount\":80,\"candidatesTokenCount\":20}}" };
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"Đoạn văn đã được viết lại mạch lạc.\"}}],\"usage\":{\"prompt_tokens\":150,\"completion_tokens\":55}}" };
        };

        var st = new AppSettings { ApiKeyEnc = Secret.Protect("sk-ant-mockKey1234567890"), ApiProvider = "anthropic", ApiModel = "claude-sonnet-4-test", AiTemp = 0.2, AiMaxTokens = 512 };
        var rSplit = AiClient.SplitLongSentence(st, "Một câu kỹ thuật rất dài cần được tách làm hai câu cho rõ.", "Đoạn báo cáo kỹ thuật.");
        Eq(rSplit.Ok && rSplit.Text.Contains("Câu 1"), true, "SplitLongSentence qua Anthropic mock");
        Eq(AiSessionUsage.TokensIn, 120L, "cộng dồn TokensIn lần 1");
        Eq(AiSessionUsage.TokensOut, 45L, "cộng dồn TokensOut lần 1");

        var stGem = new AppSettings { ApiKeyEnc = Secret.Protect("AIzaMockGeminiKey123456"), ApiProvider = "gemini", ApiModel = "gemini-2.5-flash", AiTemp = 0.3, AiMaxTokens = 512 };
        List<string> aiSyns;
        var rSyn = AiClient.SuggestSynonymsAi(stGem, "do đó", "Áp suất tăng, do đó cần mở van.", out aiSyns);
        Eq(rSyn.Ok && aiSyns.Count == 3 && aiSyns.Contains("vì vậy"), true, "SuggestSynonymsAi qua Gemini mock");
        Eq(AiSessionUsage.TokensIn, 200L, "cộng dồn TokensIn lần 2 (120 + 80)");
        Eq(AiSessionUsage.TokensOut, 65L, "cộng dồn TokensOut lần 2 (45 + 20)");

        // 6. Tab Viết lại: kiểm tra giới hạn 500 từ và ngữ cảnh toàn văn
        var stOai = new AppSettings { ApiKeyEnc = Secret.Protect("sk-mockOpenAiKey12345678"), ApiProvider = "openai", ApiModel = "gpt-4o-mini", AiTemp = 0.4, AiMaxTokens = 1024 };
        string w501 = string.Join(" ", Enumerable.Repeat("từ", 501).ToArray());
        int callsBefore = mock.CallCount;
        var rOver = AiClient.RewriteText(stOai, w501, "Trang trọng", "Ngữ cảnh báo cáo");
        Eq(rOver.Ok, false, "đoạn 501 từ bị từ chối");
        Eq(mock.CallCount, callsBefore, "chặn ở client nếu > 500 từ, không gửi HTTP");
        Eq(AiClient.RewriteText(stOai, "   ", "Trang trọng", "").Ok, false, "đoạn rỗng bị từ chối");

        string w500 = string.Join(" ", Enumerable.Repeat("từ", 500).ToArray());
        var rExact = AiClient.RewriteText(stOai, w500, "Trang trọng", "Tiêu đề 1: Thiết kế đường ống");
        Eq(rExact.Ok && !mock.LastBody.Contains("Tiêu đề 1"), true, "đúng 500 từ cho phép viết lại, chỉ gửi văn bản đã chọn");
        Eq(AiSessionUsage.TotalTokens, 200L + 65L + 150L + 55L, "tổng token phiên tích lũy đúng");
        Eq(AiSessionUsage.FormatFooter().Contains("350") && AiSessionUsage.FormatFooter().Contains("120"), true, "chuỗi hiển thị footer token đúng");

        AiSessionUsage.Reset();
        Eq(AiSessionUsage.TotalTokens, 0L, "Reset() đưa bộ đếm token phiên về 0");

        // 7. WordText.Replace giữ nguyên định dạng ký tự (font, cỡ, đậm, nghiêng, màu) + ApplyRewrite / UndoLast
        var doc = new FmtDoc("Hệ thống do đó đạt hiệu suất cao.\r");
        doc.SpanFont.Name = "Times New Roman"; doc.SpanFont.Size = 13f; doc.SpanFont.Bold = -1; doc.SpanFont.Italic = -1; doc.SpanFont.Color = 0xAA0000;
        var sel = new FmtSel(doc, 9, 14); // "do đó"
        var app = new FmtApp(doc, sel);
        var wtx = new WordText(app);
        SynResult sr = wtx.Suggest();
        Eq(sr.Text, "do đó", "Suggest đọc đúng cụm đang bôi đen");
        string repMsg = wtx.Replace(sr, "vì vậy");
        Eq(doc.Buf.ToString().Contains("vì vậy"), true, "Replace thay đúng văn bản");
        Eq(doc.HighlightedRange.Font.Name, "Times New Roman", "Replace giữ nguyên tên Font");
        Eq(doc.HighlightedRange.Font.Size, 13f, "Replace giữ nguyên cỡ Font");
        Eq(doc.HighlightedRange.Font.Bold, -1, "Replace giữ nguyên chữ đậm");
        Eq(doc.HighlightedRange.Font.Italic, -1, "Replace giữ nguyên chữ nghiêng");
        Eq(doc.HighlightedRange.Font.Color, 0xAA0000, "Replace giữ nguyên màu chữ");

        sel.SetRange(0, doc.Buf.Length - 1);
        RewriteSnapshot rwSnap = wtx.GetRewriteSnapshot();
        Eq(rwSnap.WordCount > 0 && rwSnap.WordCount <= 500, true, "GetRewriteSnapshot đếm đúng số từ");
        wtx.ApplyRewrite(rwSnap, "Do vậy hệ thống đạt hiệu suất rất cao.");
        Eq(doc.Buf.ToString().StartsWith("Do vậy"), true, "ApplyRewrite cập nhật đoạn văn");
        wtx.UndoLast();
        Eq(doc.UndoCalls, 1, "UndoLast gọi hoàn tác 1 bước trong Word");

        // 8. DPI awareness & cấu hình Per-Monitor V2 cho Form (Step H)
        using (var dummyForm = new Form())
        {
            DpiUtil.ConfigureForm(dummyForm, 11f);
            Eq(dummyForm.AutoScaleMode, AutoScaleMode.Dpi, "Form dùng AutoScaleMode.Dpi");
            Eq(dummyForm.Font.SizeInPoints, 11f, "Font chuẩn kích thước 11pt");
            using (DpiUtil.EnterPerMonitorV2())
            {
                // Không ném ngoại lệ ngay cả khi chạy trong console test runner
                Eq(DpiUtil.GetScaleFactor(dummyForm) >= 1f, true, "hệ số DPI >= 1.0");
            }
        }

        // 9. Kiểm tra Json tròn vẹn Unicode tiếng Việt
        string jsonRound = MiniJson.Stringify(new Dictionary<string, object> { { "text", "Tiếng Việt \"đặc biệt\"" }, { "n", 42 } });
        var parsed = MiniJson.AsDict(MiniJson.Parse(jsonRound));
        Eq(MiniJson.GetStr(parsed, "text"), "Tiếng Việt \"đặc biệt\"", "MiniJson round-trip tiếng Việt & escape");
        Eq(MiniJson.GetInt(parsed, "n"), 42, "MiniJson round-trip số nguyên");

        // 10. Rev0.9.5 Việc 1: Giữ nguyên ngôn ngữ, DetectLang, LanguageMismatch, Lời nhắc tiếng Anh và Thử lại an toàn
        // (a) DetectLang
        Eq(AiClient.DetectLang("This is an English technical specification document."), "en", "DetectLang: tiếng Anh chuẩn");
        Eq(AiClient.DetectLang("Hệ thống kiểm tra áp suất và nhiệt độ trong đường ống."), "vi", "DetectLang: tiếng Việt có dấu");
        Eq(AiClient.DetectLang("Technical analysis report"), "en", "DetectLang: tiếng Anh ngắn không có từ dừng");
        Eq(AiClient.DetectLang("Báo cáo kỹ thuật"), "vi", "DetectLang: tiếng Việt ngắn");
        Eq(AiClient.DetectLang("12345 67890 +-="), "other", "DetectLang: số và ký hiệu trả other");
        Eq(AiClient.DetectLang("   "), "other", "DetectLang: khoảng trắng trả other");
        Eq(AiClient.DetectLang("abc"), "other", "DetectLang: chuỗi dưới 4 chữ cái trả other");

        // (b) LanguageMismatch
        Eq(AiClient.LanguageMismatch("The pipeline operates at high pressure.", "Đường ống hoạt động ở áp suất cao."), true, "LanguageMismatch: vào Anh ra Việt = true");
        Eq(AiClient.LanguageMismatch("Đường ống hoạt động ở áp suất cao.", "The pipeline operates at high pressure."), true, "LanguageMismatch: vào Việt ra Anh = true");
        Eq(AiClient.LanguageMismatch("The pipeline operates at high pressure.", "The pipe operates under high pressure conditions."), false, "LanguageMismatch: vào Anh ra Anh = false");
        Eq(AiClient.LanguageMismatch("Đường ống hoạt động ở áp suất cao.", "Hệ thống đường ống vận hành dưới áp suất lớn."), false, "LanguageMismatch: vào Việt ra Việt = false");
        Eq(AiClient.LanguageMismatch("hi", "Đường ống áp suất cao."), false, "LanguageMismatch: chuỗi đầu vào quá ngắn = false");
        Eq(AiClient.LanguageMismatch("The pipeline operates at high pressure.", "ok"), false, "LanguageMismatch: chuỗi đầu ra quá ngắn = false");
        Eq(AiClient.LanguageMismatch("", ""), false, "LanguageMismatch: rỗng = false");

        // (c) BuildRewritePrompt, BuildSplitPrompt, BuildSynonymPrompt
        var prEn = AiClient.BuildRewritePrompt("Trang trọng", "The system maintains pressure below 10 bar.", "DocContext");
        Eq(prEn.Item1.Contains("SAME language"), true, "BuildRewritePrompt tiếng Anh chứa 'SAME language'");
        Eq(prEn.Item1.Contains("English"), true, "BuildRewritePrompt tiếng Anh chứa 'English'");
        Eq(prEn.Item1.Contains("NEVER translate"), true, "BuildRewritePrompt tiếng Anh chứa 'NEVER translate'");

        var prVi = AiClient.BuildRewritePrompt("Chuyên nghiệp kỹ thuật", "Hệ thống duy trì áp suất dưới 10 bar.", "");
        Eq(prVi.Item1.Contains("Vietnamese"), true, "BuildRewritePrompt tiếng Việt chứa 'Vietnamese'");
        Eq(prVi.Item1.Contains("SAME language"), true, "BuildRewritePrompt tiếng Việt chứa 'SAME language'");

        var prSplit = AiClient.BuildSplitPrompt("This is a very long sentence that needs to be divided into multiple smaller sentences for better readability.", "Context");
        Eq(prSplit.Item1.Contains("SAME language") && prSplit.Item1.Contains("English") && prSplit.Item1.Contains("NEVER translate"), true, "BuildSplitPrompt tiếng Anh có đủ mệnh lệnh ngôn ngữ");

        var prSyn = AiClient.BuildSynonymPrompt("evaluate", "We must evaluate the results.");
        Eq(prSyn.Item1.Contains("SAME language") && prSyn.Item1.Contains("English"), true, "BuildSynonymPrompt tiếng Anh chứa SAME language");

        // (d) Thử lại an toàn khi AI trả sai ngôn ngữ
        int retryCallCount = 0;
        mock.Handler = delegate(string m, string url, Dictionary<string, string> h, string b)
        {
            retryCallCount++;
            if (retryCallCount == 1)
            {
                // Lần 1: trả sai sang tiếng Việt
                return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"Hệ thống đường ống hoạt động ở áp suất rất cao.\"}}],\"usage\":{\"prompt_tokens\":100,\"completion_tokens\":30}}" };
            }
            // Lần 2 (thử lại): trả đúng tiếng Anh
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"The pipeline system operates under very high pressure.\"}}],\"usage\":{\"prompt_tokens\":120,\"completion_tokens\":25}}" };
        };
        retryCallCount = 0;
        var rRetrySuccess = AiClient.RewriteText(stOai, "The pipeline system operates at very high pressure.", "Trang trọng", "");
        Eq(rRetrySuccess.Ok && rRetrySuccess.Text.Contains("pipeline system operates"), true, "RewriteText tự động thử lại thành công khi lần đầu trả sai ngôn ngữ");
        Eq(retryCallCount, 2, "gọi đúng 2 lần (lần 1 sai + lần 2 thử lại)");

        // Thử lại lần 2 vẫn sai ngôn ngữ -> trả Ok=false
        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"Vẫn là tiếng Việt sau khi thử lại.\"}}],\"usage\":{\"prompt_tokens\":100,\"completion_tokens\":30}}" };
        };
        var rRetryFail = AiClient.RewriteText(stOai, "The pipeline system operates at very high pressure.", "Trang trọng", "");
        Eq(rRetryFail.Ok, false, "RewriteText thất bại khi cả 2 lần đều sai ngôn ngữ");
        Eq(rRetryFail.Error.Contains("AI trả lời sai ngôn ngữ (đã thử lại một lần)"), true, "báo đúng thông báo lỗi sai ngôn ngữ");


        // ===== Rev0.9.5.6: Công cụ AI, viết lại và dịch =====
        Eq(AiClient.TargetLangs.Length, 10, "Công cụ AI: có 10 ngôn ngữ đích");
        Eq(AiClient.TargetLangs[0][0] + "," + AiClient.TargetLangs[1][0], "vi,en", "Công cụ AI: tiếng Việt và tiếng Anh đứng đầu");
        var codes = new HashSet<string>(); foreach (string[] l in AiClient.TargetLangs) codes.Add(l[0]);
        Eq(codes.Count, 10, "Công cụ AI: mã ngôn ngữ không trùng");
        Eq(AiClient.LangEnglishName("ja"), "Japanese", "tên tiếng Anh của mã ja");

        string enText = "The pipeline system operates at very high pressure [1].";
        string viText = "Hệ thống đường ống vận hành ở áp suất rất cao [1].";
        var pRw = AiClient.BuildTransformPrompt("Trang trọng", enText, "Ngữ cảnh", true, false, "vi");
        Eq(pRw.Item1.Contains("ONLY") && pRw.Item1.Contains("Never translate") && pRw.Item1.Contains("English"), true, "chế độ chỉ viết lại: cấm dịch, giữ ngôn ngữ gốc English");
        Eq(pRw.Item2.Contains("REWRITE ONLY") && pRw.Item2.Contains(enText) && pRw.Item2.Contains("---TEXT START---"), true, "chế độ chỉ viết lại: mẫu người dùng có nhãn chế độ và khung văn bản");
        Eq(pRw.Item1.Contains("citation") && pRw.Item1.Contains("same number of paragraphs"), true, "chế độ chỉ viết lại: giữ trích dẫn và số đoạn");
        var pTr = AiClient.BuildTransformPrompt("Trang trọng", enText, "", false, true, "ja");
        Eq(pTr.Item1.Contains("Do NOT rewrite") && pTr.Item1.Contains("Japanese") && pTr.Item1.Contains("English"), true, "chế độ chỉ dịch: không viết lại, từ English sang Japanese");
        Eq(pTr.Item2.Contains("TRANSLATE ONLY") && !pTr.Item2.Contains("Target tone"), true, "chế độ chỉ dịch: không đưa tone vào lời nhắc");
        var pBoth = AiClient.BuildTransformPrompt("Ngắn gọn súc tích", viText, "", true, true, "en");
        Eq(pBoth.Item1.Contains("translate") && pBoth.Item1.Contains("rewrite") && pBoth.Item1.Contains("Vietnamese") && pBoth.Item1.Contains("English"), true, "chế độ dịch và viết lại: từ Vietnamese sang English");
        Eq(pBoth.Item2.Contains("TRANSLATE AND REWRITE") && pBoth.Item2.Contains("Concise"), true, "chế độ dịch và viết lại: có tone (Concise)");
        Eq(AiClient.RetryInstruction(true, "ja", "English").Contains("Japanese"), true, "câu thử lại nêu ngôn ngữ đích khi dịch");
        Eq(AiClient.RetryInstruction(false, "ja", "English").Contains("English"), true, "câu thử lại nêu ngôn ngữ gốc khi chỉ viết lại");

        Eq(AiClient.TransformText(stOai, enText, "Trang trọng", "", false, false, "vi").Ok, false, "không chọn Viết lại lẫn Dịch thì từ chối");
        var rSame = AiClient.TransformText(stOai, enText, "Trang trọng", "", false, true, "en");
        Eq(rSame.Ok == false && rSame.Error.Contains("tiếng Anh"), true, "chỉ dịch sang đúng ngôn ngữ gốc thì từ chối");

        // dịch EN -> VI: trả đúng tiếng Việt
        int trCalls = 0; string lastBody = "";
        mock.Handler = delegate(string m, string url, Dictionary<string, string> h, string b)
        {
            trCalls++; lastBody = b;
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"Hệ thống đường ống vận hành ở áp suất rất cao [1].\"}}],\"usage\":{\"prompt_tokens\":50,\"completion_tokens\":20}}" };
        };
        var rTr = AiClient.TransformText(stOai, enText, "Trang trọng", "", false, true, "vi");
        Eq(rTr.Ok && rTr.Text.Contains("[1]") && trCalls == 1, true, "dịch EN sang VI: một lần gọi, giữ nguyên trích dẫn [1]");
        Eq(lastBody.Contains("Vietnamese") && lastBody.Contains("Do NOT rewrite"), true, "dịch EN sang VI: lời nhắc gửi đi nêu Vietnamese và cấm viết lại");

        // dịch EN -> VI: lần 1 AI trả tiếng Anh (sai), lần 2 đúng tiếng Việt
        trCalls = 0;
        mock.Handler = delegate(string m, string url, Dictionary<string, string> h, string b)
        {
            trCalls++;
            string c = trCalls == 1 ? "The pipeline system operates at very high pressure [1]." : "Hệ thống đường ống vận hành ở áp suất rất cao [1].";
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"" + c + "\"}}],\"usage\":{\"prompt_tokens\":50,\"completion_tokens\":20}}" };
        };
        var rTr2 = AiClient.TransformText(stOai, enText, "Trang trọng", "", false, true, "vi");
        Eq(rTr2.Ok && trCalls == 2 && rTr2.Text.Contains("Hệ thống"), true, "dịch EN sang VI: sai ngôn ngữ lần đầu, tự thử lại và đạt");
        Eq(rTr2.TokensIn, 100, "dịch: cộng token của cả hai lần gọi");

        // dịch sang tiếng Nhật: trả chữ Latin hai lần thì từ chối; trả chữ Nhật thì đạt
        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"The pipeline system operates at high pressure.\"}}],\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":10}}" };
        };
        Eq(AiClient.TransformText(stOai, enText, "Trang trọng", "", false, true, "ja").Ok, false, "dịch sang Nhật: AI trả chữ Latin hai lần thì từ chối");
        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"パイプラインシステムは非常に高い圧力で運転される [1]。\"}}],\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":10}}" };
        };
        Eq(AiClient.TransformText(stOai, enText, "Trang trọng", "", false, true, "ja").Ok, true, "dịch sang Nhật: trả chữ Nhật thì đạt");

        // chỉ viết lại tiếng Anh: AI lỡ dịch sang tiếng Việt hai lần thì từ chối
        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":\"Hệ thống đường ống vận hành ở áp suất rất cao.\"}}],\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":10}}" };
        };
        Eq(AiClient.TransformText(stOai, enText, "Trang trọng", "", true, false, "vi").Ok, false, "chỉ viết lại EN: AI dịch sang VI hai lần thì từ chối");

        Eq(!pRw.Item2.Contains("Ngữ cảnh") && !prEn.Item2.Contains("DocContext"), true, "lời nhắc viết lại chỉ chứa vùng chọn");
        var blocks = AiReview.Blocks(40, "Clean.\rThey believes. They believes.\r");
        Eq(blocks.Count, 2, "chia tài liệu theo đoạn");
        Eq(blocks[1].Start, 47, "offset Word tính cả dấu ngắt");
        string reply = "{\"job\":\"x\",\"issues\":[{\"block\":1,\"offset\":15,\"quote\":\"They believes\",\"replacement\":\"They believe\",\"reason\":\"Chủ vị\",\"severity\":\"error\"}]}";
        var findings = AiReview.Parse(reply, "x", "gram", blocks);
        Eq(findings.Count, 1, "trích đoạn lặp được định vị bằng offset");
        Eq(findings[0].Start, 62, "lấy đúng lần thứ hai");
        string markedLong = reply.Replace("\"severity\"", "\"type\":\"long_sentence\",\"severity\"");
        Eq(AiReview.Parse(markedLong, "x", "gram", blocks, 10)[0].Issue.Cat, "ai-gram", "không mở Sửa câu dài cho cụm quá ngắn");
        var longIssue = AiReview.Parse(markedLong, "x", "gram", blocks, 1)[0].Issue;
        Eq(longIssue.Cat == "ai-long" && longIssue.Sev == "sus", true, "câu dài hợp lệ là gợi ý duyệt thủ công");
        foreach (string bad in new[] { reply.Replace("15,", "16,"), reply.Replace("They believes", "bịa"), reply + "{}",
            reply.Replace("\"job\":\"x\"", "\"job\":\"y\""), "{\"job\":\"x\",\"issues\":[bad]}" })
        {
            try { AiReview.Parse(bad, "x", "gram", blocks); Eq(false, true, "từ chối phản hồi lỗi"); }
            catch (FormatException) { Eq(true, true, "từ chối phản hồi lỗi"); }
        }
        Eq(AiReview.Parse("{\"job\":\"x\",\"issues\":[]}", "x", "gram", blocks).Count, 0, "mảng trống hợp lệ");
        Eq(AiReview.CanReplace("12,5 MPa [1]", "13,5 MPa [1]"), false, "chặn thay đổi số liệu");
        Eq(AiReview.CanReplace("12,5 MPa [1]", "Áp suất 12,5 MPa [1]"), true, "giữ số liệu và nguồn");
        int reviewCalls = 0;
        mock.Handler = delegate(string method, string url, Dictionary<string, string> headers, string body)
        {
            reviewCalls++;
            var request = MiniJson.AsDict(MiniJson.Parse(body));
            var messages = MiniJson.AsList(request["messages"]);
            var payload = MiniJson.AsDict(MiniJson.Parse(MiniJson.GetStr(MiniJson.AsDict(messages.Last()), "content")));
            string job = MiniJson.GetStr(payload, "job");
            string answer = MiniJson.Stringify(new Dictionary<string, object> { { "job", job }, { "issues", new object[0] } });
            return new AiHttpResponse { StatusCode = 200, Body = "{\"choices\":[{\"message\":{\"content\":" + MiniJson.Stringify(answer) + "}}]}" };
        };
        var longDoc = string.Join("\r", Enumerable.Repeat(new string('A', 1199) + ".", 9));
        Eq(AiReview.Run(stOai, "style", AiReview.Blocks(0, longDoc)).Count, 0, "quét nhiều lô qua API giả");
        Eq(reviewCalls, 3, "gửi toàn bộ chín đoạn qua ba lô");

        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200,
                Body = "{\"choices\":[{\"finish_reason\":\"length\",\"message\":{\"content\":\"Incomplete text\"}}]}" };
        };
        Eq(AiClient.TransformText(stOai, enText, "Trang trọng", "", true, false, "en").Ok, false,
            "không hiển thị bản viết lại bị cắt do hết token");
        mock.Handler = delegate
        {
            return new AiHttpResponse { StatusCode = 200,
                Body = "{\"choices\":[{\"message\":{\"content\":\"Looks valid\"}}]} extra" };
        };
        Eq(AiClient.Complete("sk-example", "test", "sys", "user", 0.2, 256).Ok, false,
            "không chấp nhận HTTP 200 với nội dung JSON dư");

        const string viPressure = "Áp suất thiết kế là 12,50 MPa theo Điều 5.2 [1].";
        const string enPressure = "The design pressure is 12.50 MPa according to Clause 5.2 [1].";
        mock.Handler = delegate
        {
            string result = "The design pressure is 12,50 MPa according to Clause 5.2 [1].";
            return new AiHttpResponse { StatusCode = 200,
                Body = "{\"choices\":[{\"message\":{\"content\":" + MiniJson.Stringify(result) + "}}]}" };
        };
        Eq(AiClient.TransformText(stOai, viPressure, "", "", false, true, "en").Text, enPressure,
            "dịch Việt sang Anh phải tự chuẩn hóa 12,50 thành 12.50 và giữ Điều 5.2");
        mock.Handler = delegate
        {
            string result = "Áp suất thiết kế là 12.50 MPa theo Điều 5.2 [1].";
            return new AiHttpResponse { StatusCode = 200,
                Body = "{\"choices\":[{\"message\":{\"content\":" + MiniJson.Stringify(result) + "}}]}" };
        };
        Eq(AiClient.TransformText(stOai, enPressure, "", "", true, true, "vi").Text, viPressure,
            "dịch và viết lại Anh sang Việt phải tự chuẩn hóa 12.50 thành 12,50");

        string local;
        Eq(TranslationNumbers.TryNormalize(
            "Giá trị −0,0120 × 10⁻³ MPa, 1.234,500 kPa và 45,5%.",
            "Values: −0,0120 × 10⁻³ MPa, 1,234.500 kPa and 45.5%.", "en", out local)
            && local == "Values: −0.0120 × 10⁻³ MPa, 1\u202F234.500 kPa and 45.5 %.", true,
            "giữ dấu, số mũ, chữ số 0 và gom nghìn theo kiểu SI");
        Eq(TranslationNumbers.TryNormalize(
            "The acceptable range is 10.5–12.50 MPa, or +.5 MPa above nominal.",
            "Khoảng chấp nhận là 10.5–12.50 MPa, hoặc +.5 MPa trên danh định.", "vi", out local)
            && local == "Khoảng chấp nhận là 10,5–12,50 MPa, hoặc +0,5 MPa trên danh định.", true,
            "dịch hai cận của khoảng và thêm số 0 trước dấu thập phân");
        Eq(TranslationNumbers.TryNormalize("The range is 12.5-13 MPa.",
            "Khoảng đo là 12.5-13 MPa.", "vi", out local) && local == "Khoảng đo là 12,5-13 MPa.", true,
            "giữ dấu nối của khoảng và đổi phần thập phân");
        Eq(TranslationNumbers.TryNormalize(
            "Hệ số 0,85 theo ISO 9001:2015, Điều 5.2, mục [1, 2]. Đo lúc 12:30 ngày 05/03/2026 tại P-101, H₂S.",
            "Coefficient 0,85 per ISO 9001:2015, Clause 5.2, ref. [1, 2]. Measured at 12:30 on 05/03/2026 at P-101, H₂S.",
            "en", out local) && local.Contains("Coefficient 0.85") && local.Contains("ISO 9001:2015, Clause 5.2")
            && local.Contains("12:30 on 05/03/2026 at P-101, H₂S"), true,
            "bản dịch giữ giờ, ngày, mã, điều khoản, công thức hóa học");
        Eq(TranslationNumbers.TryNormalize("Đo 1.234 kPa.", "Measured 1.234 kPa.", "en", out local)
            && local == "Measured 1\u202F234 kPa.", true, "ba chữ số sau dấu chấm trong văn bản Việt dùng quy ước nhóm nghìn");
        Eq(TranslationNumbers.TryNormalize("Đo 1.234 kPa và 12,50 MPa.", "Measured 1,234 kPa and 12.50 MPa.", "en", out local)
            && local == "Measured 1\u202F234 kPa and 12.50 MPa.", true, "chấp nhận AI đổi dấu đúng giá trị rồi chuẩn hóa nhóm nghìn");
        Eq(TranslationNumbers.TryNormalize("1,234", "1,234", "en", out local) && local == "1,234", true,
            "giữ chuỗi chỉ có số mơ hồ khi thiếu quy ước ngôn ngữ");
        Eq(TranslationNumbers.TryNormalize("Áp suất 12,50 MPa.", "The pressure is 12.5 MPa.", "en", out local), false,
            "chặn bản dịch làm mất số 0 thể hiện độ chính xác");
        foreach (string changed in new[] { enPressure.Replace("12.50", "1250"), enPressure.Replace("MPa", "mPa"),
            enPressure.Replace("5.2", "5.3"), enPressure.Replace("[1]", "[2]"), enPressure.Replace("12.50", "−12.50") })
            Eq(TranslationNumbers.TryNormalize(viPressure, changed, "en", out local), false,
                "chặn kết quả sai giá trị, đơn vị, dấu hoặc số hiệu");
        Eq(TranslationNumbers.TryNormalize("Tải 1,20e-4 N.", "Load 1.20e-5 N.", "en", out local), false,
            "chặn bản dịch đổi số mũ");
        Eq(TranslationNumbers.TryNormalize("Chọn API 579 và nhiệt độ 12,5 °C.",
            "Use ISO 579 and temperature 12.5 °C.", "en", out local), false, "chặn đổi tên tiêu chuẩn khi số hiệu giống nhau");

        mock.Handler = delegate
        {
            var changed = enPressure.Replace("MPa", "kPa");
            return new AiHttpResponse { StatusCode = 200,
                Body = "{\"choices\":[{\"message\":{\"content\":" + MiniJson.Stringify(changed) + "}}]}" };
        };
        Eq(AiClient.TransformText(stOai, viPressure, "", "", false, true, "en").Ok, false,
            "không tạo bản xem trước có sai đơn vị");

        var trDoc = new FmtDoc("START " + viPressure + " END\r");
        var trWord = new WordText(new FmtApp(trDoc, new FmtSel(trDoc, 6, 6 + viPressure.Length)));
        RewriteSnapshot trSnap = trWord.GetRewriteSnapshot();
        try { trWord.ApplyRewrite(trSnap, enPressure.Replace("12.50", "15.50"), "en"); Eq(false, true, "chặn sửa sai trong bản xem trước"); }
        catch (UserMsg) { Eq(trDoc.Buf.ToString() == "START " + viPressure + " END\r", true, "chặn sửa sai trong bản xem trước và giữ nguyên Word"); }
        trWord.ApplyRewrite(trSnap, enPressure.Replace("12.50", "12,50"), "en");
        Eq(trDoc.Buf.ToString(), "START " + enPressure + " END\r", "áp dụng đúng vùng chọn và dấu thập phân đích");

        // WrongLanguage trực tiếp
        Eq(AiClient.WrongLanguage(enText, viText, true, "vi"), false, "WrongLanguage: dịch ra đúng tiếng Việt");
        Eq(AiClient.WrongLanguage(enText, enText, true, "vi"), true, "WrongLanguage: dịch mà vẫn là tiếng Anh là sai");
        Eq(AiClient.WrongLanguage(enText, viText, false, "vi"), true, "WrongLanguage: chỉ viết lại mà ra tiếng Việt là sai");
        Eq(AiClient.WrongLanguage(enText, "Le système de pipeline fonctionne à haute pression.", true, "fr"), false, "WrongLanguage: tiếng Pháp (chữ Latin) không kiểm tra được nên không chặn");
        Eq(AiClient.WrongLanguage(enText, "Система трубопроводов работает при высоком давлении.", true, "ru"), false, "WrongLanguage: tiếng Nga đúng chữ Kirin");
        Eq(AiClient.WrongLanguage(enText, "The pipeline system operates.", true, "ru"), true, "WrongLanguage: đích tiếng Nga mà ra chữ Latin là sai");

        Console.WriteLine(fails == 0 ? "ALL PASS (" + total + ")" : "FAILED " + fails + " / " + total);
        return fails;
    }
}
