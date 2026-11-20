using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TroLyKyThuat
{
    internal static class TranslationNumbers
    {
        const string Digits = @"[0-9]+(?:[.,][0-9]+)*";
        const string Groups = @"[0-9]{1,3}(?:[ \u00A0\u202F][0-9]{3})+(?:[.,][0-9]+)*(?![0-9])";
        const string Exponent = @"(?:[eE][+\-−]?[0-9]+|[ \u00A0\u202F]*×[ \u00A0\u202F]*10(?:\^[+\-−]?[0-9]+|[⁺⁻]?[⁰¹²³⁴⁵⁶⁷⁸⁹]+))";
        static readonly Regex Tokens = new Regex(
            @"\[[^\]\r\n]+\]|" +
            @"(?<![\p{L}\p{N}_])(?:API|ASME|ASTM|ISO|IEC|EN|DNV|NACE|AMPP|TCVN|QCVN)[ \u00A0\u202F]+" +
                @"(?:[A-Z]+[- ]?)?[0-9]+[A-Z]*(?:[./:-][A-Z]*[0-9]+[A-Z]*)*|" +
            @"(?<![\p{L}\p{N}_])(?:[A-Za-z]+[0-9₀-₉][\w₀-₉]*|(?:[A-Z][A-Z0-9]*[/-])+[A-Z]*[0-9][\w]*)(?:[./:-][0-9][\w]*)*|" +
            @"\b(?:[0-9]+(?:[:/][0-9]+)+|[0-9]{4}-[0-9]{2}-[0-9]{2}|" +
                @"[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4}|[0-9]{1,2}-[0-9]{1,2}-[0-9]{4})\b|" +
            @"(?<sign>[+\-−±]?)(?<num>" + Groups + "|" + Digits + @"|(?<![\p{L}\p{N}_])[.,][0-9]+)(?<exp>" + Exponent + @")?|\p{Nd}+",
            RegexOptions.CultureInvariant);
        static readonly Regex UnitGap = new Regex(@"^[ \u00A0\u202F]*");
        static readonly Regex Adjacent = new Regex(@"[\p{L}_]");
        static readonly Regex TimeWords = new Regex(@"^(?:years?|days?|hours?|weeks?|months?|sec)$");
        static readonly Regex Qualifier = new Regex(@"^\((?:a|g|abs|gauge)\)");
        static readonly Regex RangeEnd = new Regex(@"^[ \u00A0\u202F]*(?:[–—−-])[ \u00A0\u202F]*[0-9]");
        static readonly Regex Ratio = new Regex(@"^[/:][0-9]");
        static readonly Regex ReferencePrefix = new Regex(
            @"\b(?:clause|section|figure|fig\.?|table|appendix|revision|rev\.?|version|equation|điều|mục|khoản|hình|bảng|phụ lục|phiên bản|phương trình)\s+$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex Coefficient = new Regex(@"(?:\bhệ số|\btỷ số|\btỉ số)[ \u00A0\u202F]+$", RegexOptions.IgnoreCase);

        internal static bool Supports(string target) { return target == "en" || target == "vi"; }

        sealed class Token
        {
            public int Start, Length;
            public string Original, Number, Suffix, Unit, Sign;
            public bool IsMeasurement { get { return Number != null; } }
        }

        static List<Token> Read(string text)
        {
            var result = new List<Token>();
            int end = 0;
            foreach (Match m in Tokens.Matches(text))
            {
                if (m.Index < end) continue;
                Group n = m.Groups["num"];
                end = m.Index + m.Length;
                var token = new Token { Start = m.Index, Length = m.Length, Original = m.Value };
                result.Add(token);
                if (!n.Success) continue;

                string context = text.Substring(Math.Max(0, m.Index - 32), Math.Min(32, m.Index));
                if (ReferencePrefix.IsMatch(context)) continue;
                string tail = text.Substring(end);
                int gap = UnitGap.Match(tail).Length;
                UnitHit unit = Units.Scan(tail.Substring(gap));
                bool hasUnit = unit != null && !TimeWords.IsMatch(unit.Raw) &&
                    (gap + unit.Len == tail.Length || !char.IsLetter(tail[gap + unit.Len]));
                bool adjacent = m.Index > 0 && Adjacent.IsMatch(text[m.Index - 1].ToString());
                // Keep dates, section numbers, and identifiers as written. For measurements in
                // ranges/equations, the unit or scientific notation distinguishes the number.
                bool range = RangeEnd.IsMatch(tail) ||
                    (m.Index > 0 && ("–—".IndexOf(text[m.Index - 1]) >= 0 ||
                        (m.Groups["sign"].Value == "-" && char.IsDigit(text[m.Index - 1]))));
                if (adjacent || Ratio.IsMatch(tail) ||
                    (!hasUnit && !m.Groups["exp"].Success && !range &&
                    Detector.ClassifyNumber(text, n.Index, n.Index + n.Length, new NumCtx()) == "text" &&
                    !Coefficient.IsMatch(context)))
                    continue;

                token.Number = n.Value;
                token.Sign = m.Groups["sign"].Value;
                token.Suffix = m.Groups["exp"].Value;
                if (hasUnit)
                {
                    end += gap + unit.Len;
                    string qualifier = Qualifier.Match(text.Substring(end)).Value;
                    end += qualifier.Length;
                    token.Unit = unit.Raw + qualifier;
                    token.Length = end - token.Start;
                    token.Original = text.Substring(token.Start, token.Length);
                }
            }
            return result;
        }

        static ParsedNumber Parse(string number, char mark)
        {
            if (number[0] == '.' || number[0] == ',') number = "0" + number;
            return NumberUtil.Parse(number, mark);
        }

        static string Render(ParsedNumber value, char decimalMark, bool group)
        {
            string[] parts = value.Norm.Split('.');
            string integral = parts[0];
            if (group)
                for (int at = integral.Length - 3; at > 0; at -= 3) integral = integral.Insert(at, "\u202F");
            return parts.Length == 1 ? integral : integral + decimalMark + parts[1];
        }

        internal static bool TryNormalize(string source, string translation, string target, out string result)
        {
            result = translation;
            var before = Read(source);
            var after = Read(translation);
            if (before.Count != after.Count) return false;
            char mark = target == "vi" ? ',' : '.';
            string sourceLang = AiClient.DetectLang(source);
            var marks = before.Where(t => t.IsMeasurement && (t.Number.Contains(".") || t.Number.Contains(",")))
                .Select(t => Parse(t.Number, '\0')).Where(n => n != null && !n.Amb).Select(n => n.Dec).Distinct().ToArray();
            char sourceMark = '\0';
            if (marks.Length == 1) sourceMark = marks[0];
            if (marks.Length == 0) sourceMark = sourceLang == "vi" ? ',' : sourceLang == "en" ? '.' : '\0';
            for (int i = before.Count - 1; i >= 0; i--)
            {
                Token a = before[i], b = after[i];
                if (!a.IsMeasurement || !b.IsMeasurement)
                {
                    if (a.Original != b.Original || a.Unit != b.Unit) return false;
                    continue;
                }
                if (a.Sign != b.Sign || a.Suffix != b.Suffix || a.Unit != b.Unit) return false;
                ParsedNumber value = Parse(a.Number, sourceMark);
                if (value == null || value.Amb)
                {
                    if (a.Number != b.Number) return false;
                    continue;
                }
                bool group = value.Thou != '\0' || a.Number.IndexOfAny(new[] { ' ', '\u00A0', '\u202F' }) >= 0;
                string localized = Render(value, mark, group);
                if (b.Number != a.Number)
                {
                    ParsedNumber candidate = Parse(b.Number, mark);
                    if (candidate == null || candidate.Norm != value.Norm) return false;
                }
                string gap = a.Unit == null || a.Unit == "°" ? "" : " ";
                string rendered = a.Sign + localized + a.Suffix + gap + a.Unit;
                result = result.Remove(b.Start, b.Length).Insert(b.Start, rendered);
            }
            return true;
        }
    }
}
