using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddInTest.Infrastructure.Text
{
    /// <summary>
    /// Applies ordered normalization steps to speech text so it’s friendlier for Excel/CLU.
    /// </summary>
    public static class TextNormalizer
    {
        // Reuse compiled regexes
        private static readonly Regex RxCell = new Regex(@"\b[A-Z]{1,3}[0-9]{1,4}\b", RegexOptions.Compiled);
        private static readonly Regex RxWord = new Regex(@"\b[\p{L}\-']+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxSpaces = new Regex(@"\s{2,}", RegexOptions.Compiled);
        private static readonly Regex RxLetterWord = new Regex(@"\b([a-z]{1,3})(?:\s+)(\d{1,4})\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase); // e.g. "kay 3" or "sea 10"

        /// <summary>
        /// Main entry point. Safe to call before sending to CLU or your Excel router.
        /// </summary>
        public static string Normalize(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // 1) Trim + collapse spaces
            text = text.Trim();
            text = RxSpaces.Replace(text, " ");

            // 2) Replace simple command synonyms (whole words/phrases)
            text = ReplaceVerbSynonyms(text);

            // 3) Convert homophone letters before numbers (A3, C10…)
            text = ReplaceLetterHomophonesBeforeNumbers(text);

            // 4) Convert number words when they follow a column (A three -> A3, AB ten -> AB10)
            text = ReplaceNumberWordsAfterLetters(text);
            // 5) Uppercase column letters in any cell-like token that slipped through
            text = UppercaseCellColumns(text);
            // 6) we ball
            text = ExcelCellRx(text);

            return text;
        }

        private static string ReplaceVerbSynonyms(string text)
        {
            // Replace longer keys first (e.g., "make bold" before "bold")
            var ordered = SpeechLexicon.VerbSynonyms
                .OrderByDescending(kv => kv.Key.Length);

            foreach (var kv in ordered)
            {
                // \b … \b is not great for multi-word phrases; use simple replace on boundaries.
                var pattern = @"(?<!\w)" + Regex.Escape(kv.Key) + @"(?!\w)";
                text = Regex.Replace(text, pattern, kv.Value, RegexOptions.IgnoreCase);
            }
            return text;
        }

        private static string ReplaceLetterHomophonesBeforeNumbers(string text)
        {
            // Map each word (or 1–3 words) that are letter homophones followed by digits
            // e.g., "kay 3" -> "K3", "sea ten" handled in next step
            return RxLetterWord.Replace(text, m =>
            {
                var lettersWord = m.Groups[1].Value;   // one to three letters as words
                var numberPart = m.Groups[2].Value;

                // Try to map a single word like "kay" -> "K"
                var mapped = MapLetterWord(lettersWord);
                return mapped != null ? mapped + numberPart : m.Value;
            });
        }

        private static string MapLetterWord(string word)
        {
            // Split "ab" cases: we accept things like "ab" or "cee kay" if you extend logic.
            // For now, handle single "letter words" only.
            if (SpeechLexicon.LetterHomophones.TryGetValue(word, out var single))
                return single;

            // If someone said "AB" as a single token, just uppercase it
            if (word.Length <= 3 && word.All(char.IsLetter))
                return word.ToUpperInvariant();

            return null;
        }

        private static string ReplaceNumberWordsAfterLetters(string text)
        {
            // Replace patterns like "A three" -> "A3", "AB ten" -> "AB10"
            var rx = new Regex(@"\b([A-Za-z]{1,3})\s+(zero|one|two|three|four|five|six|seven|eight|nine|ten|"
                             + @"eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|"
                             + @"twenty(?:[-\s]one|[-\s]two|[-\s]three|[-\s]four|[-\s]five|[-\s]six|[-\s]seven|[-\s]eight|[-\s]nine)?|thirty)\b",
                               RegexOptions.IgnoreCase | RegexOptions.Compiled);

            return rx.Replace(text, m =>
            {
                var col = m.Groups[1].Value.ToUpperInvariant();
                var word = m.Groups[2].Value.ToLowerInvariant().Replace('-', ' ');

                if (SpeechLexicon.NumberWords.TryGetValue(word, out var n))
                    return col + n.ToString(CultureInfo.InvariantCulture);

                return m.Value;
            });
        }

        private static string UppercaseCellColumns(string text)
        {
            // Ensure column letters are uppercase in detected cell references
            return RxCell.Replace(text, m =>
            {
                var token = m.Value;
                int i = 0;
                while (i < token.Length && char.IsLetter(token[i])) i++;
                var col = token.Substring(0, i).ToUpperInvariant();
                var row = token.Substring(i);
                return col + row;
            });
        }
        /// <summary>
        /// Add a space in between an enumeration of excel cells.
        /// (Ensure there's a space after each cell reference so CLU can parse them as separate entities.)
        /// e.g. "A1B2" -> "A1 B2"
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static string ExcelCellRx(string text)
        {
            Regex Rx = new Regex(@"[A-Z]{1,3}[0-9]{1,3}");
            text = Rx.Replace(text, m => m.Value + " ");
            //sterpelita de la david din TextNormalizer.cs 
            //text = Rx.Replace(text, " ");*/
            return text;
        }

        public static IList<string> ExcelCellRegexParser(string text) 
        {
            IList<string> list = new List<string>();
            Regex Rx = new Regex(@"[A-Z]{1,3}[0-9]{1,3}");
            var result = Rx.Matches(text);
            StringBuilder sb = new StringBuilder();
            foreach (var match in result)
            {
                list.Add(match.ToString());
            }
            return list;
        }
    }
}
