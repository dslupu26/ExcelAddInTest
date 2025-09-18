using System;
using System.Collections.Generic;

namespace ExcelAddInTest.Infrastructure.Text
{
    /// <summary>
    /// Central place for all synonym maps used by speech normalization.
    /// All keys are matched case-insensitively; all values should be UPPERCASE for cell letters.
    /// </summary>
    public static class SpeechLexicon
    {
        // 1) Verb / command synonyms → canonical verbs your router understands
        public static readonly IReadOnlyDictionary<string, string> VerbSynonyms =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // selection / navigation
                ["select"] = "select",
                ["highlight"] = "select",
                ["choose"] = "select",
                ["go to"] = "goto",
                ["open"] = "goto",

                // input / paste
                ["write"] = "write",
                ["type"] = "write",
                ["insert"] = "write",
                ["fill"] = "write",
                ["put"] = "write",
                ["paste"] = "paste",

                // math
                ["plus"] = "add",
                ["sum"] = "add",
                ["total"] = "add",
                ["subtract"] = "subtract",
                ["minus"] = "subtract",
                ["multiply"] = "multiply",
                ["times"] = "multiply",
                ["divide"] = "divide",
                ["over"] = "divide",

                // format
                ["bold"] = "bold",
                ["make bold"] = "bold",
                ["italic"] = "italic",
                ["italics"] = "italic",
                ["underline"] = "underline",

                // rows/cols
                ["insert row"] = "insert_row",
                ["add row"] = "insert_row",
                ["delete row"] = "delete_row",
                ["remove row"] = "delete_row",
                ["insert column"] = "insert_col",
                ["add column"] = "insert_col",
                ["delete column"] = "delete_col",
                ["remove column"] = "delete_col",

                // sort/filter
                ["sort"] = "sort",
                ["order"] = "sort",
                ["ascending"] = "asc",
                ["descending"] = "desc",
                ["filter"] = "filter",
                ["show only"] = "filter",
                ["hide others"] = "filter",

                // misc
                ["merge"] = "merge",
                ["unmerge"] = "unmerge",
                ["freeze"] = "freeze",
                ["unfreeze"] = "unfreeze",
                ["clear"] = "clear",
                ["delete"] = "clear"
            };

        // 2) Common letter homophones → LETTER
        //    (helps turn "kay three" into "K3", "sea ten" -> "C10")
        public static readonly IReadOnlyDictionary<string, string> LetterHomophones =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["a"] = "A",
                ["ay"] = "A",
                ["b"] = "B",
                ["bee"] = "B",
                ["c"] = "C",
                ["see"] = "C",
                ["sea"] = "C",
                ["d"] = "D",
                ["dee"] = "D",
                ["e"] = "E",
                ["f"] = "F",
                ["ef"] = "F",
                ["g"] = "G",
                ["gee"] = "G",
                ["h"] = "H",
                ["aitch"] = "H",
                ["hache"] = "H",
                ["i"] = "I",
                ["eye"] = "I",
                ["j"] = "J",
                ["jay"] = "J",
                ["k"] = "K",
                ["kay"] = "K",
                ["l"] = "L",
                ["el"] = "L",
                ["m"] = "M",
                ["em"] = "M",
                ["n"] = "N",
                ["en"] = "N",
                ["o"] = "O",
                ["oh"] = "O",
                ["p"] = "P",
                ["pee"] = "P",
                ["q"] = "Q",
                ["cue"] = "Q",
                ["queue"] = "Q",
                ["r"] = "R",
                ["ar"] = "R",
                ["s"] = "S",
                ["ess"] = "S",
                ["t"] = "T",
                ["tee"] = "T",
                ["u"] = "U",
                ["you"] = "U",
                ["yu"] = "U",
                ["v"] = "V",
                ["vee"] = "V",
                ["w"] = "W",
                ["double u"] = "W",
                ["double you"] = "W",
                ["x"] = "X",
                ["ex"] = "X",
                ["y"] = "Y",
                ["why"] = "Y",
                ["z"] = "Z",
                ["zee"] = "Z",
                ["zed"] = "Z"
            };

        // 3) Number words → digits (basic English up to 30 is enough for most cell rows in examples)
        public static readonly IReadOnlyDictionary<string, int> NumberWords =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["zero"] = 0,
                ["one"] = 1,
                ["two"] = 2,
                ["three"] = 3,
                ["four"] = 4,
                ["five"] = 5,
                ["six"] = 6,
                ["seven"] = 7,
                ["eight"] = 8,
                ["nine"] = 9,
                ["ten"] = 10,
                ["eleven"] = 11,
                ["twelve"] = 12,
                ["thirteen"] = 13,
                ["fourteen"] = 14,
                ["fifteen"] = 15,
                ["sixteen"] = 16,
                ["seventeen"] = 17,
                ["eighteen"] = 18,
                ["nineteen"] = 19,
                ["twenty"] = 20,
                ["twenty one"] = 21,
                ["twenty-two"] = 22,
                ["twenty three"] = 23,
                ["twenty four"] = 24,
                ["twenty five"] = 25,
                ["twenty six"] = 26,
                ["twenty seven"] = 27,
                ["twenty eight"] = 28,
                ["twenty nine"] = 29,
                ["thirty"] = 30
            };
    }
}
