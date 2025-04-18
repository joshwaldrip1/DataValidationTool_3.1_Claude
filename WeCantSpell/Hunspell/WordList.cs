using System.Collections.Generic;

namespace WeCantSpell.Hunspell
{
    public class WordList
    {
        public List<string> DictionaryWords { get; set; }

        public WordList(string affFile, string dicFile)
        {
            // Dummy implementation: in a real scenario, load and parse the dictionary files.
            DictionaryWords = new List<string> { "option1", "option2", "x-ray" };
        }

        public bool Check(string word)
        {
            return DictionaryWords.Contains(word.ToLower());
        }
    }
}
