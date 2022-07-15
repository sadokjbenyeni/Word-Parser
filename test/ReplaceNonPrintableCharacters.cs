using System.Text.RegularExpressions;

namespace test
{
    class ReplaceCharacters
    {
        //public static string ReplaceNonPrintableCharacters(string s)
        //{
        //    StringBuilder result = new StringBuilder();
        //    for (int i = 0; i < s.Length; i++)
        //    {
        //        char c = s[i];
        //        byte b = (byte)c;
        //        if (b < 32)
        //            result.Append("");
        //        else
        //            result.Append(c);
        //    }
        //    return result.ToString();
        //}

        public static string ReplaceNonPrintableCharacters(string value)
        {
            string pattern = "[^ -~]+";
            Regex reg_exp = new Regex(pattern);
            return reg_exp.Replace(value, "");
        }
    }
}
