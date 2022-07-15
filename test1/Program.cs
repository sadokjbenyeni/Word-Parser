


namespace test1
{
    using Microsoft.Office.Interop.Word;
    using System;
    using System.Diagnostics;
    using System.Reflection;
    using System.Runtime.InteropServices.ComTypes;

    class Program
    {
        static void Main(string[] args)
        {

            Application word = new Application();
            Microsoft.Office.Interop.Word.Document doc = new Document();

            object nullobj = Missing.Value;

            doc = word.Documents.Open(@"C:\Users\sadok.jbenyeni\Desktop\word.docx", ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);

            word.Visible = false;

            

        }
    }
}