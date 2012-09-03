using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace DocumentCommentParser.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = string.Format("{0}\\Template\\{1}", GetExecutionPath(), "Comments.docx");
            byte[] documentByte = File.ReadAllBytes(templatePath);

            DocumentParser parser = new DocumentParser();
            var comments = parser.GetCommentsFromDocument(documentByte);

            foreach (var comment in comments)
            {
                Console.WriteLine("----------------------------");
                Console.WriteLine(comment.CommentText);
                Console.WriteLine();
                Console.WriteLine(comment.CommentedText);
                Console.WriteLine("----------------------------");
            }
            Console.WriteLine("-- Ended --");
            Console.ReadLine();

        }

        private static string GetExecutionPath()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }
    }
}
