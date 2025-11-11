using System;

namespace GroupDocs.Assembly.Examples.CSharp
{
    public static class Helper
    {
        public static void WriteError(string text)
        {
            ConsoleColor foregroundColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(text);
            Console.ForegroundColor = foregroundColor;
        }
    }
}


