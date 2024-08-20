using System;

namespace TranslateWithDeepl
{
  internal class Program
  {
    public static void Main(string[] args)
    {
      var translateWithDeepl = new TranslateWithDeepl();
      try
      {
        Console.WriteLine("Translation started...");
        translateWithDeepl.Execute(args[0], args[1]);
        Console.WriteLine("Translation finished.");
        Console.WriteLine("Translated file: " + args[1]);
      }
      catch (Exception e)
      {
        Console.WriteLine(e);
        Console.WriteLine("Usage: TranslateWithDeepL.exe <source file> <target file>");
        Console.WriteLine("Example: TranslateWithDeepL.exe C:\\source.xml C:\\target.xml");
        Console.WriteLine("");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
      }
    }
  }
}