using System;
using System.IO;
using System.Security.Cryptography;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "hash")
            {
                using (var stream = File.OpenRead(args[1]))
                using (var sha256 = SHA256.Create())
                {
                    Console.WriteLine(BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", ""));
                }
                return 0;
            }

            if (args.Length == 3 && args[0] == "replace")
            {
                if (File.Exists(args[2]))
                    File.Replace(args[1], args[2], null);
                else
                    File.Move(args[1], args[2]);
                return 0;
            }

            Console.Error.WriteLine("Usage: AvdManualFileTool.exe hash <file> | replace <temporary-file> <checkpoint>");
            return 2;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine("Manual file operation failed: " + exception.Message);
            return 1;
        }
    }
}
