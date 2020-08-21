using System;
using OWAUtility;

class Program
{
    static void Main(string[] args)
    {
        // The Length property provides the number of array elements.
        Console.WriteLine($"parameter count = {args.Length}");
        OWA owa = new OWA();
        owa.Start(args);

        if (System.Diagnostics.Debugger.IsAttached)
        {
            Console.WriteLine("Hit any key to exit...");
            Console.ReadKey();
        }
    }
}