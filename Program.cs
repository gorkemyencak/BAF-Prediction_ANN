using System;
using BafOptimization;

namespace Runner
{
    class Program
    {
        static void Main(string[] args)
        {
            var process = new Optimize();

            process.ConstructModel();
            
            var result = process.Run();

            Console.WriteLine($"Result: {result}");
        }
    }
}