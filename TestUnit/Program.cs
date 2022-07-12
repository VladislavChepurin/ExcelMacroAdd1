using BenchmarkDotNet.Running;

namespace TestUnit
{
    internal class Program
    {
        static void Main(string[] args)
        {
  
            var summary = BenchmarkRunner.Run<TestBoxShield>();

        }
    }
}
