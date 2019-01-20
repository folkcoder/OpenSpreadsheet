namespace Benchmarks
{
    using BenchmarkDotNet.Running;
    using Benchmarks.Tests;

    public static class Program
    {
        private static void Main()
        {
            var read = BenchmarkRunner.Run<BenchmarkRead>();
            var write = BenchmarkRunner.Run<BenchmarkWrite>();
        }
    }
}