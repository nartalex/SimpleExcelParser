using BenchmarkDotNet.Running;
using SimpleExcelParser.Benchmark;

#if DEBUG

var benchmarkClass = new BenchmarkClass();
benchmarkClass.GlobalSetup();
benchmarkClass.ParseRows();
benchmarkClass.GlobalCleanup();

#else

BenchmarkRunner.Run<BenchmarkClass>();

#endif