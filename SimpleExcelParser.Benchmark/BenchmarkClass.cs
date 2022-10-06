using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ExcelDataReader;
using SimpleExcelParser.TestBase;

namespace SimpleExcelParser.Benchmark;

[MemoryDiagnoser]
public class BenchmarkClass
{
	private static readonly ExcelReaderConfiguration ExcelReaderConfiguration = new()
	{
		LeaveOpen = true
	};

	private const string testDataFilePath = "test-data.xlsx";
	private MemoryStream? memoryStream;
	private Stream? fileStream;

	private readonly Consumer consumer = new();

	[GlobalSetup]
	public void GlobalSetup()
	{
		memoryStream = new();
		fileStream = File.OpenRead(testDataFilePath);

		fileStream.CopyTo(memoryStream);

		Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
	}

	[GlobalCleanup]
	public void GlobalCleanup()
	{
		memoryStream?.Dispose();
		fileStream?.Dispose();
	}

	[Benchmark]
	public void ParseRowsByExcelDataReader()
	{
		memoryStream!.Seek(0, SeekOrigin.Begin);

		var rows = ExcelParser.Parse<TestDataRow>(memoryStream, "Sheet1", ExcelReaderConfiguration);

		rows.Consume(consumer);
	}
}