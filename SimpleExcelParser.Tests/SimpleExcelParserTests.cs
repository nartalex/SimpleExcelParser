using FluentAssertions;
using SimpleExcelParser.TestBase;

namespace SimpleExcelParser.Tests;

public class SimpleExcelParserTests
{
	private readonly string testDataFilePath = "test-data.xlsx";

	[Fact]
	public void ExcelParserTest()
	{
		var rowsFromFile = ExcelParser.Parse<TestDataRow>(testDataFilePath, "Sheet1").ToArray();
		rowsFromFile.Length.Should().Be(122);

		rowsFromFile.Take(2).Should().ContainInOrder(ExpectedData);
	}

	private static readonly TestDataRow[] ExpectedData = new TestDataRow[]
	{
		new TestDataRow
		{
			String = "str1",
			Bool = true,
			NullableBool = true,
			Byte = 1,
			NullableByte = 3,
			SByte = 4,
			NullableSByte = 6,
			Short = 7,
			NullableShort = 9,
			UShort = 10,
			NullableUShort = 12,
			Int = 13,
			NullableInt = 15,
			UInt = 16,
			NullableUInt = 18,
			Long = 19,
			NullableLong = 21,
			ULong = 22,
			NullableULong = 24,
			Char = 'a',
			NullableChar = 'c',
			Double = 25.5,
			NullableDouble = 27.7,
			Float = 28.8f,
			NullableFloat = 30.3f,
			Decimal = 31.1m,
			NullableDecimal = 33.3m,
			DateTime = new DateTime(2022, 01, 01, 07, 12, 00),
			NullableDateTime = new DateTime(2022, 01, 03, 07, 12, 00),
			DateOnly = new DateOnly(2022, 01, 01),
			NullableDateOnly = new DateOnly(2022, 01, 03),
			TimeOnly = new TimeOnly(11, 16, 00),
			NullableTimeOnly = new TimeOnly(13, 16, 00),
			TimeSpan = new TimeSpan(15, 16, 00),
			NullableTimeSpan = new TimeSpan(17, 16, 00)
		},
		new TestDataRow
		{
			String = "str2",
			Bool = false,
			NullableBool = null,
			Byte = 2,
			NullableByte = null,
			SByte = 5,
			NullableSByte = null,
			Short = 8,
			NullableShort = null,
			UShort = 11,
			NullableUShort = null,
			Int = 14,
			NullableInt = null,
			UInt = 17,
			NullableUInt = null,
			Long = 20,
			NullableLong = null,
			ULong = 23,
			NullableULong = null,
			Char = 'b',
			NullableChar = null,
			Double = 26.6,
			NullableDouble = null,
			Float = 29.9f,
			NullableFloat = null,
			Decimal = 32.2m,
			NullableDecimal = null,
			DateTime = new DateTime(2022, 01, 02, 07, 12, 00),
			NullableDateTime = null,
			DateOnly = new DateOnly(2022, 01, 02),
			NullableDateOnly = null,
			TimeOnly = new TimeOnly(12, 16, 00),
			NullableTimeOnly = null,
			TimeSpan = new TimeSpan(12, 16, 00),
			NullableTimeSpan = null
		},
	};
}