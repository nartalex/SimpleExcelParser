using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using ExcelDataReader;
using SimpleExcelParser.Extensions;

namespace SimpleExcelParser;

public static class ExcelParser
{
	private static readonly HashSet<string> BoolTrueValues = new()
	{
		"y",
		"yes",
		"1"
	};

	public static IEnumerable<TData> Parse<TData>(
		string filePath,
		string sheetName,
		ExcelReaderConfiguration? excelReaderConfiguration = null,
		IFormatProvider? formatProvider = null)
		where TData : new()
	{
		Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

		using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
		using var reader = ExcelReaderFactory.CreateReader(stream, excelReaderConfiguration);

		foreach (var row in ParseInternal<TData>(reader, sheetName, formatProvider ?? CultureInfo.InvariantCulture))
		{
			yield return row;
		}
	}

	public static IEnumerable<TData> Parse<TData>(
		Stream stream,
		string sheetName,
		ExcelReaderConfiguration? excelReaderConfiguration = null,
		IFormatProvider? formatProvider = null)
		where TData : new()
	{
		Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

		using var reader = ExcelReaderFactory.CreateReader(stream, excelReaderConfiguration);

		foreach (var row in ParseInternal<TData>(reader, sheetName, formatProvider ?? CultureInfo.InvariantCulture))
		{
			yield return row;
		}
	}

	private static IEnumerable<TData> ParseInternal<TData>(IExcelDataReader reader, string sheetName, IFormatProvider formatProvider)
		where TData : new()
	{
		while (reader.Name != sheetName)
		{
			if (!reader.NextResult())
			{
				throw new Exception($"Sheet {sheetName} is not found");
			}
		}

		var headerRowExists = reader.Read();
		if (!headerRowExists)
		{
			throw new Exception("The file is empty");
		}

		var headerMap = MapHeader<TData>(reader);

		while (reader.Read())
		{
			if (!IsDataInRow(reader, headerMap.Values))
			{
				continue;
			}

			var parsedRow = ParseRow<TData>(reader, headerMap, formatProvider);
			yield return parsedRow;
		}
	}

	private static Dictionary<FieldInfo, int> MapHeader<TData>(IExcelDataReader dataReader)
	{
		var fields = typeof(TData).GetFields();
		var tdataRowNames = new Dictionary<string, FieldInfo>(fields.Length);

		foreach (var field in fields)
		{
			var rowName = field.GetRowName().PrepareString();

			if (!string.IsNullOrEmpty(rowName))
			{
				tdataRowNames.Add(rowName, field);
			}
		}

		var fieldsMap = new Dictionary<FieldInfo, int>(tdataRowNames.Count);

		for (int columnIndex = 0; columnIndex < dataReader.FieldCount; columnIndex++)
		{
			var header = dataReader.GetString(columnIndex).PrepareString();

			if (!string.IsNullOrEmpty(header) && tdataRowNames.ContainsKey(header))
			{
				fieldsMap.Add(tdataRowNames[header], columnIndex);
			}
		}

		Debug.Assert(tdataRowNames.Count == fieldsMap.Count);

		return fieldsMap;
	}

	private static bool IsDataInRow(IExcelDataReader dataReader, IEnumerable<int> actualColumns)
	{
		foreach (var column in actualColumns)
		{
			if (dataReader.GetValue(column) != null)
			{
				return true;
			}
		}

		return false;
	}

	private static TData ParseRow<TData>(IExcelDataReader dataReader, Dictionary<FieldInfo, int> fieldsMap, IFormatProvider formatProvider)
		where TData : new()
	{
		var data = new TData();

		foreach (var fieldKvp in fieldsMap)
		{
			var cellValue = dataReader.GetValue(fieldKvp.Value);

			if (cellValue == null)
			{
				continue;
			}

			var value = cellValue.GetType() == fieldKvp.Key.FieldType
				? cellValue
				: ParseValue(fieldKvp.Key.FieldType, cellValue.ToString(), formatProvider);

			fieldKvp.Key.SetValue(data, value);
		}

		return data;
	}

	private static object? ParseValue(Type fieldType, string? stringData, IFormatProvider formatProvider)
	{
		if (string.IsNullOrEmpty(stringData))
		{
			return default;
		}

		if (PrimitiveTypesParseMethods.ContainsKey(fieldType))
		{
			return PrimitiveTypesParseMethods[fieldType](stringData, formatProvider);
		}

		if (fieldType == typeof(bool) || fieldType == typeof(bool?))
		{
			return BoolTrueValues.Contains(stringData.ToLowerInvariant());
		}

		if (fieldType == typeof(char) || fieldType == typeof(char?))
		{
			return stringData[0];
		}

		if (fieldType == typeof(DateTime) || fieldType == typeof(DateTime?))
		{
			if (DateTime.TryParse(stringData, formatProvider, DateTimeStyles.None, out var dateTime))
			{
				return dateTime;
			}

			if (double.TryParse(stringData, out var oaDateTime))
			{
				return DateTime.FromOADate(oaDateTime);
			}

			throw new FormatException($"Value \"{stringData}\" is not recognized as a valid datetime");
		}

		if (fieldType == typeof(DateOnly) || fieldType == typeof(DateOnly?))
		{
			if (DateTime.TryParse(stringData, formatProvider, DateTimeStyles.None, out var dateTime))
			{
				return DateOnly.FromDateTime(dateTime);
			}

			if (double.TryParse(stringData, NumberStyles.None, formatProvider, out var oaDateTime))
			{
				return DateOnly.FromDateTime(DateTime.FromOADate(oaDateTime));
			}

			throw new FormatException($"Value \"{stringData}\" is not recognized as a valid datetime");
		}

		throw new ArgumentException($"Value {stringData} cannot be casted to {fieldType.Name}", nameof(stringData));
	}

	private static readonly Dictionary<Type, Func<string, IFormatProvider, object>> PrimitiveTypesParseMethods = new()
	{
		{ typeof(byte), (s, provider) => byte.Parse(s, provider) },
		{ typeof(byte?), (s, provider) => byte.Parse(s, provider) },
		{ typeof(sbyte), (s, provider) => sbyte.Parse(s, provider) },
		{ typeof(sbyte?), (s, provider) => sbyte.Parse(s, provider) },

		{ typeof(short), (s, provider) => short.Parse(s, provider) },
		{ typeof(short?), (s, provider) => short.Parse(s, provider) },
		{ typeof(ushort), (s, provider) => ushort.Parse(s, provider) },
		{ typeof(ushort?), (s, provider) => ushort.Parse(s, provider) },

		{ typeof(int), (s, provider) => int.Parse(s, provider) },
		{ typeof(int?), (s, provider) => int.Parse(s, provider) },
		{ typeof(uint), (s, provider) => uint.Parse(s, provider) },
		{ typeof(uint?), (s, provider) => uint.Parse(s, provider) },

		{ typeof(long), (s, provider) => long.Parse(s, provider) },
		{ typeof(long?), (s, provider) => long.Parse(s, provider) },
		{ typeof(ulong), (s, provider) => ulong.Parse(s, provider) },
		{ typeof(ulong?), (s, provider) => ulong.Parse(s, provider) },

		{ typeof(float), (s, provider) => float.Parse(s, provider) },
		{ typeof(float?), (s, provider) => float.Parse(s, provider) },
		{ typeof(double), (s, provider) => double.Parse(s, provider) },
		{ typeof(double?), (s, provider) => double.Parse(s, provider) },
		{ typeof(decimal), (s, provider) => decimal.Parse(s, provider) },
		{ typeof(decimal?), (s, provider) => decimal.Parse(s, provider) },

		{ typeof(TimeOnly), (s, provider) => TimeOnly.Parse(SplitBySpaceAndGetLast(s), provider) },
		{ typeof(TimeOnly?), (s, provider) => TimeOnly.Parse(SplitBySpaceAndGetLast(s), provider) },
		{ typeof(TimeSpan), (s, provider) => TimeSpan.Parse(SplitBySpaceAndGetLast(s), provider) },
		{ typeof(TimeSpan?), (s, provider) => TimeSpan.Parse(SplitBySpaceAndGetLast(s), provider) },
	};

	private static ReadOnlySpan<char> SplitBySpaceAndGetLast(string source)
	{
		for (int i = source.Length - 1; i >= 0; i--)
		{
			if (source[i] == ' ')
			{
				return source[(i + 1)..];
			}
		}

		return source;
	}
}
