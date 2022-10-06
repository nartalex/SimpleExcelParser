using SimpleExcelParser.Attributes;

namespace SimpleExcelParser.TestBase;

public record TestDataRow
{
	[ExcelRowName("String")]
	public string String;

	[ExcelRowName("Bool")]
	public bool Bool;

	[ExcelRowName("NullableBool")]
	public bool? NullableBool;

	[ExcelRowName("Byte")]
	public byte Byte;

	[ExcelRowName("NullableByte")]
	public byte? NullableByte;

	[ExcelRowName("SByte")]
	public sbyte SByte;

	[ExcelRowName("NullableSByte")]
	public sbyte? NullableSByte;

	[ExcelRowName("Short")]
	public short Short;

	[ExcelRowName("NullableShort")]
	public short? NullableShort;

	[ExcelRowName("UShort")]
	public ushort UShort;

	[ExcelRowName("NullableUShort")]
	public ushort? NullableUShort;

	[ExcelRowName("Int")]
	public int Int;

	[ExcelRowName("NullableInt")]
	public int? NullableInt;

	[ExcelRowName("UInt")]
	public uint UInt;

	[ExcelRowName("NullableUInt")]
	public uint? NullableUInt;

	[ExcelRowName("Long")]
	public long Long;

	[ExcelRowName("NullableLong")]
	public long? NullableLong;

	[ExcelRowName("ULong")]
	public ulong ULong;

	[ExcelRowName("NullableULong")]
	public ulong? NullableULong;

	[ExcelRowName("Char")]
	public char Char;

	[ExcelRowName("NullableChar")]
	public char? NullableChar;

	[ExcelRowName("Float")]
	public float Float;

	[ExcelRowName("NullableFloat")]
	public float? NullableFloat;

	[ExcelRowName("Double")]
	public double Double;

	[ExcelRowName("NullableDouble")]
	public double? NullableDouble;

	[ExcelRowName("Decimal")]
	public decimal Decimal;

	[ExcelRowName("NullableDecimal")]
	public decimal? NullableDecimal;

	[ExcelRowName("DateTime")]
	public DateTime DateTime;

	[ExcelRowName("NullableDateTime")]
	public DateTime? NullableDateTime;

	[ExcelRowName("DateOnly")]
	public DateOnly DateOnly;

	[ExcelRowName("NullableDateOnly")]
	public DateOnly? NullableDateOnly;

	[ExcelRowName("TimeOnly")]
	public TimeOnly TimeOnly;

	[ExcelRowName("NullableTimeOnly")]
	public TimeOnly? NullableTimeOnly;

	[ExcelRowName("TimeSpan")]
	public TimeSpan TimeSpan;

	[ExcelRowName("NullableTimeSpan")]
	public TimeSpan? NullableTimeSpan;
}