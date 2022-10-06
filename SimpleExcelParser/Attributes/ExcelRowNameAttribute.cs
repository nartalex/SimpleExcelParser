namespace SimpleExcelParser.Attributes;

[AttributeUsage(AttributeTargets.Field)]
public class ExcelRowNameAttribute : Attribute
{
	public ExcelRowNameAttribute(string name)
	{
		ExcelRowName = name;
	}

	public string ExcelRowName { get; }
}
