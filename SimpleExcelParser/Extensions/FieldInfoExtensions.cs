using System.Reflection;
using SimpleExcelParser.Attributes;

namespace SimpleExcelParser.Extensions;

internal static class FieldInfoExtensions
{
	public static string? GetRowName(this FieldInfo obj)
	{
		var attr = obj.CustomAttributes.FirstOrDefault(y => y.AttributeType == typeof(ExcelRowNameAttribute));

		if(attr == null)
		{
			return null;
		}

		return attr.ConstructorArguments[0].Value!.ToString()!.PrepareString();
	}
}
