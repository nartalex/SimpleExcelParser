namespace SimpleExcelParser.Extensions;

internal static class StringExtensions
{
	public static string PrepareString(this string? s)
	{
		if (string.IsNullOrWhiteSpace(s))
		{
			return string.Empty;
		}

		return s.ToLowerInvariant()
			.Trim()
			.Replace(" ", string.Empty)
			.Replace("\\", string.Empty)
			.Replace("/", string.Empty)
			.Replace("_", string.Empty);
	}
}
