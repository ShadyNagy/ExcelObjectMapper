using System.Text;

namespace ExcelObjectMapper.Extensions
{
	internal static class StringExtensions
	{
		internal static string RemoveSpecialCharacters(this string str)
		{
			StringBuilder sb = new StringBuilder();
			foreach (char c in str)
			{
				if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
				{
					sb.Append(c);
				}
			}
			return sb.ToString();
		}

		internal static string RemoveTabAndEnter(this string str)
		{
			return str.Replace("\n", string.Empty)
				.Replace("\r", string.Empty)
				.Replace("\t", string.Empty)
				.Replace("\v", string.Empty)
				.Replace("\f", string.Empty);
		}
	}
}
