using Miracle.Arguments;

namespace Json2Excel
{
	class Program
	{
		private static void Main(string[] args)
		{
			var converter = args.ParseCommandLine<JsonToExcelConverter>();
			if (converter != null)
			{
				converter.Convert();
			}
		}
	}
}
