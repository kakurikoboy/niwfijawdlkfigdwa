using ExcelCopyTool.Services;

var applicationDirectory = AppContext.BaseDirectory;
var templatePath = Path.Combine(applicationDirectory, "sample_A.xlsx");
var sourcePath = Path.Combine(applicationDirectory, "sample_B.xlsx");

Console.Write("対象データのIDを入力してください: ");
var inputId = Console.ReadLine()?.Trim();

if (string.IsNullOrWhiteSpace(inputId))
{
	Console.Error.WriteLine("ID が入力されていません。");
	return 1;
}

if (!File.Exists(templatePath))
{
	Console.Error.WriteLine($"テンプレートファイルが見つかりません: {templatePath}");
	return 1;
}

if (!File.Exists(sourcePath))
{
	Console.Error.WriteLine($"コピー元ファイルが見つかりません: {sourcePath}");
	return 1;
}

try
{
	var excelService = new ExcelService();
	var copiedValue = excelService.TryCreateWorkbookFromTemplate(templatePath, sourcePath, inputId, applicationDirectory, out var outputPath);

	if (!copiedValue)
	{
		Console.Error.WriteLine($"ID '{inputId}' は sample_B.xlsx の sheet_B の B列に見つかりませんでした。");
		return 1;
	}

	Console.WriteLine($"ファイルを作成しました: {outputPath}");
	return 0;
}
catch (Exception exception)
{
	Console.Error.WriteLine($"処理に失敗しました: {exception.Message}");
	return 1;
}
