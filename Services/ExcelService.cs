using ClosedXML.Excel;

namespace ExcelCopyTool.Services;

internal sealed class ExcelService
{
    private const string SourceSheetName = "sheet_B";
    private const string TargetSheetName = "sheet_C";
    private static readonly Action<CopyContext>[] RowCopySteps =
    [
        CopyColumnIToCellD6,
    ];

    public bool TryCreateWorkbookFromTemplate(
        string templatePath,
        string sourcePath,
        string id,
        string outputDirectory,
        out string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);

        using var sourceWorkbook = new XLWorkbook(sourcePath);
        var sourceWorksheet = sourceWorkbook.Worksheet(SourceSheetName)
            ?? throw new InvalidOperationException($"コピー元シート '{SourceSheetName}' が見つかりません。");

        var foundSourceRow = TryFindSourceRow(sourceWorksheet, id, out var sourceRow);
        outputPath = Path.Combine(outputDirectory, $"{BuildOutputFileName(id)}.xlsx");

        if (!foundSourceRow || sourceRow is null)
        {
            return false;
        }

        if (File.Exists(outputPath))
        {
            throw new InvalidOperationException($"出力ファイルは既に存在します: {outputPath}");
        }

        File.Copy(templatePath, outputPath);

        try
        {
            using var targetWorkbook = new XLWorkbook(outputPath);
            var targetWorksheet = targetWorkbook.Worksheet(TargetSheetName)
                ?? throw new InvalidOperationException($"出力先シート '{TargetSheetName}' が見つかりません。");

            WriteId(targetWorksheet, id);
            ApplyRowCopySteps(new CopyContext(id, sourceWorksheet, sourceRow, targetWorksheet));
            targetWorkbook.Save();

            return true;
        }
        catch
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            throw;
        }
    }

    private static void ApplyRowCopySteps(CopyContext context)
    {
        foreach (var rowCopyStep in RowCopySteps)
        {
            rowCopyStep(context);
        }
    }

    private static bool TryFindSourceRow(IXLWorksheet worksheet, string id, out IXLRow? sourceRow)
    {
        foreach (var row in worksheet.RowsUsed())
        {
            if (string.Equals(row.Cell("B").GetString().Trim(), id, StringComparison.Ordinal))
            {
                sourceRow = row;
                return true;
            }
        }

        sourceRow = null;
        return false;
    }

    private static void WriteId(IXLWorksheet targetWorksheet, string id)
    {
        targetWorksheet.Cell("D1").Value = id;
    }

    private static void CopyColumnIToCellD6(CopyContext context)
    {
        var sourceValue = GetColumnIValue(context);
        WriteValueToD6(context, sourceValue);
    }

    private static XLCellValue GetColumnIValue(CopyContext context)
    {
        return context.SourceRow.Cell("I").Value;
    }

    private static void WriteValueToD6(CopyContext context, XLCellValue sourceValue)
    {
        context.TargetWorksheet.Cell("D6").Value = sourceValue;
    }

    private static string BuildOutputFileName(string id)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        if (id.IndexOfAny(invalidCharacters) >= 0)
        {
            throw new InvalidOperationException("ID にファイル名として使用できない文字が含まれています。");
        }

        return id;
    }

    private sealed record CopyContext(string Id, IXLWorksheet SourceWorksheet, IXLRow SourceRow, IXLWorksheet TargetWorksheet);
}