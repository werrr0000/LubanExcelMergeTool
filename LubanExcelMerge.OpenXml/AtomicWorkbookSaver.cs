namespace LubanExcelMerge.OpenXml;

public sealed record WorkbookSaveResult(
    string OutputPath,
    int SheetCount,
    int FormulaCount,
    WorkbookRecalculationStatus RecalculationStatus,
    string? RecalculationProvider,
    bool ProjectValidationCompleted = false,
    bool FullExportValidationCompleted = false);

public sealed class AtomicWorkbookSaver
{
    private readonly OpenXmlWorkbookReader _reader;
    private readonly OpenXmlWorkbookEditor _editor;
    private readonly IWorkbookRecalculator _recalculator;

    public AtomicWorkbookSaver(
        OpenXmlWorkbookReader? reader = null,
        OpenXmlWorkbookEditor? editor = null,
        IWorkbookRecalculator? recalculator = null)
    {
        _reader = reader ?? new OpenXmlWorkbookReader();
        _editor = editor ?? new OpenXmlWorkbookEditor();
        _recalculator = recalculator ?? CompositeWorkbookRecalculator.CreateDefault();
    }

    public WorkbookSaveResult Save(
        string localPath,
        string outputPath,
        IReadOnlyList<WorkbookEdit> edits,
        WorkbookSaveOptions? options = null)
    {
        options ??= new WorkbookSaveOptions();
        var absoluteLocal = Path.GetFullPath(localPath);
        var absoluteOutput = Path.GetFullPath(outputPath);
        if (string.Equals(absoluteLocal, absoluteOutput, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("输出路径不能与 LOCAL 源文件相同。");

        var outputDirectory = Path.GetDirectoryName(absoluteOutput)
            ?? throw new InvalidOperationException("输出路径缺少目录。");
        Directory.CreateDirectory(outputDirectory);
        var temporaryPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileName(absoluteOutput)}.{Guid.NewGuid():N}.tmp.xlsx");
        var shouldRecalculate = options.RecalculationMode == WorkbookRecalculationMode.Always ||
            options.RecalculationMode == WorkbookRecalculationMode.Auto && options.FormulaMayBeAffected;
        if (shouldRecalculate && !_recalculator.IsAvailable)
            throw new WorkbookRecalculationUnavailableException(
                $"重算模式为 {options.RecalculationMode.ToString().ToLowerInvariant()}，但 {_recalculator.ProviderName} 的自动化接口不可用。");

        var sourceSnapshot = _reader.Read(absoluteLocal);
        var sourceHashes = PackageIntegrity.HashParts(absoluteLocal);
        try
        {
            File.Copy(absoluteLocal, temporaryPath, overwrite: false);
            var editResult = edits.Count == 0
                ? new WorkbookEditResult(new HashSet<string>(StringComparer.Ordinal))
                : _editor.Apply(temporaryPath, edits);
            var touchedParts = editResult.TouchedPartPaths.ToHashSet(StringComparer.Ordinal);
            if (options.FormulaMayBeAffected)
            {
                _editor.MarkForFullCalculation(temporaryPath);
                touchedParts.Add("xl/workbook.xml");
            }

            var candidateSnapshot = _reader.Read(temporaryPath);
            ValidateCandidate(sourceSnapshot, candidateSnapshot, edits, preserveFormulaCaches: true);
            ValidateUntouchedParts(sourceHashes, PackageIntegrity.HashParts(temporaryPath), touchedParts);

            var recalculationStatus = options.FormulaMayBeAffected &&
                                      options.RecalculationMode == WorkbookRecalculationMode.Never
                ? WorkbookRecalculationStatus.SourceCachePreservedUnverified
                : WorkbookRecalculationStatus.NotNeeded;
            if (shouldRecalculate)
            {
                _recalculator.Recalculate(temporaryPath, options.EffectiveTimeout);
                var recalculatedSnapshot = _reader.Read(temporaryPath);
                ValidateRecalculatedCandidate(candidateSnapshot, recalculatedSnapshot);
                candidateSnapshot = recalculatedSnapshot;
                recalculationStatus = WorkbookRecalculationStatus.Completed;
            }

            ReplaceOutput(temporaryPath, absoluteOutput);

            return new WorkbookSaveResult(
                absoluteOutput,
                candidateSnapshot.Sheets.Count,
                candidateSnapshot.Sheets.Sum(sheet => sheet.FormulaCount),
                recalculationStatus,
                recalculationStatus == WorkbookRecalculationStatus.Completed ? _recalculator.ProviderName : null);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }
    }

    private static void ValidateCandidate(
        WorkbookSnapshot source,
        WorkbookSnapshot candidate,
        IReadOnlyList<WorkbookEdit> edits,
        bool preserveFormulaCaches)
    {
        var sourceSheetNames = source.Sheets.Select(sheet => sheet.Name).ToArray();
        var candidateSheetNames = candidate.Sheets.Select(sheet => sheet.Name).ToArray();
        if (!sourceSheetNames.SequenceEqual(candidateSheetNames, StringComparer.Ordinal))
            throw new InvalidDataException("保存验证失败：工作表名称或顺序发生意外变化。");

        foreach (var sourceSheet in source.Sheets)
        {
            var candidateSheet = candidate.GetSheet(sourceSheet.Name);
            var setCells = edits.OfType<SetCellEdit>()
                .Where(edit => edit.SheetName == sourceSheet.Name)
                .Select(edit => edit.Address)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var deletedRows = edits.OfType<DeleteRowEdit>()
                .Where(edit => edit.SheetName == sourceSheet.Name)
                .Select(edit => edit.RowNumber)
                .ToHashSet();

            foreach (var sourceFormula in sourceSheet.Rows.SelectMany(row => row.Cells)
                         .Where(cell => cell.Payload.Kind == Core.CellValueKind.Formula))
            {
                if (setCells.Contains(sourceFormula.Address) || deletedRows.Contains(sourceFormula.RowNumber))
                    continue;

                var candidateFormula = candidateSheet.GetCell(sourceFormula.Address);
                if (candidateFormula?.Payload.Kind != Core.CellValueKind.Formula ||
                    candidateFormula.Payload.FormulaText != sourceFormula.Payload.FormulaText ||
                    preserveFormulaCaches && candidateFormula.Payload.CachedValue != sourceFormula.Payload.CachedValue)
                    throw new InvalidDataException($"保存验证失败：未修改公式 {sourceSheet.Name}!{sourceFormula.Address} 或其缓存发生变化。");
            }
        }
    }

    private static void ValidateRecalculatedCandidate(
        WorkbookSnapshot beforeRecalculation,
        WorkbookSnapshot afterRecalculation)
    {
        var beforeSheetNames = beforeRecalculation.Sheets.Select(sheet => sheet.Name).ToArray();
        var afterSheetNames = afterRecalculation.Sheets.Select(sheet => sheet.Name).ToArray();
        if (!beforeSheetNames.SequenceEqual(afterSheetNames, StringComparer.Ordinal))
            throw new InvalidDataException("Excel 重算后工作表名称或顺序发生变化。");

        foreach (var beforeSheet in beforeRecalculation.Sheets)
        {
            var afterSheet = afterRecalculation.GetSheet(beforeSheet.Name);
            var beforeFormulas = beforeSheet.Rows.SelectMany(row => row.Cells)
                .Where(cell => cell.Payload.Kind == Core.CellValueKind.Formula)
                .ToArray();
            if (beforeFormulas.Length != afterSheet.FormulaCount)
                throw new InvalidDataException($"Excel 重算后工作表 {beforeSheet.Name} 的公式数量发生变化。");

            foreach (var beforeFormula in beforeFormulas)
            {
                var afterFormula = afterSheet.GetCell(beforeFormula.Address);
                if (afterFormula?.Payload.Kind != Core.CellValueKind.Formula ||
                    afterFormula.Payload.FormulaText != beforeFormula.Payload.FormulaText)
                    throw new InvalidDataException(
                        $"Excel 重算后公式 {beforeSheet.Name}!{beforeFormula.Address} 发生意外变化。");
            }
        }

        var allowedDerivedPart = "xl/calcChain.xml";
        var missingParts = beforeRecalculation.PackagePartNames
            .Except(new[] { allowedDerivedPart }, StringComparer.OrdinalIgnoreCase)
            .Except(afterRecalculation.PackagePartNames, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (missingParts.Length > 0)
            throw new InvalidDataException("Excel 重算后包部件丢失：" + string.Join(", ", missingParts));
    }

    private static void ValidateUntouchedParts(
        IReadOnlyDictionary<string, string> sourceHashes,
        IReadOnlyDictionary<string, string> candidateHashes,
        IReadOnlySet<string> touchedParts)
    {
        if (!sourceHashes.Keys.Order(StringComparer.Ordinal).SequenceEqual(candidateHashes.Keys.Order(StringComparer.Ordinal)))
            throw new InvalidDataException("保存验证失败：OpenXML 包部件集合发生意外变化。");

        foreach (var sourcePart in sourceHashes)
        {
            if (!touchedParts.Contains(sourcePart.Key) && candidateHashes[sourcePart.Key] != sourcePart.Value)
                throw new InvalidDataException($"保存验证失败：未修改包部件 {sourcePart.Key} 发生变化。");
        }
    }

    private static void ReplaceOutput(string temporaryPath, string outputPath)
    {
        if (File.Exists(outputPath))
            File.Replace(temporaryPath, outputPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
        else
            File.Move(temporaryPath, outputPath);
    }
}
