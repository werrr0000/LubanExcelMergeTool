namespace LubanExcelMerge.Cli;

public static class CliApplication
{
    public static int Run(
        IReadOnlyList<string> args,
        TextWriter output,
        TextWriter error,
        LubanMergeCoordinator? coordinator = null)
    {
        MergeDiagnosticLogger? diagnosticLogger = null;
        try
        {
            if (args.Count > 0 && string.Equals(args[0], "git-config", StringComparison.Ordinal))
                return GitConfigCommand.Run(args, output);
            if (args.Count > 0 && string.Equals(args[0], "diagnostic-package", StringComparison.Ordinal))
                return DiagnosticPackageCommand.Run(args, output);

            var options = CommandLineParser.Parse(args);
            diagnosticLogger = MergeDiagnosticLogger.Create(options.LogPath);
            diagnosticLogger?.WriteStarted(options);
            var result = (coordinator ?? new LubanMergeCoordinator()).Merge(options);
            diagnosticLogger?.WriteCompleted(result);
            if (!result.Succeeded)
            {
                error.WriteLine($"存在 {result.Conflicts.Count} 个未解决冲突，MERGED 未被修改。");
                foreach (var conflict in result.Conflicts.Take(20))
                    error.WriteLine($"[{conflict.Kind}] {conflict.Message}");
                return ExitCodes.UnresolvedConflicts;
            }

            output.WriteLine(
                $"合并完成：逻辑表={result.LogicalTable}，工作表={result.SheetName}，主键={result.KeyName}，" +
                $"写入单元格={result.ChangedCells}，新增记录={result.AddedRecords}，删除记录={result.DeletedRecords}。");
            output.WriteLine($"MERGED={result.OutputPath}");
            if (result.IgnoredFields.Count > 0)
                output.WriteLine($"已保留 LOCAL 的忽略字段：{string.Join("、", result.IgnoredFields)}。");
            if (result.LogicalTableUniquenessValidated)
                output.WriteLine("全逻辑表唯一性：已检查。");
            output.WriteLine(result.RecalculationStatus switch
            {
                LubanExcelMerge.OpenXml.WorkbookRecalculationStatus.Completed =>
                    $"公式状态：已由 {result.RecalculationProvider} 完整重算。",
                LubanExcelMerge.OpenXml.WorkbookRecalculationStatus.SourceCachePreservedUnverified =>
                    "公式状态：已保留来源缓存并标记下次打开时完整重算；当前缓存未验证。",
                _ => "公式状态：本次合并不需要重新计算。"
            });
            if (result.ProjectValidationCompleted)
                output.WriteLine("项目快速校验：通过。");
            if (result.FullExportValidationCompleted)
                output.WriteLine("完整导出校验：已在隔离副本中通过。");
            return ExitCodes.Success;
        }
        catch (MergeInputException exception)
        {
            diagnosticLogger?.WriteException(exception, ExitCodes.InvalidInput);
            error.WriteLine(exception.Message);
            error.WriteLine(CommandLineParser.Usage);
            return ExitCodes.InvalidInput;
        }
        catch (UnsafeWorkbookException exception)
        {
            diagnosticLogger?.WriteException(exception, ExitCodes.UnsafeWorkbook);
            error.WriteLine(exception.Message);
            if (exception.InnerException is not null)
                error.WriteLine(exception.InnerException.Message);
            return ExitCodes.UnsafeWorkbook;
        }
        catch (WorkbookWriteException exception)
        {
            diagnosticLogger?.WriteException(exception, ExitCodes.WriteValidationFailed);
            error.WriteLine(exception.Message);
            if (exception.InnerException is not null)
                error.WriteLine(exception.InnerException.Message);
            return ExitCodes.WriteValidationFailed;
        }
        catch (ProjectValidationException exception)
        {
            diagnosticLogger?.WriteException(exception, ExitCodes.ProjectValidationFailed);
            error.WriteLine(exception.Message);
            if (exception.InnerException is not null)
                error.WriteLine(exception.InnerException.Message);
            return ExitCodes.ProjectValidationFailed;
        }
        catch (Exception exception)
        {
            diagnosticLogger?.WriteException(exception, ExitCodes.InternalError);
            error.WriteLine($"未分类内部错误：{exception}");
            return ExitCodes.InternalError;
        }
    }
}
