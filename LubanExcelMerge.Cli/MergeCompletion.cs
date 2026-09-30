using LubanExcelMerge.Git;
using LubanExcelMerge.OpenXml;

namespace LubanExcelMerge.Cli;

public sealed record MergeCompletionResult(WorkbookSaveResult SaveResult, GitStageResult? StageResult);

public static class MergeCompletion
{
    // Windows 和跨平台界面共用保存、验证、暂存顺序及错误语义。
    public static MergeCompletionResult Save(PreparedMergeSession session, string repositoryRoot, bool stage)
    {
        var result = session.Save();
        GitStageResult? staged = null;
        if (stage)
        {
            try
            {
                staged = new GitMergedFileStager().Stage(repositoryRoot, result.OutputPath);
            }
            catch (Exception exception)
            {
                throw new WorkbookWriteException(
                    "MERGED 已保存，但 Git 暂存失败；请保留文件并检查 Git/Fork 配置。", exception);
            }
        }
        return new MergeCompletionResult(result, staged);
    }
}
