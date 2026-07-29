namespace LubanExcelMerge.Cli;

public static class ExitCodes
{
    public const int Success = 0;
    public const int UnresolvedConflicts = 1;
    public const int InvalidInput = 2;
    public const int UnsafeWorkbook = 3;
    public const int WriteValidationFailed = 4;
    public const int ProjectValidationFailed = 5;
    public const int InternalError = 10;

    public static int ForException(Exception exception) => exception switch
    {
        MergeInputException => InvalidInput,
        UnsafeWorkbookException => UnsafeWorkbook,
        WorkbookWriteException => WriteValidationFailed,
        ProjectValidationException => ProjectValidationFailed,
        _ => InternalError
    };
}

public sealed class MergeInputException : Exception
{
    public MergeInputException(string message, Exception? innerException = null) : base(message, innerException)
    {
    }
}

public sealed class UnsafeWorkbookException : Exception
{
    public UnsafeWorkbookException(string message, Exception? innerException = null) : base(message, innerException)
    {
    }
}

public sealed class WorkbookWriteException : Exception
{
    public WorkbookWriteException(string message, Exception? innerException = null) : base(message, innerException)
    {
    }
}

public sealed class ProjectValidationException : Exception
{
    public ProjectValidationException(string message, Exception? innerException = null) : base(message, innerException)
    {
    }
}
