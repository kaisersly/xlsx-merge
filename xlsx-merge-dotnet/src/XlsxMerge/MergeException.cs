using System;

namespace XlsxMerge;

/// <summary>
/// Mail-merge error caused by the data or the template (message aimed at the end user).
/// </summary>
public sealed class MergeException : Exception
{
    public MergeException(string message) : base(message)
    {
    }

    public MergeException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
