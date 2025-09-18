// Logging/ILogger.cs
using System;

namespace ExcelAddInTest.Infrastructure.Logger
{
    public interface ILogger 
    {
        void Info(string message);
        void Warn(string message);
        void Error(string message, Exception ex = null);
        void Raw(string message); // large blobs (JSON, stack traces), no extra prefix

        void Dispose();

    }
}
