using NLog;
using NLog.Targets;
using Tomlyn;
using Tomlyn.Model;
using System.Text;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Custom NLog target that writes log entries in TOML format.
/// Each log entry is written as a separate TOML section.
/// </summary>
[Target("TomlFile")]
public sealed class TomlFileTarget : FileTarget
{
    private readonly object _lockObject = new();
    private long _entryCounter = 0;

    /// <summary>
    /// Initializes a new instance of the TomlFileTarget class.
    /// </summary>
    public TomlFileTarget()
    {
        // Override the layout to use our custom TOML formatter
        Layout = "${message}"; // We'll ignore this and use our own formatting
    }

    /// <summary>
    /// Writes the log event to the file in TOML format.
    /// </summary>
    /// <param name="logEvent">The log event to write.</param>
    protected override void Write(LogEventInfo logEvent)
    {
        lock (_lockObject)
        {
            var entryId = Interlocked.Increment(ref _entryCounter);
            var tomlContent = FormatLogEventAsToml(logEvent, entryId);
            
            // Use the base FileTarget's text writing capability
            // by temporarily setting our TOML content as the rendered message
            var originalMessage = logEvent.FormattedMessage;
            try
            {
                // Create a new log event with our TOML content as the message
                var tomlLogEvent = LogEventInfo.Create(logEvent.Level, logEvent.LoggerName, tomlContent);
                tomlLogEvent.TimeStamp = logEvent.TimeStamp;
                
                // Call the base Write method with our TOML content
                base.Write(tomlLogEvent);
            }
            finally
            {
                // Restore original message (though it's not really needed)
            }
        }
    }

    /// <summary>
    /// Formats a log event as a TOML section.
    /// </summary>
    /// <param name="logEvent">The log event to format.</param>
    /// <param name="entryId">The sequential entry ID.</param>
    /// <returns>TOML formatted string.</returns>
    private static string FormatLogEventAsToml(LogEventInfo logEvent, long entryId)
    {
        var tomlTable = new TomlTable();
        
        // Create a section for this log entry
        var logSection = new TomlTable
        {
            ["id"] = entryId,
            ["timestamp"] = logEvent.TimeStamp.ToString("O"), // ISO 8601 format
            ["level"] = logEvent.Level.ToString().ToUpperInvariant(),
            ["logger"] = logEvent.LoggerName ?? "Unknown",
            ["message"] = logEvent.FormattedMessage ?? string.Empty,
            ["thread"] = Environment.CurrentManagedThreadId.ToString(),
            ["machine"] = Environment.MachineName,
            ["user"] = Environment.UserName,
            ["version"] = "1.0.0"
        };

        // Add exception information if present
        if (logEvent.Exception != null)
        {
            var exceptionSection = new TomlTable
            {
                ["type"] = logEvent.Exception.GetType().FullName ?? "Unknown",
                ["message"] = logEvent.Exception.Message,
                ["stack_trace"] = logEvent.Exception.ToString()
            };
            logSection["exception"] = exceptionSection;
        }

        // Add properties if present
        if (logEvent.Properties?.Count > 0)
        {
            var propertiesSection = new TomlTable();
            foreach (var property in logEvent.Properties)
            {
                if (property.Key?.ToString() is string key && !string.IsNullOrEmpty(key))
                {
                    propertiesSection[key] = property.Value?.ToString() ?? string.Empty;
                }
            }
            if (propertiesSection.Count > 0)
            {
                logSection["properties"] = propertiesSection;
            }
        }

        // Use a timestamp-based section name to ensure uniqueness
        var sectionName = $"log_entry_{logEvent.TimeStamp:yyyyMMdd_HHmmssfff}_{entryId}";
        tomlTable[sectionName] = logSection;

        return Toml.FromModel(tomlTable);
    }

}
