using NLog;
using NLog.Config;
using NLog.Targets;
using Tomlyn;
using Tomlyn.Model;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Service for configuring NLog from TOML configuration files.
/// </summary>
public static class LoggingConfiguration
{
    /// <summary>
    /// Initializes NLog configuration from TOML file.
    /// </summary>
    /// <param name="configPath">Path to the TOML configuration file.</param>
    public static void InitializeLogging(string configPath = "nlog.toml")
    {
        try
        {
            if (!File.Exists(configPath))
            {
                // Create default configuration if file doesn't exist
                CreateDefaultConfiguration();
                return;
            }

            var tomlContent = File.ReadAllText(configPath);
            var model = Toml.ToModel(tomlContent);
            
            var config = CreateNLogConfigurationFromToml(model);
            LogManager.Configuration = config;
            
            var logger = LogManager.GetCurrentClassLogger();
            logger.Info("NLog configuration loaded successfully from TOML file: {ConfigPath}", configPath);
        }
        catch (Exception ex)
        {
            // Fallback to default configuration if TOML parsing fails
            CreateDefaultConfiguration();
            var logger = LogManager.GetCurrentClassLogger();
            logger.Error(ex, "Failed to load TOML configuration from {ConfigPath}, using default configuration", configPath);
        }
    }

    /// <summary>
    /// Creates NLog configuration from parsed TOML model.
    /// </summary>
    /// <param name="model">The parsed TOML model.</param>
    /// <returns>NLog configuration.</returns>
    private static NLog.Config.LoggingConfiguration CreateNLogConfigurationFromToml(TomlTable model)
    {
        var config = new NLog.Config.LoggingConfiguration();
        var ruleTables = new List<TomlTable>();

        // Process variables
        if (model.TryGetValue("NLog", out var nlogSection) && nlogSection is TomlTable nlogTable)
        {
            if (nlogTable.TryGetValue("variable", out var variableSection) && variableSection is TomlTable variableTable)
            {
                foreach (var variable in variableTable)
                {
                    if (variable.Value is string variableValue)
                    {
                        config.Variables[variable.Key] = variableValue;
                    }
                }
            }

            // Process targets first
            if (nlogTable.TryGetValue("targets", out var targetsSection) && targetsSection is TomlTableArray targetsArray)
            {
                foreach (var targetItem in targetsArray)
                {
                    if (targetItem is TomlTable targetTable)
                    {
                        var target = CreateTargetFromToml(targetTable);
                        if (target != null)
                        {
                            config.AddTarget(target);
                        }
                    }
                }
            }

            // Collect rule tables for processing after targets are added
            if (nlogTable.TryGetValue("rules", out var rulesSection) && rulesSection is TomlTableArray rulesArray)
            {
                foreach (var ruleItem in rulesArray)
                {
                    if (ruleItem is TomlTable ruleTable)
                    {
                        ruleTables.Add(ruleTable);
                    }
                }
            }
        }

        // Now process rules after all targets have been added
        foreach (var ruleTable in ruleTables)
        {
            var rule = CreateRuleFromToml(ruleTable, config);
            if (rule != null)
            {
                config.LoggingRules.Add(rule);
            }
        }

        return config;
    }

    /// <summary>
    /// Creates a target from TOML configuration.
    /// </summary>
    /// <param name="targetTable">The target configuration table.</param>
    /// <returns>The configured target.</returns>
    private static Target? CreateTargetFromToml(TomlTable targetTable)
    {
        if (!targetTable.TryGetValue("name", out var nameValue) || nameValue is not string name ||
            !targetTable.TryGetValue("type", out var typeValue) || typeValue is not string type)
        {
            return null;
        }

        Target target = type.ToLowerInvariant() switch
        {
            "file" => new FileTarget(),
            "coloredconsole" => new ColoredConsoleTarget(),
            _ => throw new NotSupportedException($"Target type '{type}' is not supported")
        };

        target.Name = name;

        // Set common properties
        foreach (var property in targetTable)
        {
            if (property.Key is "name" or "type")
                continue;

            if (property.Value is string stringValue)
            {
                SetTargetProperty(target, property.Key, stringValue);
            }
            else if (property.Value is bool boolValue)
            {
                SetTargetProperty(target, property.Key, boolValue.ToString());
            }
            else if (property.Value is long longValue)
            {
                SetTargetProperty(target, property.Key, longValue.ToString());
            }
        }

        return target;
    }

    /// <summary>
    /// Creates a logging rule from TOML configuration.
    /// </summary>
    /// <param name="ruleTable">The rule configuration table.</param>
    /// <param name="config">The configuration object with targets already added.</param>
    /// <returns>The configured logging rule.</returns>
    private static LoggingRule? CreateRuleFromToml(TomlTable ruleTable, NLog.Config.LoggingConfiguration config)
    {
        if (!ruleTable.TryGetValue("logger", out var loggerValue) || loggerValue is not string loggerName ||
            !ruleTable.TryGetValue("writeTo", out var writeToValue) || writeToValue is not string writeTo)
        {
            return null;
        }

        var rule = new LoggingRule();
        rule.LoggerNamePattern = loggerName;

        // Set minimum level
        if (ruleTable.TryGetValue("minLevel", out var minLevelValue) && minLevelValue is string minLevelString)
        {
            var minLevel = minLevelString.ToLowerInvariant() switch
            {
                "trace" => LogLevel.Trace,
                "debug" => LogLevel.Debug,
                "info" => LogLevel.Info,
                "warn" => LogLevel.Warn,
                "error" => LogLevel.Error,
                "fatal" => LogLevel.Fatal,
                _ => LogLevel.Info
            };
            rule.SetLoggingLevels(minLevel, LogLevel.Fatal);
        }

        // Add target (find target by name from the current config)
        var target = config.FindTargetByName(writeTo);
        if (target == null)
        {
            throw new InvalidOperationException($"Target '{writeTo}' not found in configuration");
        }
        
        rule.Targets.Add(target);
        return rule;
    }

    /// <summary>
    /// Sets a property on a target using reflection.
    /// </summary>
    /// <param name="target">The target to configure.</param>
    /// <param name="propertyName">The property name.</param>
    /// <param name="propertyValue">The property value.</param>
    private static void SetTargetProperty(Target target, string propertyName, string propertyValue)
    {
        var property = target.GetType().GetProperty(propertyName, 
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
        
        if (property?.CanWrite == true)
        {
            try
            {
                var convertedValue = Convert.ChangeType(propertyValue, property.PropertyType);
                property.SetValue(target, convertedValue);
            }
            catch
            {
                // Ignore conversion errors for properties that can't be set
            }
        }
    }

    /// <summary>
    /// Creates a default NLog configuration when TOML file is not available.
    /// </summary>
    private static void CreateDefaultConfiguration()
    {
        var config = new NLog.Config.LoggingConfiguration();

        // Create directories
        Directory.CreateDirectory("logs");
        Directory.CreateDirectory("logs/operations");
        Directory.CreateDirectory("logs/errors");
        Directory.CreateDirectory("logs/registry");
        Directory.CreateDirectory("logs/filesystem");

        // File target
        var fileTarget = new FileTarget("fileTarget")
        {
            FileName = "logs/autodesk-cleaner-${date:format=yyyy-MM-dd}.log",
            Layout = "${longdate} [${level:uppercase=true}] ${logger} - ${message} ${exception:format=tostring}",
            ArchiveFileName = "logs/archive/autodesk-cleaner-{#}.log",
            ArchiveEvery = FileArchivePeriod.Day,
            MaxArchiveFiles = 30,
            KeepFileOpen = false,
            Encoding = System.Text.Encoding.UTF8
        };

        // Console target
        var consoleTarget = new ColoredConsoleTarget("consoleTarget")
        {
            Layout = "${time} [${level:uppercase=true:padding=-5}] ${message} ${exception:format=shorttype}",
            UseDefaultRowHighlightingRules = false
        };

        // Add color rules
        consoleTarget.RowHighlightingRules.Add(new ConsoleRowHighlightingRule
        {
            Condition = "level == LogLevel.Debug",
            ForegroundColor = ConsoleOutputColor.DarkGray
        });
        consoleTarget.RowHighlightingRules.Add(new ConsoleRowHighlightingRule
        {
            Condition = "level == LogLevel.Info",
            ForegroundColor = ConsoleOutputColor.Gray
        });
        consoleTarget.RowHighlightingRules.Add(new ConsoleRowHighlightingRule
        {
            Condition = "level == LogLevel.Warn",
            ForegroundColor = ConsoleOutputColor.Yellow
        });
        consoleTarget.RowHighlightingRules.Add(new ConsoleRowHighlightingRule
        {
            Condition = "level == LogLevel.Error",
            ForegroundColor = ConsoleOutputColor.Red
        });

        // Error target
        var errorTarget = new FileTarget("errorTarget")
        {
            FileName = "logs/errors/autodesk-cleaner-errors-${date:format=yyyy-MM-dd}.log",
            Layout = "${longdate} [${level:uppercase=true}] ${logger} - ${message} ${exception:format=tostring}",
            ArchiveFileName = "logs/archive/errors/autodesk-cleaner-errors-{#}.log",
            ArchiveEvery = FileArchivePeriod.Day,
            MaxArchiveFiles = 90
        };

        config.AddTarget(fileTarget);
        config.AddTarget(consoleTarget);
        config.AddTarget(errorTarget);

        // Add rules
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, fileTarget));
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Info, consoleTarget));
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Error, errorTarget));

        LogManager.Configuration = config;
    }
}
