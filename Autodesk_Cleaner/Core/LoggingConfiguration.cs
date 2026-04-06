using NLog;
using NLog.Config;
using NLog.Targets;
using Tomlyn;
using Tomlyn.Model;
using Spectre.Console;

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
            var resolvedConfigPath = ResolveConfigurationPath(configPath);
            var configDirectory = Path.GetDirectoryName(resolvedConfigPath) ?? AppDomain.CurrentDomain.BaseDirectory;
            var logDirectory = Path.Combine(configDirectory, "logs");
            var archiveDirectory = Path.Combine(logDirectory, "archive");

            // Register our custom TOML target with modern NLog API
            LogManager.Setup().SetupExtensions(s => s.RegisterTarget<TomlFileTarget>("TomlFile"));
            
            WriteDiagnostic($"Looking for nlog config at: {resolvedConfigPath}");
            if (!File.Exists(resolvedConfigPath))
            {
                WriteDiagnostic("TOML config file not found, using default configuration");
                // Create default configuration if file doesn't exist
                CreateDefaultConfiguration(configDirectory);
                return;
            }
            
            WriteDiagnostic($"Loading TOML configuration from: {resolvedConfigPath}");

            var tomlContent = File.ReadAllText(resolvedConfigPath);
            var model = Toml.ToModel(tomlContent);
            
            var config = CreateNLogConfigurationFromToml(model, configDirectory);
            
            // Create log directories (TOML config doesn't do this automatically)
            WriteDiagnostic("Creating log directories for TOML config...");
            CreateLogDirectories(logDirectory, archiveDirectory);

            NLog.Common.InternalLogger.LogFile = Path.Combine(logDirectory, "nlog-internal.log");
            
            LogManager.Configuration = config;
            
            // Debug: Check if targets were created
            WriteDiagnostic($"Configured targets: {string.Join(", ", config.AllTargets.Select(t => t.Name))}");
            WriteDiagnostic($"Configured rules: {config.LoggingRules.Count}");
        
        var logger = LogManager.GetCurrentClassLogger();
        logger.Info("NLog configuration loaded successfully from TOML file: {ConfigPath}", resolvedConfigPath);
        
        // Test file logging immediately
        WriteDiagnostic("Testing file logging...");
        logger.Info("TOML FILE LOGGING TEST - This should appear in log files");
        logger.Error("TOML ERROR LOGGING TEST - This should appear in log files");
        
        // Force flush to ensure files are written
        WriteDiagnostic("Flushing log files...");
        LogManager.Flush();
        
        // Debug: Check if files were created
        var logFiles = Directory.GetFiles(logDirectory, "*.log", SearchOption.AllDirectories);
        WriteDiagnostic($"Log files found: {logFiles.Length}");
        foreach (var file in logFiles)
        {
            WriteDiagnostic($"- {file} ({new FileInfo(file).Length} bytes)");
        }
        }
        catch (Exception ex)
        {
            WriteDiagnostic($"Error loading TOML config: {ex.Message}");
            // Fallback to default configuration if TOML parsing fails
            var fallbackDirectory = Path.GetDirectoryName(ResolveConfigurationPath(configPath)) ?? AppDomain.CurrentDomain.BaseDirectory;
            CreateDefaultConfiguration(fallbackDirectory);
            var logger = LogManager.GetCurrentClassLogger();
            logger.Error(ex, "Failed to load TOML configuration from {ConfigPath}, using default configuration", configPath);
        }
    }

    /// <summary>
    /// Creates NLog configuration from parsed TOML model.
    /// </summary>
    /// <param name="model">The parsed TOML model.</param>
    /// <returns>NLog configuration.</returns>
    private static NLog.Config.LoggingConfiguration CreateNLogConfigurationFromToml(TomlTable model, string configDirectory)
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
                        config.Variables[variable.Key] = variable.Key switch
                        {
                            "logDirectory" => GetAbsolutePath(configDirectory, variableValue),
                            "archiveDirectory" => GetAbsolutePath(configDirectory, variableValue),
                            _ => variableValue
                        };
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
            WriteDiagnostic("Invalid target table - missing name or type");
            return null;
        }

        WriteDiagnostic($"Creating target: {name} of type {type}");

        Target target = type.ToLowerInvariant() switch
        {
            "file" => new FileTarget(),
            "coloredconsole" => new ColoredConsoleTarget(),
            "tomlfile" => new TomlFileTarget(),
            _ => throw new NotSupportedException($"Target type '{type}' is not supported")
        };

        target.Name = name;

        // Set common properties
        foreach (var property in targetTable)
        {
            if (property.Key is "name" or "type")
                continue;

            WriteDiagnostic($"Setting property {property.Key} = {property.Value}");
            
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

        WriteDiagnostic($"Target {name} created successfully");
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
                object convertedValue;
                
                // Handle special NLog property types
                if (property.PropertyType == typeof(NLog.Layouts.Layout) || property.PropertyType.IsSubclassOf(typeof(NLog.Layouts.Layout)))
                {
                    convertedValue = NLog.Layouts.Layout.FromString(propertyValue);
                }
                else if (property.PropertyType == typeof(System.Text.Encoding))
                {
                    convertedValue = System.Text.Encoding.GetEncoding(propertyValue);
                }
                else if (property.PropertyType == typeof(FileArchivePeriod))
                {
                    convertedValue = Enum.Parse(typeof(FileArchivePeriod), propertyValue, true);
                }
                else
                {
                    convertedValue = Convert.ChangeType(propertyValue, property.PropertyType);
                }
                
                property.SetValue(target, convertedValue);
                WriteDiagnostic($"Successfully set {propertyName} = {propertyValue} on {target.Name}");
            }
            catch (Exception ex)
            {
                WriteDiagnostic($"Failed to set {propertyName} = {propertyValue} on {target.Name}: {ex.Message}");
            }
        }
        else
        {
            WriteDiagnostic($"Property {propertyName} not found or not writable on {target.GetType().Name}");
        }
    }

    /// <summary>
    /// Creates a TOML-formatted layout for modern structured logging.
    /// </summary>
    /// <returns>TOML layout string.</returns>
    private static string CreateTomlLayout()
    {
        return @"# Autodesk Cleaner Log Entry
[entry]
timestamp = ""${longdate}""
level = ""${level:uppercase=true}""
logger = ""${logger}""
message = ""${message}""
thread = ""${threadid}""
user = ""${environment-user}""
machine = ""${machinename}""
version = ""1.0.0""
${when:when='${exception}'!='':inner=[exception]
type = ""${exception:format=@}""
message = ""${exception:format=message}""
stack_trace = ""${exception:format=tostring}""}
";
    }
    
    /// <summary>
    /// Creates the log directory needed for file logging.
    /// </summary>
    private static void CreateLogDirectories(string logDirectory, string archiveDirectory)
    {
        try
        {
            Directory.CreateDirectory(logDirectory);
            Directory.CreateDirectory(archiveDirectory);
            WriteDiagnostic("Log directories created successfully");
        }
        catch (Exception ex)
        {
            WriteDiagnostic($"Error creating log directories: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Creates a default NLog configuration when TOML file is not available.
    /// </summary>
    private static void CreateDefaultConfiguration(string baseDirectory)
    {
        WriteDiagnostic("Creating default NLog configuration");
        
        // Register our custom TOML target with modern NLog API
        LogManager.Setup().SetupExtensions(s => s.RegisterTarget<TomlFileTarget>("TomlFile"));
        
        var config = new NLog.Config.LoggingConfiguration();
        var logDirectory = Path.Combine(baseDirectory, "logs");
        var archiveDirectory = Path.Combine(logDirectory, "archive");

        // Create directories
        WriteDiagnostic("Creating log directories");
        CreateLogDirectories(logDirectory, archiveDirectory);
        WriteDiagnostic($"Log directories created at: {logDirectory}");

        NLog.Common.InternalLogger.LogFile = Path.Combine(logDirectory, "nlog-internal.log");

        // Main structured log file
        var fileTarget = new FileTarget("fileTarget")
        {
            FileName = Path.Combine(logDirectory, "autodesk-cleaner-${date:format=yyyy-MM-dd}.log"),
            Layout = "${longdate} [${level:uppercase=true}] ${logger} | ${message} | thread=${threadid} user=${environment-user} machine=${machinename} ${exception:format=tostring}",
            ArchiveFileName = Path.Combine(archiveDirectory, "autodesk-cleaner-{#}.log"),
            ArchiveEvery = FileArchivePeriod.Day,
            MaxArchiveFiles = 30,
            KeepFileOpen = false,
            Encoding = System.Text.Encoding.UTF8
        };
        
        // JSON structured log file
        var jsonTarget = new FileTarget("jsonTarget")
        {
            FileName = Path.Combine(logDirectory, "autodesk-cleaner-${date:format=yyyy-MM-dd}.json"),
            Layout = "${json-encode}",
            ArchiveFileName = Path.Combine(archiveDirectory, "autodesk-cleaner-{#}.json"),
            ArchiveEvery = FileArchivePeriod.Day,
            MaxArchiveFiles = 30,
            KeepFileOpen = false,
            Encoding = System.Text.Encoding.UTF8
        };
        
        // TOML structured log file
        var tomlTarget = new TomlFileTarget
        {
            Name = "tomlTarget",
            FileName = Path.Combine(logDirectory, "autodesk-cleaner-${date:format=yyyy-MM-dd}.toml"),
            ArchiveFileName = Path.Combine(archiveDirectory, "autodesk-cleaner-{#}.toml"),
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

        config.AddTarget(fileTarget);
        config.AddTarget(jsonTarget);
        config.AddTarget(tomlTarget);
        config.AddTarget(consoleTarget);

        // Add rules for simplified logging
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Info, fileTarget));
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Info, jsonTarget));
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Info, tomlTarget));
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, consoleTarget));

        LogManager.Configuration = config;
        
        // Test file logging immediately
        WriteDiagnostic("Testing default file logging...");
        var logger = LogManager.GetCurrentClassLogger();
        logger.Info("DEFAULT FILE LOGGING TEST - This should appear in log files");
        logger.Error("DEFAULT ERROR LOGGING TEST - This should appear in log files");
        
        // Force flush to ensure files are written
        WriteDiagnostic("Flushing default log files...");
        LogManager.Flush();
        
        // Debug: Check if files were created
        var logFiles = Directory.GetFiles(logDirectory, "*.log", SearchOption.AllDirectories);
        WriteDiagnostic($"Default log files found: {logFiles.Length}");
        foreach (var file in logFiles)
        {
            WriteDiagnostic($"- {file} ({new FileInfo(file).Length} bytes)");
        }
        
        // Force a manual file write test
        try
        {
            var testFile = Path.Combine(logDirectory, "manual-test.txt");
            File.WriteAllText(testFile, $"Manual file write test at {DateTime.Now}");
            WriteDiagnostic($"Manual file write successful: {testFile}");
        }
        catch (Exception ex)
        {
            WriteDiagnostic($"Manual file write failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Resolves the logging configuration path from the current directory or application directory.
    /// </summary>
    private static string ResolveConfigurationPath(string configPath)
    {
        if (Path.IsPathRooted(configPath))
        {
            return configPath;
        }

        var candidates = new[]
        {
            Path.Combine(Environment.CurrentDirectory, configPath),
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, configPath)
        };

        return candidates.FirstOrDefault(File.Exists) ?? candidates[0];
    }

    /// <summary>
    /// Resolves a relative path against the configuration directory.
    /// </summary>
    private static string GetAbsolutePath(string baseDirectory, string path)
    {
        return Path.IsPathRooted(path)
            ? path
            : Path.GetFullPath(Path.Combine(baseDirectory, path));
    }

    /// <summary>
    /// Writes bootstrap diagnostics only when explicitly enabled during development.
    /// </summary>
    private static void WriteDiagnostic(string message)
    {
        _ = message;
    }
}
