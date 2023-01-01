using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
#if NET6_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace True.Deal.EnvironmentVariable.PowerShell.Commands
{

#if NET6_0_OR_GREATER

    /// <summary>
    /// Defines the implementation of the 'Get-WinEnviromentVariable' cmdlet.
    /// This cmdlet get the content from EnvironemtVariable.
    /// </summary>
    [SupportedOSPlatform("windows")]
    [Cmdlet(VerbsCommon.Get, "WinEnvironmentVariable", DefaultParameterSetName = "DefaultSet", HelpUri = "https://github.com/KatsutoshiOtogawa/PowerShell.Commands.True.Deal.EnvironmentVariable/Help/Get-WinEnvironmentVariable.md")]
    [OutputType(typeof(PSObject), ParameterSetName = new[] { "DefaultSet" })]
    [OutputType(typeof(string), ParameterSetName = new[] { "RawSet" })]
    public class GetWinEnvironmentVariableCommand : PSCmdlet
    {
        /// <summary>
        /// Gets or sets specifies the Name EnvironmentVariable.
        /// </summary>
        [Parameter(Position = 0, ParameterSetName = "DefaultSet", Mandatory = false, ValueFromPipelineByPropertyName = true)]
        [Parameter(Position = 0, ParameterSetName = "RawSet", Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the EnvironmentVariableTarget.
        /// </summary>
        [Parameter(Position = 1, Mandatory = false, ParameterSetName = "DefaultSet")]
        [Parameter(Position = 1, Mandatory = false, ParameterSetName = "RawSet")]
        [ValidateNotNullOrEmpty]
        public EnvironmentVariableTarget Target { get; set; } = EnvironmentVariableTarget.Process;

        /// <summary>
        /// Gets or sets property that sets delimiter.
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true, ParameterSetName = "DefaultSet")]
        [ValidateNotNullOrEmpty]
        public char? Delimiter { get; set; } = null;

        /// <summary>
        /// Gets or sets raw parameter. This will allow EnvironmentVariable return text or file list as one string.
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "RawSet")]
        public SwitchParameter Raw { get; set; }

        private static readonly List<string> DetectedDelimiterEnvrionmentVariable = new List<string> { "Path", "PATHEXT", "PSModulePath" };

        /// <summary>
        /// This method implements the ProcessRecord method for Get-WinEnvironmentVariable command.
        /// Returns the Specify Name EnvironmentVariable content as text format.
        /// </summary>
        protected override void BeginProcessing()
        {
            PSObject env;
            PSNoteProperty envname;
            PSNoteProperty envvalue;
            PSNoteProperty envtype;

            if (string.IsNullOrEmpty(Name))
            {
                foreach (DictionaryEntry kvp in Environment.WinGetEnvironmentVariables(Target))
                {
                    env = new PSObject();
                    envname = new PSNoteProperty("Name", kvp.Key.ToString());
                    envtype = Target switch
                    {
                        EnvironmentVariableTarget.Process => new PSNoteProperty("RegistryValueKind", RegistryValueKind.None),
                        _ => new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(kvp.Key.ToString()!, Target))
                    };
                    envvalue = new PSNoteProperty("Value", kvp.Value?.ToString());
                    env.Properties.Add(envname);
                    env.Properties.Add(envtype);
                    env.Properties.Add(envvalue);

                    this.WriteObject(env, true);
                }
                return;
            }
            var contentList = new List<string>();

            // try catch IOExceptionがありうる。環境変数が無い場合の
            string? textContent = Environment.WinGetEnvironmentVariable(Name, Target);
            if (string.IsNullOrEmpty(textContent))
            {
                var message = string.Format(
                    CultureInfo.CurrentCulture,
                    WinEnvironmentVariableResources.EnvironmentVariableNotFoundOrEmpty,
                    Name
                );

                ArgumentException argumentException = new ArgumentException(message);
                ErrorRecord errorRecord = new(
                    argumentException,
                    "EnvironmentVariableNotFoundOrEmpty",
                    ErrorCategory.ObjectNotFound,
                    Name);
                ThrowTerminatingError(errorRecord);
                return;
            }

            if (ParameterSetName == "RawSet")
            {
                contentList.Add(textContent);
                this.WriteObject(textContent, true);
                return;
            }
            else
            {
                if (DetectedDelimiterEnvrionmentVariable.Contains(Name))
                {
                    Delimiter = Path.PathSeparator;
                }

                contentList.AddRange(textContent.Split(Delimiter.ToString() ?? string.Empty, StringSplitOptions.None));
            }

            env = new PSObject();
            envname = new PSNoteProperty("Name", Name);
            envtype = Target switch
            {
                EnvironmentVariableTarget.Process => new PSNoteProperty("RegistryValueKind", RegistryValueKind.None),
                _ => new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(Name, Target))
            };
            envvalue = new PSNoteProperty("Value", contentList);

            env.Properties.Add(envname);
            env.Properties.Add(envtype);
            env.Properties.Add(envvalue);

            this.WriteObject(env, true);
        }
    }
#elif (NETSTANDARD2_0_OR_GREATER || NET452_OR_GREATER)

    /// <summary>
    /// Defines the implementation of the 'Get-WinEnviromentVariable' cmdlet.
    /// This cmdlet get the content from EnvironemtVariable.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "WinEnvironmentVariable", DefaultParameterSetName = "DefaultSet", HelpUri = "https://github.com/KatsutoshiOtogawa/PowerShell.Commands.True.Deal.EnvironmentVariable/Help/Get-WinEnvironmentVariable.md")]
    [OutputType(typeof(PSObject), ParameterSetName = new[] { "DefaultSet" })]
    [OutputType(typeof(string), ParameterSetName = new[] { "RawSet" })]
    public class GetWinEnvironmentVariableCommand : PSCmdlet
    {
        /// <summary>
        /// Gets or sets specifies the Name EnvironmentVariable.
        /// </summary>
        [Parameter(Position = 0, ParameterSetName = "DefaultSet", Mandatory = false, ValueFromPipelineByPropertyName = true)]
        [Parameter(Position = 0, ParameterSetName = "RawSet", Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the EnvironmentVariableTarget.
        /// </summary>
        [Parameter(Position = 1, Mandatory = false, ParameterSetName = "DefaultSet")]
        [Parameter(Position = 1, Mandatory = false, ParameterSetName = "RawSet")]
        [ValidateNotNullOrEmpty]
        public EnvironmentVariableTarget Target { get; set; } = EnvironmentVariableTarget.Process;

        /// <summary>
        /// Gets or sets property that sets delimiter.
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true, ParameterSetName = "DefaultSet")]
        [ValidateNotNullOrEmpty]
        public char? Delimiter { get; set; } = null;

        /// <summary>
        /// Gets or sets raw parameter. This will allow EnvironmentVariable return text or file list as one string.
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "RawSet")]
        public SwitchParameter Raw { get; set; }

        private static readonly List<string> DetectedDelimiterEnvrionmentVariable = new List<string> { "Path", "PATHEXT", "PSModulePath" };

        /// <summary>
        /// This method implements the ProcessRecord method for Get-WinEnvironmentVariable command.
        /// Returns the Specify Name EnvironmentVariable content as text format.
        /// </summary>
        protected override void BeginProcessing()
        {
            PSObject env;
            PSNoteProperty envname;
            PSNoteProperty envvalue;
            PSNoteProperty envtype;

            if (string.IsNullOrEmpty(Name))
            {
                foreach (DictionaryEntry kvp in Environment.WinGetEnvironmentVariables(Target))
                {
                    env = new PSObject();
                    envname = new PSNoteProperty("Name", kvp.Key.ToString());
                    // envtype = Target switch
                    // {
                    //     EnvironmentVariableTarget.Process => new PSNoteProperty("RegistryValueKind", RegistryValueKind.None),
                    //     _ => new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(kvp.Key.ToString()!, Target))
                    // };
                    switch (Target)
                    {
                        case EnvironmentVariableTarget.Process:
                            envtype = new PSNoteProperty("RegistryValueKind", RegistryValueKind.None);
                            break;
                        default:
                            envtype = new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(kvp.Key.ToString(), Target));
                            break;
                    };
                    // envvalue = new PSNoteProperty("Value", kvp.Value.?ToString());
                    if (kvp.Value != null)
                    {
                        envvalue = new PSNoteProperty("Value", kvp.Value.ToString());
                    } else
                    {
                        envvalue = new PSNoteProperty("Value", null);
                    }
                    env.Properties.Add(envname);
                    env.Properties.Add(envtype);
                    env.Properties.Add(envvalue);

                    this.WriteObject(env, true);
                }
                return;
            }
            var contentList = new List<string>();

            // try catch IOExceptionがありうる。環境変数が無い場合の
            string textContent = Environment.WinGetEnvironmentVariable(Name, Target);
            if (string.IsNullOrEmpty(textContent))
            {
                var message = string.Format(
                    CultureInfo.CurrentCulture,
                    WinEnvironmentVariableResources.EnvironmentVariableNotFoundOrEmpty,
                    Name
                );

                ArgumentException argumentException = new ArgumentException(message);
                ErrorRecord errorRecord = new ErrorRecord(
                    argumentException,
                    "EnvironmentVariableNotFoundOrEmpty",
                    ErrorCategory.ObjectNotFound,
                    Name);
                ThrowTerminatingError(errorRecord);
                return;
            }

            if (ParameterSetName == "RawSet")
            {
                contentList.Add(textContent);
                this.WriteObject(textContent, true);
                return;
            }
            else
            {
                if (DetectedDelimiterEnvrionmentVariable.Contains(Name))
                {
                    Delimiter = Path.PathSeparator;
                }

                // contentList.AddRange(textContent.Split(Delimiter.ToString() ?? string.Empty, StringSplitOptions.None));
                if (Delimiter.ToString() is null)
                {
                    contentList.AddRange(textContent.Split(string.Empty.ToCharArray(), StringSplitOptions.None));
                }
                else
                {
                    contentList.AddRange(textContent.Split(Delimiter.ToString().ToCharArray(), StringSplitOptions.None));
                }
            }

            env = new PSObject();
            envname = new PSNoteProperty("Name", Name);
            // envtype = Target switch
            // {
            //   EnvironmentVariableTarget.Process => new PSNoteProperty("RegistryValueKind", RegistryValueKind.None),
            //   _ => new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(Name, Target))
            // };
            switch (Target)
            {
                case EnvironmentVariableTarget.Process:
                    envtype = new PSNoteProperty("RegistryValueKind", RegistryValueKind.None);
                    break;
                default:
                    envtype = new PSNoteProperty("RegistryValueKind", Environment.WinGetEnvironmentValueKind(Name, Target));
                    break;
            };
            envvalue = new PSNoteProperty("Value", contentList);

            env.Properties.Add(envname);
            env.Properties.Add(envtype);
            env.Properties.Add(envvalue);

            this.WriteObject(env, true);
        }
    }
#endif

}
