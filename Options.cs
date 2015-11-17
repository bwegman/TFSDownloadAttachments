using CommandLine;
using CommandLine.Text;

namespace TFSJiraConversion
{
    internal class Options
    {
        [Option('s', "server", Required = true, HelpText = "TFS Server URL")]
        public string ServerNameUrl { get; set; }

        [Option('p', "project", Required = true, HelpText = "Project")]
        public string ProjectName { get; set; }

        [Option('q', "query", Required = true, HelpText = "Name of the query to execute")]
        public string QueryPath { get; set; }

        [Option('o', "outputpath", Required = true, HelpText = "Path where to store attachments")]
        public string OutputPath { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}