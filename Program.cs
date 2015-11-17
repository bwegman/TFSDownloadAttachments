using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using Attachment = Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment;

namespace TFSJiraConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (!CommandLine.Parser.Default.ParseArguments(args, options))

            {
                Console.WriteLine(options.GetUsage());

                return;
            }

            // Connect to the desired Team Foundation Server
            TfsTeamProjectCollection tfsServer = new TfsTeamProjectCollection(new Uri(options.ServerNameUrl));

            // Authenticate with the Team Foundation Server
            tfsServer.Authenticate();

            // Get a reference to a Work Item Store
            var workItemStore = new WorkItemStore(tfsServer);

            var project = GetProjectByName(workItemStore, options.ProjectName);

            if (project == null)
            {
                Console.WriteLine($"Could not find project '{options.ProjectName}'");
                return;
            }

            var query = GetWorkItemQueryByName(workItemStore, project, options.QueryPath);

            if (query == null)
            {
                Console.WriteLine($"Could not find query '{options.QueryPath}' in project '{options.ProjectName}'");
                return;
            }

            var queryText = query.QueryText.Replace("@project", $"'{project.Name}'");

            Console.WriteLine($"Executing query '{options.QueryPath}' with text '{queryText}'");

            var count = workItemStore.QueryCount(queryText);

            Console.WriteLine($"Exporting {count} work items");

            var workItems = workItemStore.Query(queryText);

            foreach (WorkItem workItem in workItems)
            {
                StoreAttachments(workItem, options.OutputPath);
            }
        }

        private static void StoreAttachments(WorkItem workItem, string outputPath)
        {
            // Get attachments
            var request = new System.Net.WebClient
            {
                Credentials = System.Net.CredentialCache.DefaultCredentials
            };

            // NOTE: If you use custom credentials to authenticate with TFS then you would most likely
            //       want to use those same credentials here

            if (workItem.Attachments.Count == 0)
                return;

            Console.WriteLine($"Storing {workItem.Attachments.Count} attachments for work item {workItem.Id}");

            foreach (Attachment attachment in workItem.Attachments)
            {
                // Display the name & size of the attachment
                Console.WriteLine($"  Attachment: '{attachment.Name}' ({attachment.Length} bytes)");

                var name = $"{workItem.Id} - {attachment.Name}";

                // Save the attachment to a local file
                request.DownloadFile(attachment.Uri, Path.Combine(outputPath, name));
            }
        }

        static Project GetProjectByName(WorkItemStore workItemStore, string projectName)
        {
            return workItemStore.Projects.Cast<Project>().FirstOrDefault(project => project.Name == projectName);
        }

        static QueryDefinition GetWorkItemQueryByName(WorkItemStore workItemStore, Project project, string queryPath)
        {
            var hierarchy = workItemStore.GetQueryHierarchy(project);

            foreach (var item in hierarchy)
            {
                if (item.Path == queryPath && item is QueryDefinition)
                    return item as QueryDefinition;

                if (!(item is QueryFolder))
                    continue;

                var query = GetWorkItemQueryByName(item as QueryFolder, queryPath);

                if (query != null)
                    return query;
            }

            return null;
        }

        static QueryDefinition GetWorkItemQueryByName(QueryFolder folder, string queryPath)
        {
            foreach (var item in folder)
            {
                Console.WriteLine($"{item.Name} {item.Path}");

                if (item.Path == queryPath && item is QueryDefinition)
                    return item as QueryDefinition;

                if (!(item is QueryFolder))
                    continue;

                var query = GetWorkItemQueryByName(item as QueryFolder, queryPath);

                if (query != null)
                    return query;
            }

            return null;
        }
    }
}
