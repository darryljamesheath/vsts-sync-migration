using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VstsSyncMigrator.Engine.Configuration.Processing;

namespace VstsSyncMigrator.Engine
{
    public class WorkItemQueryMigrationProcessorContext : MigrationContextBase
    {
        private int totalFolderAttempted = 0;
        private int totalQuerySkipped = 0;
        private int totalQueryAttempted = 0;
        private int totalQueryMigrated = 0;
        private int totalQueryFailed = 0;

        public override string Name
        {
            get
            {
                return "WorkItemQueryMigrationProcessorContext";
            }
        }

        public WorkItemQueryMigrationProcessorContext(MigrationEngine me, WorkItemQueryMigrationProcessorConfig config) : base(me, config)
        {

        }


        internal override void InternalExecute()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            //////////////////////////////////////////////////
            var sourceStore = new WorkItemStoreContext(me.Source, WorkItemStoreFlags.None);
            var targetStore = new WorkItemStoreContext(me.Target, WorkItemStoreFlags.None);

            var sourceQueryHierarchy = sourceStore.Store.Projects[me.Source.Name].QueryHierarchy;
            var targetQueryHierarchy = targetStore.Store.Projects[me.Target.Name].QueryHierarchy;

            Trace.WriteLine(string.Format("Found {0} root WIQ folders", sourceQueryHierarchy.Count));
            //////////////////////////////////////////////////

            foreach (QueryFolder query in sourceQueryHierarchy)
            {
                MigrateFolder(targetQueryHierarchy, query, targetQueryHierarchy);
            }

            stopwatch.Stop();
            Console.WriteLine($"Folders scanned {this.totalFolderAttempted}");
            Console.WriteLine($"Queries Found:{this.totalQueryAttempted}  Skipped:{this.totalQuerySkipped}  Migrated:{this.totalQueryMigrated}   Failed:{this.totalQueryFailed}");
            Console.WriteLine(@"DONE in {0:%h} hours {0:%m} minutes {0:s\:fff} seconds", stopwatch.Elapsed);
        }

        /// <summary>
        /// Define Query Folders under the current parent
        /// </summary>
        /// <param name="sourceFolder">The source folder (on source VSTS instance)</param>
        /// <param name="parentFolder">The target folder (on target VSTS instance)</param>
        private void MigrateFolder(QueryHierarchy targetHierarchy, QueryFolder sourceFolder, QueryFolder parentFolder)
        {
            if (sourceFolder.IsPersonal)
            {
                Trace.WriteLine($"Found a personal folder {sourceFolder.Name}. Migration only available for shared Team Query folders");
            }
            else
            {
                this.totalFolderAttempted++;

                // check if there is a folder of this name, using the path to make sure unique
                // note we need to replace the team projetc name as it returned in path
                QueryFolder targetFolder = (QueryFolder)parentFolder.FirstOrDefault(q => q.Path == sourceFolder.Path.Replace($"{me.Source.Name}/", $"{me.Target.Name}/"));
                if (targetFolder != null)
                {
                    Trace.WriteLine($"Skipping folder '{sourceFolder.Name}' as already exists");
                }
                else
                {
                    Trace.WriteLine($"Migrating a folder '{sourceFolder.Name}'");
                    targetFolder = new QueryFolder(sourceFolder.Name);
                    parentFolder.Add(targetFolder);
                    targetHierarchy.Save(); // moved the save here for better error message
                }

                foreach (QueryItem sub_query in sourceFolder)
                {
                    if (sub_query.GetType() == typeof(QueryFolder))
                    {
                        MigrateFolder(targetHierarchy, (QueryFolder)sub_query, (QueryFolder)targetFolder);
                    }
                    else
                    {
                        MigrateQuery(targetHierarchy, (QueryDefinition)sub_query, (QueryFolder)targetFolder);
                    }
                }
            }

        }

        /// <summary>
        /// Add Query Definition under a specific Query Folder.
        /// </summary>
        /// <param name="query">Query Definition - Contains the Query Details</param>
        /// <param name="QueryFolder">Parent Folder</param>
        void MigrateQuery(QueryHierarchy targetHierarchy, QueryDefinition query, QueryFolder parentFolder)
        {
            if (parentFolder.FirstOrDefault(q => q.Name == query.Name) != null)
            {
                this.totalQuerySkipped++;
                Trace.WriteLine($"Skipping query '{query.Name}' as already exists");
            }
            else
            {
                // you cannot just add an item from one store to another
                var queryCopy = new QueryDefinition(
                    query.Name,
                    query.QueryText.Replace($"'{me.Source.Name}", $"'{me.Target.Name}")); // the ' should only items at the start of areapath etc.

                this.totalQueryAttempted++;
                Trace.WriteLine($"Migrating query '{query.Name}'");
                parentFolder.Add(queryCopy);
                try
                {
                    targetHierarchy.Save(); // moved the save here for better error message
                    this.totalQueryMigrated++;
                }
                catch (Exception ex)
                {
                    this.totalQueryFailed++;
                    Trace.WriteLine($"Error saving query '{query.Name}', probably due to invalid area or iteration paths");
                    Trace.WriteLine(ex.Message);
                    targetHierarchy.Refresh(); // get the tree without the last edit
                }
            }
        }

    }
}