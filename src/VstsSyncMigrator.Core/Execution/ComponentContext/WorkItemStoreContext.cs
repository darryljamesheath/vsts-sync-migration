using System;
using System.Linq;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace VstsSyncMigrator.Engine
{
    public class WorkItemStoreContext
    {
        private readonly ITeamProjectContext targetTfs;
        private readonly WorkItemStore wistore;
        private readonly Dictionary<int, WorkItem> foundWis;
        private WorkItemCollection existingItems;

        public WorkItemStore Store { get { return wistore; } }

        public WorkItemStoreContext(ITeamProjectContext targetTfs, WorkItemStoreFlags bypassRules)
        {
            var startTime = DateTime.UtcNow;
            var timer = Stopwatch.StartNew();
            this.targetTfs = targetTfs;

            try
            {
                wistore = new WorkItemStore(targetTfs.Collection, bypassRules);
                timer.Stop();
                Telemetry.Current.TrackDependency("TeamService", "GetWorkItemStore", startTime, timer.Elapsed, true);
            }
            catch (Exception ex)
            {
                timer.Stop();
                Telemetry.Current.TrackDependency("TeamService", "GetWorkItemStore", startTime, timer.Elapsed, false);
                Telemetry.Current.TrackException(ex,
                       new Dictionary<string, string> {
                            { "CollectionUrl", targetTfs.Collection.Uri.ToString() }
                       },
                       new Dictionary<string, double> {
                            { "Time",timer.ElapsedMilliseconds }
                       });
                Trace.TraceWarning(string.Format("  [EXCEPTION] {0}", ex.Message));

                throw;
            }

            foundWis = new Dictionary<int, WorkItem>();
        }

        public Project GetProject()
        {
            return (from Project x in wistore.Projects where x.Name == targetTfs.Name select x).SingleOrDefault();
        }

        public string CreateReflectedWorkItemId(WorkItem wi)
        {
            return string.Format("{0}/{1}/{2}", wi.Store.TeamProjectCollection.Uri, wi.Project.Name, wi.Id);
        }

        public int GetReflectedWorkItemId(WorkItem wi, string reflectedWotkItemIdField)
        {
            string rwiid = wi.Fields[reflectedWotkItemIdField].Value.ToString();
            if (Regex.IsMatch(rwiid, @"(http(s)?://)?([\w-]+\.)+[\w-]+(/[\w- ;,./?%&=]*)?"))
            {
                return int.Parse(rwiid.Substring(rwiid.LastIndexOf(@"/") + 1));
            }
            return 0;
        }

        public WorkItem FindReflectedWorkItem(WorkItem workItemToFind, string reflectedWorkItemIdField, bool cache)
        {
            if (existingItems == null)
            {
                LoadAllExistingItems(reflectedWorkItemIdField);
            }

            string reflectedWorkItemId = CreateReflectedWorkItemId(workItemToFind);
            foreach(WorkItem existingItem in existingItems)
            {
                if (existingItem.Fields[reflectedWorkItemIdField].Value.ToString() == reflectedWorkItemId)
                {
                    return existingItem;
                }
            }

            return null;
        }

        public WorkItem FindReflectedWorkItemByReflectedWorkItemId(WorkItem refWi, string reflectedWotkItemIdField)
        {
            return FindReflectedWorkItemByReflectedWorkItemId(CreateReflectedWorkItemId(refWi), reflectedWotkItemIdField);
        }

        public void LoadAllExistingItems(string reflectedWorkItemIdField)
        {
            var query = new TfsQueryContext(this);
            query.Query = $"SELECT [System.Id] FROM WorkItems  WHERE [System.TeamProject]=@TeamProject AND [{reflectedWorkItemIdField}] != ''";
            query.AddParameter("TeamProject", targetTfs.Name);
            existingItems = query.Execute();
        }

        public WorkItem FindReflectedWorkItemByReflectedWorkItemId(string refId, string reflectedWotkItemIdField)
        {
            TfsQueryContext query = new TfsQueryContext(this);
            query.Query = string.Format(@"SELECT [System.Id] FROM WorkItems  WHERE [System.TeamProject]=@TeamProject AND [{0}] = @idToFind", reflectedWotkItemIdField);
            query.AddParameter("idToFind", refId.ToString());
            query.AddParameter("TeamProject", this.targetTfs.Name);
            return FindWorkItemByQuery(query);
        }

        public WorkItem FindWorkItemByQuery(TfsQueryContext query)
        {
            WorkItemCollection newFound;
            newFound = query.Execute();
            if (newFound.Count == 0)
            {
                return null;
            }
            return newFound[0];
        }
    }
}