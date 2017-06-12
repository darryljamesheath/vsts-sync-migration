﻿using Microsoft.TeamFoundation.WorkItemTracking.Client;
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
    public class WorkItemMigrationContext : MigrationContextBase
    {
        private readonly WorkItemMigrationConfig _config;
        private readonly MigrationEngine _me;
        private readonly List<String> _ignore;

        public override string Name
        {
            get
            {
                return "WorkItemMigrationContext";
            }
        }

        public WorkItemMigrationContext(MigrationEngine me, WorkItemMigrationConfig config) : base(me, config)
        {
            _me = me;
            _config = config;
            _ignore = new List<string>();
            PopulateIgnoreList();
        }

        private void PopulateIgnoreList()
        {
            _ignore.Add("System.Rev");
            _ignore.Add("System.AreaId");
            _ignore.Add("System.IterationId");
            _ignore.Add("System.Id");
            _ignore.Add("System.RevisedDate");
            _ignore.Add("System.AttachedFileCount");
            _ignore.Add("System.TeamProject");
            _ignore.Add("System.NodeName");
            _ignore.Add("System.RelatedLinkCount");
            _ignore.Add("System.WorkItemType");
            _ignore.Add("Microsoft.VSTS.Common.ActivatedDate");
            _ignore.Add("Microsoft.VSTS.Common.StateChangeDate");
            _ignore.Add("System.ExternalLinkCount");
            _ignore.Add("System.HyperLinkCount");
            _ignore.Add("System.Watermark");
            _ignore.Add("System.AuthorizedDate");
            _ignore.Add("System.BoardColumn");
            _ignore.Add("System.BoardColumnDone");
            _ignore.Add("System.BoardLane");
        }

        internal override void InternalExecute()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            WorkItemStoreContext sourceStore = new WorkItemStoreContext(me.Source, WorkItemStoreFlags.BypassRules);
            TfsQueryContext tfsqc = new TfsQueryContext(sourceStore);
            tfsqc.AddParameter("TeamProject", me.Source.Name);
            tfsqc.Query = string.Format(@"SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @TeamProject {0} ORDER BY [System.ChangedDate] desc", _config.QueryBit);
            WorkItemCollection sourceWIS = tfsqc.Execute();
            Trace.WriteLine(string.Format("Migrate {0} work items?", sourceWIS.Count), this.Name);

            WorkItemStoreContext targetStore = new WorkItemStoreContext(me.Target, WorkItemStoreFlags.BypassRules);
            Project destProject = targetStore.GetProject();
            Trace.WriteLine(string.Format("Found target project as {0}", destProject.Name), this.Name);

            int current = sourceWIS.Count;
            int count = 0;
            long elapsedms = 0;

            foreach (WorkItem sourceWI in sourceWIS)
            {
                Stopwatch witstopwatch = new Stopwatch();
                witstopwatch.Start();

                WorkItem targetFound = targetStore.FindReflectedWorkItem(sourceWI, me.ReflectedWorkItemIdFieldName, false);
                Trace.WriteLine(string.Format("{0} - Migrating: {1}-{2}", current, sourceWI.Id, sourceWI.Type.Name), Name);

                if (targetFound == null)
                {
                    // Decide on WIT
                    if (!me.WorkItemTypeDefinitions.ContainsKey(sourceWI.Type.Name))
                    {
                        Trace.WriteLine("...not supported", Name);
                        continue;
                    }

                    WorkItem newwit = CreateAndPopulateWorkItem(_config, sourceWI, destProject, me.WorkItemTypeDefinitions[sourceWI.Type.Name].Map(sourceWI));
                    if (newwit.Fields.Contains(me.ReflectedWorkItemIdFieldName))
                    {
                        newwit.Fields[me.ReflectedWorkItemIdFieldName].Value = sourceStore.CreateReflectedWorkItemId(sourceWI);
                    }

                    me.ApplyFieldMappings(sourceWI, newwit);

                    ArrayList fails = newwit.Validate();
                    foreach (Field f in fails)
                    {
                        Trace.WriteLine(string.Format("{0} - Invalid: {1}-{2}-{3}", current, sourceWI.Id, sourceWI.Type.Name, f.ReferenceName), this.Name);
                    }

                    if (newwit != null)
                    {
                        try
                        {
                            if (_config.UpdateCreatedDate)
                            {
                                newwit.Fields["System.CreatedDate"].Value = sourceWI.Fields["System.CreatedDate"].Value;
                            }

                            if (_config.UpdateCreatedBy)
                            {
                                newwit.Fields["System.CreatedBy"].Value = sourceWI.Fields["System.CreatedBy"].Value;
                            }

                            newwit.Save();
                            newwit.Close();

                            Trace.WriteLine(string.Format("...Saved as {0}", newwit.Id), this.Name);
                            if (sourceWI.Fields.Contains(me.ReflectedWorkItemIdFieldName) && _config.UpdateSoureReflectedId)
                            {
                                sourceWI.Fields[me.ReflectedWorkItemIdFieldName].Value = targetStore.CreateReflectedWorkItemId(newwit);
                            }

                            sourceWI.Save();
                            Trace.WriteLine(string.Format("...and Source Updated {0}", sourceWI.Id), this.Name);
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine("...FAILED to Save", this.Name);
                            foreach (Field f in newwit.Fields)
                            {
                                Trace.WriteLine(string.Format("{0} | {1}", f.ReferenceName, f.Value), this.Name);
                            }
                            Trace.WriteLine(ex.ToString(), this.Name);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("...Exists");
                }

                sourceWI.Close();
                witstopwatch.Stop();
                elapsedms = elapsedms + witstopwatch.ElapsedMilliseconds;
                current--;
                count++;
                TimeSpan average = new TimeSpan(0, 0, 0, 0, (int)(elapsedms / count));
                TimeSpan remaining = new TimeSpan(0, 0, 0, 0, (int)(average.TotalMilliseconds * current));
                Trace.WriteLine(string.Format("Average time of {0} per work item and {1} estimated to completion", string.Format(@"{0:s\:fff} seconds", average), string.Format(@"{0:%h} hours {0:%m} minutes {0:s\:fff} seconds", remaining)), this.Name);
                Trace.Flush();
            }

            stopwatch.Stop();
            Console.WriteLine(@"DONE in {0:%h} hours {0:%m} minutes {0:s\:fff} seconds", stopwatch.Elapsed);
        }

        private WorkItem CreateAndPopulateWorkItem(WorkItemMigrationConfig config, WorkItem oldWi, Project destProject, String destType)
        {
            Stopwatch fieldMappingTimer = new Stopwatch();
            Trace.Write("... Building", "WorkItemMigrationContext");

            var NewWorkItemstartTime = DateTime.UtcNow;
            Stopwatch NewWorkItemTimer = new Stopwatch();
            WorkItem newwit = destProject.WorkItemTypes[destType].NewWorkItem();
            NewWorkItemTimer.Stop();
            Telemetry.Current.TrackDependency("TeamService", "NewWorkItem", NewWorkItemstartTime, NewWorkItemTimer.Elapsed, true);
            Trace.WriteLine(string.Format("Dependancy: {0} - {1} - {2} - {3} - {4}", "TeamService", "NewWorkItem", NewWorkItemstartTime, NewWorkItemTimer.Elapsed, true), "WorkItemMigrationContext");

            foreach (Field f in oldWi.Fields)
            {
                if (newwit.Fields.Contains(f.ReferenceName) && !_ignore.Contains(f.ReferenceName) && newwit.Fields[f.ReferenceName].IsEditable)
                {
                    newwit.Fields[f.ReferenceName].Value = oldWi.Fields[f.ReferenceName].Value;
                }
            }

            if (config.PrefixProjectToNodes)
            {
                newwit.AreaPath = string.Format(@"{0}\{1}", newwit.Project.Name, oldWi.AreaPath);
                newwit.IterationPath = string.Format(@"{0}\{1}", newwit.Project.Name, oldWi.IterationPath);
            }
            else
            {
                var regex = new Regex(Regex.Escape(oldWi.Project.Name));
                newwit.AreaPath = regex.Replace(oldWi.AreaPath, newwit.Project.Name, 1);
                newwit.IterationPath = regex.Replace(oldWi.IterationPath, newwit.Project.Name, 1);
            }

            if (newwit.Fields.Contains("Microsoft.VSTS.Common.BacklogPriority")
                && newwit.Fields["Microsoft.VSTS.Common.BacklogPriority"].Value != null
                && !isNumeric(newwit.Fields["Microsoft.VSTS.Common.BacklogPriority"].Value.ToString(), NumberStyles.Any))
            {
                newwit.Fields["Microsoft.VSTS.Common.BacklogPriority"].Value = 10;
            }

            newwit.History = GetHistory(oldWi);

            Trace.WriteLine("...buildComplete", "WorkItemMigrationContext");

            fieldMappingTimer.Stop();
            Telemetry.Current.TrackMetric("FieldMappingTime", fieldMappingTimer.ElapsedMilliseconds);
            Trace.WriteLine(string.Format("FieldMapOnNewWorkItem: {0} - {1}", NewWorkItemstartTime, fieldMappingTimer.Elapsed.ToString("c")), "WorkItemMigrationContext");

            return newwit;
        }

        private static string GetHistory(WorkItem oldWi)
        {
            if (oldWi.Revisions.Count == 0)
            {
                return string.Empty;
            }

            var exceptedFields = new string[]
            {
                "History",
                "Changed By",
                "Changed Date",
                "Watermark",
                "Authorized Date",
                "Authorized As",
                "Revised Date"
            };

            var history = new StringBuilder();
            history.Append("<p>History from previous work item:</p>");
            history.Append("<table border='1' style='width:100%;border-color:#C0C0C0;'>");

            for (var i = oldWi.Revisions.Count - 1; i >= 0; i--)
            {
                history.Append("<tr>");
                history.Append("<tr><td style='align:right;width:100%'>");

                // add who and when
                var revision = oldWi.Revisions[i];
                history.AppendFormat("<p><b>{0} on {1}</b></p>",
                    revision.Fields["System.ChangedBy"].Value,
                    DateTime.Parse(revision.Fields["System.ChangedDate"].Value.ToString()).ToString("F"));

                // add comment if one
                var historyValue = revision.Fields["System.History"].Value as string;
                if (!string.IsNullOrEmpty(historyValue))
                {
                    history.AppendFormat("<p>{0}</p>", historyValue);
                }

                // add changed fields table (except certain fields)
                history.Append("<table border='1' style='border-color:#C0C0C0;'>");
                history.Append("<tr><th>Field</th><th>Old Value</th><th>New Value</th></tr>");

                var externalLinksChanged = false;
                var fields = revision.Fields;
                foreach (Field field in fields)
                {
                    if (exceptedFields.Contains(field.Name))
                    {
                        continue;
                    }

                    var original = field.OriginalValue == null ? string.Empty : field.OriginalValue.ToString();
                    var value = field.Value == null ? string.Empty : field.Value.ToString();
                    if (original != value)
                    {
                        if (field.Name == "External Link Count")
                        {
                            externalLinksChanged = true;
                        }

                        history.AppendFormat("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>", field.Name, original, value);
                    }
                }

                history.Append("</table>");

                var links = revision.Links;
                if (revision.Links.Count > 0 && externalLinksChanged)
                {
                    history.Append("<table border='1' style='border-color:#C0C0C0;'>");
                    history.Append("<tr><th>Link Type</th><th>Description</th></tr>");

                    foreach (Link link in links)
                    {
                        var description = string.Empty;
                        switch (link.BaseType)
                        {
                            case BaseLinkType.ExternalLink:
                                description = ((ExternalLink)link).LinkedArtifactUri;
                                break;

                            case BaseLinkType.Hyperlink:
                                description = ((Hyperlink)link).Location;
                                break;
                        }

                        history.AppendFormat("<tr><td>{0}</td><td>{1}</td></tr>", link.BaseType.ToString(), description);
                    }

                    history.Append("</table>");
                }

                history.Append("</td></tr>");
            }

            history.Append("</table>");
            history.Append("<p>Migrated by <a href='http://nkdagility.com'>naked Agility Limited's</a> open source <a href='https://github.com/nkdAgility/VstsMigrator'>VSTS/TFS Migrator</a>.</p>");

            // ensure not too big < 1048575
            if (history.Length > 1048575)
            {
                history = history.Remove(1048575, history.Length - 1048576);
            }

            return history.ToString();
        }

        static bool isNumeric(string val, NumberStyles NumberStyle)
        {
            Double result;
            return Double.TryParse(val, NumberStyle,
                System.Globalization.CultureInfo.CurrentCulture, out result);
        }
    }
}