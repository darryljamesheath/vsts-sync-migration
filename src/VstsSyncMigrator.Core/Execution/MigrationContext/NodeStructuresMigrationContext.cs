using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml;
using VstsSyncMigrator.Engine.Configuration.Processing;

namespace VstsSyncMigrator.Engine
{
    public class NodeStructuresMigrationContext : MigrationContextBase
    {
        private readonly NodeStructuresMigrationConfig config;
        private ICommonStructureService sourceCss;
        
        public override string Name
        {
            get
            {
                return "NodeStructuresMigrationContext";
            }
        }

        public NodeStructuresMigrationContext(MigrationEngine me, NodeStructuresMigrationConfig config) : base(me, config)
        {
            this.config = config;
        }

        internal override void InternalExecute()
        {
            sourceCss = me.Source.Collection.GetService<ICommonStructureService>();

            ProjectInfo sourceProjectInfo = sourceCss.GetProjectFromName(me.Source.Name);
            NodeInfo[] sourceNodes = sourceCss.ListStructures(sourceProjectInfo.Uri);
            ICommonStructureService targetCss = (ICommonStructureService)me.Target.Collection.GetService(typeof(ICommonStructureService));

            ProcessCommonStructure("Area", sourceNodes, targetCss, sourceCss);
            ProcessCommonStructureIterations(sourceNodes, sourceCss);
        }

        private void ProcessCommonStructure(string treeType, NodeInfo[] sourceNodes, ICommonStructureService targetCss, ICommonStructureService sourceCss)
        {
            NodeInfo sourceNode = (from n in sourceNodes where n.Path.Contains(treeType) select n).Single();
            XmlElement sourceTree = sourceCss.GetNodesXml(new string[] { sourceNode.Uri }, true);
            NodeInfo structureParent = targetCss.GetNodeFromPath(string.Format("\\{0}\\{1}", me.Target.Name, treeType));
            if (config.PrefixProjectToNodes)
            {
                structureParent = CreateNode(targetCss, me.Source.Name, structureParent);
            }
            if (sourceTree.ChildNodes[0].HasChildNodes)
            {
                CreateNodes(sourceTree.ChildNodes[0].ChildNodes[0].ChildNodes, targetCss, structureParent);
            }
        }

        private void CreateNodes(XmlNodeList nodeList, ICommonStructureService css, NodeInfo parentPath)
        {
            foreach (XmlNode item in nodeList)
            {
                string newNodeName = item.Attributes["Name"].Value;
                NodeInfo targetNode = CreateNode(css, newNodeName, parentPath);
                if (item.HasChildNodes)
                {
                    CreateNodes(item.ChildNodes[0].ChildNodes, css, targetNode);
                }
            }
        }

        private NodeInfo CreateNode(ICommonStructureService css, string name, NodeInfo parent)
        {
            string nodePath = string.Format(@"{0}\{1}", parent.Path, name);
            NodeInfo node = null;
            Trace.Write(string.Format("--CreateNode: {0}", nodePath));
            try
            {
                node = css.GetNodeFromPath(nodePath);
                Trace.Write("...found");
            }
            catch (CommonStructureSubsystemException ex)
            {
                Telemetry.Current.TrackException(ex);
                Trace.Write("...missing");
                string newPathUri = css.CreateNode(name, parent.Uri);
                Trace.Write("...created");
                node = css.GetNode(newPathUri);
            }
            return node;
        }

        private void ProcessCommonStructureIterations(NodeInfo[] sourceNodes, ICommonStructureService sourceCss)
        {
            var creds = new VssClientCredentials(true) { PromptType = CredentialPromptType.PromptIfNeeded };
            var connection = new VssConnection(new Uri(me.Target.Collection.Name), creds);
            var client = connection.GetClient<WorkItemTrackingHttpClient>();

            var sourceNode = (from n in sourceNodes where n.Path.Contains("Iteration") select n).Single();
            var sourceTree = sourceCss.GetNodesXml(new string[] { sourceNode.Uri }, true);

            var rootPath = string.Empty;
            if (config.PrefixProjectToNodes)
            {
                CreateIterationNode(client, me.Source.Name, sourceNode.StartDate, sourceNode.FinishDate, null);
                rootPath = me.Source.Name;
            }

            CreateIterationNodes(client, sourceTree.ChildNodes[0].ChildNodes[0].ChildNodes, rootPath);
        }

        private void CreateIterationNodes(WorkItemTrackingHttpClient client, XmlNodeList nodeList, string path)
        {
            foreach (XmlNode xmlNode in nodeList)
            {
                var nodeId = GetNodeID(xmlNode.OuterXml);
                var nodeInfo = sourceCss.GetNode(nodeId);
                var parentPath = CreateIterationNode(client, nodeInfo.Name, nodeInfo.StartDate, nodeInfo.FinishDate, path);

                if (xmlNode.HasChildNodes)
                {
                    CreateIterationNodes(client, xmlNode.ChildNodes[0].ChildNodes, parentPath);
                }
            }
        }

        private string CreateIterationNode(WorkItemTrackingHttpClient client, string nodeName, DateTime? startDate, DateTime? finishDate, string path)
        {
            Trace.WriteLine($"node:{nodeName} for path:{path} - checking...", "Iterations");

            var node = new WorkItemClassificationNode
            {
                StructureType = TreeNodeStructureType.Iteration,
                Name = nodeName,
            };

            if (startDate.HasValue && finishDate.HasValue)
            {
                node.Attributes = new Dictionary<string, object>();
                node.Attributes.Add("startDate", startDate);
                node.Attributes.Add("finishDate", finishDate);
            }

            var newPath = string.Concat(path ?? string.Empty, "\\", nodeName);

            try
            {
                var newPathNode = client.GetClassificationNodeAsync(me.Target.Name, TreeStructureGroup.Iterations, newPath).Result;
                if (newPathNode != null)
                {
                    Trace.WriteLine($"node:{nodeName} for path:{path} - checking...found", "Iterations");
                }
            }
            catch (Exception)
            {
                // not the best but only way to determine if node doesn't already exist.
                Trace.WriteLine($"node:{nodeName} for path:{path} - checking...not found so creating...", "Iterations");

                var newnode = client.CreateOrUpdateClassificationNodeAsync(node, me.Target.Name, TreeStructureGroup.Iterations, path).Result;
                if (newnode == null)
                {
                    Trace.WriteLine($"node:{nodeName} for path:{path} - checking...not found so creating...failed", "Iterations");
                }
                else
                {
                    Trace.WriteLine($"node:{nodeName} for path:{path} - checking...not found so creating...success", "Iterations");
                }
            }

            return newPath;
        }

        private string GetNodeID(string xml)
        {
            var first = "NodeID=\"";
            var start = xml.IndexOf(first) + first.Length;
            var end = xml.IndexOf("\"", start);
            return xml.Substring(start, (end - start));
        }
    }
}