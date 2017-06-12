using Microsoft.TeamFoundation.TestManagement.Client;
using System.Diagnostics;
using System.Linq;
using VstsSyncMigrator.Engine.ComponentContext;
using VstsSyncMigrator.Engine.Configuration.Processing;

namespace VstsSyncMigrator.Engine
{
    public class TestConfigurationsMigrationContext : MigrationContextBase
    {
        public override string Name
        {
            get
            {
                return "TestConfigurationsMigrationContext";
            }
        }

        public TestConfigurationsMigrationContext(MigrationEngine me, TestConfigurationsMigrationConfig config) : base(me, config)
        {
        }

        internal override void InternalExecute()
        {
            var sourceTmc = new TestManagementContext(me.Source);
            var targetTmc = new TestManagementContext(me.Target);

            ITestConfigurationCollection tc = sourceTmc.Project.TestConfigurations.Query("Select * From TestConfiguration");
            Trace.WriteLine($"Plan to copy {tc.Count} Configurations", "TestConfigurationMigration");

            foreach (var sourceTestConf in tc)
            {
                Trace.WriteLine($"{sourceTestConf.Name} - Attempting to copy configuration", "TestConfigurationMigration");

                var targetTc = GetCon(targetTmc.Project.TestConfigurations, sourceTestConf.Name);
                if (targetTc != null)
                {
                    Trace.WriteLine($"{sourceTestConf.Name} - Already exists in target", "TestConfigurationMigration");
                    continue;
                }

                Trace.WriteLine($"{sourceTestConf.Name} - Creating new configuration in target", "TestConfigurationMigration");
                targetTc = targetTmc.Project.TestConfigurations.Create();
                targetTc.AreaPath = sourceTestConf.AreaPath.Replace(me.Source.Name, me.Target.Name);
                targetTc.Description = sourceTestConf.Description;
                targetTc.IsDefault = sourceTestConf.IsDefault;
                targetTc.Name = sourceTestConf.Name;
                targetTc.State = sourceTestConf.State;

                foreach (var val in sourceTestConf.Values)
                {
                    targetTc.Values.Add(val);
                }

                targetTc.Save();

                Trace.WriteLine($"{sourceTestConf.Name} - Created in target as {targetTc.Name}", "TestConfigurationMigration");
            }
        }

        internal ITestConfiguration GetCon(ITestConfigurationHelper tch, string configToFind)
        {
            return (from tv in tch.Query("Select * From TestConfiguration") where tv.Name == configToFind select tv).SingleOrDefault();
        }
    }
}