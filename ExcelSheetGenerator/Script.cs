using EsapiEssentials.Plugin;
using ExcelSheetGenerator;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using VMS.TPS.Common.Model.API;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script : ScriptBase
    {
        const string SCRIPT_NAME = "Excel Sheet Generator";
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public override void Run(PluginScriptContext context)
        {

            PlanSetup planSetup = context.PlanSetup;

            //check loading data 
            if (context.Patient == null || context.Course == null || planSetup == null)
            {
                MessageBox.Show("Please load a plan before running this script.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            //Plan approval status check
            if (planSetup.ApprovalStatus != VMS.TPS.Common.Model.Types.PlanSetupApprovalStatus.PlanningApproved &&
                planSetup.ApprovalStatus != VMS.TPS.Common.Model.Types.PlanSetupApprovalStatus.TreatmentApproved)
            {
                MessageBox.Show("Plan approval status is not PlanningApproved or TreatmentApproved.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            var mainWindow = new MainWindow(context);
            mainWindow.ShowDialog();
        }
    }
}
