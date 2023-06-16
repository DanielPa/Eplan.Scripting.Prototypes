using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Eplan.EplApi.Scripting;

namespace DanielPa.Scripts
{
    /// <summary>
    /// With the help of System.Reflection this script creates an instance of ProjectManager, sets its property
    /// LockProjectByDefault to false (else you get issues because Lock can't be created), gets the selected project as
    /// object and invokes its RemoveAllPages method.
    /// The DataModelu assembly can be picked form the current AppDomain because it seems that EPLAN only has a single
    /// AppDomain for the main App, API Add-Ins and Scripts. 
    /// </summary>
    public class ScriptDataModelTest
    {
        [Start]
        public void RemoveAllPages()
        {
            var dataModelAssembly = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => a.FullName.StartsWith("Eplan.EplApi.DataModelu"));
            if (dataModelAssembly != null)
            {
                var projectManagerType = dataModelAssembly.ExportedTypes.FirstOrDefault(t => t.Name == "ProjectManager");
                if (projectManagerType != null)
                {
                    var projectManager = Activator.CreateInstance(projectManagerType);
                    MethodInfo getCurrentProjectWithDialog = projectManagerType.GetMethod("GetCurrentProjectWithDialog");
                    var lockProjectByDefault = projectManagerType.GetProperty("LockProjectByDefault", BindingFlags.Public | BindingFlags.Instance);
                    if (lockProjectByDefault != null)
                    {
                        lockProjectByDefault.SetValue(projectManager, false);
                    }

                    try
                    {
                        if (getCurrentProjectWithDialog != null)
                        {
                            var project = getCurrentProjectWithDialog.Invoke(projectManager, new object[] { });
                            var removeAllPages = project.GetType().GetMethod("RemoveAllPages");
                            if (removeAllPages != null)
                            {
                                removeAllPages.Invoke(project, new object[] { });
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Debugger.Break();
                    }
                }
            }
        }
    }
}