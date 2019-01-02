//------------------------------------------------------------------------------
// <copyright file="Scriptor.cs" company="LabSolvay">
//     Copyright (c) LabSolvay.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Reflection;
using System.Text;
using Microsoft.VisualStudio.Shell;
using Microsoft.SqlServer.Management.SqlStudio.Explorer;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace SSMS2017_AddIn
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Scriptor
    {
        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private readonly ContextService contextService;
        private static bool IsTableMenuAdded = false;
        private static bool IsColumnMenuAdded = false;
        private static bool IsServerMenuAdded = false;
        /// <summary>
        /// Initializes a new instance of the <see cref="Scriptor"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Scriptor(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            this.package = package;
            AddWindowMenu();
            contextService = (ContextService)this.ServiceProvider.GetService(typeof(IContextService));
            contextService.ActionContext.CurrentContextChanged += ActionContextOnCurrentContextChanged;
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Scriptor Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => this.package;

        private DTE2 Dte => (DTE2)this.ServiceProvider.GetService(typeof(DTE));
        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Scriptor(package);
        }

        private void AddWindowMenu()
        {
            var commandBars = (CommandBars)Dte.CommandBars;
            CommandBar tabContext = commandBars["SQL Files Editor Context"];
            var btnScripts =
                tabContext.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true)
                    as CommandBarButton;
            if (btnScripts != null)
            {
                btnScripts.Caption = "Curve Name Scripts";
                btnScripts.Click += (CommandBarButton ctrl, ref bool cancelDefault) =>
                {
                    var doc = (TextDocument)this.Dte.ActiveDocument.Object("TextDocument");
                    var content = doc.StartPoint.CreateEditPoint().GetText(doc.EndPoint);
                    if (!string.IsNullOrEmpty(content))
                    {
                        content = content.Replace("\r", "");
                        var arr = content.Split('\n');
                        StringBuilder builder = new StringBuilder();
                        foreach (string s in arr)
                        {
                            var arr2 = s.Split(',');
                            if (arr2.Length == 2)
                            {
                                builder.AppendLine("INSERT INTO [ESTRADE].[dbo].[CurlingCurveNameMatching] VALUES ('" + arr2[1] + "','" + arr2[1] + "','NN')");
                                builder.AppendLine("INSERT INTO [ESTRADE].[dbo].[CurlingCurveIdMatching] VALUES ('" + arr2[0] + "','" + arr2[1] + "')");
                                builder.AppendLine("INSERT INTO [ESTRADE].[dbo].[CurlingCurveHeader] VALUES ('" + arr2[1] + "','EUR','',1,1)");
                            }
                        }
                        if (builder.Length > 0)
                        {
                            doc.StartPoint.CreateEditPoint().Delete(doc.EndPoint);
                            content = "/*\n" + content + "\n*/\n" + builder;
                            doc.EndPoint.CreateEditPoint().Insert(content);
                        }
                    }
                };
            }
        }

        private void ActionContextOnCurrentContextChanged(object sender, EventArgs e)
        {
            try
            {
                INodeInformation[] nodes;
                int nodeCount;
                IObjectExplorerService objectExplorer = (ObjectExplorerService)this.ServiceProvider.GetService(typeof(IObjectExplorerService));
                objectExplorer.GetSelectedNodes(out nodeCount, out nodes);
                var node = nodeCount > 0 ? nodes[0] : null;
                if (node != null)
                {
                    if (node.Parent != null && node.Parent.InvariantName == "UserTables")
                    {
                        if (!IsTableMenuAdded)
                        {
                            var tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
                            var item = new SqlTableMenuItem(ScriptorPackage.applicationObject);
                            tableMenu.AddChild(string.Empty, item);
                            IsTableMenuAdded = true;
                        }
                    }
                    else
                    if (node.UrnPath == "Server")
                    {
                        if (!IsServerMenuAdded)
                        {
                            var tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
                            var treeViewProp = objectExplorer.GetType().GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);
                            if (treeViewProp != null)
                            {
                                var treeView = (TreeView)treeViewProp.GetValue(objectExplorer, null);
                                var item = new ServerMenuItem(treeView);
                                tableMenu.AddChild(string.Empty, item);
                            }
                            IsServerMenuAdded = true;
                        }
                    }
                    if (IsServerMenuAdded && IsTableMenuAdded)
                        contextService.ActionContext.CurrentContextChanged -= ActionContextOnCurrentContextChanged;
                }
            }
            catch (Exception objectExplorerContextException)
            {
                MessageBox.Show(objectExplorerContextException.Message);
            }
        }
    }
}