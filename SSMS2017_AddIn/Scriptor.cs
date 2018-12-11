//------------------------------------------------------------------------------
// <copyright file="Scriptor.cs" company="LabSolvay">
//     Copyright (c) LabSolvay.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Reflection;
using Microsoft.VisualStudio.Shell;
using Microsoft.SqlServer.Management.SqlStudio.Explorer;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using System.Windows.Forms;

namespace SSMS2017_AddIn
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Scriptor
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("e658de75-add3-4bcb-9013-4fd905b55a59");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;


        private CommandID menuCommandID;
        private MenuCommand menuItem;
        private ContextService contextService;
        private static bool IsTableMenuAdded = false;
        private static bool IsColumnMenuAdded = false;
        private static bool IsServerMenuAdded = false;
        private static HierarchyObject tableMenu = null;
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

            var commandService = (OleMenuCommandService)this.ServiceProvider.GetService(typeof(IMenuCommandService));

            if (commandService != null)
            {
                //menuCommandID = new CommandID(CommandSet, CommandId);
                //menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                //commandService.AddCommand(menuItem);
                contextService = (ContextService)this.ServiceProvider.GetService(typeof(IContextService));
                contextService.ActionContext.CurrentContextChanged += ActionContextOnCurrentContextChanged;
            }
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
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Scriptor(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            //var commandService = (OleMenuCommandService)this.ServiceProvider.GetService(typeof(IMenuCommandService));
            //commandService.RemoveCommand(menuItem);
            //contextService = (ContextService)this.ServiceProvider.GetService(typeof(IContextService));
            //contextService.ActionContext.CurrentContextChanged += ActionContextOnCurrentContextChanged;
        }
     
        private void ActionContextOnCurrentContextChanged(object sender, EventArgs e)
        {
            try
            {
                INodeInformation[] nodes;
                INodeInformation node;
                int nodeCount;
                IObjectExplorerService objectExplorer = (ObjectExplorerService)this.ServiceProvider.GetService(typeof(IObjectExplorerService));
                objectExplorer.GetSelectedNodes(out nodeCount, out nodes);
                node = nodeCount > 0 ? nodes[0] : null;
                if (node != null)
                {
                    if (node.Parent != null && node.Parent.InvariantName == "UserTables")
                    {
                        if (!IsTableMenuAdded)
                        {
                            tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
                            SqlTableMenuItem item = new SqlTableMenuItem(ScriptorPackage.applicationObject);
                            tableMenu.AddChild(string.Empty, item);
                            IsTableMenuAdded = true;
                        }
                    }
                    else if (node.UrnPath == "Server")
                    {
                        if (!IsServerMenuAdded)
                        {
                            tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
                            var treeViewProp = objectExplorer.GetType().GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);
                            if (treeViewProp != null)
                            {
                                var treeView = (TreeView)treeViewProp.GetValue(objectExplorer, null);
                                ServerMenuItem item = new ServerMenuItem(treeView);
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