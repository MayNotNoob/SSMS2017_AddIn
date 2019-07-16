//------------------------------------------------------------------------------
// <copyright file="SearchBar.cs" company="LabSolvay">
//     Copyright (c) LabSolvay.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.SqlStudio.Explorer;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SearchBar
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SearchBar
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("40dde1c2-3123-4c0a-8656-32deafcdf0fc");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchBar"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private SearchBar(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            this.package = package;

            //OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            //if (commandService != null)
            //{
            //    var menuCommandID = new CommandID(CommandSet, CommandId);
            //    var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
            //    commandService.AddCommand(menuItem);
            //}
            AddControlEdit();
        }

        private void AddControlEdit()
        {
            IObjectExplorerService objectExplorer = (ObjectExplorerService)this.ServiceProvider.GetService(typeof(IObjectExplorerService));
            var treeViewProp = objectExplorer.GetType().GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);
            var treeView = (TreeView)treeViewProp.GetValue(objectExplorer, null);
            TreeNode node=new TreeNode("xxx");

            
            treeView.Nodes.Add(node);
            node.BeginEdit();
            //try
            //{
            //    var Dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            //    var commandBars = (CommandBars)Dte.CommandBars;
            //    CommandBar standard = commandBars["MenuBar"];
            //    var control = standard.Controls.Add(MsoControlType.msoControlEdit, Type.Missing, Type.Missing, 1, true);
            //    control.Caption = "Test";
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show(e.Message);
            //}

        }
        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SearchBar Instance
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
            Instance = new SearchBar(package);
        }
    }
}
