//------------------------------------------------------------------------------
// <copyright file="Scriptor.cs" company="LabSolvay">
//     Copyright (c) LabSolvay.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
        private static bool IsSearchBarAdded = false;
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

        private static Dictionary<String, String> dataBaseMap = new Dictionary<String, String>
        {
            ["estrade_risk"] = "ESTRADE",
            ["Dispatching"] = "ENERGY",
            ["ELDSES"] = "ELDSES"
        };
        private static List<TreeNode> filteredNodes = new List<TreeNode>();

        private static String exText = "";

        private static TreeNode selectedNode = null;

        private static TreeNode filteredNode = null;

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
                    if (node.UrnPath == "Server")
                    {
                        if (!IsServerMenuAdded)
                        {
                            var tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
                            PropertyInfo treeViewProp = objectExplorer.GetType().GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);
                            if (treeViewProp != null)
                            {
                                var treeView = (TreeView)treeViewProp.GetValue(objectExplorer, null);
                                var item = new ServerMenuItem(treeView);
                                tableMenu.AddChild(string.Empty, item);
                                treeView.AfterExpand += FilterDataBase;
                                treeView.AfterSelect += (o, args) =>
                                {
                                    selectedNode = args.Node;
                                };
                            }
                            IsServerMenuAdded = true;
                        }
                    }
                }
                if (!IsSearchBarAdded)
                {
                    PropertyInfo treeViewProp = objectExplorer.GetType().GetProperty("Tree", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    var treeView = (TreeView)treeViewProp.GetValue(objectExplorer, null);
                    var searchNode = new SearchNode(treeView, "search bar");
                    treeView.Nodes.Insert(0, searchNode);
                    treeView.Controls.Add(searchNode.TextBox);
                    searchNode.TextBox.AutoSize = false;
                    searchNode.TextBox.Size = new Size(treeView.Bounds.Width, searchNode.Bounds.Height);
                    searchNode.TextBox.KeyDown += OnKeyDown;
                    IsSearchBarAdded = true;
                }
                if (IsSearchBarAdded && IsServerMenuAdded)
                    contextService.ActionContext.CurrentContextChanged -= ActionContextOnCurrentContextChanged;
            }
            catch (Exception objectExplorerContextException)
            {
                MessageBox.Show(objectExplorerContextException.Message);
            }
        }

        private void OnKeyDown(Object sender, KeyEventArgs args)
        {
            try
            {
                if (args.KeyCode != Keys.Enter) return;
                var text = ((TextBox)sender).Text.ToLower();
                if (exText == text && selectedNode == filteredNode) return;
                if (selectedNode == null) return;
                if (selectedNode != filteredNode && filteredNodes.Any())
                {
                    ClearFilter(filteredNode);
                }
                if (text == "" && filteredNodes.Any())
                {
                    ClearFilter(selectedNode);
                }
                else if (text.StartsWith(exText))
                {
                    FilterNode(text);
                }
                else if (exText.StartsWith(text))
                {
                    if (filteredNodes.Any())
                    {
                        filteredNodes.ForEach(n =>
                        {
                            if (n.Text.ToLower().Contains(text))
                            {
                                selectedNode.Nodes.Add(n);
                            }
                        });
                        filteredNodes.RemoveAll(n => n.Text.Contains(text));
                    }
                }
                else
                {
                    ClearFilter(selectedNode);
                    FilterNode(text);
                }
                exText = text;
                filteredNode = selectedNode;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void ClearFilter(TreeNode node)
        {
            if (node == null) return;
            node.Nodes.AddRange(filteredNodes.ToArray());
            filteredNodes.Clear();
        }

        private void FilterNode(String text)
        {
            for (var i = 0; i < selectedNode.Nodes.Count; i++)
            {
                if (selectedNode.Nodes[i].Text.ToLower().Contains(text)) continue;
                filteredNodes.Add(selectedNode.Nodes[i]);
            }
            filteredNodes.ForEach(n => n.Remove());
        }

        private async void FilterDataBase(Object sender, TreeViewEventArgs args)
        {
            ((TreeView)sender).SelectedNode = args.Node;
            if (args.Node.Text.StartsWith("Databases"))
            {
                await System.Threading.Tasks.Task.Delay(100);
                if (args.Node.Nodes.Count == 1) return;
                String baseName = GetBaseName(args.Node.Parent);
                if (!String.IsNullOrEmpty(baseName))
                {
                    var toDelete = new List<TreeNode>();
                    for (var i = 0; i < args.Node.Nodes.Count; i++)
                    {
                        if (!args.Node.Nodes[i].Text.StartsWith(baseName))
                            toDelete.Add(args.Node.Nodes[i]);
                    }
                    toDelete.ForEach(n => n.Remove());
                }
            }
        }

        private String GetBaseName(TreeNode node)
        {
            String baseName = "";
            String text = node.Text;
            String tag = node.Tag == null ? "" : node.Tag.ToString();
            foreach (String key in dataBaseMap.Keys)
            {
                if (text.Contains(key) || tag.Contains(key))
                {
                    baseName = dataBaseMap[key];
                    break;
                }
            }
            return baseName;
        }

        public class SearchNode : TreeNode
        {
            public TextBox TextBox { get; set; }

            public SearchNode(TreeView treeView, String text) : base(text)
            {
                TextBox = new TextBox { Tag = treeView };
            }
        }
    }
}