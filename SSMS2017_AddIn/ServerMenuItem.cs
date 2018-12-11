﻿using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;

namespace SSMS2017_AddIn
{
    public class ServerMenuItem : ToolsMenuItemBase, IWinformsMenuHandler
    {
        //private INodeInformation node;

        private TreeView treeView;

        public ServerMenuItem(TreeView treeView)
        {
            //this.node = node;
            this.treeView = treeView;
        }

        public override object Clone()
        {
            return new ServerMenuItem(null);
        }

        protected override void Invoke()
        {
        }

        public ToolStripItem[] GetMenuItems()
        {
            ToolStripMenuItem item = new ToolStripMenuItem("Rename");
            item.Click += Item_Click;
            return new ToolStripItem[] { item };
        }

        private void Item_Click(object sender, System.EventArgs e)
        {
            TreeNode treeNode = treeView.SelectedNode;
            if (treeNode != null)
            {
                string input =
                    Microsoft.VisualBasic.Interaction.InputBox("Please input a name", "Rename node");
                if (!string.IsNullOrEmpty(input))
                {
                    treeNode.Text = input.ToUpper();
                }
            }
        }
    }
}