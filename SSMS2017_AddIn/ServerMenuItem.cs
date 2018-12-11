using System.Linq;
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
            var navigableItem = (INavigableItem)this.Parent.GetType()
                .GetProperty("NavigableItem", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.GetValue(this.Parent);
            TreeNode treeNode = null;
            for (var i = 0; i < this.treeView.Nodes.Count; i++)
            {
                if (navigableItem != null && this.treeView.Nodes[i].Text == navigableItem.DisplayName)
                {
                    treeNode = this.treeView.Nodes[i];
                }
            }
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