using EnvDTE80;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Sdk.Sfc;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SSMS2017_AddIn
{
    class SqlTableMenuItem : ToolsMenuItemBase, IWinformsMenuHandler
    {
        #region Class Variables
        private readonly DTE2 applicationObject;
        private readonly DTEApplicationController dteController = null;
        private readonly Regex tableRegex = null;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="SqlTableMenuItem"/> class.
        /// </summary>
        /// <param name="applicationObject">The application object.</param>
        public SqlTableMenuItem(DTE2 applicationObject)
        {
            this.applicationObject = applicationObject;
            this.dteController = new DTEApplicationController();
            this.tableRegex = new Regex(@"^Server\[[^\]]*\]/Database\[[^\]]*\]/Table\[[^\]]*\]$");
        }
        #endregion

        #region Override Methods
        /// <summary>
        /// Invokes this instance.
        /// </summary>
        protected override void Invoke() { }

        /// <summary>
        /// Clones this instance.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return new SqlTableMenuItem(null);
        }
        #endregion

        #region IWinformsMenuHandler Members
        /// <summary>
        /// Gets the menu items.
        /// </summary>
        /// <returns></returns>
        public ToolStripItem[] GetMenuItems()
        {
            /*context menu*/
            var item = new ToolStripMenuItem("Custom SQL")
            {
                Image = Image.FromFile(
                    @"C:\SES_DEV_GIT\SSMS2017_AddIn\SSMS2017_AddIn\Resources\1453477392_hammer_screwdriver.png")
            };

            /*context submenu item - generate inserts*/
            var insertItem = new ToolStripMenuItem("Insertion")
            {
                Image = Image.FromFile(
                    @"C:\SES_DEV_GIT\SSMS2017_AddIn\SSMS2017_AddIn\Resources\1453480812_database_save.png"),
                Tag = false
            };
            insertItem.Click += OnInsertItemClick;
            item.DropDownItems.Add(insertItem);

            /*context submenu item - count*/
            insertItem = new ToolStripMenuItem("Count(*)")
            {
                Image = Image.FromFile(
                    @"C:\SES_DEV_GIT\SSMS2017_AddIn\SSMS2017_AddIn\Resources\1453477379_Calculator.png"),
                Tag = false
            };
            insertItem.Click += OnCountClick;
            item.DropDownItems.Add(insertItem);

            /*context submenu */
            var scriptIt = new ToolStripMenuItem("Create table")
            {
                Image = Image.FromFile(
                    @"C:\SES_DEV_GIT\SSMS2017_AddIn\SSMS2017_AddIn\Resources\1453480695_table-add.png")
            };
            scriptIt.Click += OnScriptItClick;
            item.DropDownItems.Add(scriptIt);
            return new ToolStripItem[] { item };
        }
        #endregion

        #region Custom Click Events
        /// <summary>
        /// Handles the Click event of the Count control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void OnCountClick(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItem item = (ToolStripMenuItem)sender;
                bool generateColumnNames = (bool)item.Tag;
                Match match = this.tableRegex.Match(this.Parent.Context);

                Dictionary<string, string> dic = this.ExtractTableInfo(match.Groups[0].Value);
                string connectionString = this.Parent.Connection.ConnectionString + ";Database=" + dic["database"];
                string sqlStatement = string.Format(@"SELECT COUNT(*) AS COUNT FROM [{0}].[{1}]", dic["schema"], dic["table"]);

                SqlCommand command = new SqlCommand(sqlStatement);
                command.Connection = new SqlConnection(connectionString);
                command.Connection.Open();
                int tableCount = int.Parse(command.ExecuteScalar().ToString());
                command.Connection.Close();
                StringBuilder resultCaption = new StringBuilder().AppendFormat("{0}", sqlStatement);
                this.dteController.CreateNewScriptWindow(resultCaption); // create new document
                this.applicationObject.ExecuteCommand("Query.Execute"); // get query analyzer window to execute query
            }
            catch (Exception ex)
            {
                this.dteController.CreateNewScriptWindow(new StringBuilder(ex.Message));
            }
        }

        /// <summary>
        /// Handles the Click event of the ScriptIt control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void OnScriptItClick(object sender, EventArgs e)
        {
            Match match = this.tableRegex.Match(this.Parent.Context);
            Dictionary<string, string> dic = this.ExtractTableInfo(match.Groups[0].Value);
            string serverName = this.Parent.Connection.ServerName;
            string username = this.Parent.Connection.UserName;
            string password = this.Parent.Connection.Password;
            StringBuilder output = SmoGenerateSql(serverName, username, password, dic["database"], dic["table"], dic["schema"]);
            this.dteController.CreateNewScriptWindow(output);
        }

        private Dictionary<string, string> ExtractTableInfo(string data)
        {
            string[] info = data.Split('/');
            string tableName = info[2].Split('\'')[1];
            string schema = info[2].Split('\'')[3];
            string database = info[1].Split('\'')[1];
            var retour = new Dictionary<string, string>
            {
                ["table"] = tableName,
                ["schema"] = schema,
                ["database"] = database
            };
            return retour;
        }

        /// <summary>
        /// Handles the Click event of the InsertItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void OnInsertItemClick(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            bool generateColumnNames = (bool)item.Tag;

            Match match = this.tableRegex.Match(this.Parent.Context);
            if (match != null)
            {
                Dictionary<string, string> dic = this.ExtractTableInfo(match.Groups[0].Value);

                string connectionString = this.Parent.Connection.ConnectionString + ";Database=" + dic["database"];

                SqlCommand command = new SqlCommand(string.Format(@"SELECT * FROM [{0}].[{1}]", dic["schema"], dic["table"]));
                command.Connection = new SqlConnection(connectionString);
                command.Connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable table = new DataTable();
                adapter.Fill(table);

                command.Connection.Close();

                StringBuilder buffer = new StringBuilder();

                // generate INSERT prefix
                StringBuilder prefix = new StringBuilder();
                if (generateColumnNames)
                {
                    prefix.AppendFormat("INSERT INTO {0} (", dic["table"]);
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        if (i > 0) prefix.Append(", ");
                        prefix.AppendFormat("[{0}]", table.Columns[i].ColumnName);
                    }
                    prefix.Append(") VALUES (");
                }
                else
                    prefix.AppendFormat("INSERT INTO {0} VALUES (", dic["table"]);

                // generate INSERT statements
                foreach (DataRow row in table.Rows)
                {
                    StringBuilder values = new StringBuilder();
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        if (i > 0) values.Append(", ");

                        if (row.IsNull(i)) values.Append("NULL");
                        else if (table.Columns[i].DataType == typeof(int) ||
                            table.Columns[i].DataType == typeof(decimal) ||
                            table.Columns[i].DataType == typeof(long) ||
                            table.Columns[i].DataType == typeof(double) ||
                            table.Columns[i].DataType == typeof(float) ||
                            table.Columns[i].DataType == typeof(byte))
                            values.Append(row[i].ToString());
                        else
                            values.AppendFormat("'{0}'", row[i].ToString());
                    }
                    values.AppendFormat(")");

                    buffer.AppendLine(prefix.ToString() + values.ToString());
                }
                // create new sql page
                this.dteController.CreateNewScriptWindow(buffer);
            }
        }

        /// <summary>
        /// Smoes the generate SQL.
        /// </summary>
        /// <param name="serverName">Name of the curren server.</param>
        /// <param name="dbName">Name of the database.</param>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="tableSchema">The table schema.</param>
        /// <returns></returns>
        private static StringBuilder SmoGenerateSql(string serverName, string username, string password, string dbName, string tableName, string tableSchema)
        {
            var serverCon = new ServerConnection(serverName)
            {
                LoginSecure = false,
                Login = username,
                Password = password
            };
            var server = new Server(serverCon);
            Database db = server.Databases[dbName];
            var list = new List<Urn> { db.Tables[tableName, tableSchema].Urn };

            foreach (Index index in db.Tables[tableName, tableSchema].Indexes)
            {
                list.Add(index.Urn);
            }

            foreach (ForeignKey foreignKey in db.Tables[tableName, tableSchema].ForeignKeys)
            {
                list.Add(foreignKey.Urn);
            }

            foreach (Trigger triggers in db.Tables[tableName, tableSchema].Triggers)
            {
                list.Add(triggers.Urn);
            }

            Scripter scripter = new Scripter
            {
                Server = server,
                Options =
                {
                    IncludeHeaders = true,
                    SchemaQualify = true,
                    SchemaQualifyForeignKeysReferences = true,
                    NoCollation = true,
                    DriAllConstraints = true,
                    DriAll = true,
                    DriAllKeys = true,
                    DriIndexes = true,
                    ClusteredIndexes = true,
                    NonClusteredIndexes = true,
                    ToFileOnly = false
                }
            };
            StringCollection scriptedSql = scripter.Script(list.ToArray());

            var sb = new StringBuilder();

            foreach (string s in scriptedSql)
            {
                sb.AppendLine(s);
            }
            return sb;
        }

        #endregion        
    }
}