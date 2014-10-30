using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using System.IO;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using System.Diagnostics;
using System.Drawing;
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Eplanwiki.Scripting.MacroNavi
{

    public partial class MacroNavi : System.Windows.Forms.Form
    {
        private SplitContainer splitContainer1;
        private TreeView treeView1;
        private ListView listView1;
        private ColumnHeader columnHeader1;
        private ColumnHeader columnHeader2;
        private ColumnHeader columnHeader3;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem alleErweiternToolStripMenuItem;
        private ToolStripMenuItem alleReduzierenToolStripMenuItem;
        private CheckBox checkBox1;
        private ContextMenuStrip contextMenuStrip2;
        private ToolStripMenuItem aufteilungToolStripMenuItem;
        private ToolStripMenuItem horizontalToolStripMenuItem;
        private ToolStripMenuItem vertikalToolStripMenuItem;
        private ContextMenuStrip contextMenuStrip3;
        private ToolStripMenuItem allpoligToolStripMenuItem;
        private ToolStripMenuItem einpoligToolStripMenuItem;
        private ToolStripMenuItem paarquerverweisToolStripMenuItem;
        private ToolStripMenuItem übersichtToolStripMenuItem;
        private ToolStripMenuItem grafikToolStripMenuItem;
        private ToolStripMenuItem schaltschrankaufbauToolStripMenuItem;
        private ToolStripMenuItem rIFließbildToolStripMenuItem;
        private ToolStripMenuItem allpoligFluidToolStripMenuItem;
        private ToolStripMenuItem dMontageaufbauToolStripMenuItem;
        private ToolStripMenuItem einstellungenLöschenToolStripMenuItem;       
        private string macropath;
        private bool userDeletedSettings = false;
        private string currentProject = string.Empty;
        private static string USER_SETTING_FORMLOCATION = "USER.SCRIPTS.MACRONAVI.FORMLOCATION";
        private static string USER_SETTING_FORMSIZE = "USER.SCRIPTS.MACRONAVI.FORMSIZE";
        private static string USER_SETTING_ORIENTATION = "USER.SCRIPTS.MACRONAVI.ORIENTATION";
        private static string USER_SETTING_SPLITDISTANCE = "USER.SCRIPTS.MACRONAVI.SPLITDISTANCE";
        private static string USER_SETTING_PREVIEW = "USER.SCRIPTS.MACRONAVI.PREVIEW";
        private static string[] USER_SETTINGS = {USER_SETTING_FORMLOCATION, 
                                                USER_SETTING_FORMSIZE,
                                                USER_SETTING_ORIENTATION,
                                                USER_SETTING_SPLITDISTANCE,
                                                USER_SETTING_PREVIEW};
        private ToolStripMenuItem projektAktualisierenToolStripMenuItem;        

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen
        /// gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor
        /// geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.alleErweiternToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.alleReduzierenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.contextMenuStrip3 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.allpoligToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.einpoligToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.paarquerverweisToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.übersichtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.grafikToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.schaltschrankaufbauToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rIFließbildToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.allpoligFluidToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dMontageaufbauToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.aufteilungToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.horizontalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.vertikalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.einstellungenLöschenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.projektAktualisierenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.contextMenuStrip3.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(0, 35);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeView1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.listView1);
            this.splitContainer1.Size = new System.Drawing.Size(259, 437);
            this.splitContainer1.SplitterDistance = 218;
            this.splitContainer1.TabIndex = 0;
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.Location = new System.Drawing.Point(3, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(253, 196);
            this.treeView1.TabIndex = 0;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.alleErweiternToolStripMenuItem,
            this.alleReduzierenToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 48);
            // 
            // alleErweiternToolStripMenuItem
            // 
            this.alleErweiternToolStripMenuItem.Name = "alleErweiternToolStripMenuItem";
            this.alleErweiternToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.alleErweiternToolStripMenuItem.Text = "Alle erweitern";
            this.alleErweiternToolStripMenuItem.Click += new System.EventHandler(this.alleErweiternToolStripMenuItem_Click);
            // 
            // alleReduzierenToolStripMenuItem
            // 
            this.alleReduzierenToolStripMenuItem.Name = "alleReduzierenToolStripMenuItem";
            this.alleReduzierenToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.alleReduzierenToolStripMenuItem.Text = "Alle reduzieren";
            this.alleReduzierenToolStripMenuItem.Click += new System.EventHandler(this.alleReduzierenToolStripMenuItem_Click);
            // 
            // listView1
            // 
            this.listView1.AllowDrop = true;
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.listView1.ContextMenuStrip = this.contextMenuStrip3;
            this.listView1.Location = new System.Drawing.Point(3, 12);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(253, 203);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.listView1_ItemDrag);
            this.listView1.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.listView1_ItemSelectionChanged);
            this.listView1.DragLeave += new System.EventHandler(this.listView1_DragLeave);
            this.listView1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.listView1_KeyPress);
            this.listView1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseDoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Name";
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Typ";
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "letzte Änderung";
            // 
            // contextMenuStrip3
            // 
            this.contextMenuStrip3.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.allpoligToolStripMenuItem,
            this.einpoligToolStripMenuItem,
            this.paarquerverweisToolStripMenuItem,
            this.übersichtToolStripMenuItem,
            this.grafikToolStripMenuItem,
            this.schaltschrankaufbauToolStripMenuItem,
            this.rIFließbildToolStripMenuItem,
            this.allpoligFluidToolStripMenuItem,
            this.dMontageaufbauToolStripMenuItem});
            this.contextMenuStrip3.Name = "contextMenuStrip3";
            this.contextMenuStrip3.Size = new System.Drawing.Size(185, 202);
            // 
            // allpoligToolStripMenuItem
            // 
            this.allpoligToolStripMenuItem.Name = "allpoligToolStripMenuItem";
            this.allpoligToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.allpoligToolStripMenuItem.Text = "Allpolig";
            this.allpoligToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // einpoligToolStripMenuItem
            // 
            this.einpoligToolStripMenuItem.Name = "einpoligToolStripMenuItem";
            this.einpoligToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.einpoligToolStripMenuItem.Text = "Einpolig";
            this.einpoligToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // paarquerverweisToolStripMenuItem
            // 
            this.paarquerverweisToolStripMenuItem.Name = "paarquerverweisToolStripMenuItem";
            this.paarquerverweisToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.paarquerverweisToolStripMenuItem.Text = "Paarquerverweis";
            this.paarquerverweisToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // übersichtToolStripMenuItem
            // 
            this.übersichtToolStripMenuItem.Name = "übersichtToolStripMenuItem";
            this.übersichtToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.übersichtToolStripMenuItem.Text = "Übersicht";
            this.übersichtToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // grafikToolStripMenuItem
            // 
            this.grafikToolStripMenuItem.Name = "grafikToolStripMenuItem";
            this.grafikToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.grafikToolStripMenuItem.Text = "Grafik";
            this.grafikToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // schaltschrankaufbauToolStripMenuItem
            // 
            this.schaltschrankaufbauToolStripMenuItem.Name = "schaltschrankaufbauToolStripMenuItem";
            this.schaltschrankaufbauToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.schaltschrankaufbauToolStripMenuItem.Text = "Schaltschrankaufbau";
            this.schaltschrankaufbauToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // rIFließbildToolStripMenuItem
            // 
            this.rIFließbildToolStripMenuItem.Name = "rIFließbildToolStripMenuItem";
            this.rIFließbildToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.rIFließbildToolStripMenuItem.Text = "RI-Fließbild";
            this.rIFließbildToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // allpoligFluidToolStripMenuItem
            // 
            this.allpoligFluidToolStripMenuItem.Name = "allpoligFluidToolStripMenuItem";
            this.allpoligFluidToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.allpoligFluidToolStripMenuItem.Text = "Allpolig Fluid";
            this.allpoligFluidToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // dMontageaufbauToolStripMenuItem
            // 
            this.dMontageaufbauToolStripMenuItem.Name = "dMontageaufbauToolStripMenuItem";
            this.dMontageaufbauToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.dMontageaufbauToolStripMenuItem.Text = "3D-Montageaufbau";
            this.dMontageaufbauToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(188, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(71, 17);
            this.checkBox1.TabIndex = 3;
            this.checkBox1.Text = "Vorschau";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aufteilungToolStripMenuItem,
            this.einstellungenLöschenToolStripMenuItem,
            this.projektAktualisierenToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(190, 92);
            // 
            // aufteilungToolStripMenuItem
            // 
            this.aufteilungToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.horizontalToolStripMenuItem,
            this.vertikalToolStripMenuItem});
            this.aufteilungToolStripMenuItem.Name = "aufteilungToolStripMenuItem";
            this.aufteilungToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.aufteilungToolStripMenuItem.Text = "Aufteilung";
            // 
            // horizontalToolStripMenuItem
            // 
            this.horizontalToolStripMenuItem.Name = "horizontalToolStripMenuItem";
            this.horizontalToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.horizontalToolStripMenuItem.Text = "Horizontal";
            this.horizontalToolStripMenuItem.Click += new System.EventHandler(this.horizontalToolStripMenuItem_Click);
            // 
            // vertikalToolStripMenuItem
            // 
            this.vertikalToolStripMenuItem.Name = "vertikalToolStripMenuItem";
            this.vertikalToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.vertikalToolStripMenuItem.Text = "Vertikal";
            this.vertikalToolStripMenuItem.Click += new System.EventHandler(this.vertikalToolStripMenuItem_Click);
            // 
            // einstellungenLöschenToolStripMenuItem
            // 
            this.einstellungenLöschenToolStripMenuItem.Name = "einstellungenLöschenToolStripMenuItem";
            this.einstellungenLöschenToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.einstellungenLöschenToolStripMenuItem.Text = "Einstellungen löschen";
            this.einstellungenLöschenToolStripMenuItem.Click += new System.EventHandler(this.einstellungenLöschenToolStripMenuItem_Click);
            // 
            // projektAktualisierenToolStripMenuItem
            // 
            this.projektAktualisierenToolStripMenuItem.Name = "projektAktualisierenToolStripMenuItem";
            this.projektAktualisierenToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.projektAktualisierenToolStripMenuItem.Text = "Projekt aktualisieren";
            this.projektAktualisierenToolStripMenuItem.Click += new System.EventHandler(this.projektAktualisierenToolStripMenuItem_Click);
            // 
            // MacroNavi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(259, 474);
            this.ContextMenuStrip = this.contextMenuStrip2;
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.splitContainer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MacroNavi";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Makros";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frm_FormClosing);
            this.Load += new System.EventHandler(this.frm_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.contextMenuStrip3.ResumeLayout(false);
            this.contextMenuStrip2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public MacroNavi()
        {
            InitializeComponent();
        }

        #endregion

        [DeclareMenu]
        public void MenuFunction()
        {
            Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
            uint presMenuId = oMenu.GetPersistentMenuId("Symbole");
            oMenu.AddMenuItem("Fenstermakros", "ShowMacroNavi", "Makros", presMenuId, 0, false, false);
        }      

        [DeclareAction("ShowMacroNavi")]
        public void ShowMacroNavi()
        {              
            MacroNavi frm = new MacroNavi();
            Process oCurrent = Process.GetCurrentProcess();
            var ww = new WindowWrapper(oCurrent.MainWindowHandle);            

            setFormProperties(frm);                       

            if (!frm_IsLoaded(frm.Name))
            {
                frm.Show(ww);
            }
            
                        
        }

        /// <summary>
        /// Determines the current project via the Eplan-Action "SelectinSet". Doesn't work if the GED is not in focus.
        /// </summary>
        /// <returns>Absolute projectaname as string</returns>
        private string getCurrentProject()
        {
            CommandLineInterpreter oCli = new CommandLineInterpreter();
            ActionCallingContext acc = new ActionCallingContext();
            string project = string.Empty;
            acc.AddParameter("Type", "PROJECT");
            oCli.Execute("selectionset", acc);
            acc.GetParameter("PROJECT", ref project);
            return project;
        }

        /// <summary>
        /// Sets the form-size and location, orientation an size of the splitcontainer and the value of the checkbox from usersetting.
        /// </summary>
        /// <param name="_frm"></param>
        private void setFormProperties(MacroNavi _frm)
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

            if (oSettings.ExistSetting(USER_SETTING_FORMLOCATION))
            {
                _frm.Top = oSettings.GetNumericSetting(USER_SETTING_FORMLOCATION, 0);
                _frm.Left = oSettings.GetNumericSetting(USER_SETTING_FORMLOCATION, 1);
            }
            if (oSettings.ExistSetting(USER_SETTING_FORMSIZE))
            {
                _frm.Height = oSettings.GetNumericSetting(USER_SETTING_FORMSIZE, 0);
                _frm.Width = oSettings.GetNumericSetting(USER_SETTING_FORMSIZE, 1);
            }
            if (oSettings.ExistSetting(USER_SETTING_ORIENTATION))
            {
                _frm.splitContainer1.Orientation = (Orientation)oSettings.GetNumericSetting(USER_SETTING_ORIENTATION, 0);
            }
            if (oSettings.ExistSetting(USER_SETTING_SPLITDISTANCE))
            {
                _frm.splitContainer1.SplitterDistance = oSettings.GetNumericSetting(USER_SETTING_SPLITDISTANCE, 0);
            }
            if (oSettings.ExistSetting(USER_SETTING_PREVIEW))
            {
                _frm.checkBox1.Checked = oSettings.GetBoolSetting(USER_SETTING_PREVIEW, 0);
            }            
        }

        /// <summary>
        /// Determines if the form is already loaded
        /// </summary>
        /// <param name="sName">Name of the MacroNavi object</param>
        /// <returns></returns>
        private bool frm_IsLoaded(string sName)
        {
            bool bResult = false;
            foreach (Form oForm in Application.OpenForms)
            {
                if (oForm.Name.ToLower() == sName.ToLower())
                {
                    bResult = true;
                    break;
                }
            }
            return (bResult);
        }        

        private void frm_Load(object sender, System.EventArgs e)
        {
            currentProject = getCurrentProject();
            setPreview("", currentProject, checkBox1.Checked);
            populateTreeView();
        }

        /// <summary>
        /// Populates the tree with foldernames of the eplan macropath.
        /// <seealso cref="http://msdn.microsoft.com/en-us/library/ms171645(v=vs.90).aspx"/>
        /// </summary>
        private void populateTreeView()
        {
            TreeNode rootNode;

            macropath = PathMap.SubstitutePath("$(MD_MACROS)");
            DirectoryInfo info = new DirectoryInfo(macropath + @"\");
            if (info.Exists)
            {
                rootNode = new TreeNode(info.Name);
                rootNode.Tag = info;
                getDirectories(info.GetDirectories(), rootNode);
                treeView1.Nodes.Add(rootNode);
            }
        }

        /// <summary>
        /// <see cref="http://msdn.microsoft.com/en-us/library/ms171645(v=vs.90).aspx"/>
        /// </summary>
        /// <param name="subDirs"></param>
        /// <param name="nodeToAddTo"></param>
        private void getDirectories(DirectoryInfo[] subDirs, TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            DirectoryInfo[] subSubDirs;
            foreach (DirectoryInfo subDir in subDirs)
            {
                aNode = new TreeNode(subDir.Name, 0, 0);
                aNode.Tag = subDir;
                //aNode.ImageKey = "folder";
                subSubDirs = subDir.GetDirectories();
                if (subSubDirs.Length != 0)
                {
                    getDirectories(subSubDirs, aNode);
                }
                nodeToAddTo.Nodes.Add(aNode);
            }
        }

        /// <summary>
        /// Called from event if the selected item in listview has changed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (listView1.SelectedItems.Count > 0 && checkBox1.Checked)
            {                
                string absoluteMacroName = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + listView1.SelectedItems[0].Text;                
                setPreview(absoluteMacroName, currentProject, this.checkBox1.Checked);
            }
            listView1.Focus();
        }

        /// <summary>
        /// Opens or closes the preview window via the eplan action "XSDPreviewAction"
        /// </summary>
        /// <param name="_absoluteMacroName"></param>
        /// <param name="_absoluteProjectName"></param>
        /// <param name="_show"></param>
        private void setPreview(string _absoluteMacroName, string _absoluteProjectName, bool _show)
        {
            CommandLineInterpreter oCli = new CommandLineInterpreter();
            ActionCallingContext acc = new ActionCallingContext();            
            acc.AddParameter("PROJECTNAME", _absoluteProjectName);
            if (File.Exists(_absoluteMacroName))
            {
                acc.AddParameter("MACRONAME", _absoluteMacroName);
            }            
            acc.AddParameter("SHOW", Convert.ToInt16(_show).ToString());            
            oCli.Execute("XSDPreviewAction", acc);
        }

        /// <summary>
        /// Inserts macro by doubleclicking an item in listview.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            bool isDirectory;
            if (listView1.SelectedItems != null)
            {
                isDirectory = listView1.SelectedItems[0].SubItems[1].Text.Equals("Directory");


                if (isDirectory)
                {
                    foreach (TreeNode node in treeView1.SelectedNode.Nodes)
                    {
                        if (node.Text.Equals(listView1.SelectedItems[0].SubItems[0].Text))
                        {
                            //treeView1.SelectedNode = node;
                        }
                    }
                }
                if(!isDirectory)
                {
                    instertMacro(getAbsoluteMacroPath(), WindowMacro.RepresentationType.MultiLine, 0);
                }
            }
        }

        /// <summary>
        /// Inserts macro by hittig enter on selected item in listview.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == (char)Keys.Return)
            {
                instertMacro(getAbsoluteMacroPath() , WindowMacro.RepresentationType.MultiLine, 0);
            }

        }

        /// <summary>
        /// Inserts macro via the eplan interaction "XGedStartInteractionAction".
        /// </summary>
        /// <param name="reptype">Use the enumerators from the static class WindowMacro</param>
        /// <param name="variant">Variant_A = 0 Variant_B = 1 ... Variant_P = 15 </param>
        private void instertMacro(string absoluteMacroName, WindowMacro.RepresentationType reptype, Int32 variant)
        {

            if (listView1.SelectedItems.Count > 0 && File.Exists(absoluteMacroName))
            {                
                new CommandLineInterpreter().Execute("XGedStartInteractionAction /Name:XMIaInsertMacro /filename:"
                + "\"" + absoluteMacroName + "\"" +
                "/variant:" + variant.ToString() + " /RepresentationType:" + ((int)reptype).ToString());
            }
        }

        private string getAbsoluteMacroPath()
        {
            return macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + listView1.SelectedItems[0].Text;
        }

        /// <summary>
        /// Called by event if form is closing. Writes the current satus of the form to usersettigs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!userDeletedSettings)
            {
                settingsWrite();
            }
            setPreview("", currentProject, false);
        }

        /// <summary>
        /// Shows and hides the preview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            setPreview("", currentProject, checkBox1.Checked);
        }

        /// <summary>
        /// Expands the hole treeview.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void alleErweiternToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.ExpandAll();
        }

        /// <summary>
        /// Collapses the hole treeview.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void alleReduzierenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.CollapseAll();
        }

        /// <summary>
        /// <see cref="http://msdn.microsoft.com/en-us/library/ms171645(v=vs.90).aspx"/>
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {            
            TreeNode newSelected = e.Node;
            listView1.Items.Clear();
            DirectoryInfo nodeDirInfo = (DirectoryInfo)newSelected.Tag;

            ListViewItem.ListViewSubItem[] subItems;
            ListViewItem item = null;

            foreach (DirectoryInfo dir in nodeDirInfo.GetDirectories())
            {
                item = new ListViewItem(dir.Name, 0);
                subItems = new ListViewItem.ListViewSubItem[]
                    {new ListViewItem.ListViewSubItem(item, "Directory"), 
                     new ListViewItem.ListViewSubItem(item, 
						dir.LastAccessTime.ToShortDateString())};
                item.SubItems.AddRange(subItems);
                listView1.Items.Add(item);
            }
            foreach (FileInfo file in nodeDirInfo.GetFiles("*.ema"))
            {
                item = new ListViewItem(file.Name, 1);
                subItems = new ListViewItem.ListViewSubItem[]
                    { new ListViewItem.ListViewSubItem(item, "File"), 
                     new ListViewItem.ListViewSubItem(item, 
						file.LastAccessTime.ToShortDateString())};
                
                item.SubItems.AddRange(subItems);
                listView1.Items.Add(item);
            }

            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        /// <summary>
        /// Changes the oriantation of the splitcontainer to horizontal.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void horizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitContainer1.Orientation = Orientation.Horizontal;
        }

        /// <summary>
        /// Changes the oriantation of the splitcontainer to vertical.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void vertikalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitContainer1.Orientation = Orientation.Vertical;
        }

        /// <summary>
        /// Inserts macro in representationtype MultiLine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WindowMacro.RepresentationType repType;
            WindowMacro.representaiontypes.TryGetValue(((ToolStripItem)sender).Text, out repType);            
            instertMacro("", repType, 0);
        }
        
        /// <summary>
        /// Writes the properties of the form to eplan user settings.
        /// <seealso cref="http://ww3.cad.de/foren/ubb/Forum467/HTML/005894.shtml"/>
        /// </summary>
        private void settingsWrite()
        {

            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
            if (!oSettings.ExistSetting(USER_SETTING_FORMLOCATION))
            {
                oSettings.AddNumericSetting(USER_SETTING_FORMLOCATION,
                    new int[] { 0 },
                    new Range[] { new Range { FromValue = -32768, ToValue = 32768 } },
                    "Default value of test setting",
                    new int[] { 0 },
                    ISettings.CreationFlag.Insert);
            }

            oSettings.SetNumericSetting(USER_SETTING_FORMLOCATION, this.Top, 0);
            oSettings.SetNumericSetting(USER_SETTING_FORMLOCATION, this.Left, 1);
            if (!oSettings.ExistSetting(USER_SETTING_FORMSIZE))
            {
                oSettings.AddNumericSetting(USER_SETTING_FORMSIZE,
                    new int[] { 0 },
                    new Range[] { new Range { FromValue = -32768, ToValue = 32768 } },
                    "Default value of test setting",
                    new int[] { 0 },
                    ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_FORMSIZE, this.Height, 0);
            oSettings.SetNumericSetting(USER_SETTING_FORMSIZE, this.Width, 1);


            if (!oSettings.ExistSetting(USER_SETTING_ORIENTATION))
            {
                oSettings.AddNumericSetting(USER_SETTING_ORIENTATION,
                    new int[] { 0 },
                    new Range[] { new Range { FromValue = 0, ToValue = 1 } },
                    "Default value of test setting",
                    new int[] { 0 },
                    ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_ORIENTATION, (int)this.splitContainer1.Orientation, 0);

            if (!oSettings.ExistSetting(USER_SETTING_SPLITDISTANCE))
            {
                oSettings.AddNumericSetting(USER_SETTING_SPLITDISTANCE,
                    new int[] { 0 },
                    new Range[] { new Range { FromValue = -32768, ToValue = 32768 } },
                    "Default value of test setting",
                    new int[] { 0 },
                    ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_SPLITDISTANCE, (int)this.splitContainer1.SplitterDistance, 0);

            if (!oSettings.ExistSetting(USER_SETTING_PREVIEW))
            {
                oSettings.AddBoolSetting(USER_SETTING_PREVIEW, new bool[] { false },
                    "Default value of test setting",
                    new bool[] { false },
                    ISettings.CreationFlag.Insert);
            }
            oSettings.SetBoolSetting(USER_SETTING_PREVIEW, this.checkBox1.Checked, 0);
        }

        /// <summary>
        /// Deletes the user settings of this script
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void einstellungenLöschenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Sollen die Einstellungen dieses Scripts aus den User-Settings gelöscht werden?", "Einstellungen löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
                foreach (string setting in USER_SETTINGS)
                {
                    if (oSettings.ExistSetting(setting))
                    {
                        oSettings.DeleteSetting(setting);
                    }
                }
                userDeletedSettings = true;
            }

        }

        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(sender, DragDropEffects.All);
        }

        private void listView1_DragLeave(object sender, EventArgs e)
        {
            instertMacro(getAbsoluteMacroPath(), WindowMacro.RepresentationType.MultiLine, 0);
        }

        private void projektAktualisierenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentProject = getAbsoluteMacroPath();
        }
    
    }


    /// <summary>
    /// Contains the enums of macro represetationtypes
    /// </summary>
    public static class WindowMacro
    {
        public static SortedDictionary<string, WindowMacro.RepresentationType> representaiontypes = new SortedDictionary<string, WindowMacro.RepresentationType>()
        {
            {"Allpolig", WindowMacro.RepresentationType.MultiLine},
            {"Einpolig" ,WindowMacro.RepresentationType.SingleLine},
            {"Paarquerverweis", WindowMacro.RepresentationType.PairCrossReference},
            {"Übersicht", WindowMacro.RepresentationType.Overview},
            {"Grafik", WindowMacro.RepresentationType.Graphics},
            {"Schaltschrankaufbau", WindowMacro.RepresentationType.ArticlePlacement},
            {"RI-Fließbild", WindowMacro.RepresentationType.PI_FlowChart},
            {"Allpolig Fluid", WindowMacro.RepresentationType.Fluid_MultiLine},
            {"Cabeling", WindowMacro.RepresentationType.Cabling},
            {"3D-Montageaufbau", WindowMacro.RepresentationType.ArticlePlacement3D},
            {"Funktional", WindowMacro.RepresentationType.Functional}        
        };

        public enum RepresentationType
        {
            MultiLine = 1,
            SingleLine = 2,
            PairCrossReference = 3,
            Overview = 4,
            Graphics = 5,
            ArticlePlacement = 6,
            PI_FlowChart = 7,
            Fluid_MultiLine = 8,
            Cabling = 9,
            ArticlePlacement3D = 10,
            Functional = 11
        }
    }

    /// <summary>
    /// Class to handle the owner of the form.
    /// </summary>
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }
        private IntPtr _hwnd;
    }
}