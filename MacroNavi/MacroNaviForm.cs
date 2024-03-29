/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */
 
using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using System.IO;
using System.Xml;
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
        private ToolStripMenuItem expandAllToolStripMenuItem;
        private ToolStripMenuItem collapseAllToolStripMenuItem;
        private CheckBox checkBox1;
        private ContextMenuStrip contextMenuStrip2;
        private ToolStripMenuItem splitOrientationToolStripMenuItem;
        private ToolStripMenuItem horizontalToolStripMenuItem;
        private ToolStripMenuItem verticalToolStripMenuItem;
        private ContextMenuStrip contextMenuStrip3;
        private ToolStripMenuItem multiLineToolStripMenuItem;
        private ToolStripMenuItem singleLineToolStripMenuItem;
        private ToolStripMenuItem pairCrossReferenceToolStripMenuItem;
        private ToolStripMenuItem overviewToolStripMenuItem;
        private ToolStripMenuItem graphicsToolStripMenuItem;
        private ToolStripMenuItem articlePlacementToolStripMenuItem;
        private ToolStripMenuItem pi_FlowChartToolStripMenuItem;
        private ToolStripMenuItem fluid_MultiLineToolStripMenuItem;
        private ToolStripMenuItem articlePlacement3DToolStripMenuItem;
        private ToolStripMenuItem deleteUserSettingsToolStripMenuItem;       
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
        private ToolStripMenuItem refreshCurrentProjectToolStripMenuItem;
        private ToolStripMenuItem openPathToolStripMenuItem;
        private int nIndex = 0;
        private ToolStripMenuItem eigenschaftenToolStripMenuItem;

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
            this.expandAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.collapseAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openPathToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.contextMenuStrip3 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.multiLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.singleLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pairCrossReferenceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.overviewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.graphicsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.articlePlacementToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pi_FlowChartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fluid_MultiLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.articlePlacement3DToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.splitOrientationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.horizontalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verticalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteUserSettingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshCurrentProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.eigenschaftenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.splitContainer1.Location = new System.Drawing.Point(0, 43);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(4);
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
            this.splitContainer1.Size = new System.Drawing.Size(345, 538);
            this.splitContainer1.SplitterDistance = 268;
            this.splitContainer1.SplitterWidth = 5;
            this.splitContainer1.TabIndex = 0;
            this.splitContainer1.TabStop = false;
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.Location = new System.Drawing.Point(4, 14);
            this.treeView1.Margin = new System.Windows.Forms.Padding(4);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(336, 240);
            this.treeView1.TabIndex = 0;
            this.treeView1.TabStop = false;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            this.treeView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.treeView1_KeyDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.expandAllToolStripMenuItem,
            this.collapseAllToolStripMenuItem,
            this.openPathToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(193, 76);
            // 
            // expandAllToolStripMenuItem
            // 
            this.expandAllToolStripMenuItem.Name = "expandAllToolStripMenuItem";
            this.expandAllToolStripMenuItem.Size = new System.Drawing.Size(192, 24);
            this.expandAllToolStripMenuItem.Text = "Alle erweitern";
            this.expandAllToolStripMenuItem.Click += new System.EventHandler(this.expandAllToolStripMenuItem_Click);
            // 
            // collapseAllToolStripMenuItem
            // 
            this.collapseAllToolStripMenuItem.Name = "collapseAllToolStripMenuItem";
            this.collapseAllToolStripMenuItem.Size = new System.Drawing.Size(192, 24);
            this.collapseAllToolStripMenuItem.Text = "Alle reduzieren";
            this.collapseAllToolStripMenuItem.Click += new System.EventHandler(this.collapseAllToolStripMenuItem_Click);
            // 
            // openPathToolStripMenuItem
            // 
            this.openPathToolStripMenuItem.Name = "openPathToolStripMenuItem";
            this.openPathToolStripMenuItem.Size = new System.Drawing.Size(192, 24);
            this.openPathToolStripMenuItem.Text = "Dateipfad öffnen";
            this.openPathToolStripMenuItem.Click += new System.EventHandler(this.openPathToolStripMenuItem_Click);
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
            this.listView1.Location = new System.Drawing.Point(4, 14);
            this.listView1.Margin = new System.Windows.Forms.Padding(4);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.ShowItemToolTips = true;
            this.listView1.Size = new System.Drawing.Size(336, 238);
            this.listView1.TabIndex = 0;
            this.listView1.TabStop = false;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.listView1_ItemDrag);
            this.listView1.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.listView1_ItemSelectionChanged);
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
            this.multiLineToolStripMenuItem,
            this.singleLineToolStripMenuItem,
            this.pairCrossReferenceToolStripMenuItem,
            this.overviewToolStripMenuItem,
            this.graphicsToolStripMenuItem,
            this.articlePlacementToolStripMenuItem,
            this.pi_FlowChartToolStripMenuItem,
            this.fluid_MultiLineToolStripMenuItem,
            this.articlePlacement3DToolStripMenuItem,
            this.eigenschaftenToolStripMenuItem});
            this.contextMenuStrip3.Name = "contextMenuStrip3";
            this.contextMenuStrip3.Size = new System.Drawing.Size(214, 266);
            // 
            // multiLineToolStripMenuItem
            // 
            this.multiLineToolStripMenuItem.Name = "multiLineToolStripMenuItem";
            this.multiLineToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.multiLineToolStripMenuItem.Text = "Allpolig";
            this.multiLineToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // singleLineToolStripMenuItem
            // 
            this.singleLineToolStripMenuItem.Name = "singleLineToolStripMenuItem";
            this.singleLineToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.singleLineToolStripMenuItem.Text = "Einpolig";
            this.singleLineToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // pairCrossReferenceToolStripMenuItem
            // 
            this.pairCrossReferenceToolStripMenuItem.Name = "pairCrossReferenceToolStripMenuItem";
            this.pairCrossReferenceToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.pairCrossReferenceToolStripMenuItem.Text = "Paarquerverweis";
            this.pairCrossReferenceToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // overviewToolStripMenuItem
            // 
            this.overviewToolStripMenuItem.Name = "overviewToolStripMenuItem";
            this.overviewToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.overviewToolStripMenuItem.Text = "Übersicht";
            this.overviewToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // graphicsToolStripMenuItem
            // 
            this.graphicsToolStripMenuItem.Name = "graphicsToolStripMenuItem";
            this.graphicsToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.graphicsToolStripMenuItem.Text = "Grafik";
            this.graphicsToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // articlePlacementToolStripMenuItem
            // 
            this.articlePlacementToolStripMenuItem.Name = "articlePlacementToolStripMenuItem";
            this.articlePlacementToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.articlePlacementToolStripMenuItem.Text = "Schaltschrankaufbau";
            this.articlePlacementToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // pi_FlowChartToolStripMenuItem
            // 
            this.pi_FlowChartToolStripMenuItem.Name = "pi_FlowChartToolStripMenuItem";
            this.pi_FlowChartToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.pi_FlowChartToolStripMenuItem.Text = "RI-Fließbild";
            this.pi_FlowChartToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // fluid_MultiLineToolStripMenuItem
            // 
            this.fluid_MultiLineToolStripMenuItem.Name = "fluid_MultiLineToolStripMenuItem";
            this.fluid_MultiLineToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.fluid_MultiLineToolStripMenuItem.Text = "Allpolig Fluid";
            this.fluid_MultiLineToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // articlePlacement3DToolStripMenuItem
            // 
            this.articlePlacement3DToolStripMenuItem.Name = "articlePlacement3DToolStripMenuItem";
            this.articlePlacement3DToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.articlePlacement3DToolStripMenuItem.Text = "3D-Montageaufbau";
            this.articlePlacement3DToolStripMenuItem.Click += new System.EventHandler(this.repTypeToolStripMenuItem_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(255, 14);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(90, 21);
            this.checkBox1.TabIndex = 1;
            this.checkBox1.TabStop = false;
            this.checkBox1.Text = "Vorschau";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.splitOrientationToolStripMenuItem,
            this.deleteUserSettingsToolStripMenuItem,
            this.refreshCurrentProjectToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(221, 76);
            // 
            // splitOrientationToolStripMenuItem
            // 
            this.splitOrientationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.horizontalToolStripMenuItem,
            this.verticalToolStripMenuItem});
            this.splitOrientationToolStripMenuItem.Name = "splitOrientationToolStripMenuItem";
            this.splitOrientationToolStripMenuItem.Size = new System.Drawing.Size(220, 24);
            this.splitOrientationToolStripMenuItem.Text = "Aufteilung";
            // 
            // horizontalToolStripMenuItem
            // 
            this.horizontalToolStripMenuItem.Name = "horizontalToolStripMenuItem";
            this.horizontalToolStripMenuItem.Size = new System.Drawing.Size(148, 24);
            this.horizontalToolStripMenuItem.Text = "Horizontal";
            this.horizontalToolStripMenuItem.Click += new System.EventHandler(this.horizontalToolStripMenuItem_Click);
            // 
            // verticalToolStripMenuItem
            // 
            this.verticalToolStripMenuItem.Name = "verticalToolStripMenuItem";
            this.verticalToolStripMenuItem.Size = new System.Drawing.Size(148, 24);
            this.verticalToolStripMenuItem.Text = "Vertikal";
            this.verticalToolStripMenuItem.Click += new System.EventHandler(this.verticalToolStripMenuItem_Click);
            // 
            // deleteUserSettingsToolStripMenuItem
            // 
            this.deleteUserSettingsToolStripMenuItem.Name = "deleteUserSettingsToolStripMenuItem";
            this.deleteUserSettingsToolStripMenuItem.Size = new System.Drawing.Size(220, 24);
            this.deleteUserSettingsToolStripMenuItem.Text = "Einstellungen löschen";
            this.deleteUserSettingsToolStripMenuItem.Click += new System.EventHandler(this.deleteUserSettingsToolStripMenuItem_Click);
            // 
            // refreshCurrentProjectToolStripMenuItem
            // 
            this.refreshCurrentProjectToolStripMenuItem.Name = "refreshCurrentProjectToolStripMenuItem";
            this.refreshCurrentProjectToolStripMenuItem.Size = new System.Drawing.Size(220, 24);
            this.refreshCurrentProjectToolStripMenuItem.Text = "Projekt aktualisieren";
            this.refreshCurrentProjectToolStripMenuItem.Click += new System.EventHandler(this.refreshCurrentProjectToolStripMenuItem_Click);
            // 
            // eigenschaftenToolStripMenuItem
            // 
            this.eigenschaftenToolStripMenuItem.Name = "eigenschaftenToolStripMenuItem";
            this.eigenschaftenToolStripMenuItem.Size = new System.Drawing.Size(213, 24);
            this.eigenschaftenToolStripMenuItem.Text = "Eigenschaften...";
            this.eigenschaftenToolStripMenuItem.Click += new System.EventHandler(this.eigenschaftenToolStripMenuItem_Click);
            // 
            // MacroNavi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(345, 583);
            this.ContextMenuStrip = this.contextMenuStrip2;
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.splitContainer1);
            this.Margin = new System.Windows.Forms.Padding(4);
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

        /// <summary>
        /// Inserts menupoint above the "Symbole" (GER) entry in "Projektdaten"
        /// </summary>
        [DeclareMenu]
        public void MenuFunction()
        {
            Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
            uint presMenuId = oMenu.GetPersistentMenuId("Symbole");
            oMenu.AddMenuItem("Fenstermakros", "ShowMacroNavi", "Makros", presMenuId, 0, false, false);
        }      

        /// <summary>
        /// The action called from menu "Fenstermakros", shows the form
        /// </summary>
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

        /// <summary>
        /// first treeview population, preview visibility and current project set
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            treeView1.Nodes.Clear();
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
            treeView1.Sort();
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
            try
            {
                DirectoryInfo nodeDirInfo = (DirectoryInfo)newSelected.Tag;
                ListViewItem.ListViewSubItem[] subItems;
                ListviewFileItem item = null;

                foreach (DirectoryInfo dir in nodeDirInfo.GetDirectories())
                {
                    item = new ListviewFileItem(dir.Name, 0);
                    item.FullFilePath = dir.FullName;
                    subItems = new ListViewItem.ListViewSubItem[]
                    {new ListViewItem.ListViewSubItem(item, "Directory"), 
                     new ListViewItem.ListViewSubItem(item, 
						dir.LastAccessTime.ToShortDateString())};
                    item.SubItems.AddRange(subItems);
                    listView1.Items.Add(item);
                }
                foreach (FileInfo file in nodeDirInfo.GetFiles("*.ema"))
                {
                    item = new ListviewFileItem(file.Name, 1);
                    item.ToolTipText = getListViewItemToolTipText(file);
                    item.FullFilePath = file.FullName;
                    subItems = new ListViewItem.ListViewSubItem[]
                    { new ListViewItem.ListViewSubItem(item, "File"), 
                     new ListViewItem.ListViewSubItem(item, 
						file.LastAccessTime.ToShortDateString())};

                    item.SubItems.AddRange(subItems);
                    listView1.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehlgeschlagen: " + ex.Message, "Pfad nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
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
                //string absoluteMacroName = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + listView1.SelectedItems[0].Text;                
                setPreview(((ListviewFileItem)e.Item).FullFilePath, currentProject, this.checkBox1.Checked);
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
                string equals = listView1.SelectedItems[0].SubItems[0].Text;

                if (isDirectory)
                {                    
                    foreach (TreeNode node in treeView1.SelectedNode.Nodes)
                    {
                        if (node.Text.Equals(equals))
                        {
                            treeView1.SelectedNode = node;
                        }
                    }
                }
                if(!isDirectory)
                {                    
                    instertMacro(((ListviewFileItem)listView1.SelectedItems[0]).FullFilePath, WindowMacro.RepresentationType.MultiLine, 0);
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
                instertMacro(((ListviewFileItem)listView1.SelectedItems[0]).FullFilePath, WindowMacro.RepresentationType.MultiLine, 0);
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

        /// <summary>
        /// Returns the full filepath accumulated from macropath, selected treenode an listviewitem
        /// </summary>
        /// <returns></returns>        
        //private string getAbsoluteMacroPath()
        //{
        //    return macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + listView1.SelectedItems[0].Text;
        //}

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
        private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.ExpandAll();
        }

        /// <summary>
        /// Collapses the hole treeview.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void collapseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.CollapseAll();
        }

        /// <summary>
        /// Reads the description of a windowmacro and returns one language
        /// </summary>
        /// <param name="file">has to be a macrofile</param>
        /// <returns>description from multilanguagestring</returns>
        private string getListViewItemToolTipText(FileInfo file)
        {
            string description = "keine Beschreibung";

            try
            {
                XmlTextReader xreader = new XmlTextReader(file.FullName);
                while (xreader.Read())
                {
                    if (xreader.HasAttributes)
                    {
                        while (xreader.MoveToNextAttribute())
                        {
                            if (xreader.Name == "P23004")
                            {
                                if (xreader.HasValue)
                                {
                                    MultiLangString mulangStr = new MultiLangString();
                                    mulangStr.SetAsString(xreader.Value);
                                    description = mulangStr.GetStringToDisplay(ISOCode.Language.L_de_DE);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lesen von Makro fehleschlagen \n" + ex.Message, "Fehler beim lesen von Makro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
            return description;
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
        private void verticalToolStripMenuItem_Click(object sender, EventArgs e)
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
            instertMacro(((ListviewFileItem)listView1.SelectedItems[0]).FullFilePath, repType, 0);
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
                                            ISettings.CreationFlag.Insert);
            }

            oSettings.SetNumericSetting(USER_SETTING_FORMLOCATION, this.Top, 0);
            oSettings.SetNumericSetting(USER_SETTING_FORMLOCATION, this.Left, 1);
            if (!oSettings.ExistSetting(USER_SETTING_FORMSIZE))
            {
                oSettings.AddNumericSetting(USER_SETTING_FORMSIZE,
                                            new int[] { 0 },
                                            new Range[] { new Range { FromValue = -32768, ToValue = 32768 } },
                                            ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_FORMSIZE, this.Height, 0);
            oSettings.SetNumericSetting(USER_SETTING_FORMSIZE, this.Width, 1);


            if (!oSettings.ExistSetting(USER_SETTING_ORIENTATION))
            {
                oSettings.AddNumericSetting(USER_SETTING_ORIENTATION,
                                            new int[] { 0 },
                                            new Range[] { new Range { FromValue = 0, ToValue = 1 } },
                                            ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_ORIENTATION, (int)this.splitContainer1.Orientation, 0);

            if (!oSettings.ExistSetting(USER_SETTING_SPLITDISTANCE))
            {
                oSettings.AddNumericSetting(USER_SETTING_SPLITDISTANCE,
                                            new int[] { 0 },
                                            new Range[] { new Range { FromValue = -32768, ToValue = 32768 } },
                                            ISettings.CreationFlag.Insert);
            }
            oSettings.SetNumericSetting(USER_SETTING_SPLITDISTANCE, (int)this.splitContainer1.SplitterDistance, 0);

            if (!oSettings.ExistSetting(USER_SETTING_PREVIEW))
            {
                oSettings.AddBoolSetting(USER_SETTING_PREVIEW, new bool[] { false },
                                         ISettings.CreationFlag.Insert);
            }
            oSettings.SetBoolSetting(USER_SETTING_PREVIEW, this.checkBox1.Checked, 0);
        }

        /// <summary>
        /// Deletes the user settings of this script
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void deleteUserSettingsToolStripMenuItem_Click(object sender, EventArgs e)
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

        /// <summary>
        /// Initiates the DragDropEvent on the listview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {            
            //string absoluteMacroName = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + ((ListViewItem)e.Item).Text;
            if (File.Exists(((ListviewFileItem)e.Item).FullFilePath))
            {
                string[] filesTodrag = { ((ListviewFileItem)e.Item).FullFilePath };
                DoDragDrop(new DataObject(DataFormats.FileDrop, filesTodrag), DragDropEffects.Copy);
            }                        
        }       

        /// <summary>
        /// workaround for currentProject value lost when form is active
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void refreshCurrentProjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentProject = getCurrentProject();
        }
        
        /// <summary>
        /// opens selected macropath in explorer-window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openPathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string folderpath = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "");

            try
            {
                Process.Start(folderpath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Pfad konnte nicht geöffnet werden\n"+ ex.Message, "Fehler beim öffnen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        /// <summary>
        /// Refresh treeview with F5 key
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {              
                populateTreeView();
            }
        }

        private void eigenschaftenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WindowMacro wm = new WindowMacro(((ListviewFileItem)listView1.SelectedItems[0]).FullFilePath);
        }        
    }

    public class ListviewFileItem : ListViewItem
    {
        private string p1;
        private int p2;

        public ListviewFileItem(string text, int imageIndex) : base(text, imageIndex) { }
        
        public string FullFilePath { get; set; }
    }

    /// <summary>
    /// Contains the enums of macro represetationtypes
    /// </summary>
    public class WindowMacro
    {
        private const string NAME_ATTRIBUTE = "EplanPxfRoot/@Name";
        private const string REPRESTYPE_ATTRIBUTE = "MacroVariant/@MacroFuncType";
        private const string VARIANT_ATTRIBUTE = "MacroVariant/@VariantId";
        private const string VERSION_ATTRIBUTE = "P11/@P23002";
        private const string SOURCE_ATTRIBUTE = "";
        private const string DESCRIPTION_ATTRIBUTE = "";
        private const string INSPOINTACTIVE_ATTRIBUTE = "";
        private const string INSPOINTX_ATTRIBUTE = "";
        private const string INSPOINTY_ATTRIBUTE = "";
        private const string REFERENCEACTIVE_ATTRIBUTE = "";
        private XmlDocument xmlDoc;
        private XmlNamespaceManager xmlNamSpaceMgr;
        private XmlAttribute xmlAttribute;

        public string Name { get; set; }
        public RepresentationType Representationtype { get; set; }
        public int[] Variants { get; set; }
        public string Version { get; set; }
        public string Source_Reference { get; set; }
        public string Description { get; set; }
        public bool InsertionPointActive { get; set; }
        public double InsertionPoint_X { get; set; }
        public double InsertionPoint_Y { get; set; }
        public bool ReferencesActive { get; set; }

         
        public WindowMacro(string filename)
        {
            xmlDoc = new XmlDocument();
            
            if (File.Exists(filename))
            {
                xmlDoc.Load(filename);
                xmlNamSpaceMgr = new XmlNamespaceManager(xmlDoc.NameTable);

                Name = ((XmlAttribute)xmlDoc.SelectSingleNode(NAME_ATTRIBUTE, xmlNamSpaceMgr)).Value;
                Variants = new int[xmlDoc.SelectNodes(VARIANT_ATTRIBUTE, xmlNamSpaceMgr).Count];
                for (int i = 0; i < xmlDoc.SelectNodes(VARIANT_ATTRIBUTE, xmlNamSpaceMgr).Count; i++)
                {
                    Variants[i] = Convert.ToInt32(xmlDoc.SelectNodes(VARIANT_ATTRIBUTE, xmlNamSpaceMgr).Item(i).Value);
                }
                
                MessageBox.Show(Name+Variants.Length);
            }
        }
        /// <summary>
        /// Dictionary with clear name and enum of representationtype of a macro
        /// </summary>
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

        /// <summary>
        /// macro representationtypes enumerators
        /// </summary>
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
