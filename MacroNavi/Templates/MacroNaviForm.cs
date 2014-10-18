﻿using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using System.IO;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using System.Diagnostics;
using System.Drawing;
using System;
using System.Runtime.InteropServices;   

public partial class MacroNavi : System.Windows.Forms.Form
{
    private SplitContainer splitContainer1;
    private TreeView treeView1;
    private ListView listView1;
    private ColumnHeader columnHeader1;
    private ColumnHeader columnHeader2;
    private ColumnHeader columnHeader3;
    string macropath;
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
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeView1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.listView1);
            this.splitContainer1.Size = new System.Drawing.Size(432, 437);
            this.splitContainer1.SplitterDistance = 110;
            this.splitContainer1.TabIndex = 0;
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.Location = new System.Drawing.Point(12, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(95, 415);
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
            this.listView1.Size = new System.Drawing.Size(303, 415);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
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
            this.allpoligToolStripMenuItem.Click += new System.EventHandler(this.allpoligToolStripMenuItem_Click);
            // 
            // einpoligToolStripMenuItem
            // 
            this.einpoligToolStripMenuItem.Name = "einpoligToolStripMenuItem";
            this.einpoligToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.einpoligToolStripMenuItem.Text = "Einpolig";
            this.einpoligToolStripMenuItem.Click += new System.EventHandler(this.einpoligToolStripMenuItem_Click);
            // 
            // paarquerverweisToolStripMenuItem
            // 
            this.paarquerverweisToolStripMenuItem.Name = "paarquerverweisToolStripMenuItem";
            this.paarquerverweisToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.paarquerverweisToolStripMenuItem.Text = "Paarquerverweis";
            this.paarquerverweisToolStripMenuItem.Click += new System.EventHandler(this.paarquerverweisToolStripMenuItem_Click);
            // 
            // übersichtToolStripMenuItem
            // 
            this.übersichtToolStripMenuItem.Name = "übersichtToolStripMenuItem";
            this.übersichtToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.übersichtToolStripMenuItem.Text = "Übersicht";
            this.übersichtToolStripMenuItem.Click += new System.EventHandler(this.übersichtToolStripMenuItem_Click);
            // 
            // grafikToolStripMenuItem
            // 
            this.grafikToolStripMenuItem.Name = "grafikToolStripMenuItem";
            this.grafikToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.grafikToolStripMenuItem.Text = "Grafik";
            this.grafikToolStripMenuItem.Click += new System.EventHandler(this.grafikToolStripMenuItem_Click);
            // 
            // schaltschrankaufbauToolStripMenuItem
            // 
            this.schaltschrankaufbauToolStripMenuItem.Name = "schaltschrankaufbauToolStripMenuItem";
            this.schaltschrankaufbauToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.schaltschrankaufbauToolStripMenuItem.Text = "Schaltschrankaufbau";
            this.schaltschrankaufbauToolStripMenuItem.Click += new System.EventHandler(this.schaltschrankaufbauToolStripMenuItem_Click);
            // 
            // rIFließbildToolStripMenuItem
            // 
            this.rIFließbildToolStripMenuItem.Name = "rIFließbildToolStripMenuItem";
            this.rIFließbildToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.rIFließbildToolStripMenuItem.Text = "RI-Fließbild";
            this.rIFließbildToolStripMenuItem.Click += new System.EventHandler(this.rIFließbildToolStripMenuItem_Click);
            // 
            // allpoligFluidToolStripMenuItem
            // 
            this.allpoligFluidToolStripMenuItem.Name = "allpoligFluidToolStripMenuItem";
            this.allpoligFluidToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.allpoligFluidToolStripMenuItem.Text = "Allpolig Fluid";
            this.allpoligFluidToolStripMenuItem.Click += new System.EventHandler(this.allpoligFluidToolStripMenuItem_Click);
            // 
            // dMontageaufbauToolStripMenuItem
            // 
            this.dMontageaufbauToolStripMenuItem.Name = "dMontageaufbauToolStripMenuItem";
            this.dMontageaufbauToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.dMontageaufbauToolStripMenuItem.Text = "3D-Montageaufbau";
            this.dMontageaufbauToolStripMenuItem.Click += new System.EventHandler(this.dMontageaufbauToolStripMenuItem_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(361, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(71, 17);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "Vorschau";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aufteilungToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(131, 26);
            // 
            // aufteilungToolStripMenuItem
            // 
            this.aufteilungToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.horizontalToolStripMenuItem,
            this.vertikalToolStripMenuItem});
            this.aufteilungToolStripMenuItem.Name = "aufteilungToolStripMenuItem";
            this.aufteilungToolStripMenuItem.Size = new System.Drawing.Size(130, 22);
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
            // test
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(432, 474);
            this.ContextMenuStrip = this.contextMenuStrip2;
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.splitContainer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "test";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Makros";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.test_FormClosing);
            this.Load += new System.EventHandler(this.test_Load);
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

    [Start]
    public void Function()
    {
        MacroNavi frm = new MacroNavi();
        Process oCurrent = Process.GetCurrentProcess();
        var ww = new WindowWrapper(oCurrent.MainWindowHandle);
        
        Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
        if (oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION"))
        {
            frm.Top = oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION", 0);
            frm.Left = oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION", 1);
        }
        else
        {
            frm.Top = 100;
            frm.Left = 500;
        }
        if (oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE"))
        {                    
            frm.Height = oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE", 0);
            frm.Width = oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE", 1);
        }
        else
        {            
            frm.Height = 411;
            frm.Width = 192;
        }

        if (oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.ORIENTATION"))
        {
            frm.splitContainer1.Orientation = (Orientation)oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.ORIENTATION", 0);             
        }
        else
        {
            frm.splitContainer1.Orientation = (Orientation)0;
        }

        if (oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.SPLITDISTANCE"))
        {
            frm.splitContainer1.SplitterDistance = oSettings.GetNumericSetting("USER.SCRIPTS.MACRONAVI.SPLITDISTANCE", 0);             
        }
        else
        {
            frm.splitContainer1.SplitterDistance = 200;
        }

        if (oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.PREVIEW"))
        {
            frm.checkBox1.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.MACRONAVI.PREVIEW", 0);
        }
        else
        {
            frm.checkBox1.Checked = true;
        }

        frm.Show(ww);        
        new CommandLineInterpreter().Execute("XSDPreviewAction /SHOW:"+ Convert.ToInt16(frm.checkBox1.Checked).ToString());
        
    }

    

    private void test_Load(object sender, System.EventArgs e)
    {
        PopulateTreeView();
    }

    private void PopulateTreeView()
    {
        TreeNode rootNode;

        macropath = PathMap.SubstitutePath("$(MD_MACROS)");
        DirectoryInfo info = new DirectoryInfo(macropath+@"\");
        if (info.Exists)
        {
            rootNode = new TreeNode(info.Name);
            rootNode.Tag = info;
            GetDirectories(info.GetDirectories(), rootNode);
            treeView1.Nodes.Add(rootNode);
        }
    }

    private void GetDirectories(DirectoryInfo[] subDirs,
            TreeNode nodeToAddTo)
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
                GetDirectories(subSubDirs, aNode);
            }
            nodeToAddTo.Nodes.Add(aNode);
        }
    }

    

    private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
    {
        if (listView1.SelectedItems.Count>0 && checkBox1.Checked)
        {
            string absoluteMacroName = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") +"\\"+ listView1.SelectedItems[0].Text;
            new CommandLineInterpreter().Execute("XSDPreviewAction /MACRONAME:" + "\"" + absoluteMacroName + "\"");            
            //MessageBox.Show("XSDPreviewAction /MACRONAME:" + "\"" + absoluteMacroName + "\"");
        }
        listView1.Focus();        
    }

    

    private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.MultiLine, 0);
    }

    private void listView1_KeyPress(object sender, KeyPressEventArgs e)
    {
        
        if (e.KeyChar == (char) Keys.Return)
        {
            instertMacro(WindowMacro.RepresentationType.MultiLine, 0);
        }
        
    }

    private void instertMacro(WindowMacro.RepresentationType reptype, Int32 variant)
    {
        
        if (listView1.SelectedItems.Count > 0)
        {
            string absoluteMacroName = macropath + treeView1.SelectedNode.FullPath.Replace(treeView1.Nodes[0].Text, "") + "\\" + listView1.SelectedItems[0].Text;
            new CommandLineInterpreter().Execute("XGedStartInteractionAction /Name:XMIaInsertMacro /filename:"
            + "\"" + absoluteMacroName + "\"" +
            "/variant:"+ variant.ToString() +" /RepresentationType:" + ((int)reptype).ToString());
        }
    }

    private void test_FormClosing(object sender, FormClosingEventArgs e)
    {
        settingsWrite();
        new CommandLineInterpreter().Execute("XSDPreviewAction /SHOW:0");
        
    }

    private void checkBox1_CheckedChanged(object sender, EventArgs e)
    {
        new CommandLineInterpreter().Execute("XSDPreviewAction /SHOW:" + Convert.ToInt32(checkBox1.Checked).ToString());
    }

    private void alleErweiternToolStripMenuItem_Click(object sender, EventArgs e)
    {
        treeView1.ExpandAll();
    }

    private void alleReduzierenToolStripMenuItem_Click(object sender, EventArgs e)
    {
        treeView1.CollapseAll();
    }

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

    private void horizontalToolStripMenuItem_Click(object sender, EventArgs e)
    {
        splitContainer1.Orientation = Orientation.Horizontal;
    }

    private void vertikalToolStripMenuItem_Click(object sender, EventArgs e)
    {
        splitContainer1.Orientation = Orientation.Vertical;
    }

    private void allpoligToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.MultiLine, 0);
    }

    private void einpoligToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.SingleLine, 0);
    }

    private void paarquerverweisToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.PairCrossReference, 0);
    }

    private void übersichtToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.Overview, 0);
    }

    private void grafikToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.Graphics, 0);
    }

    private void schaltschrankaufbauToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.ArticlePlacement, 0);
    }

    private void rIFließbildToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.PI_FlowChart, 0);
    }

    private void allpoligFluidToolStripMenuItem_Click(object sender, EventArgs e)
    {
        instertMacro(WindowMacro.RepresentationType.Fluid_MultiLine, 0);
    }

    private void dMontageaufbauToolStripMenuItem_Click(object sender, EventArgs e)
    {        
        instertMacro(WindowMacro.RepresentationType.ArticlePlacement3D, 0);
    }

    private void settingsWrite()
    {
        
        Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
        if (!oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION"))
        {
            oSettings.AddNumericSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION",
                new int[] { 0 },
                new Range[] { new Range { FromValue = 0, ToValue = 32768 } },
                "Default value of test setting",
                new int[] { 0 },
                ISettings.CreationFlag.Insert);
        }
        
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION", this.Top, 0);
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMLOCATION", this.Left, 1);
        if (!oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE"))
        {
            oSettings.AddNumericSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE",
                new int[] { 0 },
                new Range[] { new Range { FromValue = 0, ToValue = 32768 } },
                "Default value of test setting",
                new int[] { 0 },
                ISettings.CreationFlag.Insert);
        }
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE", this.Height, 0);
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.FORMSIZE", this.Width, 1);

               
        if (!oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.ORIENTATION"))
        {
            oSettings.AddNumericSetting("USER.SCRIPTS.MACRONAVI.ORIENTATION",
                new int[] { 0 },
                new Range[] { new Range { FromValue = 0, ToValue = 1 } },
                "Default value of test setting",
                new int[] { 0 },
                ISettings.CreationFlag.Insert);
        }
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.ORIENTATION", (int)this.splitContainer1.Orientation, 0);

        if (!oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.SPLITDISTANCE"))
        {
            oSettings.AddNumericSetting("USER.SCRIPTS.MACRONAVI.SPLITDISTANCE",
                new int[] { 0 },
                new Range[] { new Range { FromValue = 0, ToValue = 32768 } },
                "Default value of test setting",
                new int[] { 0 },
                ISettings.CreationFlag.Insert);
        }
        oSettings.SetNumericSetting("USER.SCRIPTS.MACRONAVI.SPLITDISTANCE", (int)this.splitContainer1.SplitterDistance, 0);

        if (!oSettings.ExistSetting("USER.SCRIPTS.MACRONAVI.PREVIEW"))
        {            
            oSettings.AddBoolSetting("USER.SCRIPTS.MACRONAVI.PREVIEW", new bool[] { false },                
                "Default value of test setting",
                new bool[] { false },
                ISettings.CreationFlag.Insert);                 
        }
        oSettings.SetBoolSetting("USER.SCRIPTS.MACRONAVI.PREVIEW", this.checkBox1.Checked, 0);
    }

}
public abstract class WindowMacro
{
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