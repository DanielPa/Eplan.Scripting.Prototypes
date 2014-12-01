using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using System;
using Eplan.EplApi.Gui;
using System.Diagnostics;
using System.Windows.Forms;

namespace Eplanwiki.Scripting.ContextMenu
{
    public class OpenProjectFilePath
    {
        Eplan.EplApi.Gui.ContextMenu menu;
        ContextMenuLocation menuLocation;
        String menuText;        

        [DeclareRegister]
        public void Register()
        {
            InitiateMenu();
        }

        [DeclareUnregister]
        public void UnRegister()
        {
            DisposeMenu();
        }

        public void InitiateMenu()
	    {	    
            menu = new Eplan.EplApi.Gui.ContextMenu();
            menuLocation = new ContextMenuLocation("PmPageObjectTreeDialog", "1007");
		    menu.AddMenuItem(menuLocation, menuText, "OpenProjectFilePath", false, false);            
	    }

        public void DisposeMenu()
        {
            if (menu != null && menuLocation != null)
            {
                menu.RemoveMenuItem(menuLocation, menuText, "OpenProjectFilePath", false, false);
            }
        }

        private string setMenuText()
        {
            ISOCode guiLanguage = new Languages().GuiLanguage;
            MultiLangString muLangMenuText = new MultiLangString();
            muLangMenuText.SetAsString("de_DE@Dateipfad öffnen;en_US@Open Path;");
            return muLangMenuText.GetString((ISOCode.Language)guiLanguage.GetNumber());
        }

        [DeclareAction("OpenProjectFilePath")]
        public void OpenProjectFilePath()
        {
            try
            {
                Process.Start(PathMap.SubstitutePath("$(P)"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
