/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

//using System;
//using System.Windows.Forms;
//using System.Xml;
//using Eplan.EplApi.Base;
//using Eplan.EplApi.ApplicationFramework;
//using Eplan.EplApi.Gui;
//using Eplan.EplApi.Scripting;

namespace EplanScriptingBBR
{		
	/// <summary>
	/// Description of NewProjectWithIdDescriptions.
	/// </summary>
	public class NewProjectWithIdDescriptions
	{
		
		
		/// <summary>
		/// Erstellt ein neues Projekt anhand einer Projekt-Exportdatei
		/// </summary>
		/// <param name="epjFilePath">Absoluter Dateipfad zur Exportdateil "C:\..\*.epj"</param>
		/// <param name="targetProject">Absoluter Dateipfad zum Zielprojekt "C:\..\*.elk"</param>
		public static void ImportProjectEPJ(string epjFilePath, string targetProject)
		{			
			CommandLineInterpreter oCLI = new CommandLineInterpreter();
			ActionCallingContext acc = new ActionCallingContext();			
			acc.AddParameter("TYPE", "PXFPROJECT");
			acc.AddParameter("IMPORTFILE", epjFilePath);
			acc.AddParameter("PROJECTNAME", targetProject);				
			oCLI.Execute("IMPORT",acc);		
		}
		
		/// <summary>
		/// Öffnet ein Projekt
		/// </summary>
		/// <param name="projectname">Absoluter Dateipfad zum Projekt "C:\..\*.elk"</param>
		public static void OpenProject(string projectname)
		{
			CommandLineInterpreter oCLI = new CommandLineInterpreter();
			ActionCallingContext acc = new ActionCallingContext();
			acc.AddParameter("PROJECTNAME", projectname);
			oCLI.Execute("EDIT",acc);
		}
		
		public static void ImportPrjSettings(string projectname, string settingsFile)
		{
			CommandLineInterpreter oCLI = new CommandLineInterpreter();
			ActionCallingContext acc = new ActionCallingContext();
			acc.AddParameter("XMLFile", settingsFile);
			acc.AddParameter("Project", projectname);
			oCLI.Execute("XSettingsImport",acc);
		}
		
		public static string WriteEpjFile(string sourcefile, Strukturekennzeichen kennzeichen)
		{
			string targetfile = sourcefile+"01";
			/*
			 * ToDo 
			 * Validierung von kennzeichen 
			 *
			 */				
			System.Xml.XmlDocument xDoc = new System.Xml.XmlDocument();
			xDoc.Load(sourcefile);
			xDoc.SelectSingleNode("//O14").AppendChild(NewProjectWithIdDescriptions.CreateO6Element(kennzeichen));
			xDoc.Save(targetfile);
			return targetfile;			
		}
		
		private static System.Xml.XmlElement CreateO6Element(Strukturekennzeichen kennzeichen)
		{
				System.Xml.XmlDocument xDoc = new System.Xml.XmlDocument();
				System.Xml.XmlElement element_O6 = xDoc.CreateElement("O6");
					System.Xml.XmlAttribute attr_Build = xDoc.CreateAttribute("Build");
						attr_Build.Value = "6360";
					element_O6.Attributes.Append(attr_Build);
					
					System.Xml.XmlAttribute attr_A1 = xDoc.CreateAttribute("A1");
						attr_A1.Value = "6/2";
						element_O6.Attributes.Append(attr_A1);
						
					System.Xml.XmlAttribute attr_A3 = xDoc.CreateAttribute("A3");
						attr_A3.Value = "0";
						element_O6.Attributes.Append(attr_A3);
						
					System.Xml.XmlAttribute attr_A13 = xDoc.CreateAttribute("A13");
						attr_A13.Value = "0";
						element_O6.Attributes.Append(attr_A13);
					
					System.Xml.XmlAttribute attr_A14 = xDoc.CreateAttribute("A14");
						attr_A14.Value = "0";
						element_O6.Attributes.Append(attr_A14);
						
					System.Xml.XmlAttribute attr_A81 = xDoc.CreateAttribute("A81");
						attr_A81.Value = "1100";
						element_O6.Attributes.Append(attr_A81);
						
					System.Xml.XmlAttribute attr_A82 = xDoc.CreateAttribute("A82");
						attr_A82.Value = kennzeichen.Bezeichnung;
						element_O6.Attributes.Append(attr_A82);
						
					System.Xml.XmlAttribute attr_A85 = xDoc.CreateAttribute("A85");
						attr_A85.Value = "10002";
						element_O6.Attributes.Append(attr_A85);
						
					
					System.Xml.XmlElement element_P11 = xDoc.CreateElement("P11");
						System.Xml.XmlAttribute attr_P10002 = xDoc.CreateAttribute("P1002");
						attr_P10002.Value = kennzeichen.Beschreibung1.GetAsString();
						element_P11.Attributes.Append(attr_P10002);
						
						System.Xml.XmlAttribute attr_P10007 = xDoc.CreateAttribute("P1007");
							attr_P10007.Value = kennzeichen.Beschreibung2.GetAsString();
						element_P11.Attributes.Append(attr_P10007);
						
						System.Xml.XmlAttribute attr_P10008 = xDoc.CreateAttribute("P1008");
							attr_P10008.Value = kennzeichen.Beschreibung3.GetAsString();
						element_P11.Attributes.Append(attr_P10008);
						
						System.Xml.XmlAttribute attr_P10009_1 = xDoc.CreateAttribute("P1009_1");
							attr_P10009_1.Value = kennzeichen.BeschreibungZusatzfeld1.GetAsString();
						element_P11.Attributes.Append(attr_P10009_1);
						
						System.Xml.XmlAttribute attr_P10009_2 = xDoc.CreateAttribute("P1009_2");
							attr_P10009_2.Value = kennzeichen.BeschreibungZusatzfeld2.GetAsString();
						element_P11.Attributes.Append(attr_P10009_2);
						
						System.Xml.XmlAttribute attr_P10009_3 = xDoc.CreateAttribute("P1009_3");
							attr_P10009_3.Value = kennzeichen.BeschreibungZusatzfeld3.GetAsString();
						element_P11.Attributes.Append(attr_P10009_3);
						
						System.Xml.XmlAttribute attr_P10009_4 = xDoc.CreateAttribute("P1009_4");
							attr_P10009_4.Value = kennzeichen.BeschreibungZusatzfeld4.GetAsString();
						element_P11.Attributes.Append(attr_P10009_4);
						
						System.Xml.XmlAttribute attr_P10009_5 = xDoc.CreateAttribute("P1009_5");
							attr_P10009_5.Value = kennzeichen.BeschreibungZusatzfeld5.GetAsString();
						element_P11.Attributes.Append(attr_P10009_5);
						
						System.Xml.XmlAttribute attr_P10009_6 = xDoc.CreateAttribute("P1009_6");
							attr_P10009_6.Value = kennzeichen.BeschreibungZusatzfeld6.GetAsString();
						element_P11.Attributes.Append(attr_P10009_6);
						
						System.Xml.XmlAttribute attr_P10009_7 = xDoc.CreateAttribute("P1009_7");
							attr_P10009_7.Value = kennzeichen.BeschreibungZusatzfeld7.GetAsString();
						element_P11.Attributes.Append(attr_P10009_7);
						
						System.Xml.XmlAttribute attr_P10009_8 = xDoc.CreateAttribute("P1009_8");
							attr_P10009_8.Value = kennzeichen.BeschreibungZusatzfeld8.GetAsString();
						element_P11.Attributes.Append(attr_P10009_8);
						
						System.Xml.XmlAttribute attr_P10009_9 = xDoc.CreateAttribute("P1009_9");
							attr_P10009_9.Value = kennzeichen.BeschreibungZusatzfeld9.GetAsString();
						element_P11.Attributes.Append(attr_P10009_9);
						
						System.Xml.XmlAttribute attr_P10009_10 = xDoc.CreateAttribute("P1009_10");
							attr_P10009_10.Value = kennzeichen.BeschreibungZusatzfeld10.GetAsString();
						element_P11.Attributes.Append(attr_P10009_10);
						
					element_O6.InnerXml = element_P11.OuterXml;
			return element_O6;
		}
		
		
		public class Strukturekennzeichen 
		{
			string _Bezeichnung;
			public Strukturekennzeichen (string bezeichnung)
			{
				_Bezeichnung = bezeichnung; 
			}
			public string Bezeichnung { get{return _Bezeichnung;}}
			public MultiLangString Beschreibung1 { get; set; }
			public MultiLangString Beschreibung2 { get; set; }
			public MultiLangString Beschreibung3 { get; set; }
			public MultiLangString BeschreibungZusatzfeld1 { get; set; }
			public MultiLangString BeschreibungZusatzfeld2 { get; set; }
			public MultiLangString BeschreibungZusatzfeld3 { get; set; }
			public MultiLangString BeschreibungZusatzfeld4 { get; set; }
			public MultiLangString BeschreibungZusatzfeld5 { get; set; }
			public MultiLangString BeschreibungZusatzfeld6 { get; set; }
			public MultiLangString BeschreibungZusatzfeld7 { get; set; }
			public MultiLangString BeschreibungZusatzfeld8 { get; set; }
			public MultiLangString BeschreibungZusatzfeld9 { get; set; }
			public MultiLangString BeschreibungZusatzfeld10 { get; set; }
				
		}
		
		/// <summary>
		/// Einstiegspunkt Script/Ausführen... 
		/// </summary>
		[Start]
		public void StartScript()
		{
			string epjsource =@"C:\Dokumente und Einstellungen\PappD\Desktop\ST10.epj";
			string epjtarget = string.Empty;
			string absPrjName =@"C:\EPLAN_P8\Projects\ST10.elk";
			string settings =@"C:\EPLAN_P8\XML\ProjectSettings.xml";
			string auswertungen =@"C:\EPLAN_P8\XML\Auswertungen.xml";
			
			Strukturekennzeichen stk = new Strukturekennzeichen("ND01");
			MultiLangString beschreibung1 = new MultiLangString();
			beschreibung1.SetAsString("de_DE@Beschreibung;en_US@Description;;");
			//beschreibung1.AddString(ISOCode.Language.L_de_DE, "Beschreibung");
			//beschreibung1.AddString(ISOCode.Language.L_en_US, "Description");
			stk.Beschreibung1 = beschreibung1;
			epjtarget = NewProjectWithIdDescriptions.WriteEpjFile(epjsource, stk);
			
			NewProjectWithIdDescriptions.ImportProjectEPJ(epjtarget,absPrjName);
							
			NewProjectWithIdDescriptions.ImportPrjSettings(absPrjName,settings);
			NewProjectWithIdDescriptions.ImportPrjSettings(absPrjName,auswertungen);
			NewProjectWithIdDescriptions.OpenProject(absPrjName);
			new CommandLineInterpreter().Execute("XPrjActionProjectCompleteMasterData");			
		}
	}
}
