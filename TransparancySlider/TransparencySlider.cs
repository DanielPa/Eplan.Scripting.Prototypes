using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

namespace DanielPa.Scripting.Prototypes
{
  public class TransparencySlider
  {
    private const string ATTRIBUTE_LAYER_NAME = "A1424";
    private const string ATTRIBUTE_LAYER_TRANSPARENCY = "A1434";

    [Start]
    [DeclareAction("TransparencySlider")]
    public void Execute()
    {
      var layer = GetLayerNameAndDescription();
      
      var percentage = GetCurrentTransparencyState(layer.Key);
      ShowSlider(percentage, layer);
    }

    [DeclareMenu]
    [DeclareRegister]
    public void AddContextMenu()
    {
      //XCabPlacerTreePage 4010
      var contextMenu = new Eplan.EplApi.Gui.ContextMenu();
      ContextMenuLocation location = new ContextMenuLocation("XCabPlacerTreePage", "4010");
      contextMenu.AddMenuItem(location, MENU_NAME, ACTION_NAME, false, false);
    }

    [DeclareUnregister]
    public void RemoveContextMenu()
    {
      var contextMenu = new Eplan.EplApi.Gui.ContextMenu();
      ContextMenuLocation location = new ContextMenuLocation("XCabPlacerTreePage", "4010");
      contextMenu.RemoveMenuItem(location, MENU_NAME, ACTION_NAME, false, false);
    }

    private const string MENU_NAME = "Transparency...";
    private const string ACTION_NAME = "TransparencySlider";

    private KeyValuePair<string, string> GetLayerNameAndDescription()
    {
      // XEsGetPropertyAction /PropertyId:? /PropertyIndex:0
      string value = null;
      var context = new ActionCallingContext();
      context.AddParameter("PropertyId", "2000");
      context.AddParameter("PropertyIndex", "0");
      var cli = new CommandLineInterpreter(true, true);
      cli.Execute("XEsGetPropertyAction", context);
      context.GetParameter("PropertyValue", ref value);
      var obj = StorableObjectWrapper.FromStringIdentifier(value);
      var placement3D = new Placement3DWrapper(obj);
      var description = placement3D.Layer.Description.Split('@').Last().TrimEnd(';');
      return new KeyValuePair<string, string>(placement3D.Layer.Name, description);
    }

    private void ShowSlider(float percentage, KeyValuePair<string,string> layer)
    {
      var form = new System.Windows.Forms.Form();
      var stackPanel = new System.Windows.Forms.FlowLayoutPanel();
      var panel = new System.Windows.Forms.Panel();
      panel.BackColor = Color.SteelBlue;
      panel.Padding = new Padding(1); // This will create a 1px border
      panel.AutoSize = true;
      
      stackPanel.Dock = DockStyle.Fill; 
      stackPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      stackPanel.BackColor = Color.White;
      
      stackPanel.AutoSize = true;
      stackPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
      panel.Controls.Add(stackPanel);
      form.Controls.Add(panel);
      var label = new System.Windows.Forms.Label();
      label.Text = string.Format("Transparency Slider [{0} - {1}]", layer.Key, layer.Value);
      label.AutoSize = true;
      label.Margin = new Padding(4, 4, 4, 4);
      stackPanel.Controls.Add(label);
      var slider = new System.Windows.Forms.TrackBar();
      slider.Width = 400;
      slider.Minimum = 0;
      slider.Maximum = 100;
      slider.TickFrequency = 10;
      slider.LargeChange = 10;
      slider.SmallChange = 10;
      slider.BackColor = Color.White;
      
      try
      {
        slider.Value = (int)(percentage * 100);
      }
      catch (Exception)
      {
        slider.Value = 0;
      }
      slider.TickStyle = System.Windows.Forms.TickStyle.Both;
      slider.ValueChanged += (sender, args) => SetTransparency(layer.Key, slider.Value / 100f);
      stackPanel.Controls.Add(slider);

      // Set form properties
      form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
      form.BackColor = Color.White;
      form.TopMost = true;
      form.AutoSize = true;
      form.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
      form.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
      form.Location = System.Windows.Forms.Cursor.Position;

      // Close form when it loses focus
      form.Deactivate += (sender, args) => form.Close();

      // Close form when Escape key is pressed
      slider.KeyDown += (sender, args) =>
      {
        if (args.KeyCode == System.Windows.Forms.Keys.Escape)
        {
          form.Close();
        }
      };
      form.KeyDown += (sender, args) =>
      {
        if (args.KeyCode == System.Windows.Forms.Keys.Escape)
        {
          form.Close();
        }
      };

      WindowWrapper windowWrapper = new WindowWrapper(Process.GetCurrentProcess().MainWindowHandle);
      form.Show(windowWrapper);
    }

    private void SetTransparency(string layerName, float sliderValue)
    {
      // changelayer /LAYER:560 /VISIBLE:1 /COLORID:9 /TRANSPARENCY:0.1
      var context = new ActionCallingContext();
      context.AddParameter("LAYER", layerName);
      context.AddParameter("TRANSPARENCY", sliderValue.ToString());
      new CommandLineInterpreter().Execute("changelayer", context);
    }

    private float GetCurrentTransparencyState(string layerName)
    {
      var fileName = ExportLayerTable();
      var transparency = GetTransparencyValue(fileName, layerName);
      return transparency;
    }

    private float GetTransparencyValue(string fileName, string layerName)
    {
      // <O76 Build="5313" A1="76/335" A3="0" A13="0" A14="0" R1421="14/1" A1422="1" A1423="1" A1424="EPLAN560" A1425="##_##@560/ESGraphics;??_??@3D-Grafik.Schrank;" A1426="560" A1427="0" A1428="274" A1429="0.5" A1430="-1" A1433="0" A1434="77" A1435="1">
      var document = new XmlDocument();
      document.Load(fileName);
      var xPathSelectElementByLayerName =
        string.Format("/EplanPxfRoot/O76[@{0}='{1}']", ATTRIBUTE_LAYER_NAME, layerName);
      var layerElement = document.SelectSingleNode(xPathSelectElementByLayerName);
      var transparencyByteValue = layerElement.Attributes[ATTRIBUTE_LAYER_TRANSPARENCY].Value;
      var transparencyPercentage = float.Parse(transparencyByteValue) / 255;
      return transparencyPercentage;
    }

    private string ExportLayerTable()
    {
      string fileName = Path.Combine(Path.GetTempPath(), "Layer.xml");
      if (File.Exists(fileName))
      {
        File.Delete(fileName);
      }
      var context = new ActionCallingContext();
      context.AddParameter("TYPE", "EXPORT");
      context.AddParameter("EXPORTFILE", fileName);
      new CommandLineInterpreter().Execute("graphicallayertable", context);
      return fileName;
    }
  }

  public class Placement3DWrapper
  {
    private readonly object _placement3D;

    public Placement3DWrapper(object o)
    {
      _placement3D = o;
    }

    public GraphicalLayerWrapper Layer
    {
      get
      {
        var layer = _placement3D.GetType().GetProperty("Layer");
        var value = layer.GetValue(_placement3D);
        return new GraphicalLayerWrapper(value);
      }
    }
  }

  public class GraphicalLayerWrapper
  {
    private readonly object _graphicalLayer;

    public GraphicalLayerWrapper(object value)
    {
      _graphicalLayer = value;
    }

    public string Name
    {
      get
      {
        var name = _graphicalLayer.GetType().GetProperty("Name");
        var value = name.GetValue(_graphicalLayer);
        return value.ToString();
      }
    }

    public string Description
    {
      get
      {
        var description = _graphicalLayer.GetType().GetProperty("Description");
        var value = description.GetValue(_graphicalLayer);
        return value.ToString();
      }
    }
  }

  public class StorableObjectWrapper
  {
    private static Assembly _dataModelAssembly;

    public static object FromStringIdentifier(string databaseId)
    {
      if (_dataModelAssembly == null)
      {
        _dataModelAssembly = AppDomain.CurrentDomain.GetAssemblies()
                                      .FirstOrDefault(a => a.FullName.StartsWith("Eplan.EplApi.DataModelu"));
      }

      var storableObjectType = _dataModelAssembly.ExportedTypes.FirstOrDefault(t => t.Name == "StorableObject");
      MethodInfo fromStringIdentifier = storableObjectType.GetMethod("FromStringIdentifier",
                                                                     BindingFlags.Public | BindingFlags.Static, null,
                                                                     new[] { typeof(string) }, null);
      var args = new object[] { databaseId };
      var storableObject = fromStringIdentifier.Invoke(null, args);
      return storableObject;
    }
  }

  public class WindowWrapper : IWin32Window
  {
    private readonly IntPtr _hwnd;

    public WindowWrapper(IntPtr handle)
    {
      _hwnd = handle;
    }

    public IntPtr Handle
    {
      get { return _hwnd; }
    }
  }
}