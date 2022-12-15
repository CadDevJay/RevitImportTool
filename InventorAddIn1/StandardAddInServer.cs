using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using InvAddIn;
using Inventor;
using Microsoft.Win32;

namespace RevitImportTool
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("576E0D55-2222-4297-928D-E875F87B1997")]
    public partial class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        ButtonDefinition oBtnDef = null;
        Ribbon oRibbon;
        RibbonTab oRibbonTab;
        RibbonPanels oPanels;
        RibbonPanel oPanel;
        CommandCategory cmdCat;
        InvAddIn.InputForm oIpForm;

        object Config16obj;
        object Config128obj;

        //InvAddIn.InputForm oIpForm;
        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;
            getIcons();

            //ClientId Declaration
            string targetID = "TargetID_2";
            cmdCat = m_inventorApplication.CommandManager.CommandCategories.Add("RVTImport", "Autodesk:ImportTool:RVTImport", targetID);
            ControlDefinitions oCmdCntrls = m_inventorApplication.CommandManager.ControlDefinitions;

            if (firstTime)
            {
                oRibbon = m_inventorApplication.UserInterfaceManager.Ribbons["ZeroDoc"];
                oRibbonTab = oRibbon.RibbonTabs.Add("ImportTool", "Autodesk:ImportTool", targetID);
                oPanels = oRibbonTab.RibbonPanels;


                oBtnDef = oCmdCntrls.AddButtonDefinition("Select File", "ImportButton", CommandTypesEnum.kShapeEditCmdType, targetID, "Import Revit Files", "auto-break tooltip", Config16obj, Config128obj, ButtonDisplayEnum.kAlwaysDisplayText);
                oBtnDef.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(InputForm);
                cmdCat.Add(oBtnDef);

                RibbonPanel oPanel = oPanels.Add("RVTImport", "Autodesk:ImportTool:RVTImport", targetID);
                oPanel.CommandControls.AddButton(oBtnDef, true);

            }
            // TODO: Add ApplicationAddInServer.Activate implementation.
            // e.g. event initialization, command creation etc.
        }

        private void InputForm(NameValueMap Context)
        {
            oIpForm = new InvAddIn.InputForm(m_inventorApplication);
            oIpForm.ShowDialog();

        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }
        
        private void getIcons()
        {
            //get current assembly
            //Assembly thisDll = Assembly.GetExecutingAssembly();

            //get icon streams
            //Stream Config16stream = thisDll.GetManifestResourceStream("//Revit16.ico");//thisDll.GetManifestResourceStream("InvAddIn.Images.Revit16.ico");//
            Stream Config128stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("//Revit128.ico");

            



            Icon Config16icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("\\Revit16.ico"));
            Icon Config128icon = new Icon(Config128stream);

            //convert to objects
            Config16obj = AxHostConverter.ImageToPictureDisp(Config16icon.ToBitmap());
            Config128obj = AxHostConverter.ImageToPictureDisp(Config128icon.ToBitmap());
        }
       
        
    } 
    #endregion




}

