using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using ETABSv17;
using static System.Windows.Forms.DialogResult;

namespace _01.CS_EtabsAPI_ConnectExcelEtabs
{
    public class EtabsClass
    {
        public cOAPI myETABSObject { get; set; }
        public cSapModel mySapModel { get; set; }
        public void SelectEtabs()
        {
            //set the following flag to true to attach to an existing instance of the program 
            //otherwise a new instance of the program will be started 
            bool AttachToInstance = false;

            //set the following flag to true to manually specify the path to ETABS.exe
            //this allows for a connection to a version of ETABS other than the latest installation
            //otherwise the latest installed version of ETABS will be launched
            bool SpecifyPath = false;

            //if the above flag is set to true, specify the path to ETABS below
            string ProgramPath = "C:\\Program Files (x86)\\Computers and Structures\\ETABS 18\\ETABS.exe"; ;

            //full path to the model 
            //set it to an already existing folder 
            string ModelDirectory = "C:\\CSi_ETABS_API_Example";
            try
            {
                System.IO.Directory.CreateDirectory(ModelDirectory);
            }
            catch (Exception)
            {
                MessageBox.Show("Could not create directory: " + ModelDirectory);
            }
            //dimension the ETABS Object as cOAPI type
            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero) 
            if (AttachToInstance)
            {
                //attach to a running instance of ETABS 
                try
                {
                    //get the active ETABS object
                    myETABSObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
                }
                catch (Exception)
                {
                    MessageBox.Show("No running instance of the program found or failed to attach.");
                    return;
                }
            }
            else
            {
                //create API helper object
                cHelper myHelper;
                try
                {
                    myHelper = new Helper();
                }
                catch (Exception)
                {
                    MessageBox.Show("Cannot create an instance of the Helper object");
                    return;
                }

                if (SpecifyPath)
                {
                    //'create an instance of the ETABS object from the specified path
                    try
                    {
                        //create ETABS object
                        myETABSObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Cannot start a new instance of the program from " + ProgramPath);
                        return;
                    }
                }
                else
                {
                    //'create an instance of the ETABS object from the latest installed ETABS
                    try
                    {
                        //create ETABS object
                        myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }
                //start ETABS application
                myETABSObject.ApplicationStart();
            }

            //Get a reference to cSapModel to access all API classes and functions
            mySapModel = myETABSObject.SapModel;
            int units = mySapModel.SetPresentUnits((eUnits) 6);
            int numberName = 0;
            string[] myName = { };
            mySapModel.LoadCases.GetNameList(ref numberName, ref myName);
            List<cLoadCases> listLoad = new List<cLoadCases>();
            for (int i = 0; i < numberName - 1; i++)
            {
                listLoad.Add(myName[i]);
            }
        }
    }
}

    

