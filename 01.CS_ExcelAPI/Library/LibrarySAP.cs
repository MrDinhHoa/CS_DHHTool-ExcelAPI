// ReSharper disable All
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using CSiAPIv1;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace _01.CS_ExcelAPI
{
    public class SapClass
    {
        public cSapModel mySapModel { get; set; }
        public void SelectSAP()
        {
            //set the following flag to true to attach to an existing instance of the program
            //otherwise a new instance of the program will be started
            bool AttachToInstance = false;
            //set the following flag to true to manually specify the path to SAP2000.exe
            //this allows for a connection to a version of SAP2000 other than the latest installation
            //otherwise the latest installed version of SAP2000 will be launched
            bool SpecifyPath = false;
            //if the above flag is set to true, specify the path to SAP2000 below
            string ProgramPath;
            ProgramPath = "C:\\Program Files\\Computers and Structures\\SAP2000 22\\SAP2000.exe";
            //full path to the model
            //set it to the desired path of your model

            string ModelDirectory = "C:\\Users\\Admin\\Documents\\Zalo Received Files\\";
            try
            {
                System.IO.Directory.CreateDirectory(ModelDirectory);
            }
            catch (Exception)
            {
                Console.WriteLine("Could not create directory: " + ModelDirectory);
            }

            string ModelName = "02B. GC TIPPING II ( SAN THEP).sdb";
            string ModelPath = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName;
            //dimension the SapObject as cOAPI type

            cOAPI mySapObject = null;
            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
            int ret = 0;
            if (AttachToInstance)
            {
                //attach to a running instance of SAP2000
                try
                {
                    //get the active SapObject
                    mySapObject =
                        (cOAPI) System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
                }
                catch (Exception)
                {
                    Console.WriteLine("No running instance of the program found or failed to attach.");
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
                    Console.WriteLine("Cannot create an instance of the Helper object");
                    return;
                }
                if (SpecifyPath)
                {
                    //'create an instance of the SapObject from the specified path
                    try
                    {
                        //ceate SapObject
                        mySapObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);
                        return;
                    }
                }
                else
                {
                    //'create an instance of the SapObject from the latest installed SAP2000
                    try
                    {
                        //create SapObject
                        mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");

                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Cannot start a new instance of the program.");

                        return;

                    }

                }
                //start SAP2000 application
                ret = mySapObject.ApplicationStart();

            }
            mySapModel = mySapObject.SapModel;
        }

        public class LoadCombinationSap
        {
            public int NumberNames { get; set; }
            public string MyNames { get; set; }
        }

        public class JointReactionSap
        {
            public string Name { get; set; }
            public string LoadCase { get; set; }
            public double F1 { get; set; }
            public double F2 { get; set; }
            public double F3 { get; set; }
            public double M1 { get; set; }
            public double M2 { get; set; }
            public double M3 { get; set; }

        }

        public class JointDisplacementSap
        {
            public string Name { get; set; }
            public string LoadCase { get; set; }
            public double F1 { get; set; }
            public double F2 { get; set; }
            public double F3 { get; set; }
            public double M1 { get; set; }
            public double M2 { get; set; }
            public double M3 { get; set; }

        }

    }
}

