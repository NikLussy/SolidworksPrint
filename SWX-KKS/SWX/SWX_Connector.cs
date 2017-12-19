using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SldWorks;
using SwConst;
using System.Runtime.InteropServices;

namespace SWX_KKS.SWX
{
    class SWX_Connector
    {
        //public static dynamic App = null;
        public static SldWorks.SldWorks App;


        public static void GetApp()
        {
            App = (SldWorks.SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            if (App == null)
            {
                KKS.KKS_Message.Show("Kein SWX am laufen");
                SWX.Settings.Run = false;
            }
        }

        public static ModelDoc2 modeldoc = null;
        public static AssemblyDoc assemblyDoc = null;
        public static Configuration configuration = null;
        public static IComponent2 iComponent = null;
        public static IComponent2 Child = null;
        public static ModelDocExtension modelDocExtension = null;
        public static DrawingDoc drawingDoc = null;
        public static Sheet sheet = null;
        public static RevisionTableAnnotation swRevTable = null;
        public static View swView = null;
        public static SldWorks.Annotation swAnot = null;

        public static void GetActualDoc(string[] Args)
        {
            try
            {
                foreach (string Arg in Args)
                {
                    switch (Arg)
                    {
                        case "modeldoc":
                            modeldoc = (ModelDoc2)App.ActiveDoc;
                            break;
                        case "assemblyDoc":
                            assemblyDoc = (AssemblyDoc)App.ActiveDoc;
                            break;
                        case "configuration":
                            configuration = (Configuration)modeldoc.GetActiveConfiguration();
                            break;
                        case "iComponent":
                            iComponent = (IComponent2)configuration.GetRootComponent();
                            break;
                        //case "Child":
                        //    iComponent = (IComponent2)configuration.GetRootComponent();
                        //    break;
                    }
                }
            }
            catch (Exception ex)
            {
                //KKS.KKS_Message.Show(ex.Message);
            }
            
            //return modeldoc;
        }
        public static void GetAssemblyDoc()
        {
            assemblyDoc = (AssemblyDoc)App.ActiveDoc;
            //return modeldoc;
        }


        public static void Close()
        {
            try
            {
                Marshal.FinalReleaseComObject(swAnot);
                Marshal.FinalReleaseComObject(swView);
                Marshal.FinalReleaseComObject(modeldoc);
                Marshal.FinalReleaseComObject(assemblyDoc);
                Marshal.FinalReleaseComObject(configuration);
                Marshal.FinalReleaseComObject(iComponent);
                Marshal.FinalReleaseComObject(Child);
                Marshal.FinalReleaseComObject(modelDocExtension);
                Marshal.FinalReleaseComObject(drawingDoc);
                Marshal.FinalReleaseComObject(sheet);
                Marshal.FinalReleaseComObject(swRevTable);
                Marshal.FinalReleaseComObject(App);
            }
            catch
            { }
            finally
            {
                System.Windows.Forms.Application.Exit();
            }
        }
    }
}
