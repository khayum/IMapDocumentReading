using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Catalog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

namespace IMapDocumentReading
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {

                //ConsoleUtility.WriteProgressBar(0);         
                string title = @"
+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+
 M|X|D| |S|a|t|e|l|l|i|t|e| |S|o|u|r|c|e| |R|e|p|a|i|r|i|n|g| |T|o|o|l|
+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+";

                string banner = "This tool will replace the old Satellite images from the current " + Environment.NewLine + "mxd and save a copy as V1 appended with new satellite images";

                Console.WriteLine(title);
                Console.WriteLine("+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+");
                Console.WriteLine(banner);
                Console.WriteLine("+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+");
                Console.WriteLine("Base Path Format : D:\\existingMxds\\");
                Console.WriteLine("+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+ +-+-+-+-+");

                Console.WriteLine();


            A: Console.WriteLine("Please enter the base path where mxds reside : ");
                string mxdBasePathFromUser = Console.ReadLine();
                string mxdSourcePath = @mxdBasePathFromUser;

                if (String.IsNullOrEmpty(mxdSourcePath) || !System.IO.Directory.Exists(mxdSourcePath))
                {
                    Console.WriteLine("MXD Source path doesn't exist");
                    goto A;

                }
                else if (!PathHasAtLeastOneFolder(mxdSourcePath))
                {
                    Console.WriteLine("MXD Source path with drive only not allowed");
                    goto A;
                }

                string logFile = Path.Combine(mxdBasePathFromUser, "log_" + DateTime.Now.Millisecond.ToString() + ".txt");
                string satellitePath = Path.Combine(Environment.CurrentDirectory, "LayerFiles", "Satellite.lyr");
                string demPath = Path.Combine(Environment.CurrentDirectory, "LayerFiles", "dem2012.lyr");


                if (!System.IO.File.Exists(satellitePath))
                {
                    Console.WriteLine("Satellite layer doesn't exist");
                    return;
                }

                if (!System.IO.File.Exists(demPath))
                {
                    Console.WriteLine("DEM layer doesn't exist");
                    return;
                }

                System.IO.FileStream fs = null;

                if (File.Exists(logFile))
                {
                    File.Delete(logFile);   
                }

                fs = File.Create(logFile);
                fs?.Close();

                deleteExistingProcessesFiles(mxdSourcePath);
                execute(mxdSourcePath, logFile, satellitePath, demPath);

            }catch(Exception ex)
            {

                Console.WriteLine(ex.ToString());   
                

            }
            finally
            {
                Console.WriteLine("Press any key to exit ..............");
                Console.ReadLine();
            }

        }

        static bool PathHasAtLeastOneFolder(string path)
        {
            return Path.GetPathRoot(path).Length < Path.GetDirectoryName(path)?.Length;
        }

        static void execute(string mxdSourcePath,string logFile,string satellitePath,string demPath)
        {

            string msg = "";
            string[] mxdFiles;

            try
            {
                               
                 mxdFiles = System.IO.Directory.GetFiles(mxdSourcePath, "*.mxd", SearchOption.AllDirectories);
                 WriteLog("Total # of files to be processed : " + mxdFiles.Length, logFile);   

                if (!ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.Engine))
                {
                    if (!ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.Desktop))
                    {
                        WriteLog("License Initialization error", logFile);
                        return;
                    }
                }

                LicenseInitializer aoLicenseInitializer = new LicenseInitializer();

                if (!aoLicenseInitializer.InitializeApplication(new esriLicenseProductCode[] { esriLicenseProductCode.esriLicenseProductCodeBasic, esriLicenseProductCode.esriLicenseProductCodeStandard, esriLicenseProductCode.esriLicenseProductCodeAdvanced }))
                {
                    WriteLog("License Initialization error", logFile);
                    aoLicenseInitializer.ShutdownApplication();
                    return;
                }

                
                ConsoleUtility.WriteProgressBar(0);
                int interval = 100 / mxdFiles.Length;
                int k = interval;

                if((interval * mxdFiles.Length) < 100)
                {
                    int diff = 100 - (interval * mxdFiles.Length);
                    k = k + diff;
                }

                foreach (string mxdFile in mxdFiles)
                {

                    try
                    {

                        ConsoleUtility.WriteProgressBar(k, true);
                        Thread.Sleep(50);

                        IMapDocument pMapDoc = new MapDocument();
                        string detinationFile = Path.Combine(System.IO.Path.GetDirectoryName(mxdFile), System.IO.Path.GetFileNameWithoutExtension(mxdFile) + "_V1.mxd");

                        //msg = "Reading from " + mxdFile + " and saving as " + detinationFile;
                        //WriteLog(msg, logFile);
                        //Console.WriteLine(msg);

                        if (File.Exists(detinationFile))
                        {
                            WriteLog("Destination file exists.Deleted", logFile);
                            Console.WriteLine("Destination file exists.Deleted");
                            File.Delete(detinationFile);
                        }

                        pMapDoc.Open(mxdFile);

                        IDocumentInfo2 pDocInfo = (IDocumentInfo2)pMapDoc;

                        for (int i = 0; i <= pMapDoc.MapCount - 1; i++)
                        {
                            IMap map = pMapDoc.Map[i];
                            IEnumLayer enumLayer = map.Layers;
                            ILayer pLayer;
                            ILayer layerToBeAdded = null;

                            while (null != (pLayer = enumLayer.Next()))
                            {
                                //Console.WriteLine(pLayer.Name);

                                layerToBeAdded = null;

                                if (pLayer.Name.ToLower().Contains("Bahrain_Image_2017".ToLower()) ||
                                    pLayer.Name.ToLower().Contains("Bahrain_Imagery_2015".ToLower()))
                                {
                                    map.DeleteLayer(pLayer);
                                    layerToBeAdded = GetLayerFromDisk(satellitePath);
                                }
                                else if (pLayer.Name.ToLower().Contains("DEM2012".ToLower()) ||
                                         pLayer.Name.ToLower().Contains("DEM_2012".ToLower()))
                                {
                                    map.DeleteLayer(pLayer);
                                    layerToBeAdded = GetLayerFromDisk(demPath);
                                }
                                else if (pLayer.Name.ToLower().Contains("Hawar".ToLower()))
                                {
                                    map.DeleteLayer(pLayer);
                                }

                                if (layerToBeAdded != null)
                                {
                                    map.AddLayer(layerToBeAdded);
                                    map.MoveLayer(layerToBeAdded, map.LayerCount);
                                }


                            }

                        }

                        pMapDoc.SaveAs(Path.Combine(Path.GetDirectoryName(mxdFile), Path.GetFileNameWithoutExtension(mxdFile) + "_V1.mxd"), true, true);
                        k= k+interval;

                    }
                    catch (Exception ex)
                    {
                        WriteLog("Failed " + msg, logFile);
                        Console.WriteLine("Failed " + msg);
                        throw ex;
                    }

                }

                //Console.WriteLine("Press any key to exit ......");
                //Console.ReadLine();

            }
            catch (Exception ex)
            {

                WriteLog("Failed " + msg, logFile);
                Console.WriteLine("Failed " + msg);
                throw ex;

            }
            finally
            {



            }

        }

        static ILayer GetLayerFromDisk(string layerPathFile)
        {
            ILayer returnLayer = null;
            IGxLayer gxLayer = new GxLayerClass();
            IGxFile gxFile = (IGxFile)gxLayer;
            gxFile.Path = layerPathFile;

            // Test if we have a valid layer and add it to the map
            if (!(gxLayer.Layer == null))
            {

                returnLayer = gxLayer.Layer;

            }

            return returnLayer;

        }

        static void WriteLog(string strLog, string logFilePath)
        {

            StreamWriter log=null;
            System.IO.FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;

            try
            {
                
                logFileInfo = new FileInfo(logFilePath);
                logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);

                if (!logDirInfo.Exists) logDirInfo.Create();
                if (!logFileInfo.Exists)
                {
                    fileStream = logFileInfo.Create();
                }
                else
                {
                    fileStream = new System.IO.FileStream(logFilePath, FileMode.Append);
                }

                log = new StreamWriter(fileStream);

                log.WriteLine(strLog);
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error", ex.ToString());
                throw ex;
            }

            finally
            {
                log?.Close();
            }
            
        }

        static void deleteExistingProcessesFiles(string mxdSourcePath)
        {

            string[]  mxdFiles = System.IO.Directory.GetFiles(mxdSourcePath, "*.mxd", SearchOption.AllDirectories);
            string detinationFile;           

            foreach (string mxdFile in mxdFiles)
            {
                detinationFile = Path.Combine(System.IO.Path.GetDirectoryName(mxdFile), System.IO.Path.GetFileNameWithoutExtension(mxdFile) + "_new.mxd");

                if (File.Exists(detinationFile))
                {
                    File.Delete(detinationFile);
                }

            }

           }
                
    }
}
