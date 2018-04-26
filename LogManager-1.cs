using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Diagnostics;

    public static class LogManager
    {
        //public static string PathToCreateLog = "C:\\InsertTrpLog\\";
        public static string PathToCreateLog = System.Configuration.ConfigurationSettings.AppSettings["logpath"].ToString();
		public static string MessageInfo = "";
		public static void CreateLogFile(string Message, long MessageId)
        {
            FileInfo fi;
            FileStream fstr;
            //Check whether dir exists or not if not then create  it
            #region  Try to create the directory. if not extits
            if (!System.IO.Directory.Exists(PathToCreateLog))
            {
                System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(PathToCreateLog);
            }
            #endregion


            string strCurrentDate = DateTime.Now.ToString("yyyyMMdd");
            string strFileName = strCurrentDate + "_Log.txt";

            //Create New File if not exists
            if (!System.IO.File.Exists(PathToCreateLog + strFileName))
            {
                fi = new FileInfo(PathToCreateLog + strFileName);
                fstr = fi.Create();
                fstr.Close();
            }

            //Write the text to the file
            StackFrame callStack = new StackFrame(1, true);
            StreamWriter strWriteText = new StreamWriter(PathToCreateLog + strFileName, true);
			
		switch (MessageId) 
		{
		case 0:
			MessageInfo = "FATAL"; 
			break ;
		case 1: 
			MessageInfo = "ERROR";
			break;
		case 2:
			MessageInfo = "WARNINNG";
		default:
			MessageInfo = "INFO";
		}
           

		try
            {
                
                
                try
	                {
	                    
	                    strWriteText.WriteLine(DateTime.Now.ToString() + MessageInfo + ", File: " + callStack.GetFileName() + ", Line: " + callStack.GetFileLineNumber() + ", Message: " + Message);
	                }
                catch (Exception ex)
	                {
						strWriteText.WriteLine(DateTime.Now.ToString() + MessageInfo +  ", File: " + callStack.GetFileName() + ", Line: " + callStack.GetFileLineNumber() + ", Message: " + ex.Message);
	                }
                finally
	                {
	                    strWriteText.Dispose();
	                    strWriteText.Close();
	                }

            }
            catch (Exception ex1)
	            {
					strWriteText.WriteLine(DateTime.Now.ToString() + MessageInfo +  ", File: " + callStack.GetFileName() + ", Line: " + callStack.GetFileLineNumber() + ", Message: " + ex1.Message);  
	            }
            finally
	            {

	            }
        }

        public static void DeleteLogFile()
        {
            DirectoryInfo source = new DirectoryInfo(PathToCreateLog);

            // Get info of each file into the directory
            foreach (FileInfo fi in source.GetFiles())
            {
                var creationTime = fi.CreationTime;

                if (creationTime < (DateTime.Now - new TimeSpan(7, 0, 0, 0)))
                {
                    fi.Delete();
                }
            }

        }
      
    }
