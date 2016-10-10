using System;
using System.Collections.Generic;
using System.IO;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

namespace ManageMasterSchedule
{
    public class RequestHandler : IExternalEventHandler
    {
        private Request m_request = new Request();

        public Request Request
        {
            get { return m_request; }
        }

        public String GetName()
        {
            return "Manage Sheet Spec Event";
        }

        public void Execute(UIApplication uiapp)
        {
            string path = Command.thisCommand.dialog.GetDocPath();
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {
                            return;
                        }
                    case RequestId.replaceImages:
                        {
                            reloadImages(uiapp, path);
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            finally
            {
            }
            return;
        }

        public void reloadImages(UIApplication uiapp, string path)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            string imagePath = Path.GetDirectoryName(path);
            int pageCount = Command.thisCommand.dialog.getPageCount();
            Command.thisCommand.dialog.SetupProgress(pageCount, "Task: Replacing Sheet Spec Images");


            FilteredElementCollector col = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RasterImages);

            int counter = 1;
            foreach (Element e in col)
            {
                var tempTest = e.GetType();
                if (tempTest.FullName == "Autodesk.Revit.DB.ImageType")
                {
                    string imageName = e.Name;
                    int index = imageName.LastIndexOf(" ");
                    if (index > 5)
                        imageName = imageName.Substring(0, index);
                    string fullImagePath = imagePath + @"\Sheet Specs (Images)\" + imageName;
                    if (imageName.StartsWith("Sheet Specs", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (File.Exists(fullImagePath))
                        {
                            using (Transaction tx = new Transaction(doc))
                            {
                                tx.Start("Replace Image");
                                ImageType image = e as ImageType;
                                image.ReloadFrom(fullImagePath);
                                tx.Commit();
                            }
                            ++counter;
                            if (counter < pageCount)
                            {
                                Command.thisCommand.dialog.IncrementProgress();
                            }
                        }
                    }
                }
            }

        }
    }
}
