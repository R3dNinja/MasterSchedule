using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ManageMasterSchedule
{
    public class ReplaceImages
    {
        public void replaceImages(Document doc, string wordDocPath, int pageCount)
        {
            Command.thisCommand.dialog.setPageCount(pageCount);
            Command.thisCommand.dialog.SetupProgress(pageCount, "Task: Replacing Master Schedule Images");

            FilteredElementCollector col = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RasterImages);

            string imagePath = Path.GetDirectoryName(wordDocPath);
            imagePath = imagePath + @"\Master Schedule (Images)\";

            int imageNumber = 1;
            foreach (Element e in col)
            {
                string imageName = e.Name;
                int second = IndexOfSecond(imageName, @" ");
                if (second != -1)
                {
                    imageName = imageName.Substring(0, second - 1);
                }
                if (imageName.StartsWith("Sheet Specs", StringComparison.InvariantCultureIgnoreCase))
                {

                    string fullImagePath = imagePath + imageName;
                    if (File.Exists(fullImagePath))
                    {
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("Transaction Name");
                            ImageType image = e as ImageType;
                            image.ReloadFrom(fullImagePath);
                            tx.Commit();
                        }
                        Command.thisCommand.dialog.IncrementProgress();
                        imageNumber++;
                    }
                }
            }
        }

        private int IndexOfSecond(string theString, string toFind)
        {
            int first = theString.IndexOf(toFind);

            if (first == -1) return -1;

            // Find the "next" occurrence by starting just past the first
            return theString.IndexOf(toFind, first + 1);
        }
    }
}
