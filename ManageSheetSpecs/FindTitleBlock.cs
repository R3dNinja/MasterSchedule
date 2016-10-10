using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit;
using Autodesk.Revit.DB;

namespace ManageMasterSchedule
{
    class FindTitleBlock
    {
        private IList<Element> m_alltitleblocks = new List<Element>();
        private IList<Element> ElementsOnSheet = new List<Element>();
        private FamilySymbol MyTitleBlock;

        public void GetAllTitleblocks(Document doc)
        {
            //get all titleblocks
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector.OfClass(typeof(FamilySymbol));

            m_alltitleblocks = collector.ToElements();
        }

        public ElementId GetViewSheet(Document doc)
        {
            FilteredElementCollector col = new FilteredElementCollector(doc);
            col.OfCategory(BuiltInCategory.OST_Sheets);
            col.OfClass(typeof(ViewSheet));
            ElementId vsID = null;

            foreach (Element v in col)
            {

                ViewSheet vs = v as ViewSheet;
                var name = vs.Name;
                //var tbb = vs.
                if (name == "MASTER SCHEDULE")
                {
                    if (vs.SheetNumber.ToString().Contains("A0.70"))
                    {
                        vsID = vs.Id;
                        return vsID;
                    }
                    //Dialog.Show("SheetNumber", vs.SheetNumber.ToString());
                }
            }
            return vsID;
        }

        private void GetTitleSheet(IList<Element> ElementsOnSheet)
        {
            foreach (Element el in ElementsOnSheet)
            {
                foreach (FamilySymbol Fs in m_alltitleblocks)
                {
                    if (el.GetTypeId().IntegerValue == Fs.Id.IntegerValue)
                    {
                        MyTitleBlock = Fs;
                    }
                }
            }
        }
    }
}
