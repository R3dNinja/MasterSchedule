using System;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ManageMasterSchedule
{
  public class ReadWriteSettingsCommand
  {
    public string ReadFilePath(Document activeDoc)
    {
        string filePath = null;
        Document doc = activeDoc;

        MyProjectSettingStorage settingStorage = new MyProjectSettingStorage();
        MyProjectSettings readSettings = settingStorage.ReadSettings( doc );
    
        if (readSettings == null)
        {
            filePath = "";
            return filePath;
        }
        else
        {
            filePath = readSettings.Parameter2;
            return filePath;
        }
    }

    public void WriteFilePath(Document activeDoc, string filePath)
    {
        Document doc = activeDoc;

        MyProjectSettings settings = new MyProjectSettings();
        settings.Parameter1 = new Random().Next();
        settings.Parameter2 = filePath;

        MyProjectSettingStorage settingStorage = new MyProjectSettingStorage();

        using(Transaction t = new Transaction(doc, "Write settings"))
        {
          t.Start();
          settingStorage.WriteSettings( doc, settings );
          t.Commit();
        }
    }
  }
}
