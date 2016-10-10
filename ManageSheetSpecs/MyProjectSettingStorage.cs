using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ManageMasterSchedule
{
  class MyProjectSettingStorage
  {
    readonly Guid settingDsId = new Guid(
      "{D6EA98F2-1F0C-4925-9B06-F6917479F25E}");

    public MyProjectSettings ReadSettings(Document doc )
    {
      //var settingDs = GetSettingsDataStorage(
      //  doc);

      //if (settingDs == null) return null;

      //settingDs.GetEntity(
      //  MyProjectSettingsSchema.GetSchema());

      var settingsEntity = GetSettingsEntity( doc );

      if( settingsEntity == null
        || !settingsEntity.IsValid() )
      {
        return null;
      }

      MyProjectSettings settings =
          new MyProjectSettings();

      settings.Parameter1 = settingsEntity.Get<int>(
        "Parameter1" );

      settings.Parameter2 = settingsEntity.Get<string>(
        "Parameter2" );

      return settings;
    }

    public void WriteSettings(
      Document doc,
      MyProjectSettings settings )
    {
      var settingDs = GetSettingsDataStorage(
        doc );

      if( settingDs == null )
      {
        settingDs = DataStorage.Create( doc );
      }

      Entity settingsEntity = new Entity(
        MyProjectSettingsSchema.GetSchema() );

      settingsEntity.Set( "Parameter1",
        settings.Parameter1 );

      settingsEntity.Set( "Parameter2",
        settings.Parameter2 );

      // Identify settings data storage

      Entity idEntity = new Entity(
        DataStorageUniqueIdSchema.GetSchema() );

      idEntity.Set( "Id", settingDsId );

      settingDs.SetEntity( idEntity );
      settingDs.SetEntity( settingsEntity );
    }

    private DataStorage GetSettingsDataStorage(
      Document doc )
    {
      // Retrieve all data storages from project

      FilteredElementCollector collector =
          new FilteredElementCollector( doc );

      var dataStorages =
          collector.OfClass( typeof( DataStorage ) );

      // Find setting data storage

      foreach( DataStorage dataStorage
        in dataStorages )
      {
        Entity settingIdEntity
          = dataStorage.GetEntity(
            DataStorageUniqueIdSchema.GetSchema() );

        if( !settingIdEntity.IsValid() ) continue;

        var id = settingIdEntity.Get<Guid>( "Id" );

        if( !id.Equals( settingDsId ) ) continue;

        return dataStorage;
      }
      return null;
    }

    private Entity GetSettingsEntity(
      Document doc )
    {
      FilteredElementCollector collector =
          new FilteredElementCollector( doc );

      var dataStorages =
          collector.OfClass( typeof( DataStorage ) );

      // Find setting data storage

      foreach( DataStorage dataStorage in dataStorages )
      {
        Entity settingEntity =
          dataStorage.GetEntity( MyProjectSettingsSchema.GetSchema() );

        // If a DataStorage contains 
        // setting entity, we found it

        if( !settingEntity.IsValid() ) continue;

        return settingEntity;
      }

      return null;
    }
  }
}
