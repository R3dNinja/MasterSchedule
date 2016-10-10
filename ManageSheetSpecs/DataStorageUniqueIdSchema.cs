using System;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ManageMasterSchedule
{
  static class DataStorageUniqueIdSchema
  {
    static readonly Guid schemaGuid = new Guid(
      "{D15C65AE-11F3-4560-B66F-163B92880C18}");

    public static Schema GetSchema()
    {
      Schema schema = Schema.Lookup( schemaGuid );

      if( schema != null )
        return schema;

      SchemaBuilder schemaBuilder = new SchemaBuilder( 
        schemaGuid );

      schemaBuilder.SetSchemaName( 
        "DataStorageUniqueId" );

      schemaBuilder.AddSimpleField( 
        "Id", typeof( Guid ) );

      return schemaBuilder.Finish();
    }
  }
}
