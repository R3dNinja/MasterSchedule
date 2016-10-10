using System;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace ManageMasterSchedule
{
  public static class MyProjectSettingsSchema
  {
    readonly static Guid schemaGuid = new Guid(
      "{763A6FA2-A28F-4C85-A323-ED3150408552}");

    public static Schema GetSchema()
    {
      Schema schema = Schema.Lookup( schemaGuid );

      if( schema != null ) return schema;

      SchemaBuilder schemaBuilder =
          new SchemaBuilder( schemaGuid );

      schemaBuilder.SetSchemaName( 
        "MyProjectSettings" );

      schemaBuilder.AddSimpleField( 
        "Parameter1", typeof( int ) );

      schemaBuilder.AddSimpleField( 
        "Parameter2", typeof( string ) );

      return schemaBuilder.Finish();
    }
  }
}
