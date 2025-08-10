//
// --- FILE: ExtensibleStorageManager.cs ---
//
// Description:
// A utility class for managing Revit extensible storage operations.
// Provides centralized methods for saving, recalling, and manipulating
// data stored in Revit's extensible storage system.
//
// Change Log:
// - August 16, 2025: Initial creation. Consolidated extensible storage functionality
//                    from PC_Extensible.cs, RT_TrayOccupancy.cs, and RTS_RevitUtils.cs.
//

#region Namespaces
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
#endregion

namespace RTS.Utilities
{
    /// <summary>
    /// Provides methods to work with Revit's extensible storage system
    /// </summary>
    public static class ExtensibleStorageManager
    {
        // Standard vendor ID used across all schemas
        private const string VendorId = "ReTick_Solutions";

        #region Storage Schema GUIDs and Names

        // PowerCAD Primary Data
        public static readonly Guid PrimarySchemaGuid = new Guid("A3F6D2AF-6702-4B9C-9DEF-336EBAA87336");
        public const string PrimarySchemaName = "PC_ExtensibleDataSchema";
        public const string PrimaryFieldName = "PC_DataJson";
        public const string PrimaryDataStorageElementName = "PC_Extensible_PC_Data_Storage";

        // PowerCAD Consultant Data
        public static readonly Guid ConsultantSchemaGuid = new Guid("B5E7F8C0-1234-5678-9ABC-DEF012345678");
        public const string ConsultantSchemaName = "PC_ExtensibleConsultantDataSchema";
        public const string ConsultantFieldName = "PC_ConsultantDataJson";
        public const string ConsultantDataStorageElementName = "PC_Extensible_Consultant_Data_Storage";

        // Model Generated Data
        public static readonly Guid ModelGeneratedSchemaGuid = new Guid("C7D8E9F0-1234-5678-9ABC-DEF012345678");
        public const string ModelGeneratedSchemaName = "PC_ExtensibleModelGeneratedDataSchema";
        public const string ModelGeneratedFieldName = "PC_ModelGeneratedDataJson";
        public const string ModelGeneratedDataStorageElementName = "PC_Extensible_Model_Generated_Data_Storage";

        #endregion

        #region Core Extensible Storage Methods

        /// <summary>
        /// Gets or creates the Schema for storing generic data
        /// </summary>
        public static Schema GetOrCreateSchema(Guid schemaGuid, string schemaName, string fieldName)
        {
            Schema schema = Schema.Lookup(schemaGuid);

            if (schema == null)
            {
                SchemaBuilder schemaBuilder = new SchemaBuilder(schemaGuid);
                schemaBuilder.SetSchemaName(schemaName);
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Vendor);
                schemaBuilder.SetVendorId(VendorId);
                schemaBuilder.AddSimpleField(fieldName, typeof(string));
                schema = schemaBuilder.Finish();
            }
            return schema;
        }

        /// <summary>
        /// Gets the existing DataStorage element, or creates a new one if it doesn't exist
        /// </summary>
        public static DataStorage GetOrCreateDataStorage(Document doc, Guid schemaGuid, string dataStorageElementName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            // Check for both name and schema entity
            Schema schema = Schema.Lookup(schemaGuid);
            if (schema != null)
            {
                foreach (DataStorage ds in collector)
                {
                    if (ds.Name == dataStorageElementName && ds.GetEntity(schema) != null)
                    {
                        dataStorage = ds;
                        break;
                    }
                }
            }

            // Fallback to finding by name only, in case an empty DataStorage element already exists
            if (dataStorage == null)
            {
                dataStorage = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            }

            // If it's still null, create a new one
            if (dataStorage == null)
            {
                dataStorage = DataStorage.Create(doc);
                dataStorage.Name = dataStorageElementName;
            }

            return dataStorage;
        }

        /// <summary>
        /// Saves a list of objects to extensible storage
        /// </summary>
        public static void SaveDataToExtensibleStorage<T>(
            Document doc,
            List<T> dataList,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName)
        {
            Schema schema = GetOrCreateSchema(schemaGuid, schemaName, fieldName);
            DataStorage dataStorage = GetOrCreateDataStorage(doc, schemaGuid, dataStorageElementName);

            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(dataList, options);

            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(fieldName), jsonString);

            dataStorage.SetEntity(entity);
        }

        /// <summary>
        /// Recalls a list of objects from extensible storage
        /// </summary>
        public static List<T> RecallDataFromExtensibleStorage<T>(
            Document doc,
            Guid schemaGuid,
            string schemaName,
            string fieldName,
            string dataStorageElementName) where T : new()
        {
            Schema schema = Schema.Lookup(schemaGuid);

            if (schema == null)
            {
                return new List<T>();
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorage = null;
            foreach (DataStorage ds in collector)
            {
                if (ds.Name == dataStorageElementName && ds.GetEntity(schema) != null)
                {
                    dataStorage = ds;
                    break;
                }
            }

            if (dataStorage == null)
            {
                return new List<T>();
            }

            Entity entity = dataStorage.GetEntity(schema);

            if (!entity.IsValid())
            {
                return new List<T>();
            }

            string jsonString = entity.Get<string>(schema.GetField(fieldName));

            if (string.IsNullOrEmpty(jsonString))
            {
                return new List<T>();
            }

            try
            {
                return JsonSerializer.Deserialize<List<T>>(jsonString) ?? new List<T>();
            }
            catch
            {
                return new List<T>();
            }
        }

        /// <summary>
        /// Clears data from extensible storage by deleting the DataStorage element
        /// </summary>
        public static bool ClearData(Document doc, Guid schemaGuid, string dataStorageElementName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage));

            DataStorage dataStorageToDelete = null;
            Schema schema = Schema.Lookup(schemaGuid);

            if (schema != null)
            {
                // Find the specific DataStorage element by name that has an entity of our schema
                foreach (DataStorage ds in collector)
                {
                    if (ds.Name == dataStorageElementName && ds.GetEntity(schema) != null)
                    {
                        dataStorageToDelete = ds;
                        break;
                    }
                }
            }

            // Fallback to finding by name only, in case an empty DataStorage element already exists
            if (dataStorageToDelete == null)
            {
                dataStorageToDelete = collector.Cast<DataStorage>().FirstOrDefault(ds => ds.Name == dataStorageElementName);
            }

            // If we found a DataStorage element, delete it
            if (dataStorageToDelete != null)
            {
                doc.Delete(dataStorageToDelete.Id);
                return true;
            }

            return false; // Nothing to delete
        }

        #endregion

        #region Cable Data Specific Methods

        /// <summary>
        /// Saves primary cable data to extensible storage
        /// </summary>
        public static void SavePrimaryData<T>(Document doc, List<T> data)
        {
            SaveDataToExtensibleStorage(doc, data, PrimarySchemaGuid, PrimarySchemaName, PrimaryFieldName, PrimaryDataStorageElementName);
        }

        /// <summary>
        /// Recalls primary cable data from extensible storage
        /// </summary>
        public static List<T> RecallPrimaryData<T>(Document doc) where T : new()
        {
            return RecallDataFromExtensibleStorage<T>(doc, PrimarySchemaGuid, PrimarySchemaName, PrimaryFieldName, PrimaryDataStorageElementName);
        }

        /// <summary>
        /// Saves consultant cable data to extensible storage
        /// </summary>
        public static void SaveConsultantData<T>(Document doc, List<T> data)
        {
            SaveDataToExtensibleStorage(doc, data, ConsultantSchemaGuid, ConsultantSchemaName, ConsultantFieldName, ConsultantDataStorageElementName);
        }

        /// <summary>
        /// Recalls consultant cable data from extensible storage
        /// </summary>
        public static List<T> RecallConsultantData<T>(Document doc) where T : new()
        {
            return RecallDataFromExtensibleStorage<T>(doc, ConsultantSchemaGuid, ConsultantSchemaName, ConsultantFieldName, ConsultantDataStorageElementName);
        }

        /// <summary>
        /// Saves model generated data to extensible storage
        /// </summary>
        public static void SaveModelGeneratedData<T>(Document doc, List<T> data)
        {
            SaveDataToExtensibleStorage(doc, data, ModelGeneratedSchemaGuid, ModelGeneratedSchemaName, ModelGeneratedFieldName, ModelGeneratedDataStorageElementName);
        }

        /// <summary>
        /// Recalls model generated data from extensible storage
        /// </summary>
        public static List<T> RecallModelGeneratedData<T>(Document doc) where T : new()
        {
            return RecallDataFromExtensibleStorage<T>(doc, ModelGeneratedSchemaGuid, ModelGeneratedSchemaName, ModelGeneratedFieldName, ModelGeneratedDataStorageElementName);
        }

        #endregion
    }
}