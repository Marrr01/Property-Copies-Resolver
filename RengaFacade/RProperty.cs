using Renga;
using System;
using System.Collections.Generic;

namespace RengaFacade
{
    // todo: CSVExportFlag, Expression
    public class RProperty
    {
        private readonly RengaFacade mFacade;
        private readonly IPropertyManager mPropManager;
        public Guid Id { get; private set; }
        public string Name { get => mPropManager.GetPropertyName(Id); }
        public PropertyType PropertyType { get => mPropManager.GetPropertyType(Id); }
        public IEnumerable<(Guid Id, string Name)> ObjectTypes
        {
            get {
                var result = new List<(Guid, string)>();
                foreach (var type in mFacade.ObjectTypes)
                {
                    if (mPropManager.IsPropertyAssignedToType(Id, type.Id))
                    {
                        result.Add(type);
                    }
                }
                return result;
            }
        }
        public bool HasDefaultValue
        {
            get
            {
                var value = Value;
                switch (PropertyType)
                {
                    case PropertyType.PropertyType_Boolean:
                        return value == "False" ? true : false;

                    case PropertyType.PropertyType_Logical:
                        return value == "Logical_Indeterminate" ? true : false;

                    case PropertyType.PropertyType_Enumeration:
                        var defaultValue = (string)mFacade.Project.PropertyManager.GetPropertyDescription2(Id).GetEnumerationItems().GetValue(0);
                        return value == defaultValue ? true : false;

                    default:
                        if (string.IsNullOrEmpty(value))
                        {
                            return true;
                        }
                        if (double.TryParse(value, out double valueAsDouble))
                        {
                            if (valueAsDouble == 0) return true;
                        }
                        return false;

                }
            }
        }
        public string Value
        {
            get
            {
                if (IsDefinition) throw new Exception("У определения свойства не может быть значения");
                return mFacade.GetPropertyValue(Id, (Guid)mObjectId);
            }
            set
            {
                if (IsDefinition) throw new Exception("Нельзя присвоить значение определению свойства");
                mFacade.SetPropertyValue(Id, (Guid)mObjectId, value);
            }
        }
        
        private readonly Guid? mObjectId;

        /// <summary> Это определение свойства или свойство объекта? </summary>
        public bool IsDefinition => mObjectId is null;

        internal RProperty(RengaFacade facade, Guid id, Guid? objectId = null)
        {
            mFacade = facade;
            mPropManager = facade.Project.PropertyManager;
            Id = id;
            mObjectId = objectId;
        }
        public void AddToObjectType(Guid ObjectTypeId) => mPropManager.AssignPropertyToType(Id, ObjectTypeId);
        public void AddToObjectType(IEnumerable<Guid> ObjectTypeIds)
        {
            foreach (var objectTypeId in ObjectTypeIds)
            {
                AddToObjectType(objectTypeId);
            }
        }
        public void DeleteFromObjectType(Guid ObjectTypeId) => mPropManager.UnassignPropertyFromType(Id, ObjectTypeId);
        public void DeleteFromObjectType(IEnumerable<Guid> ObjectTypeIds)
        {
            foreach (var objectTypeId in ObjectTypeIds)
            {
                DeleteFromObjectType(objectTypeId);
            }
        }
    }
}
