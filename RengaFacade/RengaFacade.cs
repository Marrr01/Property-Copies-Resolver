using Renga;
using System.Collections.Generic;
using System;

namespace RengaFacade
{
    // todo: вынести в библиотеку
    public class RengaFacade
    {
        public readonly IProject Project;
        public RengaFacade(IProject project) => Project = project;

        #region Объекты

        private IEnumerable<(Guid Id, string Name)> mObjectTypes;
        public IEnumerable<(Guid Id, string Name)> ObjectTypes
        {
            get
            {
                if (mObjectTypes == null)
                {
                    var result = new List<(Guid Id, string Name)>();
                    var props = typeof(ObjectTypes).GetProperties();
                    foreach (var prop in props)
                    {
                        result.Add(((Guid)prop.GetValue(null), prop.Name));
                    }
                    mObjectTypes = result;
                }
                return mObjectTypes;
            }
        }

        public IEnumerable<RObject> GetObjects()
        {
            var objCollection = Project.Model.GetObjects();
            var result = new List<RObject>(objCollection.Count);
            for (var i = 0; i < objCollection.Count; i++)
            {
                var obj = objCollection.GetByIndex(i);
                result.Add(new RObject(this, obj.UniqueId));
            }
            return result;
        }

        #endregion

        #region Свойства

        public List<RProperty> GetProperties()
        {
            var props = new List<RProperty>();
            var propManager = Project.PropertyManager;
            for (var i = 0; i < propManager.PropertyCount; i++)
            {
                var propId = propManager.GetPropertyId(i);
                props.Add(new RProperty(this, propId));
            }
            return props;
        }

        public RProperty CreateProperty(string name, PropertyType type)
        {
            var propManager = Project.PropertyManager;
            var propId = Guid.NewGuid();
            propManager.RegisterPropertyS(propId.ToString(), name, type);
            return new RProperty(this, propId);
        }

        public void DeleteProperty(Guid id) => Project.PropertyManager.UnregisterProperty(id);

        public string GetPropertyValue(Guid propertyId, Guid objectId)
        {
            var obj = Project.Model.GetObjects().GetByUniqueId(objectId);
            var prop = obj.GetProperties().Get(propertyId);

            switch (prop.Type)
            {
                case PropertyType.PropertyType_Angle:
                    return prop.GetAngleValue(AngleUnit.AngleUnit_Degrees).ToString();

                case PropertyType.PropertyType_Area:
                    return prop.GetAreaValue(AreaUnit.AreaUnit_Meters2).ToString();

                case PropertyType.PropertyType_Boolean:
                    return prop.GetBooleanValue().ToString();

                case PropertyType.PropertyType_Double:
                    return prop.GetDoubleValue().ToString();

                case PropertyType.PropertyType_Enumeration:
                    return prop.GetEnumerationValue();

                case PropertyType.PropertyType_Integer:
                    return prop.GetIntegerValue().ToString();

                case PropertyType.PropertyType_Length:
                    return prop.GetLengthValue(LengthUnit.LengthUnit_Millimeters).ToString();

                case PropertyType.PropertyType_Logical:
                    return prop.GetLogicalValue().ToString();

                case PropertyType.PropertyType_Mass:
                    return prop.GetMassValue(MassUnit.MassUnit_Kilograms).ToString();

                case PropertyType.PropertyType_String:
                    return prop.GetStringValue();

                case PropertyType.PropertyType_Undefined:
                    throw new NotImplementedException();

                case PropertyType.PropertyType_Volume:
                    return prop.GetVolumeValue(VolumeUnit.VolumeUnit_Meters3).ToString();

                default:
                    throw new NotImplementedException();
            }
        }

        public void SetPropertyValue(Guid propertyId, Guid objectId, string value)
        {
            var obj = Project.Model.GetObjects().GetByUniqueId(objectId);
            var prop = obj.GetProperties().Get(propertyId);

            switch (prop.Type)
            {
                case PropertyType.PropertyType_Angle:
                    if (double.TryParse(value, out double angle))
                    {
                        prop.SetAngleValue(angle, AngleUnit.AngleUnit_Degrees);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Area:
                    if (double.TryParse(value, out double area))
                    {
                        prop.SetAreaValue(area, AreaUnit.AreaUnit_Meters2);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Boolean:
                    if (value.ToLower() == "true")
                    {
                        prop.SetBooleanValue(true);
                        break;
                    }
                    if (value.ToLower() == "false")
                    {
                        prop.SetBooleanValue(false);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Double:
                    if (double.TryParse(value, out double d))
                    {
                        prop.SetDoubleValue(d);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Enumeration:
                    prop.SetEnumerationValue(value);
                    break;

                case PropertyType.PropertyType_Integer:
                    if (int.TryParse(value, out int i))
                    {
                        prop.SetIntegerValue(i);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Length:
                    if (double.TryParse(value, out double lenght))
                    {
                        prop.SetLengthValue(lenght, LengthUnit.LengthUnit_Millimeters);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Logical:
                    var low = value.ToLower();
                    if (low == "true" || low == "logical_true")
                    {
                        prop.SetLogicalValue(Logical.Logical_True);
                        break;
                    }
                    if (low == "false" || low == "logical_false")
                    {
                        prop.SetLogicalValue(Logical.Logical_False);
                        break;
                    }
                    if (low == "indeterminate" || low == "logical_indeterminate")
                    {
                        prop.SetLogicalValue(Logical.Logical_Indeterminate);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_Mass:
                    if (double.TryParse(value, out double mass))
                    {
                        prop.SetMassValue(mass, MassUnit.MassUnit_Kilograms);
                        break;
                    }
                    goto default;

                case PropertyType.PropertyType_String:
                    prop.SetStringValue(value);
                    break;

                case PropertyType.PropertyType_Undefined:
                    throw new NotImplementedException();

                case PropertyType.PropertyType_Volume:
                    if (double.TryParse(value, out double volume))
                    {
                        prop.SetVolumeValue(volume, VolumeUnit.VolumeUnit_Meters3);
                        break;
                    }
                    goto default;

                default:
                    throw new Exception($"Не удалось преобразовать значение \"{value}\" из строки для атрибута \"{prop.Name}\"");
            }
        }

        #endregion
    }
}
