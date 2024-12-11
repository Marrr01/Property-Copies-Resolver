using Renga;
using RengaFacade;
using System;
using System.Collections.Generic;

namespace RengaTemplate
{
    internal class PropertyRule
    {
        public RProperty Property;

        public string PropertyName;
        public PropertyType PropertyType;
        public IEnumerable<(Guid Id, string Name)> PropertyObjectTypes;

        public string GetValueFromName;
        public IEnumerable<RProperty> GetValueFrom;
    }
}
