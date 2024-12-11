using Renga;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RengaFacade
{
    public class RObject
    {
        private readonly RengaFacade mFacade;
        private readonly IModelObjectCollection mCollection;
        public Guid Id { get; private set; }
        public string Name { get => mCollection.GetByUniqueId(Id).Name; }
        public (Guid Id, string Name) ObjectType
        { 
            get {
                var objTypeId = mCollection.GetByUniqueId(Id).ObjectType;
                return mFacade.ObjectTypes.First(x => x.Id == objTypeId);
            }
        }
        public IEnumerable<RProperty> Properties
        {
            get
            {
                var propIdsCollection = mCollection.GetByUniqueId(Id).GetProperties().GetIds();
                var propIds = new List<Guid>(propIdsCollection.Count);
                for (var i = 0; i < propIdsCollection.Count; i++)
                {
                    propIds.Add(propIdsCollection.Get(i));
                }
                var result = new List<RProperty>(propIds.Count);
                foreach(var propId in propIds)
                {
                    result.Add(new RProperty(mFacade, propId, Id));
                }
                return result;
            }
        }

        internal RObject(RengaFacade facade, Guid id)
        {
            mFacade = facade;
            mCollection = facade.Project.Model.GetObjects();
            Id = id;
        }
    }
}
