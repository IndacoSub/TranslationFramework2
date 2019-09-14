﻿using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using UnderRailLib.AssemblyResolver;
using UnderRailLib.Models.EventArgs;
using UnderRailLib.Models.Interfaces;

namespace UnderRailLib.Models
{
    [EncodedTypeName("IC")]
    [Serializable]
    public sealed class ItemContainer : iITC, ISerializable
    {
        public ItemContainer()
        {

        }

        public ItemContainer(SerializationInfo info, StreamingContext ctx)
        {
            if (DataModelVersion.MinorVersion != 0)
            {
                SerializationHelper.ReadList("IC:I", ref _items, info);
                if (DataModelVersion.MajorVersion < 5)
                {
                    _itemAdded = (EventHandler<ItemEventArgs>) info.GetValue("IC:IA", typeof(EventHandler<ItemEventArgs>));
                    _itemRemoved = (EventHandler<ItemEventArgs>) info.GetValue("IC:IR", typeof(EventHandler<ItemEventArgs>));
                    return;
                }

                SerializationHelper.ReadEvent("IC:IA", ref _itemAdded, info);
                SerializationHelper.ReadEvent("IC:IR", ref _itemRemoved, info);
                if (DataModelVersion.MajorVersion >= 52)
                {
                    _ix = info.GetBoolean("IC:IX");
                    _ibc =
                        (info.GetValue("IC:IBC",
                            typeof(Dictionary<string, List<ItemInstance>>)) as Dictionary<string, List<ItemInstance>>);
                    _ibi =
                        (info.GetValue("IC:IBI", typeof(Dictionary<Guid, List<ItemInstance>>)) as Dictionary<Guid, List<ItemInstance>>);
                }
            }
            else
            {
                if (GetType() == typeof(ItemContainer))
                {
                    _items = (List<ItemInstance>) info.GetValue("_Items", typeof(List<ItemInstance>));
                    _itemAdded = (EventHandler<ItemEventArgs>) info.GetValue("ItemAdded", typeof(EventHandler<ItemEventArgs>));
                    _itemRemoved = (EventHandler<ItemEventArgs>) info.GetValue("ItemRemoved", typeof(EventHandler<ItemEventArgs>));
                    return;
                }

                _items = (List<ItemInstance>) info.GetValue("ItemContainer+_Items", typeof(List<ItemInstance>));
                _itemAdded = (EventHandler<ItemEventArgs>) info.GetValue("ItemContainer+ItemAdded", typeof(EventHandler<ItemEventArgs>));
                _itemRemoved = (EventHandler<ItemEventArgs>) info.GetValue("ItemContainer+ItemRemoved", typeof(EventHandler<ItemEventArgs>));
            }
        }

        public void GetObjectData(SerializationInfo info, StreamingContext ctx)
        {
            SerializationHelper.WriteList("IC:I", _items, info);
            SerializationHelper.WriteEvent("IC:IA", _itemAdded, info);
            SerializationHelper.WriteEvent("IC:IR", _itemRemoved, info);
            info.AddValue("IC:IX", _ix);
            info.AddValue("IC:IBC", _ibc);
            info.AddValue("IC:IBI", _ibi);
        }

        private List<ItemInstance> _items = new List<ItemInstance>();

        private Dictionary<string, List<ItemInstance>> _ibc;

        private Dictionary<Guid, List<ItemInstance>> _ibi;

        private bool _ix;

        private EventHandler<ItemEventArgs> _itemAdded;

        private EventHandler<ItemEventArgs> _itemRemoved;
    }
}