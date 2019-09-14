﻿using System;
using System.Runtime.Serialization;
using UnderRailLib.AssemblyResolver;
using UnderRailLib.Models.Interfaces;

namespace UnderRailLib.Models
{
    [EncodedTypeName("XPBLPIRRTCC")]
    [Serializable]
    public sealed class XPBLPIRRTCC : iLCPP, ISerializable
    {
        private XPBLPIRRTCC(SerializationInfo info, StreamingContext ctx)
        {
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
        }
    }
}