using System;
using System.Collections.Generic;

namespace YoutubeIPv6BlockForVRChat
{
    public class NetblocksConfig
    {
        public DateTime LastUpdated { get; set; }
        public List<string> GoogleIPv6Netblocks { get; set; } = new List<string>();
    }
}