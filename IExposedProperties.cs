using System;
using System.Runtime.InteropServices;

namespace ComControlTemplate
{
    [ComVisible(true)]
    [Guid("10675027-46E8-496E-B611-9820D394EF12")]
    public interface IExposedProperties
    {
        [DispId(1)]
        string CustomText { get; set; }
    }
}
