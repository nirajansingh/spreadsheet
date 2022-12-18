using System;
using DocumentFormat.OpenXml;

namespace Excel.SpreadSheet.OpenXml
{
	public enum VerticalAlignment
	{
        [EnumString("top")]
        Top,
        [EnumString("center")]
        Center,
        [EnumString("bottom")]
        Bottom,
        [EnumString("justify")]
        Justify,
        [EnumString("distributed")]
        Distributed
    }
}

