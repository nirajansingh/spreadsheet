using System;
using DocumentFormat.OpenXml;

namespace Excel.SpreadSheet.OpenXml;

public enum HorizontalAlignment
{
    [EnumString("general")]
    General,
    [EnumString("left")]
    Left,
    [EnumString("center")]
    Center,
    [EnumString("right")]
    Right,
    [EnumString("fill")]
    Fill,
    [EnumString("justify")]
    Justify,
    [EnumString("centerContinuous")]
    CenterContinuous,
    [EnumString("distributed")]
    Distributed
}

