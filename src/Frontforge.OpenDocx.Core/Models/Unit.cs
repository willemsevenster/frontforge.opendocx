﻿using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Frontforge.OpenDocx.Core.Models
{
    public class Unit
    {
        public static implicit operator Unit(double value)
        {
            return new Unit {Value = value};
        }

        public static implicit operator string(Unit value)
        {
            return value?.ToString();
        }

        public static implicit operator StringValue(Unit value)
        {
            return new StringValue(value.ToString());
        }

        public override string ToString()
        {
            var result = Value;
            var unit = Enum.GetName(typeof(UnitType), Type);

            switch (Type)
            {
                case UnitType.pt:
                    result = Value * 20d;
                    unit = "pt";
                    break;
                case UnitType.pct:
                    result = Value * 50d;
                    unit = "pct";
                    break;
                case UnitType.inch:
                    result = Value * 1440d;
                    unit = "in";
                    break;
                case UnitType.mm:
                    result = Value * 1440d / 25.4d;
                    unit = "mm";
                    break;
                case UnitType.cm:
                    result = Value * 1440d / 2.54d;
                    unit = "cm";
                    break;
            }

            return $"{result}{unit}";
        }
        
        internal double ToDxa()
        {
            var result = Value;

            switch (Type)
            {
                case UnitType.pt:
                    result = Value * 20d;
                    break;
                case UnitType.pct:
                    result = Value * 50d;
                    break;
                case UnitType.inch:
                    result = Value * 1440d;
                    break;
                case UnitType.mm:
                    result = Value * 1440d / 25.4d;
                    break;
                case UnitType.cm:
                    result = Value * 1440d / 2.54d;
                    break;
            }

            return result;
        }

        private Unit()
        {
        }

        public Unit(double value, UnitType type = UnitType.pt)
        {
            Value = value;
            Type = type;
        }

        public double Value { get; set; }
        public UnitType Type { get; set; } = UnitType.pt;

        internal T ToTableWidthType<T>()
            where T : TableWidthType, new()
        {
            var result = new T
            {
                Width = ToDxa().ToString(CultureInfo.InvariantCulture),
                Type = Type == UnitType.pct ? TableWidthUnitValues.Pct : TableWidthUnitValues.Dxa
            };

            return result;
        }

        internal T ToTableMargin<T>()
            where T : TableWidthDxaNilType, new()
        {
            var result = new T
            {
                Width = Convert.ToInt16(ToDxa()),
                Type = TableWidthValues.Dxa
            };

            return result;
        }
    }
}