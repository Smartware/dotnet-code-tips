using System;

namespace aspnet_modelbinder.Utility.Excel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ExcelManagerCellPropertyAttribute : Attribute
    {
        protected internal string Name { get; set; }

        public ExcelManagerCellPropertyAttribute(string Name)
        {
            this.Name = Name;
        }

        public ExcelManagerCellPropertyAttribute()
        {
        }
    }
}