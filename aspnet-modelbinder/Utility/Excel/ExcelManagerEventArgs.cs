using System;

namespace aspnet_modelbinder.Utility.Excel
{
    public class ExcelManagerEventArgs : EventArgs
    {
        public ExcelSheet Sheet { get; set; }
        public ExcelManagerEventArgs()
        {
        }
        public ExcelManagerEventArgs(ExcelSheet sheet)
        {
            Sheet = sheet;
        }
    }
}