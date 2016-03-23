using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class AppSetting
    {
        private string key;
        private string keyvalue;

        public string Key
        {
            get { return key; }
            set { key = value; }
        }

        public string KeyValue
        {
            get { return keyvalue; }
            set { keyvalue = value; }
        }
    }
}
