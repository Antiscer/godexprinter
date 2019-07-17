using System;
using System.Runtime.InteropServices;

namespace AddIn
{
    // для установки дополнительных атрибутов методов и свойств, в данном случае название по русски
    [ComVisible(false)]
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    class Alias : Attribute
    {
        private string rusName;
        public string RusName => rusName;

        public Alias(string rusName)
        {
            this.rusName = rusName;
        }
    }
}
