using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.EnterpriseServices;
using System.Reflection;
using System.Collections.Generic;


public enum HRESULT : int
{
    S_FALSE = 0x0001,
    S_OK = 0x0000,
    E_FAIL = unchecked((int)0x80004005),
    E_INVALIDARG = unchecked((int)0x80070057),
    E_OUTOFMEMORY = unchecked((int)0x8007000E)
}

namespace AddIn
{
    [ComVisible(false), Guid("7DA558DD-4DF7-46B8-8C16-40ACAF29F593")]
    public class ComponentBase : IInitDone, ILanguageExtender
    {
        public IErrorLog pErrorInfo = null;
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            try
            {
                if (pErrorInfo == null)
                {
                    pErrorInfo = (IErrorLog)connection;
                }
            }
            catch(Exception e)
            {
                if (pErrorInfo != null)
                {
                    ShowError(this.GetType().Name + ": " + e.Message, 1);
                }
                else
                {
                    throw new COMException(e.Message, (int)HRESULT.S_FALSE);
                }
            }
        }
        public void Done()
        {
           
        }
        public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
        }
        public ComponentBase() { }
        private Hashtable methodNameToNumberLat;
        private Hashtable methodNameToNumberRus;
        private Hashtable methodNumberToNameLat;
        private Hashtable methodNumberToNameRus;
        private Hashtable numberToParamsCount;
        private Hashtable propertyNameToNumberLat;
        private Hashtable propertyNameToNumberRus;
        private Hashtable propertyNumberToNameLat;
        private Hashtable propertyNumberToNameRus;
        MethodInfo[] allMethodInfo;

        private List<MethodInfo> methodsInfo;
        private List<PropertyInfo> propertiesInfo;

        protected string errorMessage;
        protected int messagesLevel;
        protected bool showMеssages;
        protected bool ignoreWarnings;

        // инициализирующий метод, который передает перечисление методов и свойств в 1С при создании объекта компоненты в 1С
        // в нашем случае в перечисление попадают все методы и свойства отмеченные расширенными свойствами [Alias("Название")]
        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            try
            {
                methodNameToNumberLat = new Hashtable();
                methodNameToNumberRus = new Hashtable();
                methodNumberToNameLat = new Hashtable();
                methodNumberToNameRus = new Hashtable();
                numberToParamsCount = new Hashtable();

                propertyNameToNumberLat = new Hashtable();
                propertyNameToNumberRus = new Hashtable();
                propertyNumberToNameLat = new Hashtable();
                propertyNumberToNameRus = new Hashtable();
                methodsInfo = new List<MethodInfo>();
                propertiesInfo = new List<PropertyInfo>();
                Type type = this.GetType();
                //Type[] allInterfaceTypes = type.get;

                allMethodInfo = type.GetMethods();
                int Identifier = 0;
                foreach (MethodInfo method in allMethodInfo)
                {
                    if (method.DeclaringType == type && !method.IsConstructor)
                    {
                        foreach (object attr in method.GetCustomAttributes(true))
                        {
                            if (attr is Alias)
                            {
                                methodsInfo.Add(method);
                                //methodNameToNumberLat.Add(((Alias)attr).LatName, Identifier);
                                methodNameToNumberLat.Add(method.Name, Identifier);
                                methodNameToNumberRus.Add(((Alias)attr).RusName, Identifier);
                                //methodNumberToNameLat.Add(Identifier, ((Alias)attr).LatName);
                                methodNumberToNameLat.Add(Identifier, method.Name);
                                methodNumberToNameRus.Add(Identifier, ((Alias)attr).RusName);
                                numberToParamsCount.Add(Identifier, method.GetParameters().Length);
                                Identifier++;
                                break;
                            }
                        }
                    }
                }
                    Identifier = 0;
                    foreach(PropertyInfo property in this.GetType().GetProperties())
                    {
                        if (property.DeclaringType == this.GetType())
                        {
                            foreach(object attr in property.GetCustomAttributes(true))
                            {
                                if(attr is Alias)
                                {
                                    propertiesInfo.Add(property);
                                    //propertyNameToNumberLat.Add(((Alias)attr).LatName, Identifier);
                                    propertyNameToNumberLat.Add(property.Name, Identifier);
                                    propertyNameToNumberRus.Add(((Alias)attr).RusName, Identifier);
                                    //propertyNumberToNameLat.Add(Identifier, ((Alias)attr).LatName);
                                    propertyNumberToNameLat.Add(Identifier, property.Name);
                                    propertyNumberToNameRus.Add(Identifier, ((Alias)attr).RusName);
                                    Identifier++;
                                    break;
                                }
                            }
                        }
                    }
                

                extensionName = type.Name;
            }
            catch(Exception)
            {
                return;
            }
        }
        // количество свойств
        public void GetNProps(ref Int32 props)
        {
            props = propertiesInfo.Count;
        }
        // Поиск свойств по имени
        public void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum)
        {
            propNum = -1;
            if (propertyNameToNumberRus.ContainsKey(propName))
                propNum = (int)propertyNameToNumberRus[propName];
            else if (propertyNameToNumberLat.ContainsKey(propName))
                propNum = (int)propertyNameToNumberLat[propName];
        }
        // Поиск имени свойства по номеру
        public void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName)
        {
            propName = propAlias == 0 ? (string)propertyNumberToNameLat[propNum] : (string)propertyNumberToNameRus[propNum];
        }
        // запрашивает значение свойства
        public void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            propVal = propertiesInfo[propNum].GetValue(this, null);
        }
        // устанавливает значение свойства
        public void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {            
            propertiesInfo[propNum].SetValue(this, propVal, null);            
        }
        // можно ли читать свойство (имеет ли get{})
        public void IsPropReadable(Int32 propNum, ref bool propRead)
        {
            propRead = propertiesInfo[propNum].CanRead;
        }
        // можно ли изменять свойство (имеет ли set{})
        public void IsPropWritable(Int32 propNum, ref Boolean propWrite)
        {
            propWrite = propertiesInfo[propNum].CanWrite;
        }
        // возвращает количество методов
        public void GetNMethods(ref Int32 pMethods)
        {
            pMethods = methodsInfo.Count;
        }
        // поиск метода по имени
        public void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum)
        {
            methodNum = -1;
            if (methodNameToNumberRus.ContainsKey(methodName))
                methodNum = (Int32)methodNameToNumberRus[methodName];
            else if (methodNameToNumberLat.ContainsKey(methodName))
                methodNum = (Int32)methodNameToNumberLat[methodName];
        }
        // поиск имени метода по номеру
        public void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName)
        {
            methodName = methodAlias == 0 ? (string)methodNumberToNameLat[methodNum] : (string)methodNameToNumberRus[methodNum];
        }
        // количество параметров метода
        public void GetNParams(Int32 methodNum, ref Int32 pParams)
        {
            pParams = (Int32)numberToParamsCount[methodNum];
        }
        // возвращает значение выбранного параметра по умолчанию
        public void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue)
        {
            if (!DBNull.Value.Equals(methodsInfo[methodNum].GetParameters()[paramNum].DefaultValue))
            {
                paramDefValue = methodsInfo[methodNum].GetParameters()[paramNum].DefaultValue;
            }
        }
        // имеет возвращаемое значение
        public void HasRetVal(Int32 methodNum, ref Boolean retValue)
        {
            retValue = methodsInfo[methodNum].ReturnType != typeof(void);
        }
        // вызываем как процедуру (нет возвращаемого значения)
        public void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                methodsInfo[methodNum].Invoke(this, pParams);
            }
            catch(Exception e)
            {
                errorMessage += GetExceptionMessage(e.InnerException == null ? e : e.InnerException, messagesLevel);
                if (showMеssages) ShowError(errorMessage);
                throw new COMException(errorMessage, (int)HRESULT.S_FALSE);
            }
            if (errorMessage != string.Empty && !ignoreWarnings) throw new COMException(errorMessage, (int)HRESULT.S_FALSE);
        }
        // вызываем как функцию (с возвращаемым значением)
        public void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            if (!methodsInfo[methodNum].Name.Equals("GetErrorMessage")) errorMessage = string.Empty;
            try
            {
                retValue = methodsInfo[methodNum].Invoke(this, pParams);
            }
            catch(Exception e)
            {
                errorMessage += GetExceptionMessage(e.InnerException == null ? e : e.InnerException, messagesLevel);
                if (showMеssages) ShowError(errorMessage);
                throw new COMException(errorMessage, (int)HRESULT.S_FALSE);
            }
            if (errorMessage != string.Empty && !ignoreWarnings && methodsInfo[methodNum].Name.Equals("GetErrorMessage")) throw new COMException(errorMessage, (int)HRESULT.S_FALSE);
        }
        
        public void ShowError(String message, int withException = 0)
        {
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exepInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            exepInfo.wCode = 1006;
            exepInfo.scode = withException;
            exepInfo.bstrDescription = message;
            exepInfo.bstrSource = "AddIn." + this.GetType().Name;
            pErrorInfo.AddError(null, ref exepInfo);
        }        
        public string GetExceptionMessage(Exception e, int level)
        {
            string res = showMеssages ? string.Empty : ("AddIn." + this.GetType().Name);
            switch (level)
            {
                case 0:
                    res += e.Message;
                    break;
                case 1:
                    res += e.Source + ":" + e.Message;
                    break;
                case 2:
                    res += e.ToString();
                    break;
                default:
                    res += e.Message;
                    break;                    
            }
            return res;
        } 
    }
}
