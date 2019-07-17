using System;
using System.Runtime.InteropServices;

namespace AddIn
{
    [Guid("AB634001-F13D-11d0-A459-004095E1DAEA")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitDone
    {
        void Init([MarshalAs(UnmanagedType.IDispatch)] object connection);
        void Done();
        void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info);
    }

    [Guid("AB634003-F13D-11d0-A459-004095E1DAEA")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILanguageExtender
    {
        void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName);
        void GetNProps(ref Int32 props);
        void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum);
        void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName);
        void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void IsPropReadable(Int32 propNum, ref bool propRead);
        void IsPropWritable(Int32 propNum, ref Boolean propWrite);
        void GetNMethods(ref Int32 pMethods);
        void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum);
        void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName);
        void GetNParams(Int32 methodNum, ref Int32 pParams);
        void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue);
        void HasRetVal(Int32 methodNum, ref Boolean retValue);
        void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
        void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
    }
// Интерфейс для обмена информацией об ошибках
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851")]
    [ComVisible(false), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IErrorLog
    {
        void AddError(string pszPropName, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExepInfo);
    }
}

