
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace COMProcessor
{
    public static class TypeScriptGenerator
    {
        public static void Generate(ITypeLib typeLib, string outputPath)
        {
            StringBuilder output = new StringBuilder();

            output.AppendLine("// Generated COM declarations");
            output.AppendLine();

            for (int i = 0; i < typeLib.GetTypeInfoCount(); i++)
            {
                TYPEKIND kind;
                ITypeInfo info;

                typeLib.GetTypeInfoType(i, out kind);
                typeLib.GetTypeInfo(i, out info);

                try
                {
                    string name;
                    string doc;
                    int helpContext;
                    string helpFile;

                    info.GetDocumentation(-1, out name, out doc,
                        out helpContext, out helpFile);

                    if (kind == TYPEKIND.TKIND_DISPATCH ||
                        kind == TYPEKIND.TKIND_INTERFACE)
                    {
                        GenerateInterface(info, name, output);
                    }
                    else if (kind == TYPEKIND.TKIND_ENUM)
                    {
                        GenerateEnum(info, name, output);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(info);
                }
            }

            File.WriteAllText(outputPath, output.ToString(),
                new UTF8Encoding(false));
        }

        private static void GenerateInterface(
            ITypeInfo info,
            string name,
            StringBuilder output)
        {
            IntPtr attrPtr;
            info.GetTypeAttr(out attrPtr);

            try
            {
                TYPEATTR attr = (TYPEATTR)Marshal.PtrToStructure(
                    attrPtr, typeof(TYPEATTR));

                output.AppendLine("declare interface " + name + " {");

                Dictionary<int, string> properties =
                    new Dictionary<int, string>();

                Dictionary<int, string> propertyTypes =
                    new Dictionary<int, string>();

                HashSet<int> writable =
                    new HashSet<int>();

                for (int i = 0; i < attr.cFuncs; i++)
                {
                    IntPtr funcPtr;
                    info.GetFuncDesc(i, out funcPtr);

                    try
                    {
                        FUNCDESC func = (FUNCDESC)
                            Marshal.PtrToStructure(
                                funcPtr, typeof(FUNCDESC));

                        string[] names =
                            new string[func.cParams + 1];

                        int found;
                        info.GetNames(func.memid, names,
                            names.Length, out found);

                        string member = found > 0
                            ? names[0] : "Unknown";

                        if (func.invkind ==
                            INVOKEKIND.INVOKE_PROPERTYGET)
                        {
                            properties[func.memid] = member;
                            propertyTypes[func.memid] =
                                ResolveType(info, func.elemdescFunc.tdesc);
                        }
                        else if (func.invkind ==
                            INVOKEKIND.INVOKE_PROPERTYPUT ||
                            func.invkind ==
                            INVOKEKIND.INVOKE_PROPERTYPUTREF)
                        {
                            writable.Add(func.memid);

                            if (!properties.ContainsKey(func.memid))
                                properties[func.memid] = member;

                            if (func.cParams > 0)
                            {
                                IntPtr paramPtr = IntPtr.Add(
                                    func.lprgelemdescParam,
                                    (func.cParams - 1) *
                                    Marshal.SizeOf(typeof(ELEMDESC)));

                                ELEMDESC param = (ELEMDESC)
                                    Marshal.PtrToStructure(
                                        paramPtr, typeof(ELEMDESC));

                                propertyTypes[func.memid] =
                                    ResolveType(info, param.tdesc);
                            }
                        }
                        else if (func.invkind ==
                            INVOKEKIND.INVOKE_FUNC)
                        {
                            output.Append("    " + member + "(");

                            for (int p = 0; p < func.cParams; p++)
                            {
                                if (p > 0)
                                    output.Append(", ");

                                IntPtr paramPtr = IntPtr.Add(
                                    func.lprgelemdescParam,
                                    p * Marshal.SizeOf(typeof(ELEMDESC)));

                                ELEMDESC param = (ELEMDESC)
                                    Marshal.PtrToStructure(
                                        paramPtr, typeof(ELEMDESC));

                                string paramName =
                                    p + 1 < found
                                    ? names[p + 1]
                                    : "arg" + p;

                                bool optional =
                                    (param.desc.paramdesc.wParamFlags &
                                    PARAMFLAG.PARAMFLAG_FOPT) != 0;

                                output.Append(paramName);

                                if (optional)
                                    output.Append("?");

                                output.Append(": ");
                                output.Append(ResolveType(
                                    info, param.tdesc));
                            }

                            output.Append("): ");
                            output.Append(ResolveType(
                                info, func.elemdescFunc.tdesc));

                            output.AppendLine(";");
                        }
                    }
                    finally
                    {
                        info.ReleaseFuncDesc(funcPtr);
                    }
                }

                foreach (var property in properties)
                {
                    string type = propertyTypes.ContainsKey(property.Key)
                        ? propertyTypes[property.Key]
                        : "any";

                    output.Append("    ");

                    if (!writable.Contains(property.Key))
                        output.Append("readonly ");

                    output.Append(property.Value);
                    output.Append(": ");
                    output.Append(type);
                    output.AppendLine(";");
                }

                output.AppendLine("}");
                output.AppendLine();
            }
            finally
            {
                info.ReleaseTypeAttr(attrPtr);
            }
        }

        private static void GenerateEnum(
            ITypeInfo info,
            string name,
            StringBuilder output)
        {
            IntPtr attrPtr;
            info.GetTypeAttr(out attrPtr);

            try
            {
                TYPEATTR attr = (TYPEATTR)Marshal.PtrToStructure(
                    attrPtr, typeof(TYPEATTR));

                output.AppendLine("declare enum " + name + " {");

                for (int i = 0; i < attr.cVars; i++)
                {
                    IntPtr varPtr;
                    info.GetVarDesc(i, out varPtr);

                    try
                    {
                        VARDESC variable = (VARDESC)
                            Marshal.PtrToStructure(
                                varPtr, typeof(VARDESC));

                        string[] names = new string[1];
                        int found;

                        info.GetNames(variable.memid, names, 1, out found);

                        if (found == 0)
                            continue;

                        object value = null;

                        if (variable.varkind == VARKIND.VAR_CONST)
                        {
                            value = Marshal.GetObjectForNativeVariant(
                                variable.desc.lpvarValue);
                        }

                        if (value == null)
                            continue;

                        output.Append("    ");
                        output.Append(names[0]);
                        output.Append(" = ");
                        output.Append(Convert.ToString(
                            value,
                            System.Globalization.CultureInfo.InvariantCulture));
                        output.AppendLine(",");
                    }
                    finally
                    {
                        info.ReleaseVarDesc(varPtr);
                    }
                }

                output.AppendLine("}");
                output.AppendLine();
            }
            finally
            {
                info.ReleaseTypeAttr(attrPtr);
            }
        }

        private static string ResolveType(
            ITypeInfo info,
            TYPEDESC desc)
        {
            switch ((VarEnum)desc.vt)
            {
                case VarEnum.VT_VOID:
                case VarEnum.VT_HRESULT:
                    return "void";

                case VarEnum.VT_I1:
                case VarEnum.VT_I2:
                case VarEnum.VT_I4:
                case VarEnum.VT_I8:
                case VarEnum.VT_UI1:
                case VarEnum.VT_UI2:
                case VarEnum.VT_UI4:
                case VarEnum.VT_UI8:
                case VarEnum.VT_INT:
                case VarEnum.VT_UINT:
                case VarEnum.VT_R4:
                case VarEnum.VT_R8:
                case VarEnum.VT_CY:
                case VarEnum.VT_DECIMAL:
                    return "number";

                case VarEnum.VT_BSTR:
                case VarEnum.VT_LPSTR:
                case VarEnum.VT_LPWSTR:
                    return "string";

                case VarEnum.VT_BOOL:
                    return "boolean";

                case VarEnum.VT_DATE:
                    return "Date";

                case VarEnum.VT_VARIANT:
                case VarEnum.VT_DISPATCH:
                case VarEnum.VT_UNKNOWN:
                    return "any";

                case VarEnum.VT_PTR:
                case VarEnum.VT_SAFEARRAY:
                    TYPEDESC nested = (TYPEDESC)
                        Marshal.PtrToStructure(
                            desc.lpValue, typeof(TYPEDESC));

                    string nestedType = ResolveType(info, nested);

                    return (VarEnum)desc.vt == VarEnum.VT_SAFEARRAY
                        ? nestedType + "[]"
                        : nestedType;

                case VarEnum.VT_USERDEFINED:
                    ITypeInfo referenced;
                    info.GetRefTypeInfo(desc.lpValue.ToInt32(),
                        out referenced);

                    try
                    {
                        string name;
                        string doc;
                        int context;
                        string file;

                        referenced.GetDocumentation(-1,
                            out name, out doc, out context, out file);

                        return name;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(referenced);
                    }

                default:
                    return "any";
            }
        }
    }
}
