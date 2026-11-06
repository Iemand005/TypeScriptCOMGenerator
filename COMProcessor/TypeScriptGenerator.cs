
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
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

            HashSet<string> eventInterfaces;
            Dictionary<string, string> interfaceEvents;

            CollectEventInterfaces(typeLib,
                out eventInterfaces, out interfaceEvents);

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
                        GenerateInterface(info, name, output,
                            eventInterfaces.Contains(name),
                            interfaceEvents.ContainsKey(name)
                                ? interfaceEvents[name]
                                : null);
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

        private static void CollectEventInterfaces(
            ITypeLib typeLib,
            out HashSet<string> eventInterfaces,
            out Dictionary<string, string> interfaceEvents)
        {
            eventInterfaces = new HashSet<string>(StringComparer.Ordinal);
            interfaceEvents = new Dictionary<string, string>(
                StringComparer.Ordinal);

            Dictionary<string, string> preferred =
                new Dictionary<string, string>(StringComparer.Ordinal);

            Dictionary<string, string> fallback =
                new Dictionary<string, string>(StringComparer.Ordinal);

            for (int i = 0; i < typeLib.GetTypeInfoCount(); i++)
            {
                TYPEKIND kind;
                typeLib.GetTypeInfoType(i, out kind);

                if (kind != TYPEKIND.TKIND_COCLASS)
                    continue;

                ITypeInfo info;
                typeLib.GetTypeInfo(i, out info);

                try
                {
                    IntPtr attrPtr;
                    info.GetTypeAttr(out attrPtr);

                    try
                    {
                        TYPEATTR attr = (TYPEATTR)Marshal.PtrToStructure(
                            attrPtr, typeof(TYPEATTR));

                        List<string> sources = new List<string>();

                        List<KeyValuePair<string, bool>> others =
                            new List<KeyValuePair<string, bool>>();

                        string defaultSource = null;

                        for (int j = 0; j < attr.cImplTypes; j++)
                        {
                            int refType;
                            info.GetRefTypeOfImplType(j, out refType);

                            ITypeInfo refInfo;
                            info.GetRefTypeInfo(refType, out refInfo);

                            try
                            {
                                string refName;
                                string refDoc;
                                int refContext;
                                string refFile;

                                refInfo.GetDocumentation(-1,
                                    out refName, out refDoc,
                                    out refContext, out refFile);

                                IMPLTYPEFLAGS flags;
                                info.GetImplTypeFlags(j, out flags);

                                bool isSource =
                                    (flags & IMPLTYPEFLAGS
                                        .IMPLTYPEFLAG_FSOURCE) != 0;

                                bool isDefault =
                                    (flags & IMPLTYPEFLAGS
                                        .IMPLTYPEFLAG_FDEFAULT) != 0;

                                if (isSource)
                                {
                                    sources.Add(refName);
                                    eventInterfaces.Add(refName);

                                    if (isDefault && defaultSource == null)
                                        defaultSource = refName;
                                }
                                else
                                {
                                    others.Add(new KeyValuePair<string, bool>(
                                        refName, isDefault));
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(refInfo);
                            }
                        }

                        string source = defaultSource != null
                            ? defaultSource
                            : sources.Count > 0 ? sources[0] : null;

                        if (source == null)
                            continue;

                        foreach (var other in others)
                        {
                            fallback[other.Key] = source;

                            if (other.Value)
                                preferred[other.Key] = source;
                        }
                    }
                    finally
                    {
                        info.ReleaseTypeAttr(attrPtr);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(info);
                }
            }

            foreach (var item in fallback)
            {
                if (!preferred.ContainsKey(item.Key))
                    preferred[item.Key] = item.Value;
            }

            foreach (var item in preferred)
                interfaceEvents[item.Key] = item.Value;
        }

        private static void GenerateInterface(
            ITypeInfo info,
            string name,
            StringBuilder output,
            bool isEventSource,
            string eventSourceName)
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

                List<KeyValuePair<string, string>> eventHandlers =
                    new List<KeyValuePair<string, string>>();

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
                            StringBuilder parameters = new StringBuilder();

                            for (int p = 0; p < func.cParams; p++)
                            {
                                if (p > 0)
                                    parameters.Append(", ");

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

                                parameters.Append(paramName);

                                if (optional)
                                    parameters.Append("?");

                                parameters.Append(": ");
                                parameters.Append(ResolveType(
                                    info, param.tdesc));
                            }

                            string returnType = ResolveType(
                                info, func.elemdescFunc.tdesc);

                            output.Append("    " + member + "(");
                            output.Append(parameters);
                            output.Append("): ");
                            output.Append(returnType);
                            output.AppendLine(";");

                            if (isEventSource)
                            {
                                eventHandlers.Add(
                                    new KeyValuePair<string, string>(
                                        GetEventName(member),
                                        "(" + parameters + ") => "
                                        + returnType));
                            }
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

                if (eventSourceName != null)
                {
                    string eventMap = eventSourceName + "Handlers";

                    output.AppendLine("    attachEvent<T extends keyof "
                        + eventMap + ">(");
                    output.AppendLine("        event: T,");
                    output.AppendLine("        handler: "
                        + eventMap + "[T]");
                    output.AppendLine("    ): boolean;");
                    output.AppendLine("    detachEvent<T extends keyof "
                        + eventMap + ">(");
                    output.AppendLine("        event: T,");
                    output.AppendLine("        handler: "
                        + eventMap + "[T]");
                    output.AppendLine("    ): void;");
                }

                output.AppendLine("}");
                output.AppendLine();

                if (isEventSource)
                    GenerateEventHandlerMap(name, eventHandlers, output);
            }
            finally
            {
                info.ReleaseTypeAttr(attrPtr);
            }
        }

        private static void GenerateEventHandlerMap(
            string name,
            List<KeyValuePair<string, string>> eventHandlers,
            StringBuilder output)
        {
            output.AppendLine("declare interface " + name + "Handlers {");

            foreach (var handler in eventHandlers)
            {
                output.Append("    \"");
                output.Append(handler.Key);
                output.Append("\": ");
                output.Append(handler.Value);
                output.AppendLine(";");
            }

            output.AppendLine("}");
            output.AppendLine();
        }

        private static string GetEventName(string member)
        {
            string name = member;

            if (name.StartsWith("On", StringComparison.OrdinalIgnoreCase) &&
                name.Length > 2)
            {
                name = name.Substring(2);
            }

            return "on" + name.ToLowerInvariant();
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
