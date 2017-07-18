using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses
{
    static class ClassFactory
    {
        internal static object CreateInstance(ConstructorInfo ctor, params object[] parameters) {
            return ctor.Invoke(parameters);
        }

        internal static object CreateInstance(string typeName)
        {
            var type = Loader.GetType(typeName);
            return Activator.CreateInstance(type);
        }

        internal static object CreateInstance(string typeName, params object[] parameters)
        {
            var type = Loader.GetType(typeName);
            return Activator.CreateInstance(type, parameters);
        }

        internal static object CreateInstance(string typeName, string assemblyName, bool assemblyFromResource = false)
        {
            if (assemblyFromResource)
                Loader.LoadAssemblyFromResource(assemblyName);
            else
                Loader.LoadAssemblyByPath(assemblyName);
            return ClassFactory.CreateInstance(typeName);
        }

        internal static object CreateInstance(string typeName, string assemblyName, bool assemblyFromResource = false, params object[] parameters)
        {
            if (assemblyFromResource)
                Loader.LoadAssemblyFromResource(assemblyName);
            else
                Loader.LoadAssemblyByPath(assemblyName);
            return ClassFactory.CreateInstance(typeName, parameters);
        }
    }
}
