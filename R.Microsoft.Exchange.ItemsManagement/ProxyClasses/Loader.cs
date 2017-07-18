using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement
{
    static class Loader
    {
        private static Dictionary<string, Assembly> assemblies = new Dictionary<string,Assembly>();
        private static Dictionary<string, Type> types = new Dictionary<string,Type>();

        internal static Assembly LoadAssemblyByPath(string path)
        {
            if (!Loader.assemblies.ContainsKey(path))
            {
                var tmpAssembly = Assembly.LoadFile(path);
                Loader.assemblies.Add(path, tmpAssembly);
                return tmpAssembly;
            } 
            else 
                return Loader.assemblies[path];
        }

        internal static Assembly LoadAssemblyFromResource(string resourceName)
        {
            if (!Loader.assemblies.ContainsKey(resourceName))
            {
                Stream imageStream = new MemoryStream(Properties.Resources.Microsoft_Exchange_WebServices);
                long bytestreamMaxLength = imageStream.Length;
                byte[] buffer = new byte[bytestreamMaxLength];
                imageStream.Read(buffer, 0, (int)bytestreamMaxLength);
                var tmpAssembly = Assembly.Load(buffer);
                Loader.assemblies.Add(resourceName, tmpAssembly);
                return tmpAssembly;
            }
            else
                return Loader.assemblies[resourceName];
        }

        internal static Type GetType(string typeName)
        {
            if (Loader.types.ContainsKey(typeName))
                return Loader.types[typeName];
            Type type = null;
            if (assemblies != null)
            {
                foreach (var assembly in Loader.assemblies.Values)
                {
                    Logger.Write("Loader.GetType(): Looking in " + assembly.GetName().FullName, LogVerbosity.Verbose);
                    foreach (var assType in assembly.GetTypes())
                        if (assType.Name == typeName)
                        {
                            type = assType;
                            break;
                        }
                    //type = assembly.GetType(typeName, false);
                    if (type != null)
                        break;
                }
                if (null == type)
                    throw new ArgumentException("Unable to locate type: " + typeName);
            }
            Loader.types.Add(typeName, type);
            return type;
        }
    }
}
