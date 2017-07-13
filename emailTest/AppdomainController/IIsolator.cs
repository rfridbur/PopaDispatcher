using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Security.Policy;

namespace WatchTool.Core.Services.Isolation
{
    internal class Isolator : MarshalByRefObject
    {
        private static readonly ConcurrentDictionary<Guid, AppDomain> AppDomainsMap;

        static Isolator()
        {
            AppDomainsMap = new ConcurrentDictionary<Guid, AppDomain>();
        }

        public Tuple<Guid, TInterfaceType> GetIsolatedInstance<TInterfaceType, TImplemantingType>() 
               where TInterfaceType     : class
               where TImplemantingType  : MarshalByRefObject, TInterfaceType, new()
        {
            string applicationBaseDirectory = Path.GetDirectoryName(GetType().Assembly.Location);

            AppDomainSetup appDomainSetup = new AppDomainSetup
            {
                ApplicationBase = applicationBaseDirectory
            };

            Evidence baseEvidence = AppDomain.CurrentDomain.Evidence;
            Evidence evidence = new Evidence(baseEvidence);
            Guid name = Guid.NewGuid();

            AppDomain appDomain;

            lock (AppDomainsMap)
            {
                appDomain = AppDomain.CreateDomain(name.ToString(), evidence, appDomainSetup);
            }

            var remoteObj = appDomain.CreateInstanceAndUnwrap(typeof(TImplemantingType).Assembly.FullName,
                            typeof(TImplemantingType).FullName) as TInterfaceType;

            AppDomainsMap.AddOrUpdate(name, appDomain, (guid, domain) => domain);

            return Tuple.Create(name, remoteObj);
        }

        public void UnloadIsolationContext(Guid isolationId)
        {
            if (AppDomainsMap.ContainsKey(isolationId) == false) return;

            AppDomain domain;

            while (AppDomainsMap.TryRemove(isolationId, out domain) == false)
            { /* busy wait */ }

            AppDomain.Unload(domain);

            if (AppDomainsMap.Any())
            {
                GC.Collect();
            }
            else
            {
                for (int i = 0; i < GC.MaxGeneration; i++)
                {
                    GC.Collect(i, GCCollectionMode.Forced, true);
                }
            }

            GC.WaitForPendingFinalizers();
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }

        public void Dispose()
        {
        }
    }
}
