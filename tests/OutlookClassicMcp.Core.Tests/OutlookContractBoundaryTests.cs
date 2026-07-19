using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class OutlookContractBoundaryTests
    {
        private static readonly Type[] DtoTypes =
        {
            typeof(OutlookProbeSnapshot),
            typeof(OutlookDispatcherThreadProof),
            typeof(OutlookStoreProbe),
            typeof(OutlookStoreCapabilities),
            typeof(StandardFolderAvailability),
        };

        [Test]
        public void GatewayHasOneTypedCancelableProbeOperation()
        {
            var methods = typeof(IOutlookGateway).GetMethods();

            Assert.That(methods, Has.Length.EqualTo(1));
            AssertAll(() =>
            {
                Assert.That(methods[0].Name, Is.EqualTo(nameof(IOutlookGateway.ProbeAsync)));
                Assert.That(methods[0].ReturnType, Is.EqualTo(typeof(Task<OutlookProbeSnapshot>)));
                Assert.That(
                    methods[0].GetParameters().Select(parameter => parameter.ParameterType),
                    Is.EqualTo(new[] { typeof(CancellationToken) }));
            });
        }

        [Test]
        public void ProbeDtosAreSealedAndExposeNoPublicSetters()
        {
            foreach (var dtoType in DtoTypes)
            {
                Assert.That(dtoType.IsSealed, Is.True, dtoType.FullName);

                var writableProperties = dtoType
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .Where(property => property.SetMethod != null && property.SetMethod.IsPublic)
                    .Select(property => property.Name)
                    .ToArray();

                Assert.That(writableProperties, Is.Empty, dtoType.FullName);
            }
        }

        [Test]
        public void PublicOutlookContractSignaturesRecursivelyContainNoOfficeOrComTypes()
        {
            var contractTypes = typeof(IOutlookGateway).Assembly
                .GetExportedTypes()
                .Where(type => string.Equals(
                    type.Namespace,
                    typeof(IOutlookGateway).Namespace,
                    StringComparison.Ordinal))
                .ToArray();
            var visited = new HashSet<Type>();
            var prohibited = new List<string>();

            foreach (var contractType in contractTypes)
            {
                InspectContractType(contractType, visited, prohibited);
            }

            Assert.That(prohibited, Is.Empty);
        }

        [Test]
        public void PublicEnumsExposeOnlyTheReviewedValues()
        {
            AssertAll(() =>
            {
                Assert.That(
                    typeof(OutlookStoreType).GetEnumNames(),
                    Is.EqualTo(new[]
                    {
                        nameof(OutlookStoreType.PrimaryExchangeMailbox),
                        nameof(OutlookStoreType.ExchangeMailbox),
                        nameof(OutlookStoreType.ExchangePublicFolder),
                        nameof(OutlookStoreType.AdditionalExchangeMailbox),
                        nameof(OutlookStoreType.NonExchange),
                        nameof(OutlookStoreType.Unknown),
                    }));
                Assert.That(
                    typeof(OutlookFolderAvailability).GetEnumNames(),
                    Is.EqualTo(new[]
                    {
                        nameof(OutlookFolderAvailability.Unknown),
                        nameof(OutlookFolderAvailability.Available),
                        nameof(OutlookFolderAvailability.Missing),
                    }));
                Assert.That(
                    typeof(OutlookProbeWarning).GetEnumNames(),
                    Is.EqualTo(new[]
                    {
                        nameof(OutlookProbeWarning.ArchiveNotExposedByOutlookObjectModel),
                        nameof(OutlookProbeWarning.StoreMetadataIncomplete),
                        nameof(OutlookProbeWarning.StoreLimitReached),
                    }));
            });
        }

        private static void InspectContractType(
            Type type,
            ISet<Type> visited,
            ICollection<string> prohibited)
        {
            InspectSignatureType(type, type.FullName ?? type.Name, visited, prohibited);

            foreach (var constructor in type.GetConstructors(BindingFlags.Instance | BindingFlags.Public))
            {
                foreach (var parameter in constructor.GetParameters())
                {
                    InspectSignatureType(
                        parameter.ParameterType,
                        type.FullName + ".ctor(" + parameter.Name + ")",
                        visited,
                        prohibited);
                }
            }

            foreach (var property in type.GetProperties(
                BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public))
            {
                InspectSignatureType(
                    property.PropertyType,
                    type.FullName + "." + property.Name,
                    visited,
                    prohibited);
            }

            foreach (var method in type.GetMethods(
                BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public))
            {
                if (method.IsSpecialName)
                {
                    continue;
                }

                InspectSignatureType(
                    method.ReturnType,
                    type.FullName + "." + method.Name + " return",
                    visited,
                    prohibited);
                foreach (var parameter in method.GetParameters())
                {
                    InspectSignatureType(
                        parameter.ParameterType,
                        type.FullName + "." + method.Name + "(" + parameter.Name + ")",
                        visited,
                        prohibited);
                }
            }
        }

        private static void InspectSignatureType(
            Type type,
            string path,
            ISet<Type> visited,
            ICollection<string> prohibited)
        {
            if (type.IsByRef || type.IsArray || type.IsPointer)
            {
                var elementType = type.GetElementType();
                if (elementType != null)
                {
                    InspectSignatureType(elementType, path, visited, prohibited);
                }

                return;
            }

            var assemblyName = type.Assembly.GetName().Name ?? string.Empty;
            var typeNamespace = type.Namespace ?? string.Empty;
            if (type == typeof(object) ||
                type.IsCOMObject ||
                type.IsImport ||
                type.GetCustomAttributes(typeof(ComImportAttribute), inherit: false).Length != 0 ||
                string.Equals(assemblyName, "Office", StringComparison.Ordinal) ||
                assemblyName.StartsWith("Microsoft.Office.", StringComparison.Ordinal) ||
                typeNamespace.StartsWith("Microsoft.Office.", StringComparison.Ordinal) ||
                typeNamespace.StartsWith("Microsoft.VisualStudio.Tools.", StringComparison.Ordinal) ||
                typeNamespace.StartsWith("System.Runtime.InteropServices.ComTypes", StringComparison.Ordinal))
            {
                prohibited.Add(path + " -> " + type.FullName);
            }

            if (!visited.Add(type))
            {
                return;
            }

            if (type.IsGenericType)
            {
                foreach (var genericArgument in type.GetGenericArguments())
                {
                    InspectSignatureType(genericArgument, path, visited, prohibited);
                }
            }
        }

        private static void AssertAll(Action assertions)
        {
            Assert.Multiple((Action)(() => assertions()));
        }

    }
}
