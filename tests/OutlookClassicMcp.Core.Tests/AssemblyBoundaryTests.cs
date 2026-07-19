using System;
using System.Linq;
using NUnit.Framework;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Core.Tests
{
    [TestFixture]
    public sealed class AssemblyBoundaryTests
    {
        [Test]
        public void CoreHasNoOfficeOrVstoReferences()
        {
            var references = typeof(ToolExposurePolicy).Assembly
                .GetReferencedAssemblies()
                .Select(reference => reference.Name ?? string.Empty)
                .ToArray();

            var prohibited = references
                .Where(reference =>
                    string.Equals(reference, "Office", StringComparison.Ordinal) ||
                    string.Equals(reference, "stdole", StringComparison.Ordinal) ||
                    string.Equals(reference, "System.Windows.Forms", StringComparison.Ordinal) ||
                    string.Equals(reference, "Microsoft.Win32.Registry", StringComparison.Ordinal) ||
                    reference.StartsWith("Microsoft.Office.", StringComparison.Ordinal) ||
                    reference.StartsWith("Microsoft.VisualStudio.Tools.", StringComparison.Ordinal))
                .ToArray();

            Assert.That(prohibited, Is.Empty);
        }

        [Test]
        public void NormalTestHostDoesNotLoadTheVstoAddIn()
        {
            var loadedAssemblyNames = AppDomain.CurrentDomain
                .GetAssemblies()
                .Select(assembly => assembly.GetName().Name ?? string.Empty)
                .ToArray();

            Assert.That(loadedAssemblyNames, Does.Not.Contain("OutlookClassicMcp.AddIn"));
        }
    }
}
