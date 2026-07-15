using System;
using System.Linq;
using NUnit.Framework;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class AssemblyBoundaryTests
    {
        [Test]
        public void TransportHasNoOfficeOrVstoReferences()
        {
            var references = typeof(LoopbackHttpServer).Assembly
                .GetReferencedAssemblies()
                .Select(reference => reference.Name ?? string.Empty)
                .ToArray();

            var prohibited = references
                .Where(reference =>
                    string.Equals(reference, "Office", StringComparison.Ordinal) ||
                    string.Equals(reference, "stdole", StringComparison.Ordinal) ||
                    reference.StartsWith("Microsoft.Office.", StringComparison.Ordinal) ||
                    reference.StartsWith("Microsoft.VisualStudio.Tools.", StringComparison.Ordinal))
                .ToArray();

            Assert.That(prohibited, Is.Empty);
        }
    }
}
