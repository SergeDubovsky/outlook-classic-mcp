using System.Collections.Generic;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using NUnit.Framework;
using OutlookClassicMcp.Core.Outlook;
using OutlookClassicMcp.Core.Policy;

namespace OutlookClassicMcp.Transport.Tests
{
    [TestFixture]
    public sealed class OutlookReadCatalogTests
    {
        private static readonly string[] ReadToolNames =
        {
            ToolNames.OutlookListMailboxes,
            ToolNames.OutlookListFolders,
            ToolNames.OutlookListMessages,
            ToolNames.OutlookSearchMessages,
            ToolNames.OutlookGetMessage,
            ToolNames.OutlookGetConversation,
            ToolNames.OutlookListAttachments,
        };

        private static readonly string[] SearchInputPropertyNames =
        {
            "scopes",
            "filter",
            "pageSize",
            "cursor",
        };

        private static readonly string[] SearchFilterPropertyNames =
        {
            "sender",
            "recipient",
            "subject",
            "text",
            "receivedFromUtc",
            "receivedToUtc",
            "isUnread",
            "category",
            "hasAttachments",
        };

        [Test]
        public void EveryPhase4DescriptorIsClosedBoundedAndReadOnly()
        {
            foreach (var toolName in ReadToolNames)
            {
                var descriptor = OutlookReadCatalog.CreateDescriptor(toolName);

                Assert.That(descriptor.Name, Is.EqualTo(toolName));
                Assert.That(descriptor.InputSchema.GetProperty("type").GetString(), Is.EqualTo("object"));
                Assert.That(
                    descriptor.InputSchema.GetProperty("additionalProperties").GetBoolean(),
                    Is.False);
                Assert.That(descriptor.OutputSchema, Is.Not.Null);
                Assert.That(
                    descriptor.OutputSchema!.Value.GetProperty("type").GetString(),
                    Is.EqualTo("object"));
                Assert.That(descriptor.Annotations, Is.Not.Null);
                Assert.That(descriptor.Annotations!.ReadOnlyHint, Is.True);
                Assert.That(descriptor.Annotations.DestructiveHint, Is.False);
                Assert.That(descriptor.Annotations.IdempotentHint, Is.True);
                Assert.That(descriptor.Annotations.OpenWorldHint, Is.False);
                AssertObjectSchemasAreClosed(descriptor.InputSchema, toolName + ".input");
                AssertObjectSchemasAreClosed(
                    descriptor.OutputSchema!.Value,
                    toolName + ".output");
            }
        }

        [Test]
        public void PagedInputsAdvertiseTransportDefaultsAndCoreCaps()
        {
            foreach (var toolName in new[]
            {
                ToolNames.OutlookListMailboxes,
                ToolNames.OutlookListFolders,
                ToolNames.OutlookListMessages,
                ToolNames.OutlookSearchMessages,
                ToolNames.OutlookGetConversation,
                ToolNames.OutlookListAttachments,
            })
            {
                var properties = OutlookReadCatalog.CreateDescriptor(toolName)
                    .InputSchema
                    .GetProperty("properties");
                var pageSize = properties.GetProperty("pageSize");
                Assert.That(
                    pageSize.GetProperty("default").GetInt32(),
                    Is.EqualTo(RequestLimits.DefaultReadPageSize));
                Assert.That(
                    pageSize.GetProperty("minimum").GetInt32(),
                    Is.EqualTo(OutlookReadLimits.MinimumPageSize));
                Assert.That(
                    pageSize.GetProperty("maximum").GetInt32(),
                    Is.EqualTo(OutlookReadLimits.MaximumPageSize));
                Assert.That(
                    properties.GetProperty("cursor").GetProperty("maxLength").GetInt32(),
                    Is.EqualTo(HmacCursorCodec.MaximumCursorLength));
            }
        }

        [Test]
        public void SearchSchemaContainsOnlyStructuredScopesAndFilters()
        {
            var input = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookSearchMessages)
                .InputSchema;
            var properties = input.GetProperty("properties");
            Assert.That(
                GetPropertyNames(properties),
                Is.EquivalentTo(SearchInputPropertyNames));

            var scopes = properties.GetProperty("scopes");
            Assert.That(scopes.GetProperty("minItems").GetInt32(), Is.EqualTo(1));
            Assert.That(
                scopes.GetProperty("maxItems").GetInt32(),
                Is.EqualTo(OutlookReadLimits.MaximumSearchScopeCount));
            Assert.That(
                scopes.GetProperty("items").GetProperty("additionalProperties").GetBoolean(),
                Is.False);

            var filter = properties.GetProperty("filter");
            Assert.That(filter.GetProperty("additionalProperties").GetBoolean(), Is.False);
            Assert.That(
                GetPropertyNames(filter.GetProperty("properties")),
                Is.EquivalentTo(SearchFilterPropertyNames));
        }

        [Test]
        public void GetMessageSchemaAdvertisesPlainTextAndMaximumBodyDefaults()
        {
            var properties = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookGetMessage)
                .InputSchema
                .GetProperty("properties");

            Assert.That(
                properties.GetProperty("bodyFormat").GetProperty("default").GetString(),
                Is.EqualTo("plainText"));
            var maximumBody = properties.GetProperty("maximumBodyCharacters");
            Assert.That(
                maximumBody.GetProperty("default").GetInt32(),
                Is.EqualTo(OutlookReadLimits.MaximumBodyCharacters));
            Assert.That(
                maximumBody.GetProperty("maximum").GetInt32(),
                Is.EqualTo(OutlookReadLimits.MaximumBodyCharacters));
        }

        [Test]
        public void FolderLocatorNullabilityMatchesReadSemantics()
        {
            var listFolders = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookListFolders)
                .InputSchema
                .GetProperty("properties")
                .GetProperty("parentFolder");
            Assert.That(listFolders.GetProperty("oneOf")[1].GetProperty("type").GetString(),
                Is.EqualTo("null"));

            var listMessages = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookListMessages)
                .InputSchema
                .GetProperty("properties")
                .GetProperty("folder");
            Assert.That(listMessages.GetProperty("type").GetString(), Is.EqualTo("object"));

            var searchFolder = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookSearchMessages)
                .InputSchema
                .GetProperty("properties")
                .GetProperty("scopes")
                .GetProperty("items")
                .GetProperty("properties")
                .GetProperty("folder");
            Assert.That(searchFolder.GetProperty("oneOf")[1].GetProperty("type").GetString(),
                Is.EqualTo("null"));
        }

        [TestCase(ToolNames.OutlookListMailboxes, "mailboxes")]
        [TestCase(ToolNames.OutlookListFolders, "folders")]
        [TestCase(ToolNames.OutlookListMessages, "messages")]
        [TestCase(ToolNames.OutlookSearchMessages, "messages")]
        [TestCase(ToolNames.OutlookGetConversation, "messages")]
        [TestCase(ToolNames.OutlookListAttachments, "attachments")]
        public void PagedOutputsUseToolSpecificArraysAndTruncationMetadata(
            string toolName,
            string arrayPropertyName)
        {
            var successProperties = OutlookReadCatalog.CreateDescriptor(toolName)
                .OutputSchema!.Value
                .GetProperty("oneOf")[0]
                .GetProperty("properties");
            Assert.That(
                successProperties.GetProperty("partial").GetProperty("type").GetString(),
                Is.EqualTo("boolean"));
            var data = successProperties.GetProperty("data").GetProperty("properties");
            Assert.That(data.TryGetProperty(arrayPropertyName, out _), Is.True);
            Assert.That(data.TryGetProperty("items", out _), Is.False);
            Assert.That(data.GetProperty("resultTruncated").GetProperty("type").GetString(),
                Is.EqualTo("boolean"));
        }

        [Test]
        public void MessageDetailOutputIncludesRecipientTotalsAndPartialMetadata()
        {
            var success = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookGetMessage)
                .OutputSchema!.Value
                .GetProperty("oneOf")[0]
                .GetProperty("properties");
            Assert.That(success.GetProperty("partial").GetProperty("type").GetString(),
                Is.EqualTo("boolean"));
            var data = success.GetProperty("data").GetProperty("properties");
            foreach (var propertyName in new[]
            {
                "message",
                "toRecipients",
                "ccRecipients",
                "bccRecipients",
                "totalToRecipientCount",
                "totalCcRecipientCount",
                "totalBccRecipientCount",
                "toRecipientsTruncated",
                "ccRecipientsTruncated",
                "bccRecipientsTruncated",
                "body",
            })
            {
                Assert.That(data.TryGetProperty(propertyName, out _), Is.True, propertyName);
            }
        }

        [Test]
        public void DescriptorFactoryDoesNotExposeMutableState()
        {
            var first = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookListMailboxes);
            first.Name = "changed";
            first.InputSchema = ParseSchema("{\"type\":\"object\"}");

            var second = OutlookReadCatalog.CreateDescriptor(ToolNames.OutlookListMailboxes);
            Assert.That(second.Name, Is.EqualTo(ToolNames.OutlookListMailboxes));
            Assert.That(second.InputSchema.GetProperty("additionalProperties").GetBoolean(), Is.False);
        }

        private static List<string> GetPropertyNames(JsonElement value)
        {
            var names = new List<string>();
            foreach (var property in value.EnumerateObject())
            {
                names.Add(property.Name);
            }

            return names;
        }

        private static JsonElement ParseSchema(string json)
        {
            using (var document = JsonDocument.Parse(json))
            {
                return document.RootElement.Clone();
            }
        }

        private static void AssertObjectSchemasAreClosed(JsonElement schema, string path)
        {
            if (schema.TryGetProperty("type", out var type) &&
                type.ValueKind == JsonValueKind.String &&
                string.Equals(type.GetString(), "object", System.StringComparison.Ordinal) &&
                !schema.TryGetProperty("oneOf", out _))
            {
                Assert.That(schema.TryGetProperty("additionalProperties", out var additional),
                    Is.True, path);
                Assert.That(additional.GetBoolean(), Is.False, path);
            }

            if (schema.TryGetProperty("properties", out var properties))
            {
                foreach (var property in properties.EnumerateObject())
                {
                    AssertObjectSchemasAreClosed(
                        property.Value,
                        path + ".properties." + property.Name);
                }
            }

            if (schema.TryGetProperty("items", out var items))
            {
                AssertObjectSchemasAreClosed(items, path + ".items");
            }

            if (schema.TryGetProperty("oneOf", out var alternatives))
            {
                var index = 0;
                foreach (var alternative in alternatives.EnumerateArray())
                {
                    AssertObjectSchemasAreClosed(
                        alternative,
                        path + ".oneOf[" + index + "]");
                    index++;
                }
            }
        }
    }
}
