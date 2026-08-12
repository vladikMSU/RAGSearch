using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace RAGSearch
{
    [DataContract]
    internal sealed class SearchRequest
    {
        [DataMember(Name = "query")]
        public string Query { get; set; }
        [DataMember(Name = "limit")]
        public int Limit { get; set; }
    }

    [DataContract]
    internal sealed class SearchResponse
    {
        [DataMember(Name = "results")]
        public List<SearchResultDto> Results { get; set; }
        [DataMember(Name = "mode")]
        public string Mode { get; set; }
    }

    [DataContract]
    internal sealed class ResetIndexResponse
    {
        [DataMember(Name = "deleted_documents")]
        public long DeletedDocuments { get; set; }
        [DataMember(Name = "deleted_parts")]
        public long DeletedParts { get; set; }
        [DataMember(Name = "deleted_chunks")]
        public long DeletedChunks { get; set; }
    }

    [DataContract]
    internal sealed class SearchResultDto
    {
        [DataMember(Name = "source_key")]
        public string SourceKey { get; set; }
        [DataMember(Name = "kind")]
        public string Kind { get; set; }
        [DataMember(Name = "title")]
        public string Title { get; set; }
        [DataMember(Name = "metadata")]
        public Dictionary<string, object> Metadata { get; set; }
        [DataMember(Name = "locator")]
        public Dictionary<string, object> Locator { get; set; }
        [DataMember(Name = "rank")]
        public int Rank { get; set; }
        [DataMember(Name = "snippet")]
        public string Snippet { get; set; }
        [DataMember(Name = "snippet_part")]
        public string SnippetPart { get; set; }
        [DataMember(Name = "matched_parts")]
        public List<string> MatchedParts { get; set; }

        [IgnoreDataMember]
        public string Subject { get { return Title ?? string.Empty; } }
        [IgnoreDataMember]
        public string SenderName { get { return StringValue(Metadata, "sender_name"); } }
        [IgnoreDataMember]
        public string SenderEmail { get { return StringValue(Metadata, "sender_email"); } }
        [IgnoreDataMember]
        public string ReceivedAt { get { return StringValue(Metadata, "received_at"); } }
        [IgnoreDataMember]
        public string FolderPath { get { return StringValue(Metadata, "folder_path"); } }
        [IgnoreDataMember]
        public string StoreName { get { return StringValue(Metadata, "store_name"); } }
        [IgnoreDataMember]
        public string LocatorConnector { get { return StringValue(Locator, "connector"); } }
        [IgnoreDataMember]
        public string LocatorStoreId { get { return StringValue(Locator, "store_id"); } }
        [IgnoreDataMember]
        public string LocatorEntryId { get { return StringValue(Locator, "entry_id"); } }
        [IgnoreDataMember]
        public bool IsOutlookDocument
        {
            get
            {
                return string.Equals(LocatorConnector, "outlook_mapi", StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(LocatorStoreId) &&
                    !string.IsNullOrWhiteSpace(LocatorEntryId);
            }
        }

        [IgnoreDataMember]
        public string ReceivedDisplay
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ReceivedAt))
                {
                    return string.Empty;
                }
                DateTimeOffset value;
                if (!DateTimeOffset.TryParse(ReceivedAt, out value))
                {
                    return ReceivedAt;
                }
                return value.ToLocalTime().ToString("g");
            }
        }

        private static string StringValue(Dictionary<string, object> values, string key)
        {
            object value;
            if (values == null || !values.TryGetValue(key, out value))
            {
                return string.Empty;
            }
            return value as string ?? string.Empty;
        }
    }

    [DataContract]
    internal sealed class OutlookDocumentMetadata
    {
        [DataMember(Name = "sender_name", EmitDefaultValue = false)]
        public string SenderName { get; set; }
        [DataMember(Name = "sender_email", EmitDefaultValue = false)]
        public string SenderEmail { get; set; }
        [DataMember(Name = "to", EmitDefaultValue = false)]
        public string To { get; set; }
        [DataMember(Name = "cc", EmitDefaultValue = false)]
        public string Cc { get; set; }
        [DataMember(Name = "sent_at", EmitDefaultValue = false)]
        public string SentAt { get; set; }
        [DataMember(Name = "received_at", EmitDefaultValue = false)]
        public string ReceivedAt { get; set; }
        [DataMember(Name = "modified_at", EmitDefaultValue = false)]
        public string ModifiedAt { get; set; }
        [DataMember(Name = "folder_path", EmitDefaultValue = false)]
        public string FolderPath { get; set; }
        [DataMember(Name = "store_name", EmitDefaultValue = false)]
        public string StoreName { get; set; }
        [DataMember(Name = "internet_message_id", EmitDefaultValue = false)]
        public string InternetMessageId { get; set; }
        [DataMember(Name = "conversation_id", EmitDefaultValue = false)]
        public string ConversationId { get; set; }
        [DataMember(Name = "attachments_truncated", EmitDefaultValue = false)]
        public bool AttachmentsTruncated { get; set; }
    }

    [DataContract]
    internal sealed class OutlookDocumentLocator
    {
        [DataMember(Name = "connector")]
        public string Connector { get; set; }
        [DataMember(Name = "store_id")]
        public string StoreId { get; set; }
        [DataMember(Name = "entry_id")]
        public string EntryId { get; set; }
        [DataMember(Name = "folder_entry_id")]
        public string FolderEntryId { get; set; }
    }

    [DataContract]
    internal sealed class DocumentDto
    {
        [DataMember(Name = "source_key")]
        public string SourceKey { get; set; }
        [DataMember(Name = "kind")]
        public string Kind { get; set; }
        [DataMember(Name = "title")]
        public string Title { get; set; }
        [DataMember(Name = "metadata")]
        public OutlookDocumentMetadata Metadata { get; set; }
        [DataMember(Name = "locator")]
        public OutlookDocumentLocator Locator { get; set; }
        [DataMember(Name = "parts")]
        public List<DocumentPartDto> Parts { get; set; }
    }

    [DataContract]
    internal sealed class DocumentPartDto
    {
        [DataMember(Name = "key")]
        public string Key { get; set; }
        [DataMember(Name = "kind")]
        public string Kind { get; set; }
        [DataMember(Name = "name", EmitDefaultValue = false)]
        public string Name { get; set; }
        [DataMember(Name = "media_type", EmitDefaultValue = false)]
        public string MediaType { get; set; }
        [DataMember(Name = "size")]
        public long Size { get; set; }
        [DataMember(Name = "text", EmitDefaultValue = false)]
        public string Text { get; set; }
        [DataMember(Name = "content_base64", EmitDefaultValue = false)]
        public string ContentBase64 { get; set; }
        [DataMember(Name = "truncated")]
        public bool Truncated { get; set; }
    }

    [DataContract]
    internal sealed class DocumentUpsertResponse
    {
        [DataMember(Name = "source_key")]
        public string SourceKey { get; set; }
        [DataMember(Name = "status")]
        public string Status { get; set; }
    }

    [DataContract]
    internal sealed class HealthResponse
    {
        [DataMember(Name = "status")]
        public string Status { get; set; }
        [DataMember(Name = "protocol")]
        public int Protocol { get; set; }
    }

}
