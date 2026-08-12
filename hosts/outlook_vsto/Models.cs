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
        [DataMember(Name = "deleted_messages")]
        public long DeletedMessages { get; set; }
        [DataMember(Name = "deleted_attachments")]
        public long DeletedAttachments { get; set; }
        [DataMember(Name = "deleted_chunks")]
        public long DeletedChunks { get; set; }
    }

    [DataContract]
    internal sealed class SearchResultDto
    {
        [DataMember(Name = "entry_id")]
        public string EntryId { get; set; }
        [DataMember(Name = "store_id")]
        public string StoreId { get; set; }
        [DataMember(Name = "subject")]
        public string Subject { get; set; }
        [DataMember(Name = "sender_name")]
        public string SenderName { get; set; }
        [DataMember(Name = "sender_email")]
        public string SenderEmail { get; set; }
        [DataMember(Name = "received_at")]
        public string ReceivedAt { get; set; }
        [DataMember(Name = "folder_path")]
        public string FolderPath { get; set; }
        [DataMember(Name = "store_name")]
        public string StoreName { get; set; }
        [DataMember(Name = "rank")]
        public int Rank { get; set; }
        [DataMember(Name = "snippet")]
        public string Snippet { get; set; }
        [DataMember(Name = "matched_sources")]
        public List<string> MatchedSources { get; set; }

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
                    throw new FormatException("Search result received_at is not an ISO-8601 timestamp.");
                }
                return value.ToLocalTime().ToString("g");
            }
        }
    }

    [DataContract]
    internal sealed class HealthResponse
    {
        [DataMember(Name = "status")]
        public string Status { get; set; }
    }

}
