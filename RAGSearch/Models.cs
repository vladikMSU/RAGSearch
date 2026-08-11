using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace RAGSearch
{
    [DataContract]
    internal sealed class AttachmentPayload
    {
        [DataMember(Name = "name")]
        public string name { get; set; }
        [DataMember(Name = "size")]
        public long size { get; set; }
        [DataMember(Name = "content_type")]
        public string content_type { get; set; }
        [DataMember(Name = "temp_path")]
        public string temp_path { get; set; }
    }

    [DataContract]
    internal sealed class MessagePayload
    {
        [DataMember(Name = "entry_id")]
        public string entry_id { get; set; }
        [DataMember(Name = "store_id")]
        public string store_id { get; set; }
        [DataMember(Name = "folder_entry_id")]
        public string folder_entry_id { get; set; }
        [DataMember(Name = "folder_path")]
        public string folder_path { get; set; }
        [DataMember(Name = "store_name")]
        public string store_name { get; set; }
        [DataMember(Name = "subject")]
        public string subject { get; set; }
        [DataMember(Name = "sender_name")]
        public string sender_name { get; set; }
        [DataMember(Name = "sender_email")]
        public string sender_email { get; set; }
        [DataMember(Name = "to")]
        public string to { get; set; }
        [DataMember(Name = "cc")]
        public string cc { get; set; }
        [DataMember(Name = "sent_at")]
        public string sent_at { get; set; }
        [DataMember(Name = "received_at")]
        public string received_at { get; set; }
        [DataMember(Name = "modified_at")]
        public string modified_at { get; set; }
        [DataMember(Name = "internet_message_id")]
        public string internet_message_id { get; set; }
        [DataMember(Name = "conversation_id")]
        public string conversation_id { get; set; }
        [DataMember(Name = "body")]
        public string body { get; set; }
        [DataMember(Name = "attachments")]
        public List<AttachmentPayload> attachments { get; set; }
    }

    [DataContract]
    internal sealed class IngestRequest
    {
        [DataMember(Name = "messages")]
        public List<MessagePayload> messages { get; set; }
    }

    [DataContract]
    internal sealed class IngestResponse
    {
        [DataMember(Name = "accepted")]
        public int accepted { get; set; }
        [DataMember(Name = "failed")]
        public int failed { get; set; }
        [DataMember(Name = "errors")]
        public List<IngestErrorDto> errors { get; set; }
    }

    [DataContract]
    internal sealed class IngestErrorDto
    {
        [DataMember(Name = "index")]
        public int index { get; set; }
        [DataMember(Name = "error")]
        public string error { get; set; }
        [DataMember(Name = "entry_id")]
        public string entry_id { get; set; }
        [DataMember(Name = "store_id")]
        public string store_id { get; set; }
    }

    [DataContract]
    internal sealed class SearchRequest
    {
        [DataMember(Name = "query")]
        public string query { get; set; }
        [DataMember(Name = "limit")]
        public int limit { get; set; }
        [DataMember(Name = "filters")]
        public Dictionary<string, object> filters { get; set; }
    }

    [DataContract]
    internal sealed class SearchResponse
    {
        [DataMember(Name = "results")]
        public List<SearchResultDto> results { get; set; }
        [DataMember(Name = "total")]
        public int total { get; set; }
        [DataMember(Name = "mode")]
        public string mode { get; set; }
    }

    [DataContract]
    internal sealed class ResetIndexResponse
    {
        [DataMember(Name = "deleted_messages")]
        public long deleted_messages { get; set; }
        [DataMember(Name = "deleted_attachments")]
        public long deleted_attachments { get; set; }
        [DataMember(Name = "deleted_chunks")]
        public long deleted_chunks { get; set; }
    }

    [DataContract]
    internal sealed class SearchResultDto
    {
        [DataMember(Name = "entry_id")]
        public string entry_id { get; set; }
        [DataMember(Name = "store_id")]
        public string store_id { get; set; }
        [DataMember(Name = "folder_entry_id")]
        public string folder_entry_id { get; set; }
        [DataMember(Name = "subject")]
        public string subject { get; set; }
        [DataMember(Name = "sender_name")]
        public string sender_name { get; set; }
        [DataMember(Name = "sender_email")]
        public string sender_email { get; set; }
        [DataMember(Name = "received_at")]
        public string received_at { get; set; }
        [DataMember(Name = "folder_path")]
        public string folder_path { get; set; }
        [DataMember(Name = "store_name")]
        public string store_name { get; set; }
        [DataMember(Name = "internet_message_id")]
        public string internet_message_id { get; set; }
        [DataMember(Name = "conversation_id")]
        public string conversation_id { get; set; }
        [DataMember(Name = "score")]
        public double score { get; set; }
        [DataMember(Name = "snippet")]
        public string snippet { get; set; }
        [DataMember(Name = "matched_sources")]
        public List<string> matched_sources { get; set; }

        [IgnoreDataMember]
        public string ScoreDisplay
        {
            get { return score.ToString("0.000"); }
        }

        [IgnoreDataMember]
        public string ReceivedDisplay
        {
            get
            {
                DateTime value;
                return DateTime.TryParse(received_at, out value)
                    ? value.ToLocalTime().ToString("g")
                    : received_at ?? string.Empty;
            }
        }
    }

    [DataContract]
    internal sealed class HealthResponse
    {
        [DataMember(Name = "status")]
        public string status { get; set; }
        [DataMember(Name = "embedding_backend")]
        public string embedding_backend { get; set; }
        [DataMember(Name = "database")]
        public string database { get; set; }
    }

    internal sealed class IndexProgress
    {
        public int Processed { get; set; }
        public int EstimatedTotal { get; set; }
        public int Failed { get; set; }
        public string CurrentFolder { get; set; }
        public string Status { get; set; }
        public bool IsRunning { get; set; }
    }
}
