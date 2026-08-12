using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace RAGSearch
{
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
        [DataMember(Name = "candidate_count")]
        public int candidate_count { get; set; }
        [DataMember(Name = "eligible_count")]
        public int eligible_count { get; set; }
        [DataMember(Name = "lexical_match_count")]
        public int lexical_match_count { get; set; }
        [DataMember(Name = "lexical_fallback_count")]
        public int lexical_fallback_count { get; set; }
        [DataMember(Name = "lexical_gate")]
        public bool lexical_gate { get; set; }
        [DataMember(Name = "best_vector_similarity")]
        public double best_vector_similarity { get; set; }
        [DataMember(Name = "best_vector_distance")]
        public double best_vector_distance { get; set; }
        [DataMember(Name = "cutoff_similarity")]
        public double cutoff_similarity { get; set; }
        [DataMember(Name = "cutoff_distance")]
        public double cutoff_distance { get; set; }
        [DataMember(Name = "max_results")]
        public int max_results { get; set; }
        [DataMember(Name = "ranking")]
        public string ranking { get; set; }
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
        [DataMember(Name = "vector_similarity")]
        public double vector_similarity { get; set; }
        [DataMember(Name = "vector_distance")]
        public double vector_distance { get; set; }
        [DataMember(Name = "hybrid_score")]
        public double hybrid_score { get; set; }
        [DataMember(Name = "lexical_score")]
        public double lexical_score { get; set; }
        [DataMember(Name = "lexical_match_kind")]
        public string lexical_match_kind { get; set; }
        [DataMember(Name = "rank")]
        public int rank { get; set; }
        [DataMember(Name = "ranking_basis")]
        public string ranking_basis { get; set; }
        [DataMember(Name = "snippet")]
        public string snippet { get; set; }
        [DataMember(Name = "matched_sources")]
        public List<string> matched_sources { get; set; }

        [IgnoreDataMember]
        public string ScoreDisplay
        {
            get { return vector_distance.ToString("0.000"); }
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

}
