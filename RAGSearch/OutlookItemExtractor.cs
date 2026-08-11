using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RAGSearch
{
    internal sealed class OutlookItemExtractor
    {
        private const int MaxBodyCharacters = 4 * 1024 * 1024;
        private const long MaxAttachmentBytes = 64L * 1024L * 1024L;
        private const string InternetMessageIdUnicode =
            "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        private const string InternetMessageIdAnsi =
            "http://schemas.microsoft.com/mapi/proptag/0x1035001E";
        private const string SmtpAddressUnicode =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

        private readonly string spoolRoot;

        public OutlookItemExtractor()
        {
            spoolRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RAGSearch",
                "spool");
            Directory.CreateDirectory(spoolRoot);
        }

        public MessagePayload Extract(
            Outlook.MailItem mail,
            Outlook.MAPIFolder folder,
            string storeName)
        {
            if (mail == null || folder == null)
            {
                return null;
            }

            var entryId = SafeString(() => mail.EntryID);
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return null;
            }

            return new MessagePayload
            {
                entry_id = entryId,
                store_id = SafeString(() => folder.StoreID),
                folder_entry_id = SafeString(() => folder.EntryID),
                folder_path = SafeString(() => folder.FolderPath),
                store_name = storeName ?? string.Empty,
                subject = SafeString(() => mail.Subject),
                sender_name = SafeString(() => mail.SenderName),
                sender_email = ResolveSenderAddress(mail),
                to = SafeString(() => mail.To),
                cc = SafeString(() => mail.CC),
                sent_at = SafeDate(() => mail.SentOn),
                received_at = SafeDate(() => mail.ReceivedTime),
                modified_at = SafeDate(() => mail.LastModificationTime),
                internet_message_id = ReadMapiString(mail, InternetMessageIdUnicode, InternetMessageIdAnsi),
                conversation_id = SafeString(() => mail.ConversationID),
                body = Limit(SafeString(() => mail.Body), MaxBodyCharacters),
                attachments = SaveAttachments(mail, entryId)
            };
        }

        private List<AttachmentPayload> SaveAttachments(Outlook.MailItem mail, string entryId)
        {
            var result = new List<AttachmentPayload>();
            Outlook.Attachments attachments = null;
            string messageDirectory = null;

            try
            {
                attachments = mail.Attachments;
                var count = attachments == null ? 0 : attachments.Count;
                if (count == 0)
                {
                    return result;
                }

                messageDirectory = Path.Combine(spoolRoot, Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(messageDirectory);

                for (var index = 1; index <= count; index++)
                {
                    Outlook.Attachment attachment = null;
                    try
                    {
                        attachment = attachments[index];
                        var originalName = SafeString(() => attachment.FileName);
                        var size = SafeLong(() => attachment.Size);
                        string savedPath = null;

                        if (size <= MaxAttachmentBytes)
                        {
                            var fileName = string.Format(
                                CultureInfo.InvariantCulture,
                                "{0:D3}-{1}",
                                index,
                                SanitizeFileName(originalName));
                            savedPath = Path.Combine(messageDirectory, fileName);
                            attachment.SaveAsFile(savedPath);
                        }

                        result.Add(new AttachmentPayload
                        {
                            name = originalName,
                            size = size,
                            content_type = GuessContentType(originalName),
                            temp_path = savedPath
                        });
                    }
                    catch (COMException)
                    {
                        // A single corrupt or policy-blocked attachment must not abort the mailbox.
                    }
                    catch (IOException)
                    {
                        // The service records the remaining attachments; indexing continues.
                    }
                    finally
                    {
                        ComRelease.Final(attachment);
                    }
                }
            }
            finally
            {
                ComRelease.Final(attachments);
                if (messageDirectory != null && Directory.Exists(messageDirectory) &&
                    !Directory.EnumerateFileSystemEntries(messageDirectory).Any())
                {
                    Directory.Delete(messageDirectory, false);
                }
            }

            return result;
        }

        private static string ResolveSenderAddress(Outlook.MailItem mail)
        {
            var direct = SafeString(() => mail.SenderEmailAddress);
            var senderType = SafeString(() => mail.SenderEmailType);
            if (!string.Equals(senderType, "EX", StringComparison.OrdinalIgnoreCase))
            {
                return direct;
            }

            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchangeUser = null;
            Outlook.ExchangeDistributionList distributionList = null;
            Outlook.PropertyAccessor accessor = null;
            try
            {
                sender = mail.Sender;
                if (sender == null)
                {
                    return direct;
                }

                exchangeUser = sender.GetExchangeUser();
                var primary = exchangeUser == null
                    ? null
                    : SafeString(() => exchangeUser.PrimarySmtpAddress);
                if (!string.IsNullOrWhiteSpace(primary))
                {
                    return primary;
                }

                distributionList = sender.GetExchangeDistributionList();
                primary = distributionList == null
                    ? null
                    : SafeString(() => distributionList.PrimarySmtpAddress);
                if (!string.IsNullOrWhiteSpace(primary))
                {
                    return primary;
                }

                accessor = sender.PropertyAccessor;
                var smtp = accessor.GetProperty(SmtpAddressUnicode) as string;
                return string.IsNullOrWhiteSpace(smtp) ? direct : smtp;
            }
            catch (COMException)
            {
                return direct;
            }
            finally
            {
                ComRelease.Final(accessor);
                ComRelease.Final(distributionList);
                ComRelease.Final(exchangeUser);
                ComRelease.Final(sender);
            }
        }

        private static string ReadMapiString(Outlook.MailItem mail, params string[] schemas)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = mail.PropertyAccessor;
                foreach (var schema in schemas)
                {
                    try
                    {
                        var value = accessor.GetProperty(schema) as string;
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            return value;
                        }
                    }
                    catch (COMException)
                    {
                        // Try the ANSI/Unicode alternative.
                    }
                }
            }
            catch (COMException)
            {
                return string.Empty;
            }
            finally
            {
                ComRelease.Final(accessor);
            }

            return string.Empty;
        }

        private static string SafeString(Func<string> reader)
        {
            try
            {
                return reader() ?? string.Empty;
            }
            catch (COMException)
            {
                return string.Empty;
            }
        }

        private static long SafeLong(Func<int> reader)
        {
            try
            {
                return reader();
            }
            catch (COMException)
            {
                return 0;
            }
        }

        private static string SafeDate(Func<DateTime> reader)
        {
            try
            {
                var value = reader();
                return value == DateTime.MinValue
                    ? null
                    : value.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
            }
            catch (COMException)
            {
                return null;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static string Limit(string value, int maxCharacters)
        {
            return value != null && value.Length > maxCharacters
                ? value.Substring(0, maxCharacters)
                : value ?? string.Empty;
        }

        private static string SanitizeFileName(string value)
        {
            var source = string.IsNullOrWhiteSpace(value) ? "attachment.bin" : value;
            var invalid = new HashSet<char>(Path.GetInvalidFileNameChars());
            var cleaned = new string(source.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
            if (cleaned.Length == 0)
            {
                cleaned = "attachment.bin";
            }

            return cleaned.Length <= 180 ? cleaned : cleaned.Substring(cleaned.Length - 180);
        }

        private static string GuessContentType(string fileName)
        {
            switch ((Path.GetExtension(fileName) ?? string.Empty).ToLowerInvariant())
            {
                case ".txt": return "text/plain";
                case ".md": return "text/markdown";
                case ".csv": return "text/csv";
                case ".htm":
                case ".html": return "text/html";
                case ".pdf": return "application/pdf";
                case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                default: return "application/octet-stream";
            }
        }
    }
}
