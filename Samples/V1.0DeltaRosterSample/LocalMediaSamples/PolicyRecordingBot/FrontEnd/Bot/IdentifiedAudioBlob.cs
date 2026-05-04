// <copyright file="IdentifiedAudioBlob.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    using System;

    /// <summary>
    /// Represents a publishable audio blob with participant identity and ordering metadata attached.
    /// </summary>
    public class IdentifiedAudioBlob
    {
        /// <summary>
        /// Gets or sets the unique blob identifier.
        /// </summary>
        public string BlobId { get; set; }

        /// <summary>
        /// Gets or sets the call identifier.
        /// </summary>
        public string CallId { get; set; }

        /// <summary>
        /// Gets or sets the stable stream identifier used to partition audio for STT.
        /// </summary>
        public string StreamId { get; set; }

        /// <summary>
        /// Gets or sets the sequence number within the stream.
        /// </summary>
        public long SequenceNumber { get; set; }

        /// <summary>
        /// Gets or sets the UTC time when the bot received the media buffer.
        /// </summary>
        public DateTimeOffset ReceivedAtUtc { get; set; }

        /// <summary>
        /// Gets or sets the media source identifier.
        /// </summary>
        public string MediaSourceId { get; set; }

        /// <summary>
        /// Gets or sets the media timestamp.
        /// </summary>
        public long MediaTimestamp { get; set; }

        /// <summary>
        /// Gets or sets the original sender timestamp for unmixed participant audio.
        /// </summary>
        public long? OriginalSenderTimestamp { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the blob is mixed call audio rather than one participant's unmixed audio.
        /// </summary>
        public bool IsMixed { get; set; }

        /// <summary>
        /// Gets or sets the buffer length.
        /// </summary>
        public long Length { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the buffer contains silence.
        /// </summary>
        public bool IsSilence { get; set; }

        /// <summary>
        /// Gets or sets the audio format.
        /// </summary>
        public string AudioFormat { get; set; }

        /// <summary>
        /// Gets or sets the copied audio buffer.
        /// </summary>
        public byte[] Buffer { get; set; }

        /// <summary>
        /// Gets or sets the participant user identifier, when available.
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// Gets or sets the participant display name, when available.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets the participant tenant identifier, when available.
        /// </summary>
        public string ParticipantTenantId { get; set; }

        /// <summary>
        /// Gets or sets the configured org or tenant identifier supplied by the bot operator.
        /// </summary>
        public string ConfiguredOrgId { get; set; }

        /// <summary>
        /// Gets or sets the identity type, such as User, AdditionalData, or Unknown.
        /// </summary>
        public string IdentityType { get; set; }

        /// <summary>
        /// Gets or sets the participant identity metadata.
        /// </summary>
        public ParticipantIdentityMetadata Identity { get; set; }
    }

}