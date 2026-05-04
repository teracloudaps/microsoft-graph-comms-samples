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
        /// Gets or sets the participant identity metadata.
        /// </summary>
        public ParticipantIdentityMetadata Identity { get; set; }
    }

    /// <summary>
    /// Participant identity metadata emitted with audio blobs.
    /// </summary>
    public class ParticipantIdentityMetadata
    {
        /// <summary>
        /// Gets or sets the Graph participant identifier.
        /// </summary>
        public string ParticipantId { get; set; }

        /// <summary>
        /// Gets or sets the media source identifier used to map audio to the participant.
        /// </summary>
        public string MediaSourceId { get; set; }

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
    }

    /// <summary>
    /// Receives identity-enriched audio blobs.
    /// </summary>
    public interface IAudioBlobSink
    {
        /// <summary>
        /// Gets a value indicating whether audio blobs should be built and published.
        /// </summary>
        bool IsEnabled { get; }

        /// <summary>
        /// Gets a value indicating whether the mixed audio buffer should be copied into the payload.
        /// </summary>
        bool IncludeMixedAudioBuffer { get; }

        /// <summary>
        /// Gets a value indicating whether the sink can currently accept another blob without blocking.
        /// </summary>
        bool CanAccept { get; }

        /// <summary>
        /// Tries to publish an audio blob without blocking the media callback.
        /// </summary>
        /// <param name="audioBlob">The audio blob to publish.</param>
        /// <returns>True when the blob was accepted; otherwise, false.</returns>
        bool TryPublish(IdentifiedAudioBlob audioBlob);
    }

    /// <summary>
    /// Default sink used when no downstream publisher is configured.
    /// </summary>
    internal class NullAudioBlobSink : IAudioBlobSink
    {
        /// <summary>
        /// Gets the shared null sink instance.
        /// </summary>
        public static readonly IAudioBlobSink Instance = new NullAudioBlobSink();

        /// <inheritdoc/>
        public bool IsEnabled => false;

        /// <inheritdoc/>
        public bool IncludeMixedAudioBuffer => false;

        /// <inheritdoc/>
        public bool CanAccept => false;

        private NullAudioBlobSink()
        {
        }

        /// <inheritdoc/>
        public bool TryPublish(IdentifiedAudioBlob audioBlob)
        {
            return true;
        }
    }
}