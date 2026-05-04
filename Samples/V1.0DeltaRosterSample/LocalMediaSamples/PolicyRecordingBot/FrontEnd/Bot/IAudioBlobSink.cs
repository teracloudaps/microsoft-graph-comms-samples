// <copyright file="IAudioBlobSink.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
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
}