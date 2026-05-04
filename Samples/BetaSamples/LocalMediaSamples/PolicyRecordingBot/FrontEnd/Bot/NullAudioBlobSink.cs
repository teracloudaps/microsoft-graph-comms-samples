// <copyright file="NullAudioBlobSink.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    /// <summary>
    /// Default sink used when no downstream publisher is configured.
    /// </summary>
    internal class NullAudioBlobSink : IAudioBlobSink
    {
        /// <summary>
        /// The shared null sink instance.
        /// </summary>
        public static readonly IAudioBlobSink Instance = new NullAudioBlobSink();

        /// <summary>
        /// Initializes a new instance of the <see cref="NullAudioBlobSink"/> class.
        /// </summary>
        private NullAudioBlobSink()
        {
        }

        /// <inheritdoc/>
        public bool IsEnabled => false;

        /// <inheritdoc/>
        public bool IncludeMixedAudioBuffer => false;

        /// <inheritdoc/>
        public bool CanAccept => false;

        /// <inheritdoc/>
        public bool TryPublish(IdentifiedAudioBlob audioBlob)
        {
            return true;
        }
    }
}
