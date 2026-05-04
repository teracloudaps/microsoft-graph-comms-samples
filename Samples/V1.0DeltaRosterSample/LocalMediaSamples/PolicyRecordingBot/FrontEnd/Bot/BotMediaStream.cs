// --------------------------------------------------------------------------------------------------------------------
// <copyright file="BotMediaStream.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>
// <summary>
//   The bot media stream.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Skype.Bots.Media;
    using Microsoft.Skype.Internal.Media.Services.Common;

    /// <summary>
    /// Class responsible for streaming audio and video.
    /// </summary>
    public class BotMediaStream : ObjectRootDisposable
    {
        private readonly IAudioSocket audioSocket;
        private readonly IVideoSocket vbssSocket;
        private readonly List<IVideoSocket> videoSockets;
        private readonly ILocalMediaSession mediaSession;
        private readonly CallHandler callHandler;
        private readonly IAudioBlobSink audioBlobSink;
        private readonly ConcurrentDictionary<uint, long> audioSequenceNumbers = new ConcurrentDictionary<uint, long>();
        private long mixedAudioSequenceNumber;

        /// <summary>
        /// Initializes a new instance of the <see cref="BotMediaStream"/> class.
        /// </summary>
        /// <param name="mediaSession">The media session.</param>
        /// <param name="callHandler">The call handler.</param>
        /// <param name="logger">Graph logger.</param>
        /// <param name="audioBlobSink">Audio blob sink.</param>
        /// <exception cref="InvalidOperationException">Throws when no audio socket is passed in.</exception>
        internal BotMediaStream(ILocalMediaSession mediaSession, CallHandler callHandler, IGraphLogger logger, IAudioBlobSink audioBlobSink = null)
            : base(logger)
        {
            ArgumentVerifier.ThrowOnNullArgument(mediaSession, nameof(mediaSession));
            ArgumentVerifier.ThrowOnNullArgument(callHandler, nameof(callHandler));
            ArgumentVerifier.ThrowOnNullArgument(logger, nameof(logger));

            this.mediaSession = mediaSession;
            this.callHandler = callHandler;
            this.audioBlobSink = audioBlobSink ?? NullAudioBlobSink.Instance;

            // Subscribe to the audio media.
            this.audioSocket = mediaSession.AudioSocket;
            if (this.audioSocket == null)
            {
                throw new InvalidOperationException("A mediaSession needs to have at least an audioSocket");
            }

            this.audioSocket.AudioMediaReceived += this.OnAudioMediaReceived;

            // Subscribe to the video media.
            this.videoSockets = this.mediaSession.VideoSockets?.ToList();
            if (this.videoSockets?.Any() == true)
            {
                this.videoSockets.ForEach(videoSocket => videoSocket.VideoMediaReceived += this.OnVideoMediaReceived);
            }

            // Subscribe to the VBSS media.
            this.vbssSocket = this.mediaSession.VbssSocket;
            if (this.vbssSocket != null)
            {
                this.mediaSession.VbssSocket.VideoMediaReceived += this.OnVbssMediaReceived;
            }
        }

        /// <summary>
        /// Subscription for video and vbss.
        /// </summary>
        /// <param name="mediaType">vbss or video.</param>
        /// <param name="mediaSourceId">The video source Id.</param>
        /// <param name="videoResolution">The preferred video resolution.</param>
        /// <param name="socketId">Socket id requesting the video. For vbss it is always 0.</param>
        public void Subscribe(MediaType mediaType, uint mediaSourceId, VideoResolution videoResolution, uint socketId = 0)
        {
            try
            {
                this.ValidateSubscriptionMediaType(mediaType);

                this.GraphLogger.Info($"Subscribing to the video source: {mediaSourceId} on socket: {socketId} with the preferred resolution: {videoResolution} and mediaType: {mediaType}");
                if (mediaType == MediaType.Vbss)
                {
                    if (this.vbssSocket == null)
                    {
                        this.GraphLogger.Warn($"vbss socket not initialized");
                    }
                    else
                    {
                        this.vbssSocket.Subscribe(videoResolution, mediaSourceId);
                    }
                }
                else if (mediaType == MediaType.Video)
                {
                    if (this.videoSockets == null)
                    {
                        this.GraphLogger.Warn($"video sockets were not created");
                    }
                    else
                    {
                        this.videoSockets[(int)socketId].Subscribe(videoResolution, mediaSourceId);
                    }
                }
            }
            catch (Exception ex)
            {
                this.GraphLogger.Error(ex, $"Video Subscription failed for the socket: {socketId} and MediaSourceId: {mediaSourceId} with exception");
            }
        }

        /// <summary>
        /// Unsubscribe to video.
        /// </summary>
        /// <param name="mediaType">vbss or video.</param>
        /// <param name="socketId">Socket id. For vbss it is always 0.</param>
        public void Unsubscribe(MediaType mediaType, uint socketId = 0)
        {
            try
            {
                this.ValidateSubscriptionMediaType(mediaType);

                this.GraphLogger.Info($"Unsubscribing to video for the socket: {socketId} and mediaType: {mediaType}");

                if (mediaType == MediaType.Vbss)
                {
                    this.vbssSocket?.Unsubscribe();
                }
                else if (mediaType == MediaType.Video)
                {
                    this.videoSockets[(int)socketId]?.Unsubscribe();
                }
            }
            catch (Exception ex)
            {
                this.GraphLogger.Error(ex, $"Unsubscribing to video failed for the socket: {socketId} with exception");
            }
        }

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            this.audioSocket.AudioMediaReceived -= this.OnAudioMediaReceived;

            if (this.videoSockets?.Any() == true)
            {
                this.videoSockets.ForEach(videoSocket => videoSocket.VideoMediaReceived -= this.OnVideoMediaReceived);
            }

            // Subscribe to the VBSS media.
            if (this.vbssSocket != null)
            {
                this.mediaSession.VbssSocket.VideoMediaReceived -= this.OnVbssMediaReceived;
            }
        }

        /// <summary>
        /// Ensure media type is video or VBSS.
        /// </summary>
        /// <param name="mediaType">Media type to validate.</param>
        private void ValidateSubscriptionMediaType(MediaType mediaType)
        {
            if (mediaType != MediaType.Vbss && mediaType != MediaType.Video)
            {
                throw new ArgumentOutOfRangeException($"Invalid mediaType: {mediaType}");
            }
        }

        /// <summary>
        /// Receive audio from subscribed participant.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The audio media received arguments.
        /// </param>
        private void OnAudioMediaReceived(object sender, AudioMediaReceivedEventArgs e)
        {
            try
            {
                if (!this.audioBlobSink.IsEnabled)
                {
                    return;
                }

                this.PublishIdentifiedAudioBlobs(e.Buffer);
            }
            catch (Exception ex)
            {
                this.GraphLogger.Error(ex, "Failed to process received audio buffer.");
            }
            finally
            {
                e.Buffer.Dispose();
            }
        }

        /// <summary>
        /// Publishes identity-enriched audio blobs from a received media buffer.
        /// </summary>
        /// <param name="audioBuffer">The received audio buffer.</param>
        private void PublishIdentifiedAudioBlobs(AudioMediaBuffer audioBuffer)
        {
            var callId = this.callHandler.Call.Id;
            var receivedAtUtc = DateTimeOffset.UtcNow;
            var audioFormat = audioBuffer.AudioFormat.ToString();
            var unmixedAudioBuffers = audioBuffer.UnmixedAudioBuffers;

            if (unmixedAudioBuffers != null)
            {
                foreach (var unmixedBuffer in unmixedAudioBuffers)
                {
                    var mediaSourceId = unmixedBuffer.ActiveSpeakerId;
                    var sequenceNumber = this.GetNextSequenceNumber(mediaSourceId);
                    var streamId = CreateStreamId(callId, mediaSourceId.ToString());
                    if (!this.audioBlobSink.CanAccept)
                    {
                        continue;
                    }

                    this.audioBlobSink.TryPublish(new IdentifiedAudioBlob
                    {
                        BlobId = CreateBlobId(streamId, sequenceNumber),
                        CallId = callId,
                        StreamId = streamId,
                        SequenceNumber = sequenceNumber,
                        ReceivedAtUtc = receivedAtUtc,
                        MediaSourceId = mediaSourceId.ToString(),
                        MediaTimestamp = audioBuffer.Timestamp,
                        OriginalSenderTimestamp = unmixedBuffer.OriginalSenderTimestamp,
                        IsMixed = false,
                        IsSilence = audioBuffer.IsSilence,
                        Length = unmixedBuffer.Length,
                        AudioFormat = audioFormat,
                        Buffer = CopyBuffer(unmixedBuffer.Data, unmixedBuffer.Length),
                        Identity = this.callHandler.GetParticipantIdentityMetadata(mediaSourceId),
                    });
                }
            }

            if (this.audioBlobSink.IncludeMixedAudioBuffer)
            {
                var mediaSourceId = "mixed";
                var sequenceNumber = Interlocked.Increment(ref this.mixedAudioSequenceNumber);
                var streamId = CreateStreamId(callId, mediaSourceId);
                if (!this.audioBlobSink.CanAccept)
                {
                    return;
                }

                this.audioBlobSink.TryPublish(new IdentifiedAudioBlob
                {
                    BlobId = CreateBlobId(streamId, sequenceNumber),
                    CallId = callId,
                    StreamId = streamId,
                    SequenceNumber = sequenceNumber,
                    ReceivedAtUtc = receivedAtUtc,
                    MediaSourceId = mediaSourceId,
                    MediaTimestamp = audioBuffer.Timestamp,
                    IsMixed = true,
                    IsSilence = audioBuffer.IsSilence,
                    Length = audioBuffer.Length,
                    AudioFormat = audioFormat,
                    Buffer = CopyBuffer(audioBuffer.Data, audioBuffer.Length),
                });
            }
        }

        private long GetNextSequenceNumber(uint mediaSourceId)
        {
            return this.audioSequenceNumbers.AddOrUpdate(mediaSourceId, 1, (key, currentValue) => currentValue + 1);
        }

        private static string CreateStreamId(string callId, string mediaSourceId)
        {
            return $"{callId}:{mediaSourceId}";
        }

        private static string CreateBlobId(string streamId, long sequenceNumber)
        {
            return $"{streamId}:{sequenceNumber:D20}";
        }

        /// <summary>
        /// Copies unmanaged media buffer data into managed memory.
        /// </summary>
        /// <param name="data">The unmanaged buffer pointer.</param>
        /// <param name="length">The buffer length.</param>
        /// <returns>The copied buffer, or null when no data is present.</returns>
        private static byte[] CopyBuffer(IntPtr data, long length)
        {
            if (data == IntPtr.Zero || length <= 0)
            {
                return null;
            }

            var buffer = new byte[checked((int)length)];
            Marshal.Copy(data, buffer, 0, buffer.Length);
            return buffer;
        }

        /// <summary>
        /// Receive video from subscribed participant.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The video media received arguments.
        /// </param>
        private void OnVideoMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            this.GraphLogger.Verbose($"[{e.SocketId}]: Received Video: [VideoMediaReceivedEventArgs(Data=<{e.Buffer.Data.ToString()}>, Length={e.Buffer.Length}, Timestamp={e.Buffer.Timestamp}, Width={e.Buffer.VideoFormat.Width}, Height={e.Buffer.VideoFormat.Height}, ColorFormat={e.Buffer.VideoFormat.VideoColorFormat}, FrameRate={e.Buffer.VideoFormat.FrameRate})]");

            // TBD: Policy Recording bots can record the Video here
            e.Buffer.Dispose();
        }

        /// <summary>
        /// Receive vbss from subscribed participant.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The video media received arguments.
        /// </param>
        private void OnVbssMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            this.GraphLogger.Verbose($"[{e.SocketId}]: Received VBSS: [VideoMediaReceivedEventArgs(Data=<{e.Buffer.Data.ToString()}>, Length={e.Buffer.Length}, Timestamp={e.Buffer.Timestamp}, Width={e.Buffer.VideoFormat.Width}, Height={e.Buffer.VideoFormat.Height}, ColorFormat={e.Buffer.VideoFormat.VideoColorFormat}, FrameRate={e.Buffer.VideoFormat.FrameRate})]");

            // TBD: Policy Recording bots can record the VBSS here
            e.Buffer.Dispose();
        }
    }
}
