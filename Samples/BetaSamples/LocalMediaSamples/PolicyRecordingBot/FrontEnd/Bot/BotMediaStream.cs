// <copyright file="BotMediaStream.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

#pragma warning disable SA1600 // Recovered sample internals are intentionally kept close to the upstream PoC shape.

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Sockets;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Timers;
    using Microsoft.Graph.Beta.Models;
    using Microsoft.Graph.Communications.Calls;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Graph.Communications.Resources;
    using Microsoft.Skype.Bots.Media;
    using Microsoft.Skype.Internal.Media.Services.Common;

    /// <summary>
    /// Receives audio frames and participant updates from the Teams media SDK and relays
    /// them to local TCP consumers. The bot performs no business classification (agent
    /// vs. caller, internal vs. external, etc.) — it forwards raw identity data only,
    /// and the consumer decides how to route.
    /// </summary>
    internal class BotMediaStream : ObjectRootDisposable
    {
        private const int ENERGYHISTORYSIZE = 10;
        private const double SILENCEENERGYTHRESHOLD = 100.0;
        private const double SPEAKERPERSISTENCEMS = 1000;
        private const double SPEAKERCHANGEDEBOUNCEMS = 500;
        private const int MAXQUEUEDAUDIOFRAMES = 1000;
        private const int MAXQUEUEDMETADATAMESSAGES = 100;
        private const int IAB1HEADERLEN = 12;
        private const int AUDIOSTREAMPORT = 5001;
        private const int METADATASTREAMPORT = 5002;

        // Attribution key used in the pre-roll buffer for mixed-mode frames (which have no
        // per-MSI key). uint.MaxValue is outside the SDK's MSI range so it cannot collide with
        // a real unmixed-mode MSI key.
        private const uint MIXEDMODEATTRIBUTIONKEY = uint.MaxValue;

        // Bounds memory growth of the per-key pre-roll buffer and the size of the drain
        // burst at resolution. ~5 seconds at 50fps — well past typical SDK resolution
        // latency (a few hundred ms after speech onset) and aligned with SPEAKERPERSISTENCEMS
        // (which already covers brief mid-call gaps). Larger values create GC + audioFrameQueue
        // pressure when the buffer drains in one go.
        private const int MAXPREROLLFRAMESPERKEY = 250;

        private static readonly byte[] Iab1Magic = { 0x49, 0x41, 0x42, 0x31 }; // "IAB1"
        private static readonly char[] HexDigits = "0123456789abcdef".ToCharArray();
        private static readonly Func<uint, Queue<byte[]>> CreatePreRollQueue = _ => new Queue<byte[]>();

        private readonly IAudioSocket audioSocket;
        private readonly IVideoSocket vbssSocket;
        private readonly List<IVideoSocket> videoSockets;
        private readonly ILocalMediaSession mediaSession;
        private readonly ICall call;
        private readonly BlockingCollection<byte[]> audioFrameQueue = new BlockingCollection<byte[]>(MAXQUEUEDAUDIOFRAMES);
        private readonly BlockingCollection<byte[]> metadataMessageQueue = new BlockingCollection<byte[]>(MAXQUEUEDMETADATAMESSAGES);
        private readonly ConcurrentDictionary<TcpClient, byte> connectedClients = new ConcurrentDictionary<TcpClient, byte>();
        private readonly ConcurrentDictionary<TcpClient, byte> connectedMetadataClients = new ConcurrentDictionary<TcpClient, byte>();
        private readonly ConcurrentDictionary<string, ParticipantInfo> participantsById = new ConcurrentDictionary<string, ParticipantInfo>();
        private readonly ConcurrentDictionary<uint, string> msiToParticipantId = new ConcurrentDictionary<uint, string>();
        private readonly ConcurrentDictionary<string, string> mergedParticipantTarget = new ConcurrentDictionary<string, string>();
        private readonly string currentSessionId = Guid.NewGuid().ToString();
        private readonly Queue<double> recentAudioEnergy = new Queue<double>();
        private readonly CancellationTokenSource shutdownCts = new CancellationTokenSource();

        // Audio frames whose attribution is not yet resolved, keyed by MSI (unmixed mode) or
        // MIXEDMODEATTRIBUTIONKEY (mixed mode). Drained to the resolving participant on the
        // next attributed emit. No sticky "resolved" state — mid-call binding races (new
        // speaker's MSI arrives via DominantSpeakerChanged but isn't in msiToParticipantId yet)
        // re-buffer and re-drain naturally, so utterance onsets are preserved.
        private readonly ConcurrentDictionary<uint, Queue<byte[]>> preRollAudioByKey = new ConcurrentDictionary<uint, Queue<byte[]>>();

        private TcpListener audioStreamServer;
        private TcpListener metadataStreamServer;
        private int nextChannelIndex; // Atomic; monotonic; never reused after a participant leaves
        private volatile ParticipantInfo cachedSdkUnboundCaller; // Non-null only when exactly one IsSdkUnbound real caller exists; lets the audio hot path adopt unknown MSIs without a LINQ scan per frame.
        private uint? currentDominantSpeakerMsi;
        private DateTime lastSpeakerChangeTime = DateTime.UtcNow;
        private string lastKnownSpeakerId;
        private DateTime lastKnownSpeakerTime = DateTime.UtcNow;
        private long mixedPacketCount;
        private long unmixedPacketCount;
        private long silencePacketCount;
        private long droppedAudioFrameCount;
        private long droppedMetadataMessageCount;
        private DateTime lastStatsLog = DateTime.UtcNow;
        private bool unmixedAudioEnabled;
        private bool detailedLogging = false;
        private System.Timers.Timer participantCheckTimer;
        private int participantCheckCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="BotMediaStream"/> class.
        /// </summary>
        /// <param name="mediaSession">The local media session.</param>
        /// <param name="logger">The graph logger.</param>
        /// <param name="call">The active call.</param>
        public BotMediaStream(ILocalMediaSession mediaSession, IGraphLogger logger, ICall call)
            : base(logger)
        {
            ArgumentVerifier.ThrowOnNullArgument(mediaSession, nameof(mediaSession));
            ArgumentVerifier.ThrowOnNullArgument(logger, nameof(logger));
            ArgumentVerifier.ThrowOnNullArgument(call, nameof(call));

            this.mediaSession = mediaSession;
            this.call = call;

            this.audioSocket = mediaSession.AudioSocket;
            if (this.audioSocket == null)
            {
                throw new InvalidOperationException("A mediaSession needs to have at least an audioSocket");
            }

            this.audioSocket.AudioMediaReceived += this.OnAudioMediaReceived;
            this.audioSocket.DominantSpeakerChanged += this.OnDominantSpeakerChanged;
            Console.WriteLine("Subscribed to AudioMediaReceived and DominantSpeakerChanged events");

            this.videoSockets = this.mediaSession.VideoSockets?.ToList();
            if (this.videoSockets?.Any() == true)
            {
                this.videoSockets.ForEach(videoSocket => videoSocket.VideoMediaReceived += this.OnVideoMediaReceived);
            }

            this.vbssSocket = this.mediaSession.VbssSocket;
            if (this.vbssSocket != null)
            {
                this.mediaSession.VbssSocket.VideoMediaReceived += this.OnVbssMediaReceived;
            }

            this.call.Participants.OnUpdated += this.OnParticipantsUpdated;

            Console.WriteLine($"Initial participants: {this.call.Participants.Count}");
            foreach (var participant in this.call.Participants)
            {
                this.ProcessParticipant(participant, isNew: false);
            }

            this.participantCheckTimer = new System.Timers.Timer(1000) { AutoReset = true };
            this.participantCheckTimer.Elapsed += this.CheckForParticipants;
            this.participantCheckTimer.Start();
            Console.WriteLine("Started participant polling timer");

            this.StartAudioStreamServer();
            this.StartMetadataStreamServer();
        }

        /// <summary>
        /// Registers a participant discovered by the call handler.
        /// </summary>
        /// <param name="participant">The participant to register.</param>
        public void RegisterParticipantFromCallHandler(IParticipant participant)
        {
            try
            {
                Console.WriteLine($"BotMediaStream received participant from CallHandler: {participant.Id}");

                if (!this.participantsById.ContainsKey(participant.Id))
                {
                    this.ProcessParticipant(participant, isNew: true);
                }
                else
                {
                    Console.WriteLine("Participant already tracked");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering participant from CallHandler: {ex.Message}");
            }
        }

        /// <summary>
        /// Subscribes a video socket to a media source.
        /// </summary>
        /// <param name="mediaType">The media type to subscribe.</param>
        /// <param name="mediaSourceId">The media source id.</param>
        /// <param name="videoResolution">The requested video resolution.</param>
        /// <param name="socketId">The socket id to use for video subscriptions.</param>
        public void Subscribe(MediaType mediaType, uint mediaSourceId, VideoResolution videoResolution, uint socketId = 0)
        {
            try
            {
                this.ValidateSubscriptionMediaType(mediaType);
                if (mediaType == MediaType.Vbss)
                {
                    this.vbssSocket?.Subscribe(videoResolution, mediaSourceId);
                }
                else if (mediaType == MediaType.Video)
                {
                    this.videoSockets?[(int)socketId]?.Subscribe(videoResolution, mediaSourceId);
                }
            }
            catch (Exception ex)
            {
                this.GraphLogger.Error(ex, "Video subscription failed");
            }
        }

        /// <summary>
        /// Unsubscribes a video socket from its current media source.
        /// </summary>
        /// <param name="mediaType">The media type to unsubscribe.</param>
        /// <param name="socketId">The socket id to use for video subscriptions.</param>
        public void Unsubscribe(MediaType mediaType, uint socketId = 0)
        {
            try
            {
                this.ValidateSubscriptionMediaType(mediaType);
                if (mediaType == MediaType.Vbss)
                {
                    this.vbssSocket?.Unsubscribe();
                }
                else if (mediaType == MediaType.Video)
                {
                    this.videoSockets?[(int)socketId]?.Unsubscribe();
                }
            }
            catch (Exception ex)
            {
                this.GraphLogger.Error(ex, "Unsubscribing failed");
            }
        }

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            this.shutdownCts.Cancel();

            if (this.participantCheckTimer != null)
            {
                this.participantCheckTimer.Stop();
                this.participantCheckTimer.Elapsed -= this.CheckForParticipants;
                this.participantCheckTimer.Dispose();
            }

            this.audioSocket.AudioMediaReceived -= this.OnAudioMediaReceived;
            this.audioSocket.DominantSpeakerChanged -= this.OnDominantSpeakerChanged;

            if (this.call != null)
            {
                this.call.Participants.OnUpdated -= this.OnParticipantsUpdated;
            }

            if (this.videoSockets?.Any() == true)
            {
                this.videoSockets.ForEach(videoSocket => videoSocket.VideoMediaReceived -= this.OnVideoMediaReceived);
            }

            if (this.vbssSocket != null)
            {
                this.vbssSocket.VideoMediaReceived -= this.OnVbssMediaReceived;
            }

            this.StopAudioStreamServer();
            this.StopMetadataStreamServer();
            this.shutdownCts.Dispose();
        }

        private static byte[] BuildFrameMetadata(string sessionId, ParticipantInfo info)
        {
            var sb = new StringBuilder(192);
            sb.Append('{');
            AppendJsonString(sb, "callId", sessionId);
            sb.Append(',');
            AppendJsonString(sb, "streamId", info.ParticipantId);
            sb.Append(',');
            AppendJsonString(sb, "participantId", info.ParticipantId);
            sb.Append(',');
            AppendJsonString(sb, "userId", info.UserId ?? string.Empty);
            sb.Append(',');
            AppendJsonString(sb, "displayName", info.DisplayName ?? string.Empty);
            sb.Append(',');
            AppendJsonString(sb, "tenantId", info.TenantId ?? string.Empty);
            sb.Append(',');
            sb.Append("\"channelId\":").Append(info.ChannelId);
            sb.Append(',');
            sb.Append("\"isMsiFallback\":").Append(info.IsMsiFallback ? "true" : "false");
            sb.Append('}');
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static string BuildParticipantEventJson(string sessionId, ParticipantInfo info, string eventType)
        {
            var sb = new StringBuilder(256);
            sb.Append('{');
            AppendJsonString(sb, "eventType", eventType);
            sb.Append(',');
            AppendJsonString(sb, "sessionId", sessionId);
            sb.Append(',');
            AppendJsonString(sb, "timestamp", DateTime.UtcNow.ToString("o"));
            sb.Append(',');
            AppendJsonString(sb, "participantId", info.ParticipantId ?? string.Empty);
            sb.Append(',');
            AppendJsonString(sb, "userId", info.UserId ?? string.Empty);
            sb.Append(',');
            AppendJsonString(sb, "displayName", info.DisplayName ?? string.Empty);
            sb.Append(',');
            AppendJsonString(sb, "tenantId", info.TenantId ?? string.Empty);
            sb.Append(',');
            sb.Append("\"channelId\":").Append(info.ChannelId);
            sb.Append(',');
            sb.Append("\"isMsiFallback\":").Append(info.IsMsiFallback ? "true" : "false");
            sb.Append(',');
            sb.Append("\"mediaStreamIds\":[");
            var first = true;
            foreach (var msi in info.MediaStreamIds)
            {
                if (!first)
                {
                    sb.Append(',');
                }

                sb.Append(msi);
                first = false;
            }

            sb.Append(']');
            sb.Append('}');
            return sb.ToString();
        }

        private static void AppendJsonString(StringBuilder sb, string name, string value)
        {
            sb.Append('"').Append(name).Append("\":");
            AppendEscapedJsonString(sb, value);
        }

        private static void AppendEscapedJsonString(StringBuilder sb, string value)
        {
            sb.Append('"');
            if (!string.IsNullOrEmpty(value))
            {
                for (int i = 0; i < value.Length; i++)
                {
                    var c = value[i];
                    switch (c)
                    {
                        case '\\': sb.Append("\\\\"); break;
                        case '"': sb.Append("\\\""); break;
                        case '\b': sb.Append("\\b"); break;
                        case '\f': sb.Append("\\f"); break;
                        case '\n': sb.Append("\\n"); break;
                        case '\r': sb.Append("\\r"); break;
                        case '\t': sb.Append("\\t"); break;
                        default:
                            if (c < 0x20)
                            {
                                sb.Append("\\u00")
                                  .Append(HexDigits[(c >> 4) & 0xF])
                                  .Append(HexDigits[c & 0xF]);
                            }
                            else
                            {
                                sb.Append(c);
                            }

                            break;
                    }
                }
            }

            sb.Append('"');
        }

        private void CheckForParticipants(object sender, ElapsedEventArgs e)
        {
            try
            {
                this.participantCheckCount++;

                if (this.call?.Participants != null && this.participantsById.IsEmpty)
                {
                    try
                    {
                        foreach (var participant in this.call.Participants)
                        {
                            if (!this.participantsById.ContainsKey(participant.Id))
                            {
                                this.ProcessParticipant(participant, isNew: true);
                            }
                        }
                    }
                    catch (Exception enumEx)
                    {
                        if (this.participantCheckCount <= 5)
                        {
                            Console.WriteLine($"Cannot enumerate participants: {enumEx.Message}");
                        }
                    }
                }

                if (!this.participantsById.IsEmpty || this.participantCheckCount > 30)
                {
                    if (!this.participantsById.IsEmpty)
                    {
                        Console.WriteLine($"Successfully tracked {this.participantsById.Count} participant(s) - stopping polling");
                    }
                    else
                    {
                        Console.WriteLine("No participants found after 30 checks - relying on MSI fallback");
                    }

                    this.participantCheckTimer?.Stop();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in CheckForParticipants: {ex.Message}");
            }
        }

        private void OnParticipantsUpdated(IParticipantCollection sender, CollectionEventArgs<IParticipant> args)
        {
            try
            {
                foreach (var participant in args.AddedResources)
                {
                    if (this.participantsById.ContainsKey(participant.Id))
                    {
                        Console.WriteLine($"Participant already tracked in update event: {participant.Id}");
                        continue;
                    }

                    Console.WriteLine("Participant JOINED");
                    this.ProcessParticipant(participant, isNew: true);
                }

                bool removedAny = false;
                foreach (var participant in args.RemovedResources)
                {
                    var participantId = participant.Id;
                    if (this.participantsById.TryRemove(participantId, out var info))
                    {
                        foreach (var msi in info.MediaStreamIds)
                        {
                            this.msiToParticipantId.TryRemove(msi, out _);
                        }

                        Console.WriteLine($"Participant LEFT: {info.DisplayName} (CH{info.ChannelId})");
                        this.SendParticipantEvent(info, "LEAVE");

                        if (this.lastKnownSpeakerId == participantId)
                        {
                            this.lastKnownSpeakerId = null;
                        }

                        removedAny = true;
                    }
                }

                if (removedAny)
                {
                    this.RefreshCachedSdkUnboundCaller();
                }

                Console.WriteLine($"Total participants: {this.call.Participants.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in OnParticipantsUpdated: {ex.Message}");
            }
        }

        private void ProcessParticipant(IParticipant participant, bool isNew)
        {
            try
            {
                var participantId = participant.Id;

                if (this.mergedParticipantTarget.ContainsKey(participantId))
                {
                    return;
                }

                var identity = participant.Resource?.Info?.Identity;
                if (identity == null)
                {
                    Console.WriteLine($" Participant {participantId} has no identity");
                    return;
                }

                if (identity.Application != null)
                {
                    Console.WriteLine(" Skipping bot participant");
                    return;
                }

                string displayName;
                string userId = null;
                string tenantId = null;

                if (identity.User != null)
                {
                    displayName = identity.User.DisplayName ?? "Teams User";
                    userId = identity.User.Id;
                    if (identity.User.AdditionalData != null &&
                        identity.User.AdditionalData.TryGetValue("tenantId", out var tenantObj))
                    {
                        tenantId = tenantObj?.ToString()?.Trim();
                    }
                }
                else
                {
                    displayName = "External Caller";
                }

                var mediaStreamIds = new List<uint>();
                if (participant.Resource.MediaStreams != null)
                {
                    foreach (var stream in participant.Resource.MediaStreams)
                    {
                        if (stream.MediaType != Modality.Audio || !uint.TryParse(stream.SourceId, out uint msi))
                        {
                            continue;
                        }

                        // Only the participant's own voice (Send/SendReceive). ReceiveOnly is
                        // what this participant is hearing from others — attributing it to them
                        // would route incoming audio to the wrong channel.
                        if (stream.Direction != MediaDirection.SendOnly &&
                            stream.Direction != MediaDirection.SendReceive)
                        {
                            Console.WriteLine($"  Identity '{displayName}' skipped audio MSI={msi} direction={stream.Direction}");
                            continue;
                        }

                        mediaStreamIds.Add(msi);
                        Console.WriteLine($"  Identity '{displayName}' audio MSI={msi} direction={stream.Direction}");
                    }
                }

                // Capture SDK exposure now, before placeholder claiming mutates mediaStreamIds.
                // Sticky: a participant who had no Send/SendReceive MediaStream from the SDK
                // remains the absorber for unbound MSIs even after claiming a placeholder, so
                // a later distinct MSI for their voice routes to them rather than a new placeholder.
                bool isSdkUnbound = !string.IsNullOrWhiteSpace(userId) && mediaStreamIds.Count == 0;

                // If a real participant arrives with no MSI in MediaStreams, claim ALL outstanding
                // MSI-fallback placeholders. This is the reverse of the dominant-speaker adoption
                // path and covers the timing where audio arrives first (one or more placeholders
                // created), then identity resolves. Aggressive (all-placeholders) because in the
                // policy-bot scenario every unbound MSI belongs to the same single absorber.
                if (isSdkUnbound)
                {
                    var orphanPlaceholders = this.participantsById.Values
                        .Where(p => p.IsMsiFallback)
                        .OrderBy(p => p.JoinTime)
                        .ToList();

                    foreach (var orphanPlaceholder in orphanPlaceholders)
                    {
                        if (orphanPlaceholder.MediaStreamIds != null)
                        {
                            foreach (var msi in orphanPlaceholder.MediaStreamIds)
                            {
                                if (!mediaStreamIds.Contains(msi))
                                {
                                    mediaStreamIds.Add(msi);
                                }
                            }
                        }

                        if (this.participantsById.TryRemove(orphanPlaceholder.ParticipantId, out var evicted))
                        {
                            this.SendParticipantEvent(evicted, "LEAVE");
                            Console.WriteLine($"Claimed MSI placeholder '{evicted.DisplayName}' (was CH{evicted.ChannelId}) for incoming real caller '{displayName}'");
                        }
                    }
                }

                // Anonymous shadow participants (no userId) are merged into a real caller.
                // The Teams SDK frequently reports a real participant first (with no MSI in
                // MediaStreams) and then a separate anonymous shadow that mirrors them. The
                // shadow's most likely owner is the SDK-unbound real caller (the policy-attached
                // user) — even if that caller has already absorbed unbound MSIs, they remain
                // the right merge target. Fall back to the most-recently-joined real caller
                // only if no SDK-unbound caller exists. mergedParticipantTarget locks the
                // decision so a later OnParticipantsUpdated re-fire cannot flip the target.
                if (string.IsNullOrWhiteSpace(userId))
                {
                    var realCallers = this.participantsById.Values
                        .Where(p => !p.IsMsiFallback && !string.IsNullOrWhiteSpace(p.UserId))
                        .ToList();

                    var mergeTarget = realCallers
                        .Where(p => p.IsSdkUnbound)
                        .OrderBy(p => p.JoinTime)
                        .FirstOrDefault()
                        ?? realCallers
                            .OrderByDescending(p => p.JoinTime)
                            .FirstOrDefault();

                    if (mergeTarget != null)
                    {
                        var heuristic = mergeTarget.IsSdkUnbound
                            ? "sdk-unbound"
                            : "most-recent-fallback";

                        foreach (var msi in mediaStreamIds)
                        {
                            this.msiToParticipantId[msi] = mergeTarget.ParticipantId;
                            if (!mergeTarget.MediaStreamIds.Contains(msi))
                            {
                                mergeTarget.MediaStreamIds.Add(msi);
                            }
                        }

                        Console.WriteLine($" Merged anonymous participant '{participantId}' into '{mergeTarget.DisplayName}' (CH{mergeTarget.ChannelId}) via {heuristic}");
                        this.mergedParticipantTarget.TryAdd(participantId, mergeTarget.ParticipantId);
                        this.RefreshCachedSdkUnboundCaller();
                        return;
                    }
                }

                int channelId = Interlocked.Increment(ref this.nextChannelIndex) - 1;

                var info = new ParticipantInfo
                {
                    ParticipantId = participantId,
                    UserId = userId,
                    DisplayName = displayName,
                    TenantId = tenantId,
                    ChannelId = channelId,
                    IsSdkUnbound = isSdkUnbound,
                    MediaStreamIds = mediaStreamIds,
                    JoinTime = DateTime.UtcNow,
                };
                info.FrameMetadataBytes = BuildFrameMetadata(this.currentSessionId, info);

                // Re-bind MSIs to this real participant; evict any MSI-fallback placeholder
                // that previously held the same MSI.
                foreach (var msi in mediaStreamIds)
                {
                    if (this.msiToParticipantId.TryGetValue(msi, out var existingId) &&
                        existingId != participantId &&
                        this.participantsById.TryGetValue(existingId, out var existing) &&
                        existing.IsMsiFallback)
                    {
                        if (this.participantsById.TryRemove(existingId, out var evicted))
                        {
                            Console.WriteLine($"Evicted MSI fallback '{evicted.DisplayName}' (CH{evicted.ChannelId}) — superseded by '{displayName}'");
                            this.SendParticipantEvent(evicted, "LEAVE");
                        }
                    }

                    this.msiToParticipantId[msi] = participantId;
                    Console.WriteLine($"Mapped MSI {msi} to {participantId} (CH{channelId})");
                }

                this.participantsById[participantId] = info;
                this.SendParticipantEvent(info, "JOIN");
                this.RefreshCachedSdkUnboundCaller();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing participant: {ex.Message}");
            }
        }

        private void OnDominantSpeakerChanged(object sender, DominantSpeakerChangedEventArgs e)
        {
            try
            {
                var now = DateTime.UtcNow;
                var elapsed = (now - this.lastSpeakerChangeTime).TotalMilliseconds;
                if (elapsed < SPEAKERCHANGEDEBOUNCEMS && this.currentDominantSpeakerMsi.HasValue)
                {
                    return;
                }

                if (e.CurrentDominantSpeaker == DominantSpeakerChangedEventArgs.None)
                {
                    this.currentDominantSpeakerMsi = null;
                    return;
                }

                this.currentDominantSpeakerMsi = e.CurrentDominantSpeaker;
                this.lastSpeakerChangeTime = now;

                if (!this.msiToParticipantId.ContainsKey(e.CurrentDominantSpeaker))
                {
                    // The Teams SDK lights up audio MSIs via the dominant-speaker event for
                    // policy-attached users whose MediaStreams are never populated, and the
                    // SDK may use *different* MSIs over the call's lifetime for the same
                    // speaker (renegotiation, codec change, etc). Route every unbound MSI to
                    // the same SDK-unbound real caller (sticky by IsSdkUnbound, not by current
                    // MediaStreamIds — once we adopt MSI N, the caller is no longer "orphan"
                    // by stream count but still owns subsequent unbound MSIs).
                    //
                    // If 0 or 2+ SDK-unbound real callers exist the attribution is ambiguous —
                    // fall back to a placeholder and let the consumer resolve it.
                    var sdkUnboundCandidates = this.participantsById.Values
                        .Where(p => !p.IsMsiFallback && p.IsSdkUnbound && !string.IsNullOrWhiteSpace(p.UserId))
                        .ToList();

                    if (sdkUnboundCandidates.Count == 1)
                    {
                        var target = sdkUnboundCandidates[0];
                        if (target.MediaStreamIds == null)
                        {
                            target.MediaStreamIds = new List<uint>();
                        }

                        if (!target.MediaStreamIds.Contains(e.CurrentDominantSpeaker))
                        {
                            target.MediaStreamIds.Add(e.CurrentDominantSpeaker);
                        }

                        this.msiToParticipantId[e.CurrentDominantSpeaker] = target.ParticipantId;
                        Console.WriteLine($"Adopted MSI {e.CurrentDominantSpeaker} for SDK-unbound real caller '{target.DisplayName}' (CH{target.ChannelId})");
                    }
                    else
                    {
                        var placeholderId = $"MSI-{e.CurrentDominantSpeaker}";
                        int channelId = Interlocked.Increment(ref this.nextChannelIndex) - 1;
                        var placeholder = new ParticipantInfo
                        {
                            ParticipantId = placeholderId,
                            DisplayName = $"Speaker-{e.CurrentDominantSpeaker}",
                            ChannelId = channelId,
                            IsMsiFallback = true,
                            MediaStreamIds = new List<uint> { e.CurrentDominantSpeaker },
                            JoinTime = now,
                        };
                        placeholder.FrameMetadataBytes = BuildFrameMetadata(this.currentSessionId, placeholder);
                        this.msiToParticipantId[e.CurrentDominantSpeaker] = placeholderId;
                        this.participantsById[placeholderId] = placeholder;
                        Console.WriteLine($"MSI-FALLBACK: Registered placeholder for MSI {e.CurrentDominantSpeaker} on CH{channelId}");
                    }
                }

                if (this.msiToParticipantId.TryGetValue(e.CurrentDominantSpeaker, out var pid) &&
                    this.participantsById.TryGetValue(pid, out var info))
                {
                    var speakerChanged = this.lastKnownSpeakerId != pid;
                    this.lastKnownSpeakerId = pid;
                    this.lastKnownSpeakerTime = now;
                    if (speakerChanged)
                    {
                        Console.WriteLine($"Speaker resolved: MSI {e.CurrentDominantSpeaker} -> {info.DisplayName} (userId={info.UserId ?? "<empty>"}, CH{info.ChannelId}, fallback={info.IsMsiFallback})");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in OnDominantSpeakerChanged: {ex.Message}");
            }
        }

        private void OnAudioMediaReceived(object sender, AudioMediaReceivedEventArgs e)
        {
            try
            {
                if (this.connectedClients.IsEmpty)
                {
                    return;
                }

                if (e.Buffer.UnmixedAudioBuffers != null)
                {
                    try
                    {
                        if (e.Buffer.UnmixedAudioBuffers.Any())
                        {
                            this.ProcessUnmixedAudio(e);
                            return;
                        }
                    }
                    catch (ArgumentNullException)
                    {
                        // Buffer not initialized — fall through to mixed mode
                    }
                }

                if (e.Buffer.IsSilence && this.unmixedAudioEnabled)
                {
                    this.ProcessSilenceForUnmixedMode(e);
                    return;
                }

                this.ProcessMixedAudio(e);
            }
            catch (Exception ex)
            {
                Console.WriteLine($" Audio processing error: {ex.Message}");
            }
            finally
            {
                e.Buffer.Dispose();
            }
        }

        private void ProcessUnmixedAudio(AudioMediaReceivedEventArgs e)
        {
            Interlocked.Increment(ref this.unmixedPacketCount);

            if (!this.unmixedAudioEnabled)
            {
                this.unmixedAudioEnabled = true;
                Console.WriteLine($" UNMIXED AUDIO MODE ENABLED ({e.Buffer.UnmixedAudioBuffers.Count()} streams)");
            }

            foreach (var unmixedBuffer in e.Buffer.UnmixedAudioBuffers)
            {
                uint msi = unmixedBuffer.ActiveSpeakerId;
                byte[] audioData = new byte[unmixedBuffer.Length];
                System.Runtime.InteropServices.Marshal.Copy(
                    unmixedBuffer.Data, audioData, 0, (int)unmixedBuffer.Length);

                // ResolveParticipantByMsi may return null when audio arrives for a brand-new
                // MSI before DominantSpeakerChanged fires (the SDK races these two streams).
                // Eagerly adopt onto the cached sole SdkUnbound real caller so the FIRST
                // frame is attributed correctly. If both paths return null (0 or 2+ SdkUnbound
                // callers), EmitOrBufferFrame holds the frame in the per-MSI pre-roll buffer
                // until DominantSpeakerChanged binds the MSI.
                var participant = this.ResolveParticipantByMsi(msi)
                    ?? this.TryEagerAdoptMsi(msi);
                this.EmitOrBufferFrame(msi, audioData, participant);
            }

            this.MaybeLogStatistics();
        }

        private void ProcessSilenceForUnmixedMode(AudioMediaReceivedEventArgs e)
        {
            if (this.connectedClients.IsEmpty || this.participantsById.IsEmpty)
            {
                return;
            }

            Interlocked.Increment(ref this.silencePacketCount);

            byte[] silenceData = new byte[e.Buffer.Length];
            System.Runtime.InteropServices.Marshal.Copy(
                e.Buffer.Data, silenceData, 0, (int)e.Buffer.Length);

            foreach (var info in this.participantsById.Values)
            {
                this.SendAudioFrame(silenceData, info);
            }
        }

        private void ProcessMixedAudio(AudioMediaReceivedEventArgs e)
        {
            Interlocked.Increment(ref this.mixedPacketCount);

            byte[] audioData = new byte[e.Buffer.Length];
            System.Runtime.InteropServices.Marshal.Copy(
                e.Buffer.Data, audioData, 0, (int)e.Buffer.Length);

            double audioEnergy = this.CalculateAudioEnergy(audioData);
            this.recentAudioEnergy.Enqueue(audioEnergy);
            if (this.recentAudioEnergy.Count > ENERGYHISTORYSIZE)
            {
                this.recentAudioEnergy.Dequeue();
            }

            double avgEnergy = this.recentAudioEnergy.Count > 0 ? this.recentAudioEnergy.Average() : 0;
            bool isTrueSilence = avgEnergy < SILENCEENERGYTHRESHOLD;

            if (isTrueSilence)
            {
                Interlocked.Increment(ref this.silencePacketCount);
            }

            var participant = this.ResolveSpeaker();
            this.EmitOrBufferFrame(MIXEDMODEATTRIBUTIONKEY, audioData, participant);

            this.MaybeLogStatistics();
        }

        // Unified emission path used by both mixed (key = MIXEDMODEATTRIBUTIONKEY) and
        // unmixed (key = MSI) audio. Replaces the original ship-as-streamId=unknown fallback,
        // which created phantom streams on the consumer side that collided with real channels
        // (AWS Transcribe saw "restart" mid-call and dropped the session).
        //
        // Semantics: if attribution succeeded, drain any frames buffered for this key and emit
        // the current frame. If attribution failed, buffer the current frame. The buffer is
        // bounded by MAXPREROLLFRAMESPERKEY (~5s); overflow drops the oldest. There is no
        // sticky "resolved once" state — a mid-call gap (e.g., a new speaker's MSI is reported
        // dominant before OnDominantSpeakerChanged binds it) re-buffers and re-drains on the
        // very next attributed emit, so utterance onsets are not lost.
        //
        // Hot-path cost: when there is no pending buffer (the steady state), the only work is
        // a lock-free ContainsKey check, then SendAudioFrame. No dictionary locks taken.
        private void EmitOrBufferFrame(uint attributionKey, byte[] audioData, ParticipantInfo participant)
        {
            if (participant != null)
            {
                if (this.preRollAudioByKey.ContainsKey(attributionKey) &&
                    this.preRollAudioByKey.TryRemove(attributionKey, out var pending))
                {
                    lock (pending)
                    {
                        while (pending.Count > 0)
                        {
                            this.SendAudioFrame(pending.Dequeue(), participant);
                        }
                    }
                }

                this.SendAudioFrame(audioData, participant);
                return;
            }

            var queue = this.preRollAudioByKey.GetOrAdd(attributionKey, CreatePreRollQueue);
            lock (queue)
            {
                if (queue.Count >= MAXPREROLLFRAMESPERKEY)
                {
                    queue.Dequeue();
                    Interlocked.Increment(ref this.droppedAudioFrameCount);
                }

                queue.Enqueue(audioData);
            }
        }

        private ParticipantInfo ResolveParticipantByMsi(uint msi)
        {
            if (this.msiToParticipantId.TryGetValue(msi, out var pid) &&
                this.participantsById.TryGetValue(pid, out var info))
            {
                return info;
            }

            return null;
        }

        // Refresh the cached sole SdkUnbound real caller. Called whenever the participant
        // set changes (join/leave/merge/placeholder claim). If zero or 2+ SdkUnbound real
        // callers exist, the cache is null and the audio path will not eager-adopt unknown
        // MSIs — EmitOrBufferFrame then holds those frames in the per-MSI pre-roll buffer
        // until DominantSpeakerChanged binds the MSI.
        private void RefreshCachedSdkUnboundCaller()
        {
            ParticipantInfo found = null;
            foreach (var p in this.participantsById.Values)
            {
                if (p.IsMsiFallback || !p.IsSdkUnbound || string.IsNullOrWhiteSpace(p.UserId))
                {
                    continue;
                }

                if (found != null)
                {
                    found = null;
                    break;
                }

                found = p;
            }

            this.cachedSdkUnboundCaller = found;
        }

        // Bind an MSI to the cached SdkUnbound caller from the audio hot path. The atomic
        // ConcurrentDictionary insert is the source of truth for routing; the diagnostic
        // MediaStreamIds list on ParticipantInfo is intentionally NOT mutated here — it's
        // touched without locking from the SDK callback threads (ProcessParticipant merge,
        // OnDominantSpeakerChanged adoption), and adding hot-path mutation would race.
        private ParticipantInfo TryEagerAdoptMsi(uint msi)
        {
            var target = this.cachedSdkUnboundCaller;
            if (target == null)
            {
                return null;
            }

            if (this.msiToParticipantId.TryAdd(msi, target.ParticipantId))
            {
                Console.WriteLine($"Eagerly adopted MSI {msi} for SDK-unbound caller '{target.DisplayName}' (CH{target.ChannelId}) from audio path");
                return target;
            }

            // Another thread won the race; resolve and return whatever is now bound.
            return this.ResolveParticipantByMsi(msi);
        }

        private ParticipantInfo ResolveSpeaker()
        {
            if (this.currentDominantSpeakerMsi.HasValue)
            {
                var info = this.ResolveParticipantByMsi(this.currentDominantSpeakerMsi.Value);
                if (info != null)
                {
                    return info;
                }
            }

            if (!string.IsNullOrEmpty(this.lastKnownSpeakerId))
            {
                var elapsed = (DateTime.UtcNow - this.lastKnownSpeakerTime).TotalMilliseconds;
                if (elapsed < SPEAKERPERSISTENCEMS &&
                    this.participantsById.TryGetValue(this.lastKnownSpeakerId, out var info))
                {
                    return info;
                }
            }

            if (this.participantsById.Count == 1)
            {
                foreach (var info in this.participantsById.Values)
                {
                    return info;
                }
            }

            return null;
        }

        private void SendAudioFrame(byte[] audioData, ParticipantInfo participant)
        {
            // participant is non-null on every path that reaches here — EmitOrBufferFrame
            // buffers unresolved frames rather than sending them, and ProcessSilenceForUnmixedMode
            // iterates over participantsById.Values. The null guard is defensive.
            if (this.connectedClients.IsEmpty || participant == null)
            {
                return;
            }

            byte[] metadataBytes = participant.FrameMetadataBytes;

            var frame = new byte[IAB1HEADERLEN + metadataBytes.Length + audioData.Length];
            Buffer.BlockCopy(Iab1Magic, 0, frame, 0, 4);
            Buffer.BlockCopy(BitConverter.GetBytes((uint)metadataBytes.Length), 0, frame, 4, 4);
            Buffer.BlockCopy(BitConverter.GetBytes((uint)audioData.Length), 0, frame, 8, 4);
            Buffer.BlockCopy(metadataBytes, 0, frame, IAB1HEADERLEN, metadataBytes.Length);
            Buffer.BlockCopy(audioData, 0, frame, IAB1HEADERLEN + metadataBytes.Length, audioData.Length);

            this.QueueAudioFrame(frame);
        }

        private void QueueAudioFrame(byte[] frame)
        {
            try
            {
                if (this.audioFrameQueue.IsAddingCompleted || !this.audioFrameQueue.TryAdd(frame))
                {
                    Interlocked.Increment(ref this.droppedAudioFrameCount);
                }
            }
            catch (InvalidOperationException)
            {
                Interlocked.Increment(ref this.droppedAudioFrameCount);
            }
        }

        private double CalculateAudioEnergy(byte[] audioData)
        {
            double sum = 0;
            for (int i = 0; i < audioData.Length - 1; i += 2)
            {
                short sample = BitConverter.ToInt16(audioData, i);
                sum += sample * sample;
            }

            return Math.Sqrt(sum / (audioData.Length / 2));
        }

        private void MaybeLogStatistics()
        {
            var now = DateTime.UtcNow;
            if ((now - this.lastStatsLog).TotalSeconds < 10)
            {
                return;
            }

            this.lastStatsLog = now;
            Console.WriteLine(
                $"STATS: mixed={Interlocked.Read(ref this.mixedPacketCount)} " +
                $"unmixed={Interlocked.Read(ref this.unmixedPacketCount)} " +
                $"silence={Interlocked.Read(ref this.silencePacketCount)} " +
                $"droppedAudio={Interlocked.Read(ref this.droppedAudioFrameCount)} " +
                $"droppedMeta={Interlocked.Read(ref this.droppedMetadataMessageCount)} " +
                $"participants={this.participantsById.Count}");
        }

        private void ValidateSubscriptionMediaType(MediaType mediaType)
        {
            if (mediaType != MediaType.Vbss && mediaType != MediaType.Video)
            {
                throw new ArgumentOutOfRangeException($"Invalid mediaType: {mediaType}");
            }
        }

        private void StartAudioStreamServer()
        {
            try
            {
                this.audioStreamServer = new TcpListener(IPAddress.Any, AUDIOSTREAMPORT);
                this.audioStreamServer.Start();
                Console.WriteLine($"LISTENING: 0.0.0.0:{AUDIOSTREAMPORT} (Audio Stream — IAB1 frames)");

                Task.Run(() => this.ProcessAudioFrameQueue(), this.shutdownCts.Token);
                _ = Task.Run(async () => await this.AcceptAudioClientsAsync().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start audio server: {ex.Message}");
            }
        }

        private async Task AcceptAudioClientsAsync()
        {
            while (this.audioStreamServer != null && !this.shutdownCts.IsCancellationRequested)
            {
                try
                {
                    var client = await this.audioStreamServer.AcceptTcpClientAsync().ConfigureAwait(false);
                    client.NoDelay = true;

                    // Without a send timeout, a consumer whose receive buffer fills (slow
                    // downstream like AWS Transcribe back-pressure) causes the synchronous
                    // stream.Write in WriteAudioFrameToClients to block forever, freezing
                    // audio relay for ALL clients. 5s is well past any legitimate scheduling
                    // hiccup; on timeout, IOException bubbles up and RemoveAudioClient runs.
                    client.SendTimeout = 5000;

                    Console.WriteLine($"Audio client connected: {((IPEndPoint)client.Client.RemoteEndPoint).Address}");
                    this.connectedClients[client] = 0;
                }
                catch (Exception ex)
                {
                    if (!this.shutdownCts.IsCancellationRequested)
                    {
                        Console.WriteLine($"Audio accept error: {ex.Message}");
                    }

                    break;
                }
            }
        }

        private void ProcessAudioFrameQueue()
        {
            try
            {
                foreach (var frame in this.audioFrameQueue.GetConsumingEnumerable(this.shutdownCts.Token))
                {
                    this.WriteAudioFrameToClients(frame);
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (ObjectDisposedException)
            {
            }
        }

        private void WriteAudioFrameToClients(byte[] frame)
        {
            foreach (var client in this.connectedClients.Keys)
            {
                try
                {
                    if (client.Connected)
                    {
                        var stream = client.GetStream();
                        stream.Write(frame, 0, frame.Length);
                    }
                    else
                    {
                        this.RemoveAudioClient(client);
                    }
                }
                catch
                {
                    this.RemoveAudioClient(client);
                }
            }
        }

        private void RemoveAudioClient(TcpClient client)
        {
            if (this.connectedClients.TryRemove(client, out _))
            {
                try
                {
                    client.Close();
                }
                catch
                {
                }
            }
        }

        private void StopAudioStreamServer()
        {
            try
            {
                var closeMeta = Encoding.UTF8.GetBytes(
                    $"{{\"callId\":\"{this.currentSessionId}\",\"streamId\":\"eos\",\"endOfStream\":true}}");
                var closeFrame = new byte[IAB1HEADERLEN + closeMeta.Length];
                Buffer.BlockCopy(Iab1Magic, 0, closeFrame, 0, 4);
                Buffer.BlockCopy(BitConverter.GetBytes((uint)closeMeta.Length), 0, closeFrame, 4, 4);
                Buffer.BlockCopy(BitConverter.GetBytes(0u), 0, closeFrame, 8, 4);
                Buffer.BlockCopy(closeMeta, 0, closeFrame, IAB1HEADERLEN, closeMeta.Length);

                foreach (var client in this.connectedClients.Keys)
                {
                    try
                    {
                        client.GetStream().Write(closeFrame, 0, closeFrame.Length);
                        client.Close();
                    }
                    catch
                    {
                    }
                }

                if (!this.audioFrameQueue.IsAddingCompleted)
                {
                    this.audioFrameQueue.CompleteAdding();
                }

                this.audioStreamServer?.Stop();
                Console.WriteLine("Audio stream server stopped");
            }
            catch
            {
            }
        }

        private void StartMetadataStreamServer()
        {
            try
            {
                this.metadataStreamServer = new TcpListener(IPAddress.Any, METADATASTREAMPORT);
                this.metadataStreamServer.Start();
                Console.WriteLine($"LISTENING: 0.0.0.0:{METADATASTREAMPORT} (Participant Metadata Stream)");

                Task.Run(() => this.ProcessMetadataMessageQueue(), this.shutdownCts.Token);
                _ = Task.Run(async () => await this.AcceptMetadataClientsAsync().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start metadata server: {ex.Message}");
            }
        }

        private async Task AcceptMetadataClientsAsync()
        {
            while (this.metadataStreamServer != null && !this.shutdownCts.IsCancellationRequested)
            {
                try
                {
                    var client = await this.metadataStreamServer.AcceptTcpClientAsync().ConfigureAwait(false);
                    client.NoDelay = true;
                    Console.WriteLine($"Metadata client connected: {((IPEndPoint)client.Client.RemoteEndPoint).Address}");
                    this.connectedMetadataClients[client] = 0;
                    this.SendCurrentParticipantsSnapshot(client);
                }
                catch (Exception ex)
                {
                    if (!this.shutdownCts.IsCancellationRequested)
                    {
                        Console.WriteLine($"Metadata accept error: {ex.Message}");
                    }

                    break;
                }
            }
        }

        private void SendCurrentParticipantsSnapshot(TcpClient client)
        {
            try
            {
                foreach (var info in this.participantsById.Values)
                {
                    if (info.IsMsiFallback)
                    {
                        continue;
                    }

                    var json = BuildParticipantEventJson(this.currentSessionId, info, "CURRENT");
                    var bytes = Encoding.UTF8.GetBytes(json + "\n");
                    client.GetStream().Write(bytes, 0, bytes.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending snapshot: {ex.Message}");
            }
        }

        private void SendParticipantEvent(ParticipantInfo info, string eventType)
        {
            try
            {
                if (info.IsMsiFallback && eventType != "LEAVE")
                {
                    return;
                }

                var json = BuildParticipantEventJson(this.currentSessionId, info, eventType);
                var bytes = Encoding.UTF8.GetBytes(json + "\n");
                Console.WriteLine($"FLOW[Bot->Consumer] session={this.currentSessionId} event={eventType} participantId={info.ParticipantId} userId={info.UserId ?? string.Empty} name={info.DisplayName ?? string.Empty} channel={info.ChannelId}");

                try
                {
                    if (this.metadataMessageQueue.IsAddingCompleted || !this.metadataMessageQueue.TryAdd(bytes))
                    {
                        Interlocked.Increment(ref this.droppedMetadataMessageCount);
                    }
                }
                catch (InvalidOperationException)
                {
                    Interlocked.Increment(ref this.droppedMetadataMessageCount);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending participant event: {ex.Message}");
            }
        }

        private void ProcessMetadataMessageQueue()
        {
            try
            {
                foreach (var message in this.metadataMessageQueue.GetConsumingEnumerable(this.shutdownCts.Token))
                {
                    this.WriteMetadataMessageToClients(message);
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (ObjectDisposedException)
            {
            }
        }

        private void WriteMetadataMessageToClients(byte[] message)
        {
            foreach (var client in this.connectedMetadataClients.Keys)
            {
                try
                {
                    if (client.Connected)
                    {
                        var stream = client.GetStream();
                        stream.Write(message, 0, message.Length);
                    }
                    else
                    {
                        this.RemoveMetadataClient(client);
                    }
                }
                catch
                {
                    this.RemoveMetadataClient(client);
                }
            }
        }

        private void RemoveMetadataClient(TcpClient client)
        {
            if (this.connectedMetadataClients.TryRemove(client, out _))
            {
                try
                {
                    client.Close();
                }
                catch
                {
                }
            }
        }

        private void StopMetadataStreamServer()
        {
            try
            {
                foreach (var client in this.connectedMetadataClients.Keys)
                {
                    try
                    {
                        client.Close();
                    }
                    catch
                    {
                    }
                }

                this.metadataStreamServer?.Stop();
                this.metadataStreamServer = null;
                if (!this.metadataMessageQueue.IsAddingCompleted)
                {
                    this.metadataMessageQueue.CompleteAdding();
                }

                Console.WriteLine("Metadata stream server stopped");
            }
            catch
            {
            }
        }

        private void OnVideoMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            this.GraphLogger.Info($"[{e.SocketId}]: Received Video");
            e.Buffer.Dispose();
        }

        private void OnVbssMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            this.GraphLogger.Info($"[{e.SocketId}]: Received VBSS");
            e.Buffer.Dispose();
        }

        /// <summary>
        /// Immutable-ish participant record. Identity fields are set at construction.
        /// MediaStreamIds and FrameMetadataBytes may be appended/replaced during participant
        /// lifetime but only from the SDK callback thread sequence (ProcessParticipant /
        /// OnDominantSpeakerChanged), so no extra synchronization is needed beyond the
        /// outer ConcurrentDictionary.
        /// </summary>
        private class ParticipantInfo
        {
            internal string ParticipantId { get; set; }

            internal string UserId { get; set; }

            internal string DisplayName { get; set; }

            internal string TenantId { get; set; }

            internal int ChannelId { get; set; }

            internal bool IsMsiFallback { get; set; }

            // True if the SDK never exposed a Send/SendReceive audio MediaStream for this
            // participant when it was created. Policy-attached users (the bot's host) typically
            // have no such MediaStream; their audio arrives via dominant-speaker events on
            // MSIs that the SDK assigns dynamically and may change over the call's lifetime.
            // Sticky after creation so the participant remains the absorber for subsequently
            // observed unbound MSIs, not just the first one.
            internal bool IsSdkUnbound { get; set; }

            internal List<uint> MediaStreamIds { get; set; }

            internal DateTime JoinTime { get; set; }

            internal byte[] FrameMetadataBytes { get; set; }
        }
    }
}
