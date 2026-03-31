// BotMediaStream.cs - WITH UNMIXED AUDIO SUPPORT + MSI FALLBACK

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Sockets;
    using System.Runtime.Remoting.Metadata.W3cXsd2001;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading.Tasks;
    using System.Timers;
    using Microsoft.Graph.Beta.Models;
    using Microsoft.Graph.Beta.Models.TermStore;
    using Microsoft.Graph.Communications.Calls;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Graph.Communications.Resources;
    using Microsoft.Skype.Bots.Media;
    using Microsoft.Skype.Internal.Media.Services.Common;

    public class BotMediaStream : ObjectRootDisposable
    {
        private const int ENERGYHISTORYSIZE = 10;
        private const double SILENCEENERGYTHRESHOLD = 100.0;
        private const double SPEAKERPERSISTENCEMS = 1000;
        private const double SPEAKERCHANGEDEBOUNCEMS = 500;

        private readonly IAudioSocket audioSocket;
        private readonly IVideoSocket vbssSocket;
        private readonly List<IVideoSocket> videoSockets;
        private readonly ILocalMediaSession mediaSession;
        private readonly ICall call;
        private TcpListener audioStreamServer;
        private TcpListener metadataStreamServer;
        private TcpListener commandServer;
        private ConcurrentBag<TcpClient> connectedClients = new ConcurrentBag<TcpClient>();
        private ConcurrentBag<TcpClient> connectedMetadataClients = new ConcurrentBag<TcpClient>();
        private string currentSessionId = Guid.NewGuid().ToString();
        private string agentUserId;
        private string agentDisplayName;
        private string orgTenantId;

        // Unmixed audio tracking
        private bool unmixedAudioEnabled = false;
        private int unmixedPacketCount = 0;
        private int mixedPacketCount = 0;

        // Enhanced participant tracking
        private ConcurrentDictionary<string, ParticipantInfo> participantsById =
            new ConcurrentDictionary<string, ParticipantInfo>();
        private ConcurrentDictionary<uint, string> msiToParticipantId =
            new ConcurrentDictionary<uint, string>();

        // Enhanced speaker detection
        private uint? currentDominantSpeakerMsi = null;
        private uint? previousDominantSpeakerMsi = null;
        private DateTime lastSpeakerChangeTime = DateTime.UtcNow;

        // Speaker persistence
        private string lastKnownSpeakerId = null;
        private DateTime lastKnownSpeakerTime = DateTime.UtcNow;

        // Audio energy detection
        private Queue<double> recentAudioEnergy = new Queue<double>();

        // Statistics
        private int agentPacketCount = 0;
        private int callerPacketCount = 0;
        private int unknownPacketCount = 0;
        private int silencePacketCount = 0;
        private DateTime lastStatsLog = DateTime.UtcNow;

        // Detailed logging
        private bool detailedLogging = false;
        private System.Timers.Timer participantCheckTimer;
        private int participantCheckCount;

        public BotMediaStream(
            ILocalMediaSession mediaSession,
            IGraphLogger logger,
            ICall call,
            string agentUserId = null,
            string agentDisplayName = null)
            : base(logger)
        {
            ArgumentVerifier.ThrowOnNullArgument(mediaSession, nameof(mediaSession));
            ArgumentVerifier.ThrowOnNullArgument(logger, nameof(logger));
            ArgumentVerifier.ThrowOnNullArgument(call, nameof(call));

            this.mediaSession = mediaSession;
            this.call = call;
            this.agentUserId = agentUserId;
            this.agentDisplayName = agentDisplayName;

            Console.WriteLine($"Agent for this recording: {this.agentDisplayName} ({this.agentUserId})");

            // Subscribe to audio media
            this.audioSocket = mediaSession.AudioSocket;
            if (this.audioSocket == null)
            {
                throw new InvalidOperationException("A mediaSession needs to have at least an audioSocket");
            }

            this.audioSocket.AudioMediaReceived += this.OnAudioMediaReceived;
            this.audioSocket.DominantSpeakerChanged += this.OnDominantSpeakerChanged;

            Console.WriteLine($"Subscribed to AudioMediaReceived and DominantSpeakerChanged events");

            // Subscribe to video
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

            // Subscribe to participant events
            this.call.Participants.OnUpdated += this.OnParticipantsUpdated;

            // Process existing participants
            Console.WriteLine($"Initial participants: {this.call.Participants.Count}");
            foreach (var participant in this.call.Participants)
            {
                this.ProcessParticipant(participant, isNew: false);
            }

            this.participantCheckTimer = new System.Timers.Timer(1000);
            this.participantCheckTimer.AutoReset = true;
            this.participantCheckTimer.Elapsed += this.CheckForParticipants;
            this.participantCheckTimer.Start();
            Console.WriteLine($"Started participant polling timer");

            this.StartAudioStreamServer();
            this.StartMetadataStreamServer();
            this.StartCommandServer();
        }

        public string AgentUserId => this.agentUserId;

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
                    Console.WriteLine($"Participant already tracked");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering participant from CallHandler: {ex.Message}");
            }
        }

        private void CheckForParticipants(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                this.participantCheckCount++;

                if (this.call?.Participants != null && this.participantsById.Count == 0)
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

                // Stop when we find participants OR after reasonable attempts
                if (this.participantsById.Count > 0 || this.participantCheckCount > 30)
                {
                    if (this.participantsById.Count > 0)
                    {
                        Console.WriteLine($"Successfully tracked {this.participantsById.Count} participant(s) - stopping polling");
                    }
                    else
                    {
                        Console.WriteLine($"No participants found after 30 checks - relying on MSI fallback");
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

                    Console.WriteLine($"Participant JOINED");
                    this.ProcessParticipant(participant, isNew: true);
                }

                foreach (var participant in args.RemovedResources)
                {
                    var participantId = participant.Id;

                    if (this.participantsById.TryRemove(participantId, out var info))
                    {
                        foreach (var msi in info.MediaStreamIds)
                        {
                            this.msiToParticipantId.TryRemove(msi, out _);
                        }

                        Console.WriteLine($"Participant LEFT: {info.DisplayName} ({info.Role})");

                        this.SendParticipantMetadata(info, "LEAVE");

                        if (this.lastKnownSpeakerId == participantId)
                        {
                            this.lastKnownSpeakerId = null;
                        }
                    }
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
                var identity = participant.Resource?.Info?.Identity;

                if (identity == null)
                {
                    Console.WriteLine($" Participant {participantId} has no identity");
                    return;
                }

                // Skip bot itself
                if (identity.Application != null)
                {
                    Console.WriteLine($" Skipping bot participant");
                    return;
                }

                string displayName = "Unknown";
                string userId = null;
                bool isAgent = false;
                string participantTenantId = "unknown";
                string callerType = "EXTERNAL_TEAMS";

                if (identity.User != null)
                {
                    displayName = identity.User.DisplayName ?? "Teams User";
                    userId = identity.User.Id;
                    isAgent = this.IsAgentParticipant(userId, participantId, displayName);

                    if (identity.User.AdditionalData != null &&
                        identity.User.AdditionalData.TryGetValue("tenantId", out object tenantObj))
                    {
                        participantTenantId = (tenantObj?.ToString() ?? "unknown").Trim().ToLowerInvariant();
                    }

                    callerType = isAgent ? "AGENT" : "EXTERNAL_TEAMS";

                    // Capture org tenant from agent so we can classify internal callers
                    if (isAgent && string.IsNullOrEmpty(this.orgTenantId) && participantTenantId != "unknown")
                    {
                        this.orgTenantId = participantTenantId;
                        Console.WriteLine($"Org tenant ID set from agent: {this.orgTenantId}");
                        this.ReclassifyTrackedParticipantsForOrgTenant();
                    }

                    // If same tenant as org → INTERNAL (not external Teams user)
                    if (!isAgent && !string.IsNullOrEmpty(this.orgTenantId) &&
                        participantTenantId != "unknown" && participantTenantId == this.orgTenantId)
                    {
                        callerType = "INTERNAL";
                    }
                }
                else
                {
                    displayName = "External Caller";

                    if (identity.AdditionalData != null)
                    {
                        if (identity.AdditionalData.ContainsKey("phone"))
                        {
                            callerType = "PSTN";
                        }
                        else if (identity.AdditionalData.ContainsKey("guest"))
                        {
                            callerType = "GUEST";
                        }
                    }
                }

                string role = isAgent ? "AGENT" : "CALLER";
                byte channelId = isAgent ? (byte)1 : (byte)0;
                bool isInternal = isAgent || callerType == "INTERNAL";

                // Extract MSI values from media streams
                var mediaStreamIds = new List<uint>();
                if (participant.Resource.MediaStreams != null)
                {
                    foreach (var stream in participant.Resource.MediaStreams)
                    {
                        if (stream.MediaType == Modality.Audio &&
                            uint.TryParse(stream.SourceId, out uint msi))
                        {
                            mediaStreamIds.Add(msi);
                            this.msiToParticipantId[msi] = participantId;
                            Console.WriteLine($"Mapped MSI {msi} → {participantId} ({role})");
                        }
                    }
                }

                // If this is an anonymous/non-user caller but we already have a real caller on CH0,
                // merge MSI mappings into the real caller and suppress duplicate placeholder join.
                if (!isAgent && string.IsNullOrWhiteSpace(userId))
                {
                    var existingRealCaller = this.participantsById.Values.FirstOrDefault(
                        p => !p.IsAgent && p.ChannelId == 0 && !string.IsNullOrWhiteSpace(p.UserId));

                    if (existingRealCaller != null)
                    {
                        foreach (var msi in mediaStreamIds)
                        {
                            this.msiToParticipantId[msi] = existingRealCaller.ParticipantId;
                            if (!existingRealCaller.MediaStreamIds.Contains(msi))
                            {
                                existingRealCaller.MediaStreamIds.Add(msi);
                            }
                        }

                        Console.WriteLine($" Merged anonymous participant '{participantId}' into real caller '{existingRealCaller.DisplayName}'");
                        return;
                    }
                }

                var participantInfo = new ParticipantInfo
                {
                    ParticipantId = participantId,
                    UserId = userId,
                    DisplayName = displayName,
                    Role = role,
                    ChannelId = channelId,
                    IsAgent = isAgent,
                    CallerType = callerType,
                    IsInternal = isInternal,
                    TenantId = participantTenantId,
                    MediaStreamIds = mediaStreamIds,
                    JoinTime = DateTime.UtcNow,
                };

                // Evict any MSI-fallback or anonymous placeholder on the same channel
                foreach (var kvp in this.participantsById)
                {
                    if (kvp.Key != participantId && kvp.Value.ChannelId == channelId &&
                        (kvp.Value.IsMsiFallback || kvp.Value.UserId == null))
                    {
                        if (this.participantsById.TryRemove(kvp.Key, out var evicted))
                        {
                            Console.WriteLine($"Evicted placeholder '{evicted.DisplayName}' (channel {channelId}) → replaced by '{displayName}'");
                            this.SendParticipantMetadata(evicted, "LEAVE");
                        }
                    }
                }

                this.participantsById[participantId] = participantInfo;
                this.SendParticipantMetadata(participantInfo, "JOIN");

                if (isAgent && this.lastKnownSpeakerId == null)
                {
                    this.lastKnownSpeakerId = participantId;
                    Console.WriteLine($"Set as initial speaker (agent)");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing participant: {ex.Message}");
            }
        }

        private void ReclassifyTrackedParticipantsForOrgTenant()
        {
            if (string.IsNullOrWhiteSpace(this.orgTenantId))
            {
                return;
            }

            foreach (var kvp in this.participantsById)
            {
                var info = kvp.Value;
                if (info == null || info.IsAgent)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(info.TenantId) &&
                    info.TenantId != "unknown" &&
                    info.TenantId == this.orgTenantId &&
                    info.CallerType != "INTERNAL")
                {
                    info.CallerType = "INTERNAL";
                    info.IsInternal = true;
                    Console.WriteLine($"Reclassified participant as INTERNAL: {info.DisplayName} ({info.UserId ?? info.ParticipantId})");
                    this.SendParticipantMetadata(info, "UPDATE");
                }
            }
        }

        private bool IsAgentParticipant(string participantUserId, string participantId, string participantDisplayName)
        {
            var configuredAgentId = this.NormalizeIdentityValue(this.agentUserId);
            var userId = this.NormalizeIdentityValue(participantUserId);
            var pId = this.NormalizeIdentityValue(participantId);

            if (!string.IsNullOrWhiteSpace(configuredAgentId))
            {
                if (configuredAgentId == userId || configuredAgentId == pId)
                {
                    return true;
                }

                var rawConfiguredAgent = (this.agentUserId ?? string.Empty).Trim().ToLowerInvariant();
                var rawParticipantUser = (participantUserId ?? string.Empty).Trim().ToLowerInvariant();

                if (!string.IsNullOrWhiteSpace(userId) && rawConfiguredAgent.Contains(userId))
                {
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(configuredAgentId) && rawParticipantUser.Contains(configuredAgentId))
                {
                    return true;
                }
            }

            if (!string.IsNullOrWhiteSpace(this.agentDisplayName) && !string.IsNullOrWhiteSpace(participantDisplayName) &&
                string.Equals(this.agentDisplayName.Trim(), participantDisplayName.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Agent identified by display-name fallback: '{participantDisplayName}'");
                return true;
            }

            return false;
        }

        private string NormalizeIdentityValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var normalized = value.Trim().ToLowerInvariant();

            if (normalized.StartsWith("8:orgid:"))
            {
                normalized = normalized.Substring("8:orgid:".Length);
            }
            else if (normalized.StartsWith("8:teamsvisitor:"))
            {
                normalized = normalized.Substring("8:teamsvisitor:".Length);
            }
            else if (normalized.StartsWith("8:"))
            {
                normalized = normalized.Substring(2);
            }

            if (Guid.TryParse(normalized, out Guid parsedGuid))
            {
                return parsedGuid.ToString("D").ToLowerInvariant();
            }

            var parts = normalized.Split(':');
            if (parts.Length > 1)
            {
                var last = parts[parts.Length - 1];
                if (Guid.TryParse(last, out Guid lastGuid))
                {
                    return lastGuid.ToString("D").ToLowerInvariant();
                }
            }

            return normalized;
        }

        private void OnDominantSpeakerChanged(object sender, DominantSpeakerChangedEventArgs e)
        {
            try
            {
                var now = DateTime.UtcNow;
                var timeSinceLastChange = (now - this.lastSpeakerChangeTime).TotalMilliseconds;

                if (timeSinceLastChange < SPEAKERCHANGEDEBOUNCEMS &&
                    this.currentDominantSpeakerMsi.HasValue)
                {
                    if (this.detailedLogging)
                    {
                        Console.WriteLine($"Ignoring speaker change (debounce: {timeSinceLastChange:F0}ms)");
                    }

                    return;
                }

                this.previousDominantSpeakerMsi = this.currentDominantSpeakerMsi;

                if (e.CurrentDominantSpeaker != DominantSpeakerChangedEventArgs.None)
                {
                    this.currentDominantSpeakerMsi = e.CurrentDominantSpeaker;
                    this.lastSpeakerChangeTime = now;

                    if (!this.msiToParticipantId.ContainsKey(e.CurrentDominantSpeaker))
                    {
                        bool isFirstMsi = this.msiToParticipantId.Count == 0;
                        string fakeParticipantId = $"MSI-{e.CurrentDominantSpeaker}";
                        this.msiToParticipantId[e.CurrentDominantSpeaker] = fakeParticipantId;

                        var participantInfo = new ParticipantInfo
                        {
                            ParticipantId = fakeParticipantId,
                            UserId = isFirstMsi ? this.agentUserId : null,
                            DisplayName = isFirstMsi ? (this.agentDisplayName ?? "Agent (MSI-based)") : $"Caller-{e.CurrentDominantSpeaker}",
                            Role = isFirstMsi ? "AGENT" : "CALLER",
                            ChannelId = isFirstMsi ? (byte)1 : (byte)0,
                            IsAgent = isFirstMsi,
                            IsMsiFallback = true,
                            MediaStreamIds = new List<uint> { e.CurrentDominantSpeaker },
                            JoinTime = now,
                        };

                        this.participantsById[fakeParticipantId] = participantInfo;

                        if (isFirstMsi)
                        {
                            this.lastKnownSpeakerId = fakeParticipantId;
                            Console.WriteLine($"sMSI-FALLBACK: Registered MSI {e.CurrentDominantSpeaker} as AGENT (first speaker)");
                        }
                        else
                        {
                            Console.WriteLine($" MSI-FALLBACK: Registered MSI {e.CurrentDominantSpeaker} as CALLER");
                        }
                    }

                    // Look up participant by MSI
                    if (this.msiToParticipantId.TryGetValue(e.CurrentDominantSpeaker, out string participantId))
                    {
                        if (this.participantsById.TryGetValue(participantId, out var info))
                        {
                            this.lastKnownSpeakerId = participantId;
                            this.lastKnownSpeakerTime = now;

                            var channelName = info.IsAgent ? "CH1-AGENT" : "CH0-CALLER";
                            if (this.detailedLogging)
                            {
                                Console.WriteLine($"Speaker: {info.DisplayName} [{channelName}] (MSI: {e.CurrentDominantSpeaker})");
                            }
                        }
                    }
                    else
                    {
                        if (this.detailedLogging)
                        {
                            Console.WriteLine($" MSI {e.CurrentDominantSpeaker} not mapped to any participant");
                        }
                    }
                }
                else
                {
                    this.currentDominantSpeakerMsi = null;
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
                // ✅ CHECK FOR UNMIXED AUDIO BUFFERS (TRUE PER-SPEAKER AUDIO)
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
                        // UnmixedAudioBuffers collection is not properly initialized - fall through to mixed mode
                    }
                }

                // ✅ HANDLE SILENCE WITH UNMIXED MODE (write to all active tracks)
                if (e.Buffer.IsSilence && this.unmixedAudioEnabled)
                {
                    this.ProcessSilenceForUnmixedMode(e);
                    return;
                }

                // ✅ FALLBACK: MIXED AUDIO + DOMINANT SPEAKER (compliance recording mode)
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
            this.unmixedPacketCount++;

            if (!this.unmixedAudioEnabled)
            {
                this.unmixedAudioEnabled = true;
                Console.WriteLine($" UNMIXED AUDIO MODE ENABLED!");
                Console.WriteLine($" Received {e.Buffer.UnmixedAudioBuffers.Count()} unmixed audio streams");
            }

            foreach (var unmixedBuffer in e.Buffer.UnmixedAudioBuffers)
            {
                uint msi = unmixedBuffer.ActiveSpeakerId;

                // Copy audio data
                byte[] audioData = new byte[unmixedBuffer.Length];
                System.Runtime.InteropServices.Marshal.Copy(
                    unmixedBuffer.Data,
                    audioData,
                    0,
                    (int)unmixedBuffer.Length);

                // Determine channel based on MSI
                var channelResult = this.DetermineChannelByMsi(msi);
                byte channelByte = channelResult.ChannelId;
                string speakerInfo = channelResult.SpeakerInfo;

                // Track statistics
                if (channelByte == 1)
                {
                    this.agentPacketCount++;
                }
                else if (channelByte == 0)
                {
                    this.callerPacketCount++;
                }
                else
                {
                    this.unknownPacketCount++;
                }

                // Build and send frame
                this.SendAudioFrame(audioData, channelByte, speakerInfo, msi);
            }

            // Log statistics
            var now = DateTime.UtcNow;
            if ((now - this.lastStatsLog).TotalSeconds >= 10)
            {
                this.LogStatistics();
                this.lastStatsLog = now;
            }
        }

        private void ProcessSilenceForUnmixedMode(AudioMediaReceivedEventArgs e)
        {
            this.silencePacketCount++;

            // Copy silence buffer
            byte[] silenceData = new byte[e.Buffer.Length];
            System.Runtime.InteropServices.Marshal.Copy(
                e.Buffer.Data,
                silenceData,
                0,
                (int)e.Buffer.Length);

            // Write silence to ALL active speaker tracks
            foreach (var participantInfo in this.participantsById.Values)
            {
                this.SendAudioFrame(silenceData, participantInfo.ChannelId, $"{participantInfo.DisplayName} (silence)", null);
            }
        }

        private void ProcessMixedAudio(AudioMediaReceivedEventArgs e)
        {
            this.mixedPacketCount++;

            // Copy audio data
            byte[] audioData = new byte[e.Buffer.Length];
            System.Runtime.InteropServices.Marshal.Copy(
                e.Buffer.Data,
                audioData,
                0,
                (int)e.Buffer.Length);

            // Calculate audio energy
            double audioEnergy = this.CalculateAudioEnergy(audioData);
            this.recentAudioEnergy.Enqueue(audioEnergy);
            if (this.recentAudioEnergy.Count > ENERGYHISTORYSIZE)
            {
                this.recentAudioEnergy.Dequeue();
            }

            double avgEnergy = this.recentAudioEnergy.Count > 0 ? this.recentAudioEnergy.Average() : 0;
            bool isTrueSilence = avgEnergy < SILENCEENERGYTHRESHOLD;

            // Determine channel using dominant speaker logic
            var channelResult = this.DetermineChannel(isTrueSilence);
            byte channelByte = channelResult.ChannelId;
            string speakerInfo = channelResult.SpeakerInfo;

            // Track statistics
            if (isTrueSilence)
            {
                this.silencePacketCount++;
            }
            else
            {
                if (channelByte == 1)
                {
                    this.agentPacketCount++;
                }
                else if (channelByte == 0)
                {
                    this.callerPacketCount++;
                }
                else
                {
                    this.unknownPacketCount++;
                }
            }

            // Send frame
            this.SendAudioFrame(audioData, channelByte, speakerInfo, this.currentDominantSpeakerMsi);

            // Log statistics
            var now = DateTime.UtcNow;
            if (!isTrueSilence && (now - this.lastStatsLog).TotalSeconds >= 10)
            {
                this.LogStatistics();
                this.lastStatsLog = now;
            }
        }

        private ChannelDetermination DetermineChannelByMsi(uint msi)
        {
            // Look up participant by MSI
            if (this.msiToParticipantId.TryGetValue(msi, out string participantId))
            {
                if (this.participantsById.TryGetValue(participantId, out var participant))
                {
                    return new ChannelDetermination
                    {
                        ChannelId = participant.ChannelId,
                        SpeakerInfo = $"{participant.DisplayName} ({participant.Role})",
                        Confidence = "HIGH - Unmixed MSI",
                    };
                }
            }

            // Unknown MSI - default to caller
            return new ChannelDetermination
            {
                ChannelId = 0,
                SpeakerInfo = $"Unknown-MSI-{msi}",
                Confidence = "LOW - Unknown unmixed MSI",
            };
        }

        private ChannelDetermination DetermineChannel(bool isTrueSilence)
        {
            var now = DateTime.UtcNow;

            // Strategy 1: Use current dominant speaker
            if (this.currentDominantSpeakerMsi.HasValue)
            {
                if (this.msiToParticipantId.TryGetValue(this.currentDominantSpeakerMsi.Value, out string participantId))
                {
                    if (this.participantsById.TryGetValue(participantId, out var participant))
                    {
                        return new ChannelDetermination
                        {
                            ChannelId = participant.ChannelId,
                            SpeakerInfo = $"{participant.DisplayName} ({participant.Role})",
                            Confidence = "HIGH - Current dominant speaker",
                        };
                    }
                }
            }

            // Strategy 2: Use speaker persistence
            if (!string.IsNullOrEmpty(this.lastKnownSpeakerId))
            {
                var timeSinceLastSpeaker = (now - this.lastKnownSpeakerTime).TotalMilliseconds;

                if (timeSinceLastSpeaker < SPEAKERPERSISTENCEMS)
                {
                    if (this.participantsById.TryGetValue(this.lastKnownSpeakerId, out var participant))
                    {
                        return new ChannelDetermination
                        {
                            ChannelId = participant.ChannelId,
                            SpeakerInfo = $"{participant.DisplayName} ({participant.Role})",
                            Confidence = $"MEDIUM - Persistence ({timeSinceLastSpeaker:F0}ms ago)",
                        };
                    }
                }
            }

            // Strategy 3: Only one participant
            if (this.participantsById.Count == 1)
            {
                var onlyParticipant = this.participantsById.Values.First();
                return new ChannelDetermination
                {
                    ChannelId = onlyParticipant.ChannelId,
                    SpeakerInfo = $"{onlyParticipant.DisplayName} (only participant)",
                    Confidence = "MEDIUM - Only participant",
                };
            }

            // Strategy 4: Default to caller
            return new ChannelDetermination
            {
                ChannelId = 0,
                SpeakerInfo = "Unknown",
                Confidence = "VERY LOW - No speaker info",
            };
        }

        private void SendAudioFrame(byte[] audioData, byte channelByte, string speakerInfo, uint? msi)
        {
            // Calculate audio energy to detect actual speech
            double energy = this.CalculateAudioEnergy(audioData);
            bool hasAudio = energy > SILENCEENERGYTHRESHOLD;

            using (var ms = new MemoryStream())
            {
                // Session ID (36 bytes)
                byte[] sessionId = Encoding.ASCII.GetBytes(
                    this.currentSessionId.PadRight(36).Substring(0, 36));
                ms.Write(sessionId, 0, 36);

                // Channel (1 byte)
                ms.WriteByte(channelByte);

                // Length (4 bytes, little-endian)
                byte[] lengthBytes = BitConverter.GetBytes((uint)audioData.Length);
                ms.Write(lengthBytes, 0, 4);

                // Audio data
                ms.Write(audioData, 0, audioData.Length);

                byte[] frame = ms.ToArray();

                // Send to all connected clients
                this.SendToClients(frame);

                // Always log non-silence frames to debug channel routing
                if (hasAudio && this.connectedClients.Count > 0)
                {
                    var channelName = channelByte == 1 ? "CH1-AGENT" :
                                     channelByte == 0 ? "CH0-CALLER" : "CH?-UNKNOWN";
                    Console.WriteLine($" SENDING [{channelName}] {audioData.Length}B | {speakerInfo}" +
                        (msi.HasValue ? $" MSI:{msi.Value}" : string.Empty) + $" | Energy:{energy:F0}");
                }
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

        private void SendToClients(byte[] frame)
        {
            var deadClients = new List<TcpClient>();

            foreach (var client in this.connectedClients)
            {
                try
                {
                    if (client.Connected)
                    {
                        var stream = client.GetStream();
                        stream.Write(frame, 0, frame.Length);
                        stream.Flush();
                    }
                    else
                    {
                        deadClients.Add(client);
                    }
                }
                catch
                {
                    deadClients.Add(client);
                }
            }

            foreach (var dead in deadClients)
            {
                try
                {
                    dead.Close();
                }
                catch
                {
                }
            }
        }

        private void LogStatistics()
        {
            var total = this.agentPacketCount + this.callerPacketCount + this.unknownPacketCount;

            if (total == 0)
            {
                Console.WriteLine($"STATS: No audio packets yet (Silence: {this.silencePacketCount})");
                return;
            }

            var agentPercent = this.agentPacketCount * 100.0 / total;
            var callerPercent = this.callerPacketCount * 100.0 / total;

            if (this.unmixedAudioEnabled)
            {
                Console.WriteLine($"Unmixed packets: {this.unmixedPacketCount}");
            }
            else
            {
                Console.WriteLine($"Mixed packets: {this.mixedPacketCount}");
            }
        }

        // Video subscription methods (unchanged)
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
                this.GraphLogger.Error(ex, $"Video subscription failed");
            }
        }

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
                this.GraphLogger.Error(ex, $"Unsubscribing failed");
            }
        }

        private void ValidateSubscriptionMediaType(MediaType mediaType)
        {
            if (mediaType != MediaType.Vbss && mediaType != MediaType.Video)
            {
                throw new ArgumentOutOfRangeException($"Invalid mediaType: {mediaType}");
            }
        }

        private async Task AcceptClientsAsync()
        {
            while (this.audioStreamServer != null)
            {
                try
                {
                    var client = await this.audioStreamServer.AcceptTcpClientAsync().ConfigureAwait(false);
                    Console.WriteLine($"STT client connected: {((IPEndPoint)client.Client.RemoteEndPoint).Address}");
                    this.connectedClients.Add(client);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Accept error: {ex.Message}");
                    break;
                }
            }
        }

        private void StartAudioStreamServer()
        {
            Console.WriteLine("StartAudioStreamServer() CALLED!");

            try
            {
                this.audioStreamServer = new TcpListener(IPAddress.Any, 5001);
                this.audioStreamServer.Start();
                Console.WriteLine("Audio stream server listening on port 5001");

                Task.Run(async () => await this.AcceptClientsAsync().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start audio server: {ex.Message}");
            }
        }

        private void StopAudioStreamServer()
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    byte[] sessionId = Encoding.ASCII.GetBytes(
                        this.currentSessionId.PadRight(36).Substring(0, 36));
                    ms.Write(sessionId, 0, 36);
                    ms.WriteByte(255);
                    ms.Write(BitConverter.GetBytes(0u), 0, 4);

                    byte[] closeFrame = ms.ToArray();

                    foreach (var client in this.connectedClients)
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
            Console.WriteLine("StartMetadataStreamServer() CALLED!");

            try
            {
                this.metadataStreamServer = new TcpListener(IPAddress.Any, 5002);
                this.metadataStreamServer.Start();
                Console.WriteLine("LISTENING: 0.0.0.0:5002 (Metadata Stream)");

                _ = Task.Run(async () => await this.AcceptMetadataClientsAsync().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start metadata server: {ex.Message}");
            }
        }

        private async Task AcceptMetadataClientsAsync()
        {
            while (this.metadataStreamServer != null)
            {
                try
                {
                    var client = await this.metadataStreamServer.AcceptTcpClientAsync().ConfigureAwait(false);
                    Console.WriteLine($" Metadata client connected: {((IPEndPoint)client.Client.RemoteEndPoint).Address}");
                    this.connectedMetadataClients.Add(client);
                    this.SendCurrentParticipantsSnapshot(client);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Metadata accept error: {ex.Message}");
                    break;
                }
            }
        }

        private void SendCurrentParticipantsSnapshot(TcpClient client)
        {
            try
            {
                foreach (var kvp in this.participantsById)
                {
                    if (kvp.Value.IsMsiFallback)
                    {
                        continue;
                    }

                    var json = this.BuildParticipantJson(kvp.Value, "CURRENT");
                    var jsonBytes = Encoding.UTF8.GetBytes(json + "\n");
                    client.GetStream().Write(jsonBytes, 0, jsonBytes.Length);
                    Console.WriteLine($"FLOW[Bot->SPL Snapshot] session={this.currentSessionId} event=CURRENT userId={kvp.Value?.UserId ?? string.Empty} name={kvp.Value?.DisplayName ?? string.Empty} role={kvp.Value?.Role ?? string.Empty} callerType={kvp.Value?.CallerType ?? string.Empty}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending snapshot: {ex.Message}");
            }
        }

        private void SendParticipantMetadata(ParticipantInfo info, string eventType)
        {
            try
            {
                if (info.IsMsiFallback && eventType != "LEAVE")
                {
                    return;
                }

                var json = this.BuildParticipantJson(info, eventType);
                var jsonBytes = Encoding.UTF8.GetBytes(json + "\n");

                Console.WriteLine($"FLOW[Bot->SPL] session={this.currentSessionId} event={eventType} userId={info.UserId ?? string.Empty} participantId={info.ParticipantId ?? string.Empty} name={info.DisplayName ?? string.Empty} role={info.Role ?? string.Empty} callerType={info.CallerType ?? string.Empty} isInternal={info.IsInternal} isAgent={info.IsAgent} channel={info.ChannelId}");

                var deadClients = new List<TcpClient>();
                foreach (var client in this.connectedMetadataClients)
                {
                    try
                    {
                        if (client.Connected)
                        {
                            var stream = client.GetStream();
                            stream.Write(jsonBytes, 0, jsonBytes.Length);
                            stream.Flush();
                        }
                        else
                        {
                            deadClients.Add(client);
                        }
                    }
                    catch
                    {
                        deadClients.Add(client);
                    }
                }

                foreach (var dead in deadClients)
                {
                    try
                    {
                        dead.Close();
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending participant metadata: {ex.Message}");
            }
        }

        private string BuildParticipantJson(ParticipantInfo info, string eventType)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.AppendFormat("\"eventType\":\"{0}\",", eventType);
            sb.AppendFormat("\"sessionId\":\"{0}\",", this.currentSessionId);
            sb.AppendFormat("\"timestamp\":\"{0}\",", DateTime.UtcNow.ToString("o"));
            sb.AppendFormat("\"participantId\":\"{0}\",", info.ParticipantId ?? string.Empty);
            sb.AppendFormat("\"userId\":\"{0}\",", info.UserId ?? string.Empty);
            sb.AppendFormat("\"displayName\":\"{0}\",", info.DisplayName ?? string.Empty);
            sb.AppendFormat("\"role\":\"{0}\",", info.Role ?? string.Empty);
            sb.AppendFormat("\"tenantId\":\"{0}\",", info.TenantId ?? "unknown");
            sb.AppendFormat("\"callerType\":\"{0}\",", info.CallerType ?? "EXTERNAL_TEAMS");
            sb.AppendFormat("\"isInternal\":{0},", info.IsInternal ? "true" : "false");
            sb.AppendFormat("\"isAgent\":{0},", info.IsAgent ? "true" : "false");
            sb.AppendFormat("\"channelId\":{0}", info.ChannelId);
            sb.Append("}");
            return sb.ToString();
        }

        private void StopMetadataStreamServer()
        {
            try
            {
                foreach (var client in this.connectedMetadataClients)
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
                this.connectedMetadataClients = new ConcurrentBag<TcpClient>();
                Console.WriteLine("Metadata stream server stopped");
            }
            catch
            {
            }
        }

        private void StartCommandServer()
        {
            Console.WriteLine("StartCommandServer() CALLED!");

            try
            {
                this.commandServer = new TcpListener(IPAddress.Any, 5003);
                this.commandServer.Start();
                Console.WriteLine("LISTENING: 0.0.0.0:5003 (Command Server - SET_AGENT)");
                _ = Task.Run(async () => await this.AcceptCommandClientsAsync().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to start command server: {ex.Message}");
            }
        }

        private async Task AcceptCommandClientsAsync()
        {
            while (this.commandServer != null)
            {
                try
                {
                    var client = await this.commandServer.AcceptTcpClientAsync().ConfigureAwait(false);
                    Console.WriteLine($"Command client connected: {((IPEndPoint)client.Client.RemoteEndPoint).Address}");
                    _ = Task.Run(async () => await this.HandleCommandClientAsync(client).ConfigureAwait(false));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Command accept error: {ex.Message}");
                    break;
                }
            }
        }

        private async Task HandleCommandClientAsync(TcpClient client)
        {
            try
            {
                using (var reader = new System.IO.StreamReader(client.GetStream(), Encoding.UTF8))
                {
                    string line;
                    while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                    {
                        line = line.Trim();
                        if (string.IsNullOrEmpty(line))
                        {
                            continue;
                        }

                        if (line.Contains("\"SET_AGENT\""))
                        {
                            string newUserId = this.ExtractJsonStringValue(line, "userId");
                            string newDisplayName = this.ExtractJsonStringValue(line, "displayName");
                            if (!string.IsNullOrWhiteSpace(newUserId))
                            {
                                this.agentUserId = newUserId.Trim();
                            }

                            if (!string.IsNullOrWhiteSpace(newDisplayName))
                            {
                                this.agentDisplayName = newDisplayName.Trim();
                            }

                            Console.WriteLine($" FLOW[SPL->Bot Command] command=SET_AGENT userId={this.agentUserId ?? string.Empty} displayName={this.agentDisplayName ?? string.Empty}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Command client error: {ex.Message}");
            }
            finally
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

        private string ExtractJsonStringValue(string json, string key)
        {
            var search = $"\"{key}\":\"";
            int start = json.IndexOf(search, StringComparison.Ordinal);
            if (start < 0)
            {
                return null;
            }

            start += search.Length;
            int end = json.IndexOf('"', start);
            if (end < 0)
            {
                return null;
            }

            return json.Substring(start, end - start);
        }

        private void StopCommandServer()
        {
            try
            {
                this.commandServer?.Stop();
                this.commandServer = null;
                Console.WriteLine("Command server stopped");
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

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

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
            this.StopCommandServer();
        }

        private class ParticipantInfo
        {
            public string ParticipantId { get; set; }
            public string UserId { get; set; }
            public string DisplayName { get; set; }
            public string Role { get; set; }
            public byte ChannelId { get; set; }
            public bool IsAgent { get; set; }
            public bool IsMsiFallback { get; set; }
            public string CallerType { get; set; }
            public bool IsInternal { get; set; }
            public string TenantId { get; set; }
            public List<uint> MediaStreamIds { get; set; }
            public DateTime JoinTime { get; set; }
        }

        private class ChannelDetermination
        {
            public byte ChannelId { get; set; }
            public string SpeakerInfo { get; set; }
            public string Confidence { get; set; }
        }
    }
}
