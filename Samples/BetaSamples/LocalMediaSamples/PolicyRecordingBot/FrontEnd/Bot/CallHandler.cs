// <copyright file="CallHandler.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

#pragma warning disable SA1600 // Recovered sample internals are intentionally kept close to the upstream PoC shape.

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Timers;
    using Microsoft.Graph.Beta.Models;
    using Microsoft.Graph.Communications.Calls;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Graph.Communications.Resources;
    using Microsoft.Skype.Bots.Media;
    using Sample.Common;

    using RejectReason = Microsoft.Graph.Communications.Common.RejectReason;

    internal class CallHandler : HeartbeatHandler
    {
        public const uint DominantSpeakerNone = DominantSpeakerChangedEventArgs.None;

        private readonly HashSet<uint> availableSocketIds = new HashSet<uint>();
        private readonly LRUCache currentVideoSubscriptions = new LRUCache(SampleConstants.NumberOfMultiviewSockets + 1);
        private readonly object subscriptionLock = new object();
        private readonly ConcurrentDictionary<uint, uint> msiToSocketIdMapping = new ConcurrentDictionary<uint, uint>();

        // private readonly Timer recordingStatusFlipTimer;
        private int recordingStatusIndex = -1;
        private int participantsCount = 0;

        public CallHandler(
            ICall statefulCall,
            string configuredOrgId = null,
            string agentUserId = null,
            string agentDisplayName = null)
            : base(TimeSpan.FromMinutes(1), statefulCall?.GraphLogger)
        {
            var effectiveAgentUserId = string.IsNullOrWhiteSpace(agentUserId)
                ? (Environment.GetEnvironmentVariable("AGENT_USER_ID") ?? string.Empty).Trim()
                : agentUserId.Trim();

            var effectiveAgentDisplayName = string.IsNullOrWhiteSpace(agentDisplayName)
                ? (Environment.GetEnvironmentVariable("AGENT_DISPLAY_NAME") ?? string.Empty).Trim()
                : agentDisplayName.Trim();

            this.Call = statefulCall;
            this.Call.OnUpdated += this.CallOnUpdated;

            this.Call.Participants.OnUpdated += this.ParticipantsOnUpdated;
            this.Call.ParticipantLeftHandler += this.ParticipantLeft;
            this.Call.ParticipantJoiningHandler += this.ParticipantJoining;

            Console.WriteLine($"CallHandler for agent: {effectiveAgentDisplayName} ({effectiveAgentUserId})");
            if (string.IsNullOrWhiteSpace(effectiveAgentUserId))
            {
                Console.WriteLine(" AGENT_USER_ID resolved to empty. Agent may be classified as CALLER.");
            }

            if (this.Call.Resource != null)
            {
                Console.WriteLine($"   Call.Resource.State: {this.Call.Resource.State}");
                Console.WriteLine($"   Call.Resource.Direction: {this.Call.Resource.Direction}");
                Console.WriteLine($"   Call.Resource.Source: {this.Call.Resource.Source?.Identity?.User?.DisplayName ?? "N/A"}");

                if (this.Call.Resource.Targets != null)
                {
                    Console.WriteLine($"   Call.Resource.Targets.Count: {this.Call.Resource.Targets.Count}");
                }
            }

            Console.WriteLine($"   Call.Participants: {(this.Call.Participants != null ? "NOT NULL" : "NULL")}");

            if (this.Call.Participants != null)
            {
                Console.WriteLine($"   Call.Participants.Count: {this.Call.Participants.Count}");
                Console.WriteLine($"   Call.Participants is ICollection: {this.Call.Participants is System.Collections.ICollection}");

                // Try to iterate even if count is 0
                try
                {
                    int enumCount = 0;
                    foreach (var p in this.Call.Participants)
                    {
                        enumCount++;
                    }

                    Console.WriteLine($"   Actually enumerated: {enumCount} participants");
                }
                catch (Exception enumEx)
                {
                    Console.WriteLine($"     Failed to enumerate participants: {enumEx.Message}");
                    Console.WriteLine($"      Exception Type: {enumEx.GetType().Name}");
                }
            }

            var mediaSession = this.Call.GetLocalMediaSession();

            // Create BotMediaStream
            this.BotMediaStream = new BotMediaStream(
                mediaSession,
                this.GraphLogger,
                this.Call,
                effectiveAgentUserId,
                effectiveAgentDisplayName,
                configuredOrgId);

            if (this.Call.Participants.Count > 0)
            {
                Console.WriteLine($" Found {this.Call.Participants.Count} existing participant(s)");
                foreach (var participant in this.Call.Participants)
                {
                    var displayName = participant.Resource?.Info?.Identity?.User?.DisplayName ??
                                    participant.Resource?.Info?.Identity?.Application?.DisplayName ??
                                    "Unknown";
                    Console.WriteLine($" Participant: {displayName}");

                    this.BotMediaStream.RegisterParticipantFromCallHandler(participant);

                    // Subscribe for updates
                    var participantDetails = participant.Resource?.Info?.Identity?.User;
                    if (participantDetails != null)
                    {
                        participant.OnUpdated += this.OnParticipantUpdated;
                        this.SubscribeToParticipantVideo(participant, forceSubscribe: false);
                    }
                }
            }

            for (uint i = 0; i < SampleConstants.NumberOfMultiviewSockets; i++)
            {
                this.availableSocketIds.Add(i);
            }

            // DISABLED: Demo/test timer that cycles recording status - not needed for production
            // var timer = new Timer(1000 * 60 * 5);
            // timer.AutoReset = true;
            // timer.Elapsed += this.OnRecordingStatusFlip;
            // this.recordingStatusFlipTimer = timer;
        }

        public ICall Call { get; }

        public BotMediaStream BotMediaStream { get; private set; }

        protected override Task HeartbeatAsync(ElapsedEventArgs args)
        {
            Console.WriteLine($"Sending keepalive for call {this.Call.Id}");
            return this.Call.KeepAliveAsync();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            this.Call.OnUpdated -= this.CallOnUpdated;
            this.Call.Participants.OnUpdated -= this.ParticipantsOnUpdated;

            foreach (var participant in this.Call.Participants)
            {
                participant.OnUpdated -= this.OnParticipantUpdated;
            }

            // Only cleanup if timer was initialized
           // if (this.recordingStatusFlipTimer != null)
           // {
           //     this.recordingStatusFlipTimer.Enabled = false;
           //     this.recordingStatusFlipTimer.Elapsed -= this.OnRecordingStatusFlip;
           // }
            this.BotMediaStream.Dispose();
        }

        private void OnRecordingStatusFlip(object source, ElapsedEventArgs e)
        {
            _ = Task.Run(async () =>
            {
                var recordingStatus = new[] { RecordingStatus.Recording, RecordingStatus.NotRecording, RecordingStatus.Failed };
                var recordingIndex = this.recordingStatusIndex + 1;
                if (recordingIndex >= recordingStatus.Length)
                {
                    var recordedParticipantId = this.Call.Resource.IncomingContext.ObservedParticipantId;
                    this.GraphLogger.Warn($"We've rolled through all the status'... removing participant {recordedParticipantId}");
                    var recordedParticipant = this.Call.Participants[recordedParticipantId];
                    await recordedParticipant.DeleteAsync().ConfigureAwait(false);
                    return;
                }

                var newStatus = recordingStatus[recordingIndex];
                this.GraphLogger.Info($"Flipping recording status to {newStatus}");

                try
                {
                    await this.Call.UpdateRecordingStatusAsync(newStatus).ConfigureAwait(false);
                    this.recordingStatusIndex = recordingIndex;
                }
                catch (Exception exc)
                {
                    this.GraphLogger.Error(exc, $"Failed to flip the recording status to {newStatus}");
                }
            }).ForgetAndLogExceptionAsync(this.GraphLogger);
        }

        private void CallOnUpdated(ICall sender, ResourceEventArgs<Call> e)
        {
            this.GraphLogger.Info($"Call status updated to {e.NewResource.State} with ResultInfo at {e.NewResource.ResultInfo?.Code}");

            if (e.OldResource.State != e.NewResource.State)
            {
                this.GraphLogger.Info($"Call state changed from {e.OldResource.State} to {e.NewResource.State}");
            }
        }

        private void ParticipantsOnUpdated(IParticipantCollection sender, CollectionEventArgs<IParticipant> args)
        {
            foreach (var participant in args.AddedResources)
            {
                var participantDetails = participant.Resource.Info.Identity.User;
                var displayName = participantDetails?.DisplayName ??
                                 participant.Resource?.Info?.Identity?.Application?.DisplayName ??
                                 "Unknown";

                Console.WriteLine($"Participant added: {displayName}");

                this.BotMediaStream?.RegisterParticipantFromCallHandler(participant);

                if (participantDetails != null)
                {
                    participant.OnUpdated += this.OnParticipantUpdated;
                    this.SubscribeToParticipantVideo(participant, forceSubscribe: false);
                }
            }

            foreach (var participant in args.RemovedResources)
            {
                var participantDetails = participant.Resource.Info.Identity.User;
                var displayName = participantDetails?.DisplayName ?? "Unknown";
                Console.WriteLine($"Participant removed: {displayName}");

                if (participantDetails != null)
                {
                    participant.OnUpdated -= this.OnParticipantUpdated;
                    this.UnsubscribeFromParticipantVideo(participant);
                }
            }
        }

        private void ParticipantLeft(Call call, string participantId)
        {
            this.participantsCount--;
            this.GraphLogger.Info($"[{this.Call.Id}:The participant {participantId} has left with code {call?.ResultInfo?.Code} and subcode {call?.ResultInfo?.Subcode})");
        }

        private ParticipantJoiningResponse ParticipantJoining(Call call)
        {
            Console.WriteLine($"ParticipantJoining event! Current count: {this.participantsCount}");

            if (this.participantsCount < SampleConstants.GroupSize)
            {
                this.participantsCount++;
                Console.WriteLine($"Accepting participant (new count: {this.participantsCount})");
                return new AcceptJoinResponse();
            }
            else
            {
                Console.WriteLine($" Rejecting participant (room full)");
                return new RejectJoinResponse()
                {
                    Reason = RejectReason.Busy,
                };
            }
        }

        private void OnParticipantUpdated(IParticipant sender, ResourceEventArgs<Participant> args)
        {
            this.SubscribeToParticipantVideo(sender, forceSubscribe: false);
        }

        private void UnsubscribeFromParticipantVideo(IParticipant participant)
        {
            var participantSendCapableVideoStream = participant.Resource.MediaStreams.Where(x => x.MediaType == Modality.Video &&
              (x.Direction == MediaDirection.SendReceive || x.Direction == MediaDirection.SendOnly)).FirstOrDefault();

            if (participantSendCapableVideoStream != null)
            {
                var msi = uint.Parse(participantSendCapableVideoStream.SourceId);
                lock (this.subscriptionLock)
                {
                    if (this.currentVideoSubscriptions.TryRemove(msi))
                    {
                        if (this.msiToSocketIdMapping.TryRemove(msi, out uint socketId))
                        {
                            this.BotMediaStream.Unsubscribe(MediaType.Video, socketId);
                            this.availableSocketIds.Add(socketId);
                        }
                    }
                }
            }
        }

        private void SubscribeToParticipantVideo(IParticipant participant, bool forceSubscribe = true)
        {
            bool subscribeToVideo = false;
            uint socketId = uint.MaxValue;

            var participantSendCapableVideoStream = participant.Resource.MediaStreams.Where(x => x.MediaType == Modality.Video &&
               (x.Direction == MediaDirection.SendReceive || x.Direction == MediaDirection.SendOnly)).FirstOrDefault();
            if (participantSendCapableVideoStream != null)
            {
                bool updateMSICache = false;
                var msi = uint.Parse(participantSendCapableVideoStream.SourceId);
                lock (this.subscriptionLock)
                {
                    if (this.currentVideoSubscriptions.Count < this.Call.GetLocalMediaSession().VideoSockets.Count)
                    {
                        if (!this.msiToSocketIdMapping.ContainsKey(msi))
                        {
                            if (this.availableSocketIds.Any())
                            {
                                socketId = this.availableSocketIds.Last();
                                this.availableSocketIds.Remove((uint)socketId);
                                subscribeToVideo = true;
                            }
                        }

                        updateMSICache = true;
                        this.GraphLogger.Info($"[{this.Call.Id}:SubscribeToParticipant(socket {socketId} available, the number of remaining sockets is {this.availableSocketIds.Count}, subscribing to the participant {participant.Id})");
                    }
                    else if (forceSubscribe)
                    {
                        updateMSICache = true;
                        subscribeToVideo = true;
                    }

                    if (updateMSICache)
                    {
                        this.currentVideoSubscriptions.TryInsert(msi, out uint? dequeuedMSIValue);
                        if (dequeuedMSIValue != null)
                        {
                            this.msiToSocketIdMapping.TryRemove((uint)dequeuedMSIValue, out socketId);
                        }
                    }
                }

                if (subscribeToVideo && socketId != uint.MaxValue)
                {
                    this.msiToSocketIdMapping.AddOrUpdate(msi, socketId, (k, v) => socketId);
                    this.GraphLogger.Info($"[{this.Call.Id}:SubscribeToParticipant(subscribing to the participant {participant.Id} on socket {socketId})");
                    this.BotMediaStream.Subscribe(MediaType.Video, msi, VideoResolution.HD1080p, socketId);
                }
            }

            var vbssParticipant = participant.Resource.MediaStreams.SingleOrDefault(x => x.MediaType == Modality.VideoBasedScreenSharing
            && x.Direction == MediaDirection.SendOnly);
            if (vbssParticipant != null)
            {
                this.GraphLogger.Info($"[{this.Call.Id}:SubscribeToParticipant(subscribing to the VBSS sharer {participant.Id})");
                this.BotMediaStream.Subscribe(MediaType.Vbss, uint.Parse(vbssParticipant.SourceId), VideoResolution.HD1080p, socketId);
            }
        }

        private void OnDominantSpeakerChanged(object sender, DominantSpeakerChangedEventArgs e)
        {
            this.GraphLogger.Info($"[{this.Call.Id}:OnDominantSpeakerChanged(DominantSpeaker={e.CurrentDominantSpeaker})]");

            if (e.CurrentDominantSpeaker != DominantSpeakerNone)
            {
                IParticipant participant = this.GetParticipantFromMSI(e.CurrentDominantSpeaker);
                var participantDetails = participant?.Resource?.Info?.Identity?.User;
                if (participantDetails != null)
                {
                    this.SubscribeToParticipantVideo(participant, forceSubscribe: true);
                }
            }
        }

        private IParticipant GetParticipantFromMSI(uint msi)
        {
            return this.Call.Participants.SingleOrDefault(x => x.Resource.IsInLobby == false && x.Resource.MediaStreams.Any(y => y.SourceId == msi.ToString()));
        }
    }
}
