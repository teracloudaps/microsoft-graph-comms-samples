// <copyright file="ParticipantIdentityMetadata.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd.Bot
{
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
}