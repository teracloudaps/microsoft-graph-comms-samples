// <copyright file="VMConfiguration.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Skype.Bots.Media;
    using Sample.PolicyRecordingBot.FrontEnd.Http;

    /// <summary>
    /// Configuration for VM deployments that run without Azure Cloud Services runtime dependencies.
    /// </summary>
    public class VMConfiguration : IConfiguration
    {
        private const string ServiceDnsNameKey = "ServiceDnsName";
        private const string AadAppIdKey = "AadAppId";
        private const string AadAppSecretKey = "AadAppSecret";
        private const string HomeTenantIdKey = "HomeTenantId";
        private const string PlaceCallEndpointUrlKey = "PlaceCallEndpointUrl";
        private const string DefaultCertificateKey = "DefaultCertificate";
        private const string CertificateThumbprintKey = "CertificateThumbprint";
        private const string CallControlPortKey = "CallControlPort";
        private const string MediaPortKey = "MediaPort";
        private const string InstancePublicIpAddressKey = "InstancePublicIPAddress";
        private const string ServiceFqdnKey = "ServiceFqdn";
        private const string CallControlHostKey = "CallControlHost";
        private const string EnvironmentVariablePrefix = "POLICY_RECORDING_BOT_";
        private const int DefaultCallControlPort = 9442;
        private const int DefaultMediaPort = 8445;

        private readonly IGraphLogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="VMConfiguration"/> class.
        /// </summary>
        /// <param name="logger">Graph logger.</param>
        public VMConfiguration(IGraphLogger logger)
        {
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
            this.Initialize();
        }

        /// <inheritdoc/>
        public string ServiceDnsName { get; private set; }

        /// <inheritdoc/>
        public IEnumerable<Uri> CallControlListeningUrls { get; private set; }

        /// <inheritdoc/>
        public Uri CallControlBaseUrl { get; private set; }

        /// <inheritdoc/>
        public Uri PlaceCallEndpointUrl { get; private set; }

        /// <inheritdoc/>
        public string AadAppId { get; private set; }

        /// <inheritdoc/>
        public string AadAppSecret { get; private set; }

        /// <inheritdoc/>
        public string HomeTenantId { get; private set; }

        /// <inheritdoc/>
        public MediaPlatformSettings MediaPlatformSettings { get; private set; }

        /// <inheritdoc/>
        public void Dispose()
        {
        }

        private void Initialize()
        {
            this.ServiceDnsName = GetRequiredSetting(ServiceDnsNameKey);
            this.AadAppId = GetRequiredSetting(AadAppIdKey);
            this.AadAppSecret = GetRequiredSetting(AadAppSecretKey);
            this.HomeTenantId = GetOptionalSetting(HomeTenantIdKey)?.Trim();

            var callControlPort = GetIntSetting(CallControlPortKey, DefaultCallControlPort);
            var mediaPort = GetIntSetting(MediaPortKey, DefaultMediaPort);
            var callControlHost = GetOptionalSetting(CallControlHostKey) ?? this.ServiceDnsName;
            var serviceFqdn = GetOptionalSetting(ServiceFqdnKey) ?? this.ServiceDnsName;
            var endpoint = GetOptionalSetting(PlaceCallEndpointUrlKey) ?? "https://graph.microsoft.com/v1.0";

            this.PlaceCallEndpointUrl = new Uri(endpoint, UriKind.Absolute);
            this.CallControlBaseUrl = new Uri($"https://{callControlHost}:{callControlPort}/{HttpRouteConstants.CallSignalingRoutePrefix}/{HttpRouteConstants.OnNotificationRequestRoute}");
            this.CallControlListeningUrls = new[] { new Uri($"https://+:{callControlPort}/") };

            var certificate = this.GetCertificateFromStore(GetCertificateThumbprint());
            var publicIpAddress = this.GetPublicIpAddress(serviceFqdn);

            this.MediaPlatformSettings = new MediaPlatformSettings
            {
                MediaPlatformInstanceSettings = new MediaPlatformInstanceSettings
                {
                    CertificateThumbprint = certificate.Thumbprint,
                    InstanceInternalPort = mediaPort,
                    InstancePublicIPAddress = publicIpAddress,
                    InstancePublicPort = mediaPort,
                    ServiceFqdn = serviceFqdn,
                },
                ApplicationId = this.AadAppId,
            };

            this.TraceConfigValue(ServiceDnsNameKey, this.ServiceDnsName);
            this.TraceConfigValue("CallControlBaseUrl", this.CallControlBaseUrl);
            this.TraceConfigValue("CallControlListeningUrls", string.Join(", ", this.CallControlListeningUrls));
            this.TraceConfigValue("MediaPublicEndpoint", $"{publicIpAddress}:{mediaPort}");
            this.TraceConfigValue("MediaPlatformFqdn", serviceFqdn);
        }

        private static string GetRequiredSetting(string key)
        {
            var value = GetOptionalSetting(key);
            if (string.IsNullOrWhiteSpace(value) || IsPlaceholderValue(value))
            {
                throw new ConfigurationErrorsException($"Missing required setting '{key}'. Set it in FrontEnd App.config or environment variable '{EnvironmentVariablePrefix}{key}'.");
            }

            return value.Trim();
        }

        private static string GetOptionalSetting(string key)
        {
            var value = ConfigurationManager.AppSettings[key];
            if (!string.IsNullOrWhiteSpace(value) && !IsPlaceholderValue(value))
            {
                return value.Trim();
            }

            value = Environment.GetEnvironmentVariable(key);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }

            value = Environment.GetEnvironmentVariable(EnvironmentVariablePrefix + key);
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static int GetIntSetting(string key, int defaultValue)
        {
            var value = GetOptionalSetting(key);
            if (string.IsNullOrWhiteSpace(value))
            {
                return defaultValue;
            }

            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) || parsed <= 0 || parsed > 65535)
            {
                throw new ConfigurationErrorsException($"Setting '{key}' must be a TCP port between 1 and 65535.");
            }

            return parsed;
        }

        private static bool IsPlaceholderValue(string value)
        {
            var trimmed = value.Trim();
            return (trimmed.StartsWith("%", StringComparison.Ordinal) && trimmed.EndsWith("%", StringComparison.Ordinal)) ||
                   (trimmed.StartsWith("$", StringComparison.Ordinal) && trimmed.EndsWith("$", StringComparison.Ordinal));
        }

        private static string NormalizeThumbprint(string thumbprint)
        {
            return thumbprint?.Replace(" ", string.Empty).Trim().ToUpperInvariant();
        }

        private string GetCertificateThumbprint()
        {
            return NormalizeThumbprint(GetOptionalSetting(CertificateThumbprintKey) ?? GetRequiredSetting(DefaultCertificateKey));
        }

        private X509Certificate2 GetCertificateFromStore(string thumbprint)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
                if (certs.Count == 0)
                {
                    throw new ConfigurationErrorsException($"Certificate '{thumbprint}' was not found in LocalMachine\\My.");
                }

                return certs[0];
            }
        }

        private IPAddress GetPublicIpAddress(string serviceFqdn)
        {
            var configuredIp = GetOptionalSetting(InstancePublicIpAddressKey);
            if (!string.IsNullOrWhiteSpace(configuredIp))
            {
                if (IPAddress.TryParse(configuredIp, out var parsedIp))
                {
                    return parsedIp;
                }

                throw new ConfigurationErrorsException($"Setting '{InstancePublicIpAddressKey}' is not a valid IP address.");
            }

            var resolvedAddress = Dns.GetHostAddresses(serviceFqdn).FirstOrDefault(address => address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
            if (resolvedAddress == null)
            {
                throw new ConfigurationErrorsException($"Could not resolve an IPv4 address for '{serviceFqdn}'. Set '{InstancePublicIpAddressKey}' explicitly.");
            }

            return resolvedAddress;
        }

        private void TraceConfigValue(string key, object value)
        {
            this.logger.Info($"{key} -> {value}");
        }
    }
}