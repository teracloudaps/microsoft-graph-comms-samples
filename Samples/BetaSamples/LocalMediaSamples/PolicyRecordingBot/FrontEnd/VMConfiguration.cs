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
        /// <summary>
        /// App setting key for the public service DNS name.
        /// </summary>
        private const string ServiceDnsNameKey = "ServiceDnsName";

        /// <summary>
        /// App setting key for the bot application id.
        /// </summary>
        private const string AadAppIdKey = "AadAppId";

        /// <summary>
        /// App setting key for the bot application secret.
        /// </summary>
        private const string AadAppSecretKey = "AadAppSecret";

        /// <summary>
        /// App setting key for the Microsoft Graph base endpoint.
        /// </summary>
        private const string PlaceCallEndpointUrlKey = "PlaceCallEndpointUrl";

        /// <summary>
        /// App setting key for the default certificate thumbprint.
        /// </summary>
        private const string DefaultCertificateKey = "DefaultCertificate";

        /// <summary>
        /// Alternate app setting key for the certificate thumbprint.
        /// </summary>
        private const string CertificateThumbprintKey = "CertificateThumbprint";

        /// <summary>
        /// App setting key for the call-control port.
        /// </summary>
        private const string CallControlPortKey = "CallControlPort";

        /// <summary>
        /// App setting key for the media port.
        /// </summary>
        private const string MediaPortKey = "MediaPort";

        /// <summary>
        /// App setting key for the public media IP address.
        /// </summary>
        private const string InstancePublicIpAddressKey = "InstancePublicIPAddress";

        /// <summary>
        /// App setting key for the media service FQDN.
        /// </summary>
        private const string ServiceFqdnKey = "ServiceFqdn";

        /// <summary>
        /// App setting key for the public call-control host.
        /// </summary>
        private const string CallControlHostKey = "CallControlHost";

        /// <summary>
        /// App setting key for the local call-control listener host.
        /// </summary>
        private const string CallControlListenHostKey = "CallControlListenHost";

        /// <summary>
        /// Prefix used for environment-variable fallback settings.
        /// </summary>
        private const string EnvironmentVariablePrefix = "POLICY_RECORDING_BOT_";

        /// <summary>
        /// Default local call-control listener host for VM hosting.
        /// </summary>
        private const string DefaultCallControlListenHost = "0.0.0.0";

        /// <summary>
        /// Default HTTPS call-control port for VM hosting.
        /// </summary>
        private const int DefaultCallControlPort = 9442;

        /// <summary>
        /// Default media port for VM hosting.
        /// </summary>
        private const int DefaultMediaPort = 8445;

        /// <summary>
        /// Graph logger.
        /// </summary>
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
        public MediaPlatformSettings MediaPlatformSettings { get; private set; }

        /// <inheritdoc/>
        public void Dispose()
        {
        }

        /// <summary>
        /// Gets a required setting.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <returns>The configured value.</returns>
        private static string GetRequiredSetting(string key)
        {
            var value = VMConfiguration.GetOptionalSetting(key);
            if (string.IsNullOrWhiteSpace(value) || VMConfiguration.IsPlaceholderValue(value))
            {
                throw new ConfigurationErrorsException($"Missing required setting '{key}'. Set it in FrontEnd App.config or environment variable '{EnvironmentVariablePrefix}{key}'.");
            }

            return value.Trim();
        }

        /// <summary>
        /// Gets an optional setting from app settings or environment variables.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <returns>The configured value, or null when not configured.</returns>
        private static string GetOptionalSetting(string key)
        {
            var value = ConfigurationManager.AppSettings[key];
            if (!string.IsNullOrWhiteSpace(value) && !VMConfiguration.IsPlaceholderValue(value))
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

        /// <summary>
        /// Gets a TCP port setting.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns>The configured TCP port.</returns>
        private static int GetIntSetting(string key, int defaultValue)
        {
            var value = VMConfiguration.GetOptionalSetting(key);
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

        /// <summary>
        /// Checks whether a configured value is still a deployment placeholder.
        /// </summary>
        /// <param name="value">The setting value.</param>
        /// <returns>True when the value is a placeholder; otherwise, false.</returns>
        private static bool IsPlaceholderValue(string value)
        {
            var trimmed = value.Trim();
            return (trimmed.StartsWith("%", StringComparison.Ordinal) && trimmed.EndsWith("%", StringComparison.Ordinal)) ||
                   (trimmed.StartsWith("$", StringComparison.Ordinal) && trimmed.EndsWith("$", StringComparison.Ordinal));
        }

        /// <summary>
        /// Normalizes a certificate thumbprint for certificate-store lookup.
        /// </summary>
        /// <param name="thumbprint">The certificate thumbprint.</param>
        /// <returns>The normalized thumbprint.</returns>
        private static string NormalizeThumbprint(string thumbprint)
        {
            return thumbprint?.Replace(" ", string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>
        /// Builds the local call-control listener URI.
        /// </summary>
        /// <param name="host">The listener host.</param>
        /// <param name="port">The listener port.</param>
        /// <returns>The listener URI.</returns>
        private static Uri BuildCallControlListenUri(string host, int port)
        {
            var trimmedHost = (host ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(trimmedHost) || trimmedHost == "+" || trimmedHost == "*")
            {
                throw new ConfigurationErrorsException($"Setting '{CallControlListenHostKey}' must be a URI-compatible host name or IP address. Use '+' only in the Windows URL ACL command, not in application configuration.");
            }

            return new Uri($"https://{trimmedHost}:{port}/", UriKind.Absolute);
        }

        /// <summary>
        /// Initializes configuration values from app settings and environment fallback.
        /// </summary>
        private void Initialize()
        {
            this.ServiceDnsName = VMConfiguration.GetRequiredSetting(ServiceDnsNameKey);
            this.AadAppId = VMConfiguration.GetRequiredSetting(AadAppIdKey);
            this.AadAppSecret = VMConfiguration.GetRequiredSetting(AadAppSecretKey);

            var callControlPort = VMConfiguration.GetIntSetting(CallControlPortKey, DefaultCallControlPort);
            var mediaPort = VMConfiguration.GetIntSetting(MediaPortKey, DefaultMediaPort);
            var callControlHost = VMConfiguration.GetOptionalSetting(CallControlHostKey) ?? this.ServiceDnsName;
            var callControlListenHost = VMConfiguration.GetOptionalSetting(CallControlListenHostKey) ?? DefaultCallControlListenHost;
            var serviceFqdn = VMConfiguration.GetOptionalSetting(ServiceFqdnKey) ?? this.ServiceDnsName;
            var endpoint = VMConfiguration.GetOptionalSetting(PlaceCallEndpointUrlKey) ?? "https://graph.microsoft.com/v1.0";

            this.PlaceCallEndpointUrl = new Uri(endpoint, UriKind.Absolute);
            this.CallControlBaseUrl = new Uri($"https://{callControlHost}:{callControlPort}/{HttpRouteConstants.CallSignalingRoutePrefix}/{HttpRouteConstants.OnNotificationRequestRoute}");
            this.CallControlListeningUrls = new[] { VMConfiguration.BuildCallControlListenUri(callControlListenHost, callControlPort) };

            var certificate = this.GetCertificateFromStore(this.GetCertificateThumbprint());
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

        /// <summary>
        /// Gets the configured certificate thumbprint.
        /// </summary>
        /// <returns>The normalized certificate thumbprint.</returns>
        private string GetCertificateThumbprint()
        {
            return VMConfiguration.NormalizeThumbprint(VMConfiguration.GetOptionalSetting(CertificateThumbprintKey) ?? VMConfiguration.GetRequiredSetting(DefaultCertificateKey));
        }

        /// <summary>
        /// Gets the configured certificate from the local machine certificate store.
        /// </summary>
        /// <param name="thumbprint">The certificate thumbprint.</param>
        /// <returns>The certificate.</returns>
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

        /// <summary>
        /// Gets the public IP address used by the media platform.
        /// </summary>
        /// <param name="serviceFqdn">The service FQDN to resolve when no IP address is configured.</param>
        /// <returns>The public IP address.</returns>
        private IPAddress GetPublicIpAddress(string serviceFqdn)
        {
            var configuredIp = VMConfiguration.GetOptionalSetting(InstancePublicIpAddressKey);
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

        /// <summary>
        /// Writes a non-secret configuration value to the graph logger.
        /// </summary>
        /// <param name="key">The configuration key.</param>
        /// <param name="value">The configuration value.</param>
        private void TraceConfigValue(string key, object value)
        {
            this.logger.Info($"{key} -> {value}");
        }
    }
}