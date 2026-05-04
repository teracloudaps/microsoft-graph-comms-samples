// <copyright file="Program.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.PolicyRecordingBot.FrontEnd
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Common.Telemetry;

    /// <summary>
    /// Console entry point for VM-hosted deployments.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// Starts the VM-hosted bot process.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        /// <returns>Zero when the process exits cleanly; otherwise, a non-zero value.</returns>
        public static int Main(string[] args)
        {
            IGraphLogger logger = new GraphLogger("PolicyRecordingBot", redirectToTrace: true);
            VMConfiguration configuration = null;
            var serviceStarted = false;

            try
            {
                Program.WriteBanner();

                configuration = new VMConfiguration(logger);
                Service.Instance.Initialize(configuration, logger);
                Service.Instance.Start();
                serviceStarted = true;

                Program.WriteStartupSummary(configuration);

                using (var shutdownSignal = new ManualResetEventSlim(false))
                {
                    Console.CancelKeyPress += (sender, eventArgs) =>
                    {
                        eventArgs.Cancel = true;
                        shutdownSignal.Set();
                    };

                    Task.Run(() =>
                    {
                        Console.ReadLine();
                        shutdownSignal.Set();
                    });

                    Console.WriteLine("Press ENTER or Ctrl+C to stop.");
                    shutdownSignal.Wait();
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Startup failed.");
                Program.WriteException(ex);
                return 1;
            }
            finally
            {
                if (serviceStarted)
                {
                    Program.StopService();
                }

                configuration?.Dispose();
            }
        }

        /// <summary>
        /// Writes the console banner.
        /// </summary>
        private static void WriteBanner()
        {
            Console.WriteLine("================================================");
            Console.WriteLine("Policy Recording Bot - VM Edition");
            Console.WriteLine("================================================");
            Console.WriteLine($"Starting at {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}");
            Console.WriteLine();
        }

        /// <summary>
        /// Writes the startup summary.
        /// </summary>
        /// <param name="configuration">The VM configuration.</param>
        private static void WriteStartupSummary(VMConfiguration configuration)
        {
            Console.WriteLine("Bot started successfully.");
            Console.WriteLine($"Call control callback: {configuration.CallControlBaseUrl}");
            Console.WriteLine($"Call control listeners: {string.Join(", ", configuration.CallControlListeningUrls)}");
            Console.WriteLine($"Media endpoint: {configuration.MediaPlatformSettings.MediaPlatformInstanceSettings.InstancePublicIPAddress}:{configuration.MediaPlatformSettings.MediaPlatformInstanceSettings.InstancePublicPort}");
            Console.WriteLine($"Communications client initialized: {Bot.Bot.Instance.Client != null}");
            Console.WriteLine();
        }

        /// <summary>
        /// Stops the service singleton.
        /// </summary>
        private static void StopService()
        {
            try
            {
                Service.Instance.Stop();
                Console.WriteLine("Bot stopped.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error while stopping bot: {ex.Message}");
            }
        }

        /// <summary>
        /// Writes exception details, including inner exceptions.
        /// </summary>
        /// <param name="exception">The exception to write.</param>
        private static void WriteException(Exception exception)
        {
            for (var current = exception; current != null; current = current.InnerException)
            {
                Console.Error.WriteLine(current.Message);
                Console.Error.WriteLine(current.StackTrace);
            }
        }
    }
}