//------------------------------------------------------------------------------
// <copyright file="StopEggTimerCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Lync.Model;
using System.Collections.Generic;

namespace EggTimer
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class StopEggTimerCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4129;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid MenuGroup = new Guid("98da75a3-c276-4975-816d-7aa6b0161f8f");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="StopEggTimerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private StopEggTimerCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                CommandID menuCommandID = new CommandID(MenuGroup, CommandId);
                EventHandler eventHandler = this.HangleCommandClick;
                MenuCommand menuItem = new MenuCommand(eventHandler, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static StopEggTimerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new StopEggTimerCommand(package);
        }

        private void HangleCommandClick(object sender, EventArgs e)
        {
            try
            {
                var client = LyncClient.GetClient();

                var updateInfo = new List<KeyValuePair<PublishableContactInformationType, object>>();
                updateInfo.Add(new KeyValuePair<PublishableContactInformationType, object>(PublishableContactInformationType.Availability, Microsoft.Lync.Controls.ContactAvailability.Free));
                updateInfo.Add(new KeyValuePair<PublishableContactInformationType, object>(PublishableContactInformationType.PersonalNote, string.Empty));
                client.Self.BeginPublishContactInformation(updateInfo, (ar) => { client.Self.EndPublishContactInformation(ar); }, null);
            }
            catch (ClientNotFoundException)
            {
                ShowMessageBox("The Lync client is not found");
            }
            catch (LyncClientException lce)
            {
                ShowMessageBox("Lync client exception on GetClient: " + lce.Message);
            }
            catch (Exception ex)
            {
                ShowMessageBox(ex.ToString());
            }
        }

        private void ShowMessageBox(string message)
        {
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)Package.GetGlobalService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                0,
                ref clsid,
                "StopEggTimerCommandPackage",
                string.Format(CultureInfo.CurrentCulture, message, this.GetType().FullName),
                string.Empty,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                OLEMSGICON.OLEMSGICON_INFO,
                0,        // false
                out result));
        }
    }
}
