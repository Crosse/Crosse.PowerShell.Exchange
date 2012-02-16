/*
 * Copyright (c) 2011 Seth Wright <wrightst@jmu.edu>
 *
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 *
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.Exchange.Data.Directory;
using Microsoft.Exchange.Data;
using Microsoft.Exchange.Data.Common;
using Microsoft.Exchange.Data.Directory.Management;
using System.Management.Automation;
using System.IO;

namespace Jmu.Exchange.Provisioning
{
    /// <summary>
    /// Provisions a user with a local or remote mailbox.
    /// </summary>
    [Cmdlet(VerbsCommon.Add, "ProvisionedMailbox")]
    [CmdletBinding(ConfirmImpact = ConfirmImpact.High,
        SupportsShouldProcess = true)]
    public class Add_ProvisionedMailbox : PSCmdlet
    {
        #region Parameters
        [Parameter(Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Name")]
        [Alias("id")]
        public string Identity;

        [Parameter(Mandatory = true,
            ParameterSetName = "Local")]
        public SwitchParameter Local;

        //[Parameter(Mandatory = true,
        //    ParameterSetName = "Remote")]
        //public SwitchParameter Remote;

        //[Parameter(Mandatory = true,
        //    ParameterSetName = "Remote",
        //    ValueFromPipelineByPropertyName = true)]
        //public SmtpAddress ExternalEmailAddress;

        [Parameter(Mandatory = false)]
        public SwitchParameter UseDefaultCredential;

        [Parameter(Mandatory = false)]
        [Credential()]
        public PSCredential Credential;

        #region EmailParameters
        ///// <summary>
        ///// Specifies whether or not email notifications will be sent about the new mailbox.
        ///// </summary>
        //[Parameter(Mandatory = true,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = true,
        //    ParameterSetName = "RemoteEmail")]
        //public SwitchParameter SendEmailNotification = false;

        ///// <summary>
        ///// The email address that should be used as the From: address when sending emails.
        ///// </summary>
        //[Parameter(Mandatory = true,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = true,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public SmtpAddress EmailFrom;

        ///// <summary>
        ///// The SMTP server used to send the emails.
        ///// </summary>
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public string SmtpServer;

        ///// <summary>
        ///// The path to a file containing the template used to send the "welcome" email to a user who receives a local mailbox.
        ///// </summary>
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public FileInfo LocalWelcomeEmailTemplate;

        ///// <summary>
        ///// The path to a file containing the template used to send the "welcome" email to a user who receives a remote mailbox.
        ///// </summary>
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public FileInfo RemoteWelcomeEmailTemplate;

        ///// <summary>
        ///// The path to a file containing the template used to send the "notification" email to a user who receives a local mailbox.
        ///// </summary>
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public FileInfo LocalNotificationEmailTemplate;

        ///// <summary>
        ///// The path to a file containing the template used to send the "notification" email to a user who receives a remote mailbox.
        ///// </summary>
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "LocalEmail")]
        //[Parameter(Mandatory = false,
        //    ParameterSetName = "RemoteEmail")]
        //[ValidateNotNullOrEmpty()]
        //public FileInfo RemoteNotificationEmailTemplate;
        #endregion

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty()]
        public string DomainController;
        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            WriteVerbose("Performing initialization actions");
            if (DomainController == null)
            {
                DomainController = Domain.GetCurrentDomain().FindDomainController(LocatorOptions.WriteableRequired).Name;
            }
            WriteVerbose(String.Format("Using domain controller {0}", DomainController));

            if (UseDefaultCredential && Credential != null)
            {
                ThrowTerminatingError(new ErrorRecord(
                    new PSArgumentException("Specify either LocalUseDefaultCredential or LocalCredential, but not both."),
                    "ArgumentNullException", ErrorCategory.InvalidArgument, null));
            }
            if (Credential != null)
            {
                WriteVerbose(String.Format("Using explicit credential with username of {0}", Credential.UserName));
            }

            WriteVerbose("Initialization complete.");
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            WriteVerbose("In ProcessRecord()");

            ProvisioningResult result = new ProvisioningResult();
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
            WriteVerbose("In EndProcessing()");
        }
    }
}
