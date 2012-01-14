using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.Exchange.Data.Directory;
using Microsoft.Exchange.Data;
using Microsoft.Exchange.Data.Common;
using System.Management.Automation;
using System.IO;

namespace Jmu.Exchange.Provisioning
{
    /// <summary>
    /// Provisions a user with a local or remote mailbox.
    /// </summary>
    [Cmdlet(VerbsCommon.Add, "ProvisionedMailbox")]
    [CmdletBinding(ConfirmImpact = ConfirmImpact.High,
        SupportsShouldProcess = true,
        DefaultParameterSetName = "DefaultParameterSet")]
    public class Add_ProvisionedMailbox : PSCmdlet
    {
        /// <summary>
        /// The Identity parameter specifies the user to be provisioned.  The user must already exist in Active Directory.
        /// </summary>
        [Parameter(Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        //[Parameter(Mandatory = true,
        //    Position = 0,
        //    ValueFromPipeline = true,
        //    ParameterSetName = "SendEmails")]
        [ValidateNotNullOrEmpty()]
        [Alias("Name")]
        public string Identity;

        ///// <summary>
        ///// The MailboxLocation parameter specifies the location where the mailbox should be provisioned.
        ///// </summary>
        //[Parameter(Mandatory = true,
        //    ValueFromPipelineByPropertyName = true,
        //    ParameterSetName = "DefaultParameterSet")]
        //[Parameter(Mandatory = true,
        //    ValueFromPipelineByPropertyName = true,
        //    ParameterSetName = "SendEmails")]
        //public MailboxLocation MailboxLocation;

        [Parameter(Mandatory = true,
            ParameterSetName = "LocalNoEmail")]
        [Parameter(Mandatory = true,
            ParameterSetName = "LocalEmail")]
        public SwitchParameter Local;

        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteNoEmail")]
        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteEmail")]
        public SwitchParameter Remote;

        /// <summary>
        /// When <see cref="MailboxLocation"/> is "Remote", the email address of the user's mailbox.
        /// </summary>
        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteNoEmail",
            ValueFromPipelineByPropertyName = true)]
        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteEmail",
            ValueFromPipelineByPropertyName = true)]
        public SmtpAddress ExternalEmailAddress;

        /// <summary>
        /// Specifies whether or not email notifications will be sent about the new mailbox.
        /// </summary>
        [Parameter(Mandatory = true,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteEmail")]
        public SwitchParameter SendEmailNotification = false;

        /// <summary>
        /// The email address that should be used as the From: address when sending emails.
        /// </summary>
        [Parameter(Mandatory = true,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = true,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public SmtpAddress EmailFrom;

        /// <summary>
        /// The SMTP server used to send the emails.
        /// </summary>
        [Parameter(Mandatory = false,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = false,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public string SmtpServer;

        /// <summary>
        /// The path to a file containing the template used to send the "welcome" email to a user who receives a local mailbox.
        /// </summary>
        [Parameter(Mandatory = false,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = false,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public FileInfo LocalWelcomeEmailTemplate;

        /// <summary>
        /// The path to a file containing the template used to send the "welcome" email to a user who receives a remote mailbox.
        /// </summary>
        [Parameter(Mandatory = false,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = false,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public FileInfo RemoteWelcomeEmailTemplate;

        /// <summary>
        /// The path to a file containing the template used to send the "notification" email to a user who receives a local mailbox.
        /// </summary>
        [Parameter(Mandatory = false,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = false,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public FileInfo LocalNotificationEmailTemplate;

        /// <summary>
        /// The path to a file containing the template used to send the "notification" email to a user who receives a remote mailbox.
        /// </summary>
        [Parameter(Mandatory = false,
            ParameterSetName = "LocalEmail")]
        [Parameter(Mandatory = false,
            ParameterSetName = "RemoteEmail")]
        [ValidateNotNullOrEmpty()]
        public FileInfo RemoteNotificationEmailTemplate;

        /// <summary>
        /// The DomainController parameter specifies the fully qualified domain name (FQDN) of the domain controller that retrieves data from Active Directory.
        /// </summary>
        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty()]
        public string DomainController;

        private static string modulePath;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (String.IsNullOrEmpty(modulePath))
                modulePath = SessionState.Module.Path;

            WriteVerbose("Performing initialization actions");
            if (DomainController == null)
            {
                DomainController = Domain.GetCurrentDomain().FindDomainController(LocatorOptions.WriteableRequired).Name;
            }
            WriteVerbose(String.Format("Using domain controller {0}", DomainController));

                        if (SendEmailNotification && LocalWelcomeEmailTemplate == null)
            {
                LocalWelcomeEmailTemplate = new FileInfo(Path.Combine(modulePath, "LocalWelcomeEmailTemplate.html"));
                if (!LocalWelcomeEmailTemplate.Exists)
                    throw new ParameterBindingException("LocalWelcomeEmailTemplate");
            }

            if (SendEmailNotification && RemoteWelcomeEmailTemplate == null)
            {
                RemoteWelcomeEmailTemplate = new FileInfo(Path.Combine(modulePath, "RemoteWelcomeEmailTemplate.html"));
                if (!RemoteNotificationEmailTemplate.Exists)
                    throw new ParameterBindingException("RemoteWelcomeEmailTemplate");
            }

            if (SendEmailNotification && LocalNotificationEmailTemplate == null)
            {
                LocalNotificationEmailTemplate = new FileInfo(Path.Combine(modulePath, "LocalNotificationEmailTemplate.html"));
                if (!LocalNotificationEmailTemplate.Exists)
                    throw new ParameterBindingException("LocalNotificationEmailTemplate");
            }

            if (SendEmailNotification && RemoteNotificationEmailTemplate == null)
            {
                RemoteNotificationEmailTemplate = new FileInfo(Path.Combine(modulePath, "RemoteNotificationEmailTemplate.html"));
                if (!RemoteNotificationEmailTemplate.Exists)
                    throw new ParameterBindingException("RemoteNotificationEmailTemplate");
            }

        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            WriteVerbose("In ProcessRecord()");
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
            WriteVerbose("In EndProcessing()");
        }
    }
}
