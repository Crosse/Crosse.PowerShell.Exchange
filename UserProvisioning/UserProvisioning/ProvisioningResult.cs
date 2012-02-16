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
using Microsoft.Exchange.Data;
using Microsoft.Exchange.Data.Directory;
using Microsoft.Exchange.Data.Directory.Recipient;

namespace Jmu.Exchange.Provisioning
{
    /// <summary>
    /// Contains the results of a provisioning request.
    /// </summary>
    public class ProvisioningResult
    {
        /// <summary>
        /// Gets the identity of the provisioned user.
        /// </summary>
        public ADObjectId Identity { get; internal set; }

        /// <summary>
        /// Gets the requested mailbox location.
        /// </summary>
        public MailboxLocation? RequestedMailboxLocation { get; internal set; }

        /// <summary>
        /// Gets the original state of the user object.
        /// </summary>
        public RecipientTypeDetails? OriginalState { get; internal set; }

        /// <summary>
        /// Gets the ending state of the user object.
        /// </summary>
        public RecipientTypeDetails? EndingState { get; internal set; }

        /// <summary>
        /// Gets the mail contact for the user, if one was created.
        /// </summary>
        public ADObjectId MailContact { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether a "welcome" email was sent to the user.
        /// </summary>
        public bool? WelcomeEmailSent { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether a "notification" email was sent to the user.
        /// </summary>
        public bool? NotificationEmailSent { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether the provisioning process was successful.
        /// </summary>
        public bool? ProvisioningSuccessful { get; internal set; }

        /// <summary>
        /// Gets the error, if any.
        /// </summary>
        public Exception Error { get; internal set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ProvisioningResult"/> class.
        /// </summary>
        public ProvisioningResult()
        {
            this.Identity = null;
            this.RequestedMailboxLocation = null;
            this.OriginalState = null;
            this.EndingState = null;
            this.MailContact = null;
            this.WelcomeEmailSent = null;
            this.NotificationEmailSent = null;
            this.ProvisioningSuccessful = null;
            this.Error = null;
        }
    }
}
