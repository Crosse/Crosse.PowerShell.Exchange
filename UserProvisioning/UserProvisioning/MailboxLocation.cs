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

namespace Jmu.Exchange.Provisioning
{
    /// <summary>
    /// Describes the available locations for mailbox provisioning.
    /// </summary>
    public enum MailboxLocation
    {
        /// <summary>
        /// The mailbox should be created locally; i.e., in the on-premise Exchange
        /// </summary>
        Local,

        /// <summary>
        /// The mailbox should be created remotely; i.e., Office356 or Live@edu
        /// </summary>
        Remote
    }
}
