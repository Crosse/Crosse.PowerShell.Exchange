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
