# Enable-EXOMailboxSOA-Bulk.ps1

Bulk enable/disable **Exchange mailbox SOA** (Source of Authority) for directory-synchronized mailboxes in **Exchange Online**, based on a simple CSV input file.

In hybrid environments, some Exchange mailbox attributes can be managed either from **on-premises** or from **Exchange Online**. This script helps you **switch the mailbox attribute SOA** to/from cloud management in a controlled, logged, repeatable way across many mailboxes.

---

## What this script is for

If you are:

- Running **hybrid** with directory synchronization (Entra Connect / AAD Connect / Cloud Sync),
- Migrating mailboxes to Exchange Online,
- Working toward “**last Exchange server**” retirement, or
- Wanting to manage certain Exchange attributes **in the cloud** instead of on-prem,

…then you may need to set the mailbox flag that controls whether Exchange Online is the **source of authority** for Exchange mailbox attributes.

This script bulk-updates that mailbox flag using a CSV list.
