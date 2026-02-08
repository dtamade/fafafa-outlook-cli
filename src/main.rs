//! CLI for Outlook email fetcher
//!
//! Usage:
//! ```text
//! outlook-cli [OPTIONS] <COMMAND>
//! ```
//!
//! Commands:
//!   inbox    - List inbox emails
//!   unread   - List unread emails
//!   search   - Search emails by subject
//!   read     - Mark email as read
//!   folders  - List mail folders
//!   send     - Send an email
//!   reply    - Reply to an email
//!   forward  - Forward an email
//!   drafts   - List drafts
//!   events   - Calendar commands
//!   contacts - Contact commands

use clap::{Parser, Subcommand};
use fafafa_outlook_core::{
    DateTimeTimeZone, DraftMessage, NewCalendarEvent, NewContact, NewMessage, OutlookClient,
};

const DEFAULT_CLIENT_ID: &str = fafafa_outlook_core::auth::DEFAULT_CLIENT_ID;

#[derive(Parser)]
#[command(name = "outlook-cli")]
#[command(about = "Outlook email fetcher CLI", long_about = None)]
struct Cli {
    /// Microsoft OAuth refresh token
    #[arg(short, long, env = "OUTLOOK_REFRESH_TOKEN")]
    token: String,

    /// Azure AD application ID
    #[arg(
        short,
        long,
        env = "OUTLOOK_CLIENT_ID",
        default_value_t = DEFAULT_CLIENT_ID.to_string()
    )]
    client_id: String,

    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// List inbox emails
    Inbox {
        /// Number of emails to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
    },
    /// List unread emails
    Unread {
        /// Number of emails to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
    },
    /// Search emails by subject
    Search {
        /// Search query
        query: String,
        /// Number of results to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
    },
    /// Mark email as read
    Read {
        /// Email ID
        id: String,
    },
    /// List mail folders
    Folders,
    /// Get current user info
    Me,
    /// Send an email
    Send {
        /// Recipient email address
        to: String,
        /// Email subject
        #[arg(short, long)]
        subject: String,
        /// Email body
        #[arg(short, long)]
        body: String,
        /// Send as HTML (default: plain text)
        #[arg(long)]
        html: bool,
    },
    /// Reply to an email
    Reply {
        /// Email ID to reply to
        id: String,
        /// Reply message
        #[arg(short, long)]
        message: String,
        /// Reply to all recipients
        #[arg(long)]
        all: bool,
    },
    /// Forward an email
    Forward {
        /// Email ID to forward
        id: String,
        /// Recipient email addresses (comma-separated)
        #[arg(short, long)]
        to: String,
        /// Optional comment
        #[arg(short, long)]
        comment: Option<String>,
    },
    /// Delete an email
    Delete {
        /// Email ID
        id: String,
    },
    /// Get email details
    Get {
        /// Email ID
        id: String,
    },
    /// List attachments for an email
    Attachments {
        /// Email ID
        id: String,
    },
    /// Download an attachment
    Download {
        /// Email ID
        #[arg(short, long)]
        email_id: String,
        /// Attachment ID
        #[arg(short, long)]
        attachment_id: String,
        /// Output file path
        #[arg(short, long)]
        output: String,
    },
    /// Poll for new messages
    Poll {
        /// ISO 8601 datetime to poll from
        since: String,
        /// Number of messages to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
    },
    /// Get unread message count
    UnreadCount,

    // ==================== Drafts ====================
    /// List draft emails
    Drafts {
        /// Number of drafts to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
    },
    /// Create a draft email
    CreateDraft {
        /// Email subject
        #[arg(short, long)]
        subject: Option<String>,
        /// Email body
        #[arg(short, long)]
        body: Option<String>,
        /// Recipients (comma-separated)
        #[arg(short, long)]
        to: Option<String>,
        /// Send as HTML
        #[arg(long)]
        html: bool,
    },
    /// Send a draft email
    SendDraft {
        /// Draft ID
        id: String,
    },

    // ==================== Calendar ====================
    /// List calendar events
    Events {
        /// Number of events to show
        #[arg(short, long, default_value = "10")]
        limit: u32,
        /// Start date (ISO 8601)
        #[arg(long)]
        start: Option<String>,
        /// End date (ISO 8601)
        #[arg(long)]
        end: Option<String>,
    },
    /// Get event details
    Event {
        /// Event ID
        id: String,
    },
    /// Create a calendar event
    CreateEvent {
        /// Event subject
        #[arg(short, long)]
        subject: String,
        /// Start datetime (ISO 8601)
        #[arg(long)]
        start: String,
        /// End datetime (ISO 8601)
        #[arg(long)]
        end: String,
        /// Timezone (default: UTC)
        #[arg(long, default_value = "UTC")]
        timezone: String,
        /// Location
        #[arg(short, long)]
        location: Option<String>,
        /// Attendees (comma-separated emails)
        #[arg(short, long)]
        attendees: Option<String>,
        /// All day event
        #[arg(long)]
        all_day: bool,
        /// Online meeting
        #[arg(long)]
        online: bool,
    },
    /// Delete a calendar event
    DeleteEvent {
        /// Event ID
        id: String,
    },
    /// Accept a calendar event invitation
    AcceptEvent {
        /// Event ID
        id: String,
        /// Optional comment
        #[arg(short, long)]
        comment: Option<String>,
    },
    /// Decline a calendar event invitation
    DeclineEvent {
        /// Event ID
        id: String,
        /// Optional comment
        #[arg(short, long)]
        comment: Option<String>,
    },

    // ==================== Contacts ====================
    /// List contacts
    Contacts {
        /// Number of contacts to show
        #[arg(short, long, default_value = "20")]
        limit: u32,
        /// Search query
        #[arg(short, long)]
        search: Option<String>,
    },
    /// Get contact details
    Contact {
        /// Contact ID
        id: String,
    },
    /// Create a contact
    CreateContact {
        /// First name
        #[arg(long)]
        first_name: String,
        /// Last name
        #[arg(long)]
        last_name: String,
        /// Email address
        #[arg(short, long)]
        email: Option<String>,
        /// Mobile phone
        #[arg(short, long)]
        mobile: Option<String>,
        /// Company name
        #[arg(short, long)]
        company: Option<String>,
        /// Job title
        #[arg(short, long)]
        job_title: Option<String>,
    },
    /// Delete a contact
    DeleteContact {
        /// Contact ID
        id: String,
    },
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // Load .env file
    dotenvy::dotenv().ok();

    let cli = Cli::parse();

    let client = OutlookClient::with_credentials(&cli.client_id, &cli.token).await?;

    match cli.command {
        Commands::Inbox { limit } => {
            let messages = client.inbox(limit).await?;
            print_messages(&messages);
        }
        Commands::Unread { limit } => {
            let messages = client.unread(limit).await?;
            if messages.is_empty() {
                println!("No unread messages");
            } else {
                print_messages(&messages);
            }
        }
        Commands::Search { query, limit } => {
            let messages = client.search_by_subject(&query, limit).await?;
            if messages.is_empty() {
                println!("No messages found for: {}", query);
            } else {
                print_messages(&messages);
            }
        }
        Commands::Read { id } => {
            client.mark_as_read(&id).await?;
            println!("Marked as read: {}", id);
        }
        Commands::Folders => {
            let folders = client.list_folders().await?;
            println!("Mail Folders:");
            println!();
            for folder in folders {
                let unread = folder.unread_item_count.unwrap_or(0);
                let total = folder.total_item_count.unwrap_or(0);
                let unread_str = if unread > 0 {
                    format!(" ({} unread)", unread)
                } else {
                    String::new()
                };
                println!("  {} - {} items{}", folder.display_name, total, unread_str);
            }
        }
        Commands::Me => {
            let user = client.get_me().await?;
            println!("User: {}", user.display_name.unwrap_or_default());
            println!("Email: {}", user.mail.unwrap_or_default());
        }
        Commands::Send {
            to,
            subject,
            body,
            html,
        } => {
            let message = if html {
                NewMessage::html(&to, &subject, &body)
            } else {
                NewMessage::text(&to, &subject, &body)
            };
            client.send_mail(message).await?;
            println!("Email sent to: {}", to);
        }
        Commands::Reply { id, message, all } => {
            if all {
                client.reply_all(&id, &message).await?;
                println!("Replied all to: {}", id);
            } else {
                client.reply(&id, &message).await?;
                println!("Replied to: {}", id);
            }
        }
        Commands::Forward { id, to, comment } => {
            let recipients: Vec<&str> = to.split(',').map(|s| s.trim()).collect();
            client.forward(&id, &recipients, comment.as_deref()).await?;
            println!("Forwarded to: {}", to);
        }
        Commands::Delete { id } => {
            client.delete_message(&id).await?;
            println!("Deleted: {}", id);
        }
        Commands::Get { id } => {
            let msg = client.get_message_with_body(&id).await?;
            println!(
                "Subject: {}",
                msg.subject.as_deref().unwrap_or("(no subject)")
            );
            println!(
                "From: {}",
                msg.from
                    .as_ref()
                    .map(|r| r.email_address.address.as_str())
                    .unwrap_or("unknown")
            );
            println!(
                "Date: {}",
                msg.received_date_time
                    .map(|d| d.to_string())
                    .unwrap_or_default()
            );
            println!(
                "Read: {}",
                if msg.is_read.unwrap_or(false) {
                    "Yes"
                } else {
                    "No"
                }
            );
            println!();
            if let Some(body) = msg.body {
                println!("{}", body.content);
            }
        }
        Commands::Attachments { id } => {
            let attachments = client.list_attachments(&id).await?;
            if attachments.is_empty() {
                println!("No attachments");
            } else {
                println!("Attachments:");
                for att in attachments {
                    let size = att
                        .size
                        .map(|s| format!(" ({} bytes)", s))
                        .unwrap_or_default();
                    println!("  {} - {}{}", att.id, att.name, size);
                }
            }
        }
        Commands::Download {
            email_id,
            attachment_id,
            output,
        } => {
            let bytes = client
                .download_attachment(&email_id, &attachment_id)
                .await?;
            std::fs::write(&output, &bytes)?;
            println!("Downloaded {} bytes to: {}", bytes.len(), output);
        }
        Commands::Poll { since, limit } => {
            let messages = client.poll_new_messages(&since, limit).await?;
            if messages.is_empty() {
                println!("No new messages since {}", since);
            } else {
                println!("New messages since {}:", since);
                print_messages(&messages);
            }
        }
        Commands::UnreadCount => {
            let count = client.unread_count().await?;
            println!("Unread messages: {}", count);
        }

        // ==================== Drafts ====================
        Commands::Drafts { limit } => {
            let drafts = client.list_drafts(limit).await?;
            if drafts.is_empty() {
                println!("No drafts");
            } else {
                println!("Drafts:");
                print_messages(&drafts);
            }
        }
        Commands::CreateDraft {
            subject,
            body,
            to,
            html,
        } => {
            let mut draft = DraftMessage::new();
            if let Some(s) = subject {
                draft = draft.subject(s);
            }
            if let Some(b) = body {
                if html {
                    draft = draft.body_html(b);
                } else {
                    draft = draft.body_text(b);
                }
            }
            if let Some(t) = to {
                let recipients: Vec<&str> = t.split(',').map(|s| s.trim()).collect();
                draft = draft.to(&recipients);
            }
            let created = client.create_draft(draft).await?;
            println!("Draft created: {}", created.id);
        }
        Commands::SendDraft { id } => {
            client.send_draft(&id).await?;
            println!("Draft sent: {}", id);
        }

        // ==================== Calendar ====================
        Commands::Events { limit, start, end } => {
            let events = if let (Some(s), Some(e)) = (start, end) {
                client.list_events_range(&s, &e, limit).await?
            } else {
                client.list_events(limit).await?
            };
            if events.is_empty() {
                println!("No events");
            } else {
                println!("Calendar Events:");
                for event in events {
                    let start_str = event
                        .start
                        .as_ref()
                        .map(|d| d.date_time.as_str())
                        .unwrap_or("?");
                    let subject = event.subject.as_deref().unwrap_or("(no subject)");
                    println!("  {} - {}", start_str, subject);
                    println!("    ID: {}", event.id);
                }
            }
        }
        Commands::Event { id } => {
            let event = client.get_event(&id).await?;
            println!(
                "Subject: {}",
                event.subject.as_deref().unwrap_or("(no subject)")
            );
            if let Some(start) = event.start {
                println!("Start: {} ({})", start.date_time, start.time_zone);
            }
            if let Some(end) = event.end {
                println!("End: {} ({})", end.date_time, end.time_zone);
            }
            if let Some(loc) = event.location {
                if let Some(name) = loc.display_name {
                    println!("Location: {}", name);
                }
            }
            if let Some(attendees) = event.attendees {
                println!("Attendees:");
                for att in attendees {
                    println!("  - {}", att.email_address.address);
                }
            }
        }
        Commands::CreateEvent {
            subject,
            start,
            end,
            timezone,
            location,
            attendees,
            all_day,
            online,
        } => {
            let start_dt = DateTimeTimeZone::new(&start, &timezone);
            let end_dt = DateTimeTimeZone::new(&end, &timezone);
            let mut event = NewCalendarEvent::new(&subject, start_dt, end_dt);
            if let Some(loc) = location {
                event = event.location(loc);
            }
            if let Some(att) = attendees {
                let emails: Vec<&str> = att.split(',').map(|s| s.trim()).collect();
                event = event.attendees(&emails);
            }
            if all_day {
                event = event.all_day();
            }
            if online {
                event = event.online_meeting();
            }
            let created = client.create_event(event).await?;
            println!("Event created: {}", created.id);
        }
        Commands::DeleteEvent { id } => {
            client.delete_event(&id).await?;
            println!("Event deleted: {}", id);
        }
        Commands::AcceptEvent { id, comment } => {
            client.accept_event(&id, comment.as_deref()).await?;
            println!("Event accepted: {}", id);
        }
        Commands::DeclineEvent { id, comment } => {
            client.decline_event(&id, comment.as_deref()).await?;
            println!("Event declined: {}", id);
        }

        // ==================== Contacts ====================
        Commands::Contacts { limit, search } => {
            let contacts = if let Some(q) = search {
                client.search_contacts(&q, limit).await?
            } else {
                client.list_contacts(limit).await?
            };
            if contacts.is_empty() {
                println!("No contacts");
            } else {
                println!("Contacts:");
                for contact in contacts {
                    let name = contact.display_name.as_deref().unwrap_or("(no name)");
                    let email = contact
                        .email_addresses
                        .as_ref()
                        .and_then(|e| e.first())
                        .and_then(|e| e.address.as_deref())
                        .unwrap_or("");
                    println!("  {} - {}", name, email);
                    println!("    ID: {}", contact.id);
                }
            }
        }
        Commands::Contact { id } => {
            let contact = client.get_contact(&id).await?;
            println!(
                "Name: {} {}",
                contact.given_name.as_deref().unwrap_or(""),
                contact.surname.as_deref().unwrap_or("")
            );
            if let Some(emails) = contact.email_addresses {
                for email in emails {
                    if let Some(addr) = email.address {
                        println!("Email: {}", addr);
                    }
                }
            }
            if let Some(mobile) = contact.mobile_phone {
                println!("Mobile: {}", mobile);
            }
            if let Some(company) = contact.company_name {
                println!("Company: {}", company);
            }
            if let Some(title) = contact.job_title {
                println!("Title: {}", title);
            }
        }
        Commands::CreateContact {
            first_name,
            last_name,
            email,
            mobile,
            company,
            job_title,
        } => {
            let mut contact = NewContact::new(&first_name, &last_name);
            if let Some(e) = email {
                contact = contact.email(e);
            }
            if let Some(m) = mobile {
                contact = contact.mobile(m);
            }
            if let Some(c) = company {
                contact = contact.company(c);
            }
            if let Some(j) = job_title {
                contact = contact.job_title(j);
            }
            let created = client.create_contact(contact).await?;
            println!("Contact created: {}", created.id);
        }
        Commands::DeleteContact { id } => {
            client.delete_contact(&id).await?;
            println!("Contact deleted: {}", id);
        }
    }

    Ok(())
}

fn print_messages(messages: &[fafafa_outlook_core::Message]) {
    if messages.is_empty() {
        println!("(no messages)");
        return;
    }

    for (i, msg) in messages.iter().enumerate() {
        let from = msg
            .from
            .as_ref()
            .map(|r| r.email_address.address.as_str())
            .unwrap_or("unknown");
        let subject = msg.subject.as_deref().unwrap_or("(no subject)");
        let read_mark = if msg.is_read.unwrap_or(false) {
            " "
        } else {
            "*"
        };

        println!("{:2}.{} {} - {}", i + 1, read_mark, from, subject);
        println!("      ID: {}", msg.id);
    }
    println!();
    println!("* = unread");
}
