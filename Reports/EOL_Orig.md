
# Overview

## Purpose
The purpose of this as-built-as-configured (ABAC) document is to detail each configuration item (CI) applied to the Department's Exchange Online instance. These CI's align to the design decisions captured within the Department's Office 365 Detailed Design.

## Associated Documentation
The following table lists the documents that were referenced during the creation of this ABAC.


| Name | Version | Date |
| ---- | ------- | ---- |
| GovDesk - Office 365 Design | 1.0 | 06/2019 |
| GovDesk - Platform Design | 1.0 | 06/2019 |
| GovDesk - Workstation Design | 1.0 | 06/2019 |

# Configuration

## Exchange Online
The ABAC settings for the Department's Exchange Online instance can be found below. This includes connectors, Mail Exchange (MX) records, SPF, DMARC, DKIM, Remote Domains, User mailbox configurations, Authentication Policies, Outlook on the Web policies, Mailbox Archiving, and Address lists.
Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting.

### Connectors
Exchange Online contains the following inbound mail connectors.


| Name | Status | TLS | Certificate |
| ---- | ------ | --- | ----------- |
| Not Configured | N/A | N/A | N/A |
There are no outbound mail connectors configured in Exchange Online.

### MX Records
The following MX records have been configured.


| Domain | MX Preference | Mail Exchanger |
| ------ | ------------- | -------------- |
| precisionservices.biz | MX preference = 0 | mail exchanger = precisionservices-biz.mail.protection.outlook.com


 |

### SPF Records
The following SPF records have been configured.


| Domain | SPF Record | DMARC Policy |
| ------ | ---------- | ------------ |
| precisionservices.biz | "v=spf1 include:spf.protection.outlook.com -all" | Not configured |

### Remote Domains
The following Remote Domains have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | Default |
| Remote Domain | * |
| Allowed Out Of Office Type | External |
| Automatic Reply | True |
| Automatic Forward | True |
| Delivery Reports | True |
| Non Delivery Reports | True |
| Meeting Forward Notifications | False |
| Content Type | MimeHtmlText |

### CAS Mailbox Plan
The following CAS Mailbox Plans have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | ExchangeOnline |
| ActiveSync | True |
| ActiveSync Mailbox Policy | Not Configured |
| IMAP | True |
| MAPI | True |
| Outlook Web Access | True |
| Outlook Web Access Policy | OwaMailboxPolicy-Default |
| POP | True |
| EWS | True |
| Name [TBD] | ExchangeOnlineDeskless |
| ActiveSync | True |
| ActiveSync Mailbox Policy | Not Configured |
| IMAP | False |
| MAPI | False |
| Outlook Web Access | True |
| Outlook Web Access Policy | OwaMailboxPolicy-Default |
| POP | True |
| EWS | False |
| Name [TBD] | ExchangeOnlineEnterprise |
| ActiveSync | True |
| ActiveSync Mailbox Policy | Not Configured |
| IMAP | True |
| MAPI | True |
| Outlook Web Access | True |
| Outlook Web Access Policy | OwaMailboxPolicy-Default |
| POP | True |
| EWS | True |
| Name [TBD] | ExchangeOnlineEssentials |
| ActiveSync | True |
| ActiveSync Mailbox Policy | Not Configured |
| IMAP | True |
| MAPI | True |
| Outlook Web Access | True |
| Outlook Web Access Policy | OwaMailboxPolicy-Default |
| POP | True |
| EWS | True |

### Authentication Policy
There are no Authentication Policies configured

### Outlook Web Access Policy
The following Outlook Web Access Policies have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | OwaMailboxPolicy-Default |
| Wac Editing Enabled | True |
| Print Without Download Enabled | True |
| OneDrive Attachments Enabled | True |
| Third Party File Providers Enabled | False |
| Classic Attachments Enabled | True |
| Reference Attachments Enabled | True |
| Save Attachments To Cloud Enabled | True |
| Message Previews Disabled | False |
| Direct File Access On Public Computers Enabled | True |
| Direct File Access On Private Computers Enabled | True |
| Web Ready Document Viewing On Public Computers Enabled | True |
| Web Ready Document Viewing On Private Computers Enabled | True |
| Force Web Ready Document Viewing First On Public Computers | False |
| Force Web Ready Document Viewing First On Private Computers | False |
| Wac Viewing On Public Computers Enabled | True |
| Wac Viewing On Private Computers Enabled | True |
| Force Wac Viewing First On Public Computers | False |
| Force Wac Viewing First On Private Computers | False |
| Action For Unknown File And MIME Types | True |
| Phonetic Support Enabled | False |
| Default Client Language | 0 |
| Use GB18030 | False |
| Use ISO885915 | False |
| Outbound Charset | AutoDetect |
| Global Address List Enabled | True |
| Organization Enabled | True |
| Explicit Logon Enabled | True |
| OWA Light Enabled | True |
| Delegate Access Enabled | True |
| IRM Enabled | True |
| Calendar Enabled | True |
| Contacts Enabled | True |
| Tasks Enabled | True |
| Journal Enabled | True |
| Notes Enabled | True |
| On Send Addins Enabled | False |
| Reminders And Notifications Enabled | True |
| Premium Client Enabled | True |
| Spell Checker Enabled | False |
| Classic Attachments Enabled | True |
| Search Folders Enabled | True |
| Signatures Enabled | True |
| Theme Selection Enabled | True |
| Junk Email Enabled | True |
| UM Integration Enabled | True |
| WSS Access On Public Computers Enabled | False |
| WSS Access On Private Computers Enabled | False |
| Change Password Enabled | False |
| UNC Access On Public Computers Enabled | False |
| UNC Access On Private Computers Enabled | False |
| ActiveSync Integration Enabled | True |
| All Address Lists Enabled | True |
| Rules Enabled | True |
| Public Folders Enabled | True |
| SMime Enabled | False |
| Recover Deleted Items Enabled | True |
| Instant Messaging Enabled | True |
| Text Messaging Enabled | True |
| Force Save Attachment Filtering Enabled | False |
| Silverlight Enabled | True |
| Instant Messaging Type | Ocs |
| Display Photos Enabled | True |
| Set Photo Enabled | True |
| Allow Offline On | AllComputers |
| Set Photo URL |  |
| Places Enabled | True |
| Weather Enabled | True |
| Local Events Enabled | False |
| Interesting Calendars Enabled | True |
| Allow Copy Contacts To Device Address Book | True |
| Predicted Actions Enabled | True |
| User Diagnostic Enabled | False |
| Facebook Enabled | True |
| LinkedIn Enabled | True |
| Wac External Services Enabled | True |
| Wac OMEX Enabled | False |
| Report Junk Email Enabled | True |
| Group Creation Enabled | True |
| Skip Create Unified Group Custom Sharepoint Classification | True |
| User Voice Enabled | True |
| Satisfaction Enabled | True |
| Outlook Beta Toggle Enabled | True |

### Address Lists
The following Address Lists have been configured.


| Name | Recipient Filter |
| ---- | ---------------- |
| Not Configured | N/A |
