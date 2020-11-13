# Overview


## Purpose
The purpose of this as-built-as-configured (ABAC) document is to detail each configuration item (CI) applied to the Department's Exchange Online instance. These CI's align to the design decisions captured within the Department's Office 365 Detailed Design.


## Associated Documentation
The following table lists the documents that were referenced during the creation of this ABAC.


| Name | Version | Date |
| ---- | ------- | ---- |
| GovDesk - Office 365 Design | 1.0 | 06/2019 || GovDesk - Platform Design | 1.0 | 06/2019 || GovDesk - Workstation Design | 1.0 | 06/2019 |


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
| precisionservices.biz | 0 | precisionservices-biz.mail.protection.outlook.com |

### SPF Records
The following SPF records have been configured.


| Domain | SPF Record | DMARC Policy |
| ------ | ---------- | ------------ |
| precisionservices.biz | "v=spf1 include:spf.protection.outlook.com -all" | Not configured |

### Remote Domains
The following Remote Domains have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | Default || Remote Domain | * || Allowed Out Of Office Type | External || Automatic Reply | True || Automatic Forward | True || Delivery Reports | True || Non Delivery Reports | True || Meeting Forward Notifications | False || Content Type | MimeHtmlText |

### CAS Mailbox Plan
The following CAS Mailbox Plans have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | ExchangeOnline || ActiveSync | True || ActiveSync Mailbox Policy | Not Configured || IMAP | True || MAPI | True || Outlook Web Access | True || Outlook Web Access Policy | OwaMailboxPolicy-Default || POP | True || EWS | True || Name [TBD] | ExchangeOnlineDeskless || ActiveSync | True || ActiveSync Mailbox Policy | Not Configured || IMAP | False || MAPI | False || Outlook Web Access | True || Outlook Web Access Policy | OwaMailboxPolicy-Default || POP | True || EWS | False || Name [TBD] | ExchangeOnlineEnterprise || ActiveSync | True || ActiveSync Mailbox Policy | Not Configured || IMAP | True || MAPI | True || Outlook Web Access | True || Outlook Web Access Policy | OwaMailboxPolicy-Default || POP | True || EWS | True || Name [TBD] | ExchangeOnlineEssentials || ActiveSync | True || ActiveSync Mailbox Policy | Not Configured || IMAP | True || MAPI | True || Outlook Web Access | True || Outlook Web Access Policy | OwaMailboxPolicy-Default || POP | True || EWS | True |

### Authentication Policy
There are no Authentication Policies configured

### Outlook Web Access Policy
The following Outlook Web Access Policies have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | OwaMailboxPolicy-Default || Wac Editing Enabled | True || Print Without Download Enabled | True || OneDrive Attachments Enabled | True || Third Party File Providers Enabled | False || Classic Attachments Enabled | True || Reference Attachments Enabled | True || Save Attachments To Cloud Enabled | True || Message Previews Disabled | False || Direct File Access On Public Computers Enabled | True || Direct File Access On Private Computers Enabled | True || Web Ready Document Viewing On Public Computers Enabled | True || Web Ready Document Viewing On Private Computers Enabled | True || Force Web Ready Document Viewing First On Public Computers | False || Force Web Ready Document Viewing First On Private Computers | False || Wac Viewing On Public Computers Enabled | True || Wac Viewing On Private Computers Enabled | True || Force Wac Viewing First On Public Computers | False || Force Wac Viewing First On Private Computers | False || Action For Unknown File And MIME Types | True || Phonetic Support Enabled | False || Default Client Language | 0 || Use GB18030 | False || Use ISO885915 | False || Outbound Charset | AutoDetect || Global Address List Enabled | True || Organization Enabled | True || Explicit Logon Enabled | True || OWA Light Enabled | True || Delegate Access Enabled | True || IRM Enabled | True || Calendar Enabled | True || Contacts Enabled | True || Tasks Enabled | True || Journal Enabled | True || Notes Enabled | True || On Send Addins Enabled | False || Reminders And Notifications Enabled | True || Premium Client Enabled | True || Spell Checker Enabled | False || Classic Attachments Enabled | True || Search Folders Enabled | True || Signatures Enabled | True || Theme Selection Enabled | True || Junk Email Enabled | True || UM Integration Enabled | True || WSS Access On Public Computers Enabled | False || WSS Access On Private Computers Enabled | False || Change Password Enabled | False || UNC Access On Public Computers Enabled | False || UNC Access On Private Computers Enabled | False || ActiveSync Integration Enabled | True || All Address Lists Enabled | True || Rules Enabled | True || Public Folders Enabled | True || SMime Enabled | False || Recover Deleted Items Enabled | True || Instant Messaging Enabled | True || Text Messaging Enabled | True || Force Save Attachment Filtering Enabled | False || Silverlight Enabled | True || Instant Messaging Type | Ocs || Display Photos Enabled | True || Set Photo Enabled | True || Allow Offline On | AllComputers || Set Photo URL |  || Places Enabled | True || Weather Enabled | True || Local Events Enabled | False || Interesting Calendars Enabled | True || Allow Copy Contacts To Device Address Book | True || Predicted Actions Enabled | True || User Diagnostic Enabled | False || Facebook Enabled | True || LinkedIn Enabled | True || Wac External Services Enabled | True || Wac OMEX Enabled | False || Report Junk Email Enabled | True || Group Creation Enabled | True || Skip Create Unified Group Custom Sharepoint Classification | True || User Voice Enabled | True || Satisfaction Enabled | True || Outlook Beta Toggle Enabled | True |

### Address Lists
The following Address Lists have been configured.


| Name | Recipient Filter |
| ---- | ---------------- |
| Not Configured | N/A |
## Exchange Online Protection
The ABAC settings for the Department's Exchange Online Protection instance can be found below. This includes the Connection Filtering, Anti-Malware, Policy Filtering, and Content Filtering Configuration.
Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting.
### Connection Filtering
The following Connection Filters have been configured.
| Name | IP Allow List | IP Block List | Enable Safe List | Directory Based Edge Block Mode |
| ---- | ------------- | ------------- | ---------------- | ------------------------------- |
| Default |  |  | False | Default |
### Anti-Malware
The following Malware Filters have been configured.
| Configuration Item | Value |
| ------------------ | ----- |
| Name | Default |
| Custom Notifications | False |
| Custom notification details | Not Configured |
| Internal Sender Admin Address | martin@precisionservices.biz |
| External Sender Admin Address | martin@precisionservices.biz |
| Action | DeleteAttachmentAndUseDefaultAlertText |
| Enable Internal Sender Notifications | True |
| Enable External Sender Notifications | False |
| Enable Internal Sender Admin Notifications | True |
| Enable External Sender Admin Notifications | True |
| Enable File Filter | True |
| Filter file types | ace
ani
app
docm
exe
jar
reg
scr
vbe
vbs |
### Policy Filtering
The following Policy Filters have been configured.
| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBA] | Block executable contentdoco |
| Priority | 0 |
| Description | If the message:
	includes an attachment with executable content
Take the following actions:
	reject the message and include the explanation 'The email contains an attachment that is not allowed' with the status code: '5.7.1' |
| State | Enabled |
| Mode | AuditAndNotify |
| Condition | Has no classification: False
Has sender override: False
Attachment is unsupported: False
Attachment processing limit exceeded: False
Attachment has executable content: True
Attachment is password protected: False |
| Exception | Except if has no classification: False
Except if attachment is unsupported: False
Except if attachment processing limit exceeded: False
Except if attachment has executable content: False
Except if attachment is password protected: False
Except if has sender override: False |
| Action | Moderate message by manager: False
Reject message enhanced status code: 5.7.1
Reject message reason text: The email contains an attachment that is not allowed
disconnect: False
Delete message: False
Quarantine: False
Stop rule processing: False
Route message outbound require TLS: False
Apply OME: False
Remove OME: False
Remove OMEv2: False |
| Name [TBA] | Block External Forwarding |
| Priority | 1 |
| Description | If the message:
	Is sent to 'Outside the organization'
	and Is message type 'Auto-forward'
	and Is received from 'Inside the organization'
Take the following actions:
	reject the message and include the explanation 'AutoForward to External Recipients is not allowed' with the status code: '5.7.1' |
| State | Enabled |
| Mode | Enforce |
| Condition | Has no classification: False
Has sender override: False
Attachment is unsupported: False
Attachment processing limit exceeded: False
Attachment has executable content: True
Attachment is password protected: FalseFrom scope: InOrganizationSent to scope: NotInOrganizationHas no classification: FalseHas sender override: FalseMessage type matches: AutoForwardAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: False |
| Exception | Except if has no classification: False
Except if attachment is unsupported: False
Except if attachment processing limit exceeded: False
Except if attachment has executable content: False
Except if attachment is password protected: False
Except if has sender override: FalseExcept if has no classification: FalseExcept if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: False |
| Action | Moderate message by manager: False
Reject message enhanced status code: 5.7.1
Reject message reason text: The email contains an attachment that is not allowed
disconnect: False
Delete message: False
Quarantine: False
Stop rule processing: False
Route message outbound require TLS: False
Apply OME: False
Remove OME: False
Remove OMEv2: FalseModerate message by manager: FalseReject message enhanced status code: 5.7.1Reject message reason text: AutoForward to External Recipients is not alloweddisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: False |
| Name [TBA] | BlockDRAFT |
| Priority | 2 |
| Description | If the message:
	'msip_labels' header contains ''MSIP_Label_0eff2e56-0b0a-4708-8baf-10e25fce5a89_Enabled=true''
Take the following actions:
	Prepend the subject with '[SEC=DRAFT]'
Except if the message:
	Includes these words in the message subject or body: '[SEC=DRAFT]' |
| State | Enabled |
| Mode | Enforce |
| Condition | Has no classification: False
Has sender override: False
Attachment is unsupported: False
Attachment processing limit exceeded: False
Attachment has executable content: True
Attachment is password protected: FalseFrom scope: InOrganizationSent to scope: NotInOrganizationHas no classification: FalseHas sender override: FalseMessage type matches: AutoForwardAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseHas no classification: FalseHeader contains message header: msip_labelsHeader contains words: MSIP_Label_0eff2e56-0b0a-4708-8baf-10e25fce5a89_Enabled=trueHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: False |
| Exception | Except if has no classification: False
Except if attachment is unsupported: False
Except if attachment processing limit exceeded: False
Except if attachment has executable content: False
Except if attachment is password protected: False
Except if has sender override: FalseExcept if has no classification: FalseExcept if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if subject or body contains words: [SEC=DRAFT]Except if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: False |
| Action | Moderate message by manager: False
Reject message enhanced status code: 5.7.1
Reject message reason text: The email contains an attachment that is not allowed
disconnect: False
Delete message: False
Quarantine: False
Stop rule processing: False
Route message outbound require TLS: False
Apply OME: False
Remove OME: False
Remove OMEv2: FalseModerate message by manager: FalseReject message enhanced status code: 5.7.1Reject message reason text: AutoForward to External Recipients is not alloweddisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalsePrepend Subject: [SEC=DRAFT]Moderate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: False |
| Name [TBA] | UNOFFICIAL |
| Priority | 3 |
| Description | If the message:
	'msip_labels' header contains ''MSIP_Label_610d8fb9-b026-4cfb-b829-991a2f8c7a1d_Enabled=true''
Take the following actions:
	Prepend the subject with '[SEC=UNOFFICIAL]'
	and set message header 'X-Protective-Marking' with the value 'SEC=UNOFFICIAL'
Except if the message:
	Includes these words in the message subject or body: '[SEC=UNOFFICIAL]' |
| State | Enabled |
| Mode | Enforce |
| Condition | Has no classification: False
Has sender override: False
Attachment is unsupported: False
Attachment processing limit exceeded: False
Attachment has executable content: True
Attachment is password protected: FalseFrom scope: InOrganizationSent to scope: NotInOrganizationHas no classification: FalseHas sender override: FalseMessage type matches: AutoForwardAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseHas no classification: FalseHeader contains message header: msip_labelsHeader contains words: MSIP_Label_0eff2e56-0b0a-4708-8baf-10e25fce5a89_Enabled=trueHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseHas no classification: FalseHeader contains message header: msip_labelsHeader contains words: MSIP_Label_610d8fb9-b026-4cfb-b829-991a2f8c7a1d_Enabled=trueHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: False |
| Exception | Except if has no classification: False
Except if attachment is unsupported: False
Except if attachment processing limit exceeded: False
Except if attachment has executable content: False
Except if attachment is password protected: False
Except if has sender override: FalseExcept if has no classification: FalseExcept if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if subject or body contains words: [SEC=DRAFT]Except if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if subject or body contains words: [SEC=UNOFFICIAL]Except if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: False |
| Action | Moderate message by manager: False
Reject message enhanced status code: 5.7.1
Reject message reason text: The email contains an attachment that is not allowed
disconnect: False
Delete message: False
Quarantine: False
Stop rule processing: False
Route message outbound require TLS: False
Apply OME: False
Remove OME: False
Remove OMEv2: FalseModerate message by manager: FalseReject message enhanced status code: 5.7.1Reject message reason text: AutoForward to External Recipients is not alloweddisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalsePrepend Subject: [SEC=DRAFT]Moderate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalsePrepend Subject: [SEC=UNOFFICIAL]Set header name: X-Protective-MarkingSet header value: SEC=UNOFFICIALModerate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: False |
| Name [TBA] | encrypt_external_emails |
| Priority | 4 |
| Description | If the message:
	Is sent to 'Outside the organization'
Take the following actions:
	rights protect message with RMS template:  'Encrypt' |
| State | Disabled |
| Mode | Enforce |
| Condition | Has no classification: False
Has sender override: False
Attachment is unsupported: False
Attachment processing limit exceeded: False
Attachment has executable content: True
Attachment is password protected: FalseFrom scope: InOrganizationSent to scope: NotInOrganizationHas no classification: FalseHas sender override: FalseMessage type matches: AutoForwardAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseHas no classification: FalseHeader contains message header: msip_labelsHeader contains words: MSIP_Label_0eff2e56-0b0a-4708-8baf-10e25fce5a89_Enabled=trueHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseHas no classification: FalseHeader contains message header: msip_labelsHeader contains words: MSIP_Label_610d8fb9-b026-4cfb-b829-991a2f8c7a1d_Enabled=trueHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: FalseSent to scope: NotInOrganizationHas no classification: FalseHas sender override: FalseAttachment is unsupported: FalseAttachment processing limit exceeded: FalseAttachment has executable content: FalseAttachment is password protected: False |
| Exception | Except if has no classification: False
Except if attachment is unsupported: False
Except if attachment processing limit exceeded: False
Except if attachment has executable content: False
Except if attachment is password protected: False
Except if has sender override: FalseExcept if has no classification: FalseExcept if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if subject or body contains words: [SEC=DRAFT]Except if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if subject or body contains words: [SEC=UNOFFICIAL]Except if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: FalseExcept if has no classification: FalseExcept if attachment is unsupported: FalseExcept if attachment processing limit exceeded: FalseExcept if attachment has executable content: FalseExcept if attachment is password protected: FalseExcept if has sender override: False |
| Action | Moderate message by manager: False
Reject message enhanced status code: 5.7.1
Reject message reason text: The email contains an attachment that is not allowed
disconnect: False
Delete message: False
Quarantine: False
Stop rule processing: False
Route message outbound require TLS: False
Apply OME: False
Remove OME: False
Remove OMEv2: FalseModerate message by manager: FalseReject message enhanced status code: 5.7.1Reject message reason text: AutoForward to External Recipients is not alloweddisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalsePrepend Subject: [SEC=DRAFT]Moderate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalsePrepend Subject: [SEC=UNOFFICIAL]Set header name: X-Protective-MarkingSet header value: SEC=UNOFFICIALModerate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: FalseApply rights protection template: EncryptModerate message by manager: Falsedisconnect: FalseDelete message: FalseQuarantine: FalseStop rule processing: FalseRoute message outbound require TLS: FalseApply OME: FalseRemove OME: FalseRemove OMEv2: False |
### Content Filtering
The following Content Filters have been configured.
| Configuration Item | Value |
| ------------------ | ----- |
| Name | Default |
| Add X Header Value | Not Configured |
| Modify Subject value | Not Configured |
| Redirect to recipients | Not Configured |
| False positive additional recipients | Not Configured |
| Quarantine retention period | 30 |
| End user spam notification frequency | 3 |
| Increase Score | Increase score with image links: Off
 Increase score with numeric IPs: Off
 Increase score with redirect to other port: Off
 Increase score with Biz or info URLs: Off |
| Mark as spam | Mark as spam empty messages: Off
 Mark as spam javascript in html: Off
 Mark as spam frames in HTML: Off
 Mark as spam object tags in HTML: Off
 Mark as spam embed tags in HTML: Off
 Mark as spam form tags in HTML: Off
 Mark as spam web bugs in HTML: Off
 Mark as spam sensitive word list: Off
 Mark as spam SPF record hard fail: Off
 Mark as spam from address auth fail: Off
 Mark as spam bulk mail: On
 Mark as spam NDR backscatter: Off |
| High confidence spam action | Quarantine |
| Spam action | MoveToJmf |
| Bulk spam action | MoveToJmf |
| Phish Spam action | Quarantine |
| Enable end user spam notifications | True |
| End user spam notification | Notification language: Default
 Notification limit: 0 |
| Download link | False |
| Enable region block list | False |
| Region block list | Not Configured |
| Enable language block list | Not Configured |
| Language block list | Not Configured |
| Bulk Threshold | 6 |
| Allowed Senders | Not Configured |
| Allowed sender domains | Not Configured |
| Blocked Senders | Not Configured |
| Blocked Sender Domains | Not Configured |
