
# Overview

## Purpose
The purpose of this as-built-as-configured (ABAC) document is to detail each configuration item (CI) applied to the Department's Exchange Online instance. These CI's align to the design decisions captured within the Department's Office 365 Detailed Design.

## Associated Documentation
The following table lists the documents that were referenced during the creation of this ABAC.


| Name | Version | Date |
| ---- | ------- | ---- |
| GovDesk - Office 365 Design | 1.0 | 06/2019 |

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
| precisionservices.biz | MX preference = 0 | mail exchanger = precisionservices-biz.mail.protection.outlook.com<br><br><br> |

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

### CAS Mailbox Plan
The following CAS Mailbox Plans have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | ExchangeOnline |

### Authentication Policy
There are no Authentication Policies configured

### Outlook Web Access Policy
The following Outlook Web Access Policies have been configured.


| Configuration Item | Value |
| ------------------ | ----- |
| Name [TBD] | OwaMailboxPolicy-Default |

### Address Lists
The following Address Lists have been configured.


| Name | Recipient Filter |
| ---- | ---------------- |
| Not Configured | N/A |