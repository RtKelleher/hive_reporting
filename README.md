# hive_reporting

Hive-Reporting provides easy to read case metrics supporting team contirubtions
and frequency without the need to access or create custom report in
The Hive Dashboard

Environment Variables:
* Sent_to
* SMTP Server
* From
* TheHive API Information

Process Pipeline:
Hive API > Create 30/60/90 > Pandas > Create Charts (WIP) > Send via SMTP
