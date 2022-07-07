# MailPrinter
Automatically prints e-mails (.NET)

Idea by [@florijohn](https://github.com/florijohn)

Based on [MailKit](https://github.com/jstedfast/MailKit) and [PrintHtml](https://github.com/kendallb/PrintHtml)

ToDo: Explain ``config.json``

```json
{
  "imap_server": "imap.gmx.net",
  "imap_port":  993,
  "mail": "your@e-mail.com",
  "password": "your_password",
  "printer_name": "your printer",
  "timer_interval_in_seconds": 60,
  "filter": ["order"]
}
```
