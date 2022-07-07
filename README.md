# Welcome to Automatic Mail Printer
Welcome to the official GitHub repo of Automatic Mail Printer. Automatic Mail Printer is a small program (.NET) to print emails contained in the mail account on the local printer. The idea came from my 2nd best friend (Floris J.). Together we are working on this project.

## Info about Automatic Mail Printer
The project was recently created by a new development of a website for a pizzeria. The system was built with Shopify and replaced an old website that was built on Contao.
Orders used to be received via email and sent to the local fax machine at the pizzeria via a mail to fax provider. However, the service to convert the mails to fax was very expensive. With the new website the costs should be reduced and replaced by a better system. The Automatic Mail Printer was born.
Another point why the Automatic Mail Printer could be interesting is that orders in Shopify can only be printed manually. With Automatic Mail Printer, orders are printed directly after they are received.

## Setup the Automatic Mail Printer
To install the Automatic Mail Printer we need to edit the config.json. First we need to set up the appropriate mail account. For this we need to enter the imap data and the email and password. The printer name must be set. This usually requires a little effort. The name can be found best in the device manager under Windows.
Recommended is an interval of 60 seconds in which the e-mail tray is checked for new messages. We can also set up a filter to print only orders that contain certain words in the subject, e.g. "Order" or "Message from customer".

### Nice to know
The program reads only unread e-mails. After printing, they are marked as read.

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
## Automatic Mail Printer in Action

![Automatic Mail Printer](/assets/AutomaticMailPrinter-example.jpg)

## Supported Languages
The program supports the following languages: English and German

## Use Case - Print Shopify Orders Automatically
As mentioned, the program was developed to directly print orders received by Shopify. To set this up on Shopify we just need to set up an email recipient in the "Notifications" settings. To do this, go to Settings - Notifications - Order Notifications for Employees and add a recipient. Now enter an email that the Automatic Mail Printer can access. Now you have completed the setup.

To test the process you can send a test notification directly to your email account via Shopify. Your order should now be recognized and printed by Automatic Mail Printer.

## Questions or Bugfixes?
Get in touch with us!

## Special Thanks to
- Idea by [@florijohn](https://github.com/florijohn)
- Based on [MailKit](https://github.com/jstedfast/MailKit) and [PrintHtml](https://github.com/kendallb/PrintHtml)