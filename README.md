# Goreport
Outlook Add In for reporting Gophish Mails, for old and new Outlook Client.

This Add In enables users to report supected phishing mails to a goreport instance without having to manually open the report link or forward the mail to a designated IMAP mailbox.
<br>
![Example Picture](assets/goreport_example_gif.gif)

<h2>How to set up</h2>

All of the files (except for the xml) need to be hosted on a web server, accessible for all users of the add in.
This is how Outlook Web Add Ins are designed.
The easiest way to do this is to publish a Web App on Azure for example.

The next step is to edit the manifest.xml file.
Everywhere "localhost" needs to be replaced by a domain which hosts your html and js files.

If you have compiled the Gophish binary yourself and you changed the "rid" keyword to something else, for example like "keyname", you can define that in the first line in main.js "alternateRid".

The Outlook Web Add In can then be installed for the users, see Microsoft documentation for that: <a>https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=windows</a> (Under "sideload manually")

Outlook Web Add Ins have the benefit that they work on the old and new version of the client, with that they also work in OWA.

<h2>User experience</h2>

When the user opens a mail he can find the Goreport Add In next to his other Add Ins.
He then can press a button to report the mail.
![Example picture two](assets/goreport_example2.png)

<h2>How does it work</h2>

The Add In uses regex to parse the mail content and search for a Gophish link.
If it finds a Gophish link:
<br>
It transforms the link to a report link and launches a fetch request to that report link.
The report is logged and displayed in your Gophish instance, the user receive confirmation that the mail was reported and can be deleted:
<br>
![Example picture two](assets/goreport_example3.png)
<br>
If it doesn't find a Gophish link:
<br>
No fetch request is made, but the user still receives confirmation that the mail was reported.
The reason for that is, that way the user isn't directly able to tell which phishing mails are created by your Gophish instance and which are actual phishing mails.
(There is still a DEV text shown if no Gophsih link is found, this will be removed in an upcoming update)
<br>
![Example picture two](assets/goreport_example2.png)

<h2>Why doesn't the Add In automatically remove the reported mail?</h2>

Outlook Web Add Ins can't interact with mails in that way.
To be able to move or delete mails, you would need to use the Graph API. this is doable but it would require a more complex backend and would be difficult to realize if you need to serve this to multiple customers with different tenants, for example if you're an MSP.
