# GOA - Google Organize App

This is a Google app to help someone organize their own personal GMail InBox.

You install it by adding the source code file in src/goa.ds to your Google apps project directory.
Then you can run it manually or have Google trigger it peridically (once or several times a day).

The app reads your GMail Inbox and looks up each email in your Google contacts. If if the email
is in your contents it keeps the message in your InBox otherwise it labels the message thread
with the label "UnknownSenders" and removes the message from your InBox.

## Feedback

Whenever the GOA analyzes a batch of emails when it runs it sends you an email with a summary
of what it did. If you ignore this message and do not read it, the next time the GOA runs it
adds the statistics of the previous run to new run and removes the old status email and sends
you a new one.

This means the status emails do not accumulate and the email should be close to the top of your
InBox when you check your email.

## Moving messages to Other Folders
You can also configure GOA to send messages from certain people into seperate folders.
You configure this by added a label to several people in your contact list and then telling
GOA that you want to move all message from people with that label to a folder with that name.

You configure GOA by emailing yourself a line with the subject in this form:

    GOA add folder <mylabel>

This tells GOA to move all messages from people who have this label in your contact of that person.
If you change your mind you can send GOA another email:

    GOA remove folder <mylabel>

Then GOA will stop moving message for that label. You can keep the label in the contacts.

## Missing Contacts

A common mistake you might make is receiving emails from someone important who is not
in your contact list. If this happens GOA will move messages from person to your
UnknownSenders folder in Gmail.

GOA will never remove this email, it only moves it from your InBox to your UnknownSenders folder.
You you use Gmail to Click the UnknownSends folder to you will see all emails from UnknownSends
sort with the most recent emails on top (always)

Periodically scan the names of the senders to find someone who you think is important. Most of
the emails in this list will be junk so important people will probably stand out.
Open the message. Click the name of the From and Click "Add to Contacts". If you edit the
contacts make that the email from that person is Home or Work or something not blank.

Click the checkbox next to the meail. In the top of the window click the move button and
move the message to InBox.

Then you can go back to your InBox and handle it. The next GOA run will not make this mistake
again for that person. Eventually you will have everyone important in your Contacts list.
This is a good idea anyway.

    