# Email-With-Ms-Excel
This code extracts chart and table from Ms excel file and send it an email to recipient given.

#  Working
Python comes with the built-in smtplib module for sending emails using the Simple Mail Transfer Protocol (SMTP). smtplib uses the RFC 821 protocol for SMTP.
When you send emails through Python, you should make sure that your SMTP connection is encrypted, so that your message and login credentials are not easily accessed by others. SSL (Secure Sockets Layer) and TLS (Transport Layer Security) are two protocols that can be used to encrypt an SMTP connection. It’s not necessary to use either of these when using a local debugging server.

There are two ways to start a secure connection with your email server:

1.Start an SMTP connection that is secured from the beginning using SMTP_SSL().

2.Start an unsecured SMTP connection that can then be encrypted using .starttls().

In both instances, Gmail will encrypt emails using TLS, as this is the more secure successor of SSL.
After you initiated a secure SMTP connection using either of the above methods, you can send your email using .sendmail():

server.sendmail(sender_email, receiver_email, message)

If you want to format the text in your email (bold, italics, and so on), or if you want to add any images, hyperlinks, or responsive content, then HTML comes in very handy. Today’s most common type of email is the MIME (Multipurpose Internet Mail Extensions) Multipart email, combining HTML and plain-text. MIME messages are handled by Python’s email.mime module

There are multiple libraries designed to make sending emails easier, such as Envelopes, Flanker and Yagmail. Yagmail is designed to work specifically with Gmail, and it greatly simplifies the process of sending emails through a friendly API.

You can put personalized content in a message by using str.format() to fill in curly-bracket placeholders.

For example, "hi {name}, you {result} your assignment".format(name="John", result="passed") will give you "hi John, you passed your assignment".
