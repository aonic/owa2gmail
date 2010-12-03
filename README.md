<h1>Outlook Web Access Export Tools</h1>
<p>
	<ul>
		<li>The POP daemon should be usable in the current version, I've tested it with Evolution Mail on Ubuntu 10.10.</li> 
		<li>Gmail importing works if you call the append function manually, I'm still working on fixing up gmail.py to do that seamlessly.</li>
	</ul>
</p>

<h3>Origin Documentation from <a href="http://code.google.com/p/weboutlook/">weboutlook</a></h3>
<p>
weboutlook is a Python module that retrieves full, raw e-mails from Microsoft Outlook Web Access by screen scraping. It can do the following:

<ul>
	<li>Log into a Microsoft Outlook Web Access account with a given username and password.</li>
	<li>Retrieve all e-mail IDs from the first page of your Inbox.</li>
	<li>Retrieve the full, raw source of the e-mail with a given ID.</li>
	<li>Delete an e-mail with a given ID (technically, move it to the 'Deleted Items' folder).</li>
</ul>

<b>Sample Usage:</b>

	<pre>
	>>> from scaper import OutlookWebScraper

	# Throws InvalidLogin exception for invalid username/password. 
	>>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'invalid password') 
	>>> s.login() Traceback (most recent call last):
	...
	scraper.InvalidLogin
	
	>>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'correct password') 
	>>> s.login()

	# Display IDs of messages in the inbox. 
	>>> s.inbox() '/Inbox/test-3.EML'

	# Display IDs of messages in the 'sent items' folder. 
	>>> s.get_folder('sent items') '/Sent%20Items/test-2.EML'

	# Display the raw source of a particular message. 
	>>> print s.get_message('/Inbox/Hey%20there.EML') ...

	# Delete a message. 
	>>> s.delete_message('/Inbox/Hey%20there.EML')
	</pre>
</p>
